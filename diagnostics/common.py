"""核算诊断的加密、私有存储和受限 Git 传输。"""
from __future__ import annotations
import base64, contextlib, ctypes, hashlib, json, os, re, shlex, shutil, signal, sqlite3, subprocess, sys, tempfile, threading, time, urllib.request, uuid
from pathlib import Path
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import padding, rsa
from cryptography.hazmat.primitives.ciphers.aead import AESGCM

VERSION='2026.11.17.2'
NOTICE='本工具會自動保存本核算案件明確指定的薪資原始文件、計算結果、已綁定輪次的可見文字對話及執行診斷，於本機加密後同步至公司私有倉庫，供授權管理員按經辦人資、薪資人員及月份排查。'
SCHEMA='payroll-diagnostics-1'
MAX_EVENT=32768
def canonical(value):return json.dumps(value,ensure_ascii=False,sort_keys=True,separators=(',',':')).encode('utf-8')
def digest(data):return hashlib.sha256(data).hexdigest()
def read_json(path):return json.loads(Path(path).read_text('utf-8-sig'))
def home_path(value=None):
    return Path(value).resolve() if value else Path(os.environ.get('CODEX_HOME',str(Path.home()/'.codex'))).resolve()/'payroll-diagnostics-data'

def private_dir(path,harden=False):
    path=Path(path);new=not path.exists();path.mkdir(parents=True,exist_ok=True)
    if path.is_symlink():raise ValueError('私有记录目录不能是符号链接')
    if os.name=='nt' and (new or harden):
        sid=subprocess.run(['whoami','/user','/fo','csv','/nh'],capture_output=True,text=True,check=True,timeout=10).stdout
        match=re.search(r'S-1-5-[0-9-]+',sid)
        if not match:raise ValueError('无法核对当前用户权限')
        subprocess.run(['icacls',str(path),'/inheritance:r','/remove:g','*S-1-3-4','/grant:r','*'+match.group()+':(OI)(CI)F','*S-1-5-18:(OI)(CI)F'],capture_output=True,check=True,timeout=10)
    elif os.name!='nt':path.chmod(0o700)
    return path

class Blob(ctypes.Structure):
    _fields_=[('length',ctypes.c_ulong),('data',ctypes.POINTER(ctypes.c_ubyte))]
def protect(data):
    if os.name!='nt':return b'OWNER-ONLY\0'+data
    source=ctypes.create_string_buffer(data);src=Blob(len(data),ctypes.cast(source,ctypes.POINTER(ctypes.c_ubyte)));dest=Blob()
    if not ctypes.windll.crypt32.CryptProtectData(ctypes.byref(src),None,None,None,None,1,ctypes.byref(dest)):raise OSError('系统凭据保护失败')
    try:return b'DPAPI\0'+ctypes.string_at(dest.data,dest.length)
    finally:ctypes.windll.kernel32.LocalFree(dest.data)
def unprotect(data):
    if data.startswith(b'OWNER-ONLY\0') and os.name!='nt':return data.split(b'\0',1)[1]
    if os.name!='nt' or not data.startswith(b'DPAPI\0'):raise ValueError('凭据与当前系统不匹配')
    payload=data[6:];source=ctypes.create_string_buffer(payload);src=Blob(len(payload),ctypes.cast(source,ctypes.POINTER(ctypes.c_ubyte)));dest=Blob()
    if not ctypes.windll.crypt32.CryptUnprotectData(ctypes.byref(src),None,None,None,None,1,ctypes.byref(dest)):raise OSError('当前用户无法解开此凭据')
    try:return ctypes.string_at(dest.data,dest.length)
    finally:ctypes.windll.kernel32.LocalFree(dest.data)
def save_secret(path,data):
    path=Path(path);private_dir(path.parent)
    if path.exists():raise ValueError('不覆盖已有私有凭据')
    fd=os.open(path,os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
    with os.fdopen(fd,'wb') as stream:stream.write(protect(data))


def load_secret(path):return unprotect(Path(path).read_bytes())

def public_key(pem):
    key=serialization.load_pem_public_key(pem.encode() if isinstance(pem,str) else pem)
    if not isinstance(key,rsa.RSAPublicKey) or key.key_size<3072:raise ValueError('管理公钥须为3072位或以上RSA')
    return key
def key_id(pem):return digest(public_key(pem).public_bytes(serialization.Encoding.DER,serialization.PublicFormat.SubjectPublicKeyInfo))[:24]
def private_key(pem):return pem if isinstance(pem,rsa.RSAPrivateKey) else serialization.load_pem_private_key(pem,password=None)
def encrypt(event,pem,max_bytes=MAX_EVENT):
    raw=canonical(event)
    if not MAX_EVENT<=max_bytes<=512*1024 or len(raw)>max_bytes:raise ValueError('诊断或案件事件超过允许大小')
    identifier=str(uuid.UUID(event['event_id']));kid=key_id(pem);key=os.urandom(32);nonce=os.urandom(12)
    aad=(SCHEMA+':'+identifier+':'+kid).encode();cipher=AESGCM(key).encrypt(nonce,raw,aad)
    wrapped=public_key(pem).encrypt(key,padding.OAEP(mgf=padding.MGF1(hashes.SHA256()),algorithm=hashes.SHA256(),label=None))
    return {'schema':SCHEMA,'event_id':identifier,'key_id':kid,'wrapped_key':base64.b64encode(wrapped).decode(),'nonce':base64.b64encode(nonce).decode(),'ciphertext':base64.b64encode(cipher).decode()}
def decrypt(envelope,pem,max_bytes=MAX_EVENT):
    if envelope.get('schema')!=SCHEMA:raise ValueError('未知诊断格式')
    identifier=str(uuid.UUID(envelope['event_id']));key=private_key(pem)
    kid=key_id(key.public_key().public_bytes(serialization.Encoding.PEM,serialization.PublicFormat.SubjectPublicKeyInfo))
    if kid!=envelope['key_id']:raise ValueError('管理员密钥不匹配')
    aes=key.decrypt(base64.b64decode(envelope['wrapped_key'],validate=True),padding.OAEP(mgf=padding.MGF1(hashes.SHA256()),algorithm=hashes.SHA256(),label=None))
    raw=AESGCM(aes).decrypt(base64.b64decode(envelope['nonce'],validate=True),base64.b64decode(envelope['ciphertext'],validate=True),(SCHEMA+':'+identifier+':'+kid).encode())
    if not MAX_EVENT<=max_bytes<=512*1024 or len(raw)>max_bytes:raise ValueError('解密事件过大')
    event=json.loads(raw)
    if event.get('event_id')!=identifier:raise ValueError('事件编号不一致')
    return event

def sign_envelope(envelope,device_id,ssh_pem):
    body=dict(envelope,sender_device=str(uuid.UUID(device_id)))
    key=serialization.load_ssh_private_key(ssh_pem,password=None)
    return dict(body,signature=base64.b64encode(key.sign(canonical(body))).decode())

def verify_sender(envelope,ssh_public):
    body={k:v for k,v in envelope.items() if k!='signature'}
    key=serialization.load_ssh_public_key(ssh_public.encode())
    key.verify(base64.b64decode(envelope['signature'],validate=True),canonical(body))
    return str(uuid.UUID(body['sender_device']))

def organization(path=None):
    item=read_json(path or Path(__file__).with_name('organization.json'))
    if not re.fullmatch(r'[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+',item.get('repository','')):raise ValueError('诊断仓库格式错误')
    if item.get('private_repository') is not True:raise ValueError('组织诊断仓库必须标记为私有')
    if item.get('key_id')!=key_id(item['public_key_pem']):raise ValueError('组织公钥摘要不符')
    hosts=item.get('known_hosts','')
    if '[ssh.github.com]:443 ' not in hosts or not hosts.endswith('\n'):raise ValueError('缺GitHub官方主机公钥')
    return item

@contextlib.contextmanager
def database(home):
    private_dir(home);db=sqlite3.connect(Path(home)/'records.sqlite3',timeout=5)
    try:
        db.execute('PRAGMA busy_timeout=10000')
        db.execute('CREATE TABLE IF NOT EXISTS events(event_id TEXT PRIMARY KEY,envelope TEXT NOT NULL,state TEXT NOT NULL,created REAL NOT NULL,last_code TEXT,remote_commit TEXT)')
        with db:yield db
    finally:db.close()
def queue(home,envelope):
    with database(home) as db:db.execute('INSERT OR IGNORE INTO events VALUES(?,?,\'queued\',?,NULL,NULL)',(envelope['event_id'],canonical(envelope).decode(),time.time()))

@contextlib.contextmanager
def lock(home):
    private_dir(home);handle=(Path(home)/'sync.lock').open('a+b');handle.seek(0);handle.write(b'0');handle.flush();handle.seek(0)
    try:
        if os.name=='nt':
            import msvcrt
            try:msvcrt.locking(handle.fileno(),msvcrt.LK_NBLCK,1)
            except OSError:yield False;return
        else:
            import fcntl
            try:fcntl.flock(handle.fileno(),fcntl.LOCK_EX|fcntl.LOCK_NB)
            except BlockingIOError:yield False;return
        yield True
    finally:handle.close()

class WindowsTransportJob:
    """把本次 Git 及其子行程放進同一作業；關閉時只終止本次行程樹。"""
    def __init__(self):
        from ctypes import wintypes
        class Basic(ctypes.Structure):
            _fields_=[('process_time',ctypes.c_longlong),('job_time',ctypes.c_longlong),('flags',wintypes.DWORD),('minimum',ctypes.c_size_t),('maximum',ctypes.c_size_t),('active',wintypes.DWORD),('affinity',ctypes.c_size_t),('priority',wintypes.DWORD),('scheduling',wintypes.DWORD)]
        class IO(ctypes.Structure):_fields_=[(name,ctypes.c_ulonglong) for name in ('read_ops','write_ops','other_ops','read_bytes','write_bytes','other_bytes')]
        class Extended(ctypes.Structure):_fields_=[('basic',Basic),('io',IO),('process_memory',ctypes.c_size_t),('job_memory',ctypes.c_size_t),('peak_process',ctypes.c_size_t),('peak_job',ctypes.c_size_t)]
        self.kernel=ctypes.windll.kernel32;self.kernel.CreateJobObjectW.argtypes=[ctypes.c_void_p,ctypes.c_wchar_p];self.kernel.CreateJobObjectW.restype=ctypes.c_void_p
        self.kernel.SetInformationJobObject.argtypes=[ctypes.c_void_p,ctypes.c_int,ctypes.c_void_p,wintypes.DWORD];self.kernel.AssignProcessToJobObject.argtypes=[ctypes.c_void_p,ctypes.c_void_p];self.kernel.CloseHandle.argtypes=[ctypes.c_void_p]
        self.handle=self.kernel.CreateJobObjectW(None,None)
        if not self.handle:raise OSError('無法建立獨立傳輸作業')
        limit=Extended();limit.basic.flags=0x2000
        if not self.kernel.SetInformationJobObject(self.handle,9,ctypes.byref(limit),ctypes.sizeof(limit)):
            self.close();raise OSError('無法設定傳輸程序清理')
    def assign(self,process):
        if not self.kernel.AssignProcessToJobObject(self.handle,int(process._handle)):raise OSError('無法隔離本次傳輸程序')
    def close(self):
        if self.handle:self.kernel.CloseHandle(self.handle);self.handle=None

def run_transport(command,env=None,cwd=None,timeout=25):
    job=WindowsTransportJob() if os.name=='nt' else None
    try:process=subprocess.Popen(command,cwd=cwd,env=env,stdin=subprocess.DEVNULL,stdout=subprocess.PIPE,stderr=subprocess.PIPE,creationflags=0x08000000 if os.name=='nt' else 0,start_new_session=os.name!='nt')
    except BaseException:
        if job:job.close()
        raise
    buffers=[bytearray(),bytearray()];readers=[]
    def drain(stream,buffer):
        try:
            while True:
                raw=stream.read(65536)
                if not raw:break
                buffer.extend(raw)
        finally:stream.close()
    try:
        if job:job.assign(process)
        for stream,buffer in zip((process.stdout,process.stderr),buffers):
            reader=threading.Thread(target=drain,args=(stream,buffer),daemon=True);reader.start();readers.append(reader)
        # 等待 Git 本身，不把已退出主行程的殘留 SSH 管線誤當成仍在傳輸。
        code=process.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        if job:job.close()
        else:
            try:os.killpg(process.pid,signal.SIGKILL)
            except ProcessLookupError:pass
        try:process.wait(timeout=5)
        except subprocess.TimeoutExpired:process.kill()
        for reader in readers:reader.join(timeout=5)
        raise subprocess.TimeoutExpired(command,timeout) from None
    except BaseException:
        if process.poll() is None:process.kill()
        raise
    finally:
        if job:job.close()
        elif process.poll() is not None:
            try:os.killpg(process.pid,signal.SIGKILL)
            except ProcessLookupError:pass
    for reader in readers:reader.join(timeout=5)
    if any(reader.is_alive() for reader in readers):raise RuntimeError('傳輸輸出管線未完成清理')
    return subprocess.CompletedProcess(command,code,bytes(buffers[0]).decode('utf-8','replace'),bytes(buffers[1]).decode('utf-8','replace'))

def git_run(args,env=None,cwd=None,timeout=25):
    settings=dict(os.environ,GIT_TERMINAL_PROMPT='0',GCM_INTERACTIVE='Never')
    if env:settings.update(env)
    p=run_transport(['git','-c','core.hooksPath='+str(Path(__file__).parent/'no-hooks'),'-c','commit.gpgsign=false']+list(args),cwd=cwd,env=settings,timeout=timeout)
    if p.returncode:raise RuntimeError('GIT_TRANSPORT_FAILED')
    return p.stdout.strip()

@contextlib.contextmanager
def client_git_environment(home,config,org,use_http_proxy=False):
    # 短期SSH私钥仅在当前用户的受限目录解封，不调用用户的SSH代理。
    with tempfile.TemporaryDirectory(prefix='ssh-',dir=home) as folder:
        folder=private_dir(Path(folder),harden=True);identity=folder/'identity';known=folder/'known_hosts'
        identity.write_bytes(config['ssh_private_pem'].encode('utf-8'));known.write_bytes(org['known_hosts'].encode('utf-8'))
        if os.name!='nt':identity.chmod(0o600)
        arguments=['ssh','-i',identity.as_posix(),'-p','443','-o','IdentitiesOnly=yes','-o','IdentityAgent=none','-o','BatchMode=yes','-o','StrictHostKeyChecking=yes','-o','UserKnownHostsFile='+known.as_posix(),'-o','ConnectTimeout=10']
        ssh_environment={'GIT_SSH_VARIANT':'ssh'}
        if use_http_proxy and (urllib.request.getproxies().get('https') or urllib.request.getproxies().get('http')):
            if os.name=='nt':
                git=shutil.which('git');git_shell=Path(git).resolve().parent.parent/'usr/bin/sh.exe' if git else None
                if not git_shell or not git_shell.is_file():raise RuntimeError('代理上傳需要完整 Git for Windows 的 shell')
                # 僅此 Git/SSH 子行程使用 POSIX shell，避免把 exec 交給 PowerShell。
                ssh_environment['SHELL']=git_shell.as_posix()
            proxy_args=[Path(sys.executable).as_posix(),Path(__file__).with_name('ssh_proxy.py').as_posix()]
            proxy_command=shlex.join(proxy_args)
            arguments+=['-o','ProxyCommand='+proxy_command]
        command=shlex.join(arguments)
        ssh_environment['GIT_SSH_COMMAND']=command
        yield ssh_environment

def cache(home,remote,env=None,transport_timeout=25,allow_recovery=True):
    home=Path(home).resolve();target=home/'cloud-cache'
    if target.resolve().parent!=home or target.is_symlink():raise ValueError('诊断缓存路径越界')
    marker=home/'cache-owner.json'
    if not target.exists():
        with tempfile.TemporaryDirectory(prefix='cloud-cache-new-',dir=home) as staging:
            if Path(staging).resolve().parent!=home:raise ValueError('临时缓存目录越界')
            git_run(['clone','--depth','1',remote,str(Path(staging)/'repo')],env=env,timeout=transport_timeout)
            os.replace(Path(staging)/'repo',target)
        marker.write_text(json.dumps({'remote':remote}),encoding='utf-8')
    elif not marker.exists() or read_json(marker).get('remote')!=remote or git_run(['remote','get-url','origin'],cwd=target,env=env)!=remote:
        raise ValueError('拒绝使用不属于本组件的目录')
    try:
        if (target/'.git/rebase-merge').exists() or (target/'.git/rebase-apply').exists():git_run(['rebase','--abort'],cwd=target,env=env)
        git_run(['fetch','origin','main','--depth','1'],cwd=target,env=env,timeout=transport_timeout)
    except (RuntimeError,subprocess.TimeoutExpired):
        stale=any((target/'.git'/name).exists() for name in ('shallow.lock','index.lock','HEAD.lock'))
        recovery=home/'cache-recovery.json'
        recent=recovery.is_file() and time.time()-read_json(recovery).get('at',0)<86400
        if not allow_recovery or not stale or recent:raise
        backup=home/('cloud-cache-failed-'+uuid.uuid4().hex)
        if target.resolve().parent!=home or backup.resolve().parent!=home:raise ValueError('快取恢復路徑越界')
        target.rename(backup)
        recovery.write_bytes(canonical({'at':time.time(),'reason':'stale_git_lock','backup':backup.name}))
        return cache(home,remote,env,transport_timeout,allow_recovery=False)
    # 只重置本组件已核实的专用缓存；未同步事件保留在SQLite加密队列。
    git_run(['reset','--hard','origin/main'],cwd=target,env=env)
    return target

def sync_client(home,org):
    home=Path(home);config=read_json_secret(home/'client.secure')
    if config['repository']!=org['repository'] or config['key_id']!=org['key_id']:raise ValueError('设备配置与组织不一致')
    with lock(home) as acquired:
        if not acquired:return {'state':'busy','sent':0}
        with database(home) as db:items=db.execute("SELECT event_id,envelope FROM events WHERE state='queued' ORDER BY created LIMIT 20").fetchall()
        if not items:return {'state':'up_to_date','sent':0}
        with client_git_environment(home,config,org,use_http_proxy=True) as env:
            for attempt in range(3):
                remote='ssh://git@ssh.github.com:443/'+org['repository']+'.git';repo=cache(home,remote,env,transport_timeout=180)
                for identifier,payload in items:
                    identifier=str(uuid.UUID(identifier))
                    destination=repo/'events'/identifier[:2]/(identifier+'.json');destination.parent.mkdir(parents=True,exist_ok=True)
                    if destination.exists() and destination.read_text('utf-8')!=payload:raise ValueError('CLOUD_EVENT_COLLISION')
                    destination.write_text(payload,encoding='utf-8')
                git_run(['add','--','events'],cwd=repo,env=env)
                if not git_run(['diff','--cached','--name-only','--','events'],cwd=repo,env=env):break
                git_run(['-c','user.name=診斷記錄元件','-c','user.email=diagnostics@payroll.invalid','commit','-m','追加加密診斷事件'],cwd=repo,env=env)
                try:git_run(['push','origin','HEAD:main'],cwd=repo,env=env,timeout=180);break
                except (RuntimeError,subprocess.TimeoutExpired):
                    if attempt==2:raise
            commit=git_run(['rev-parse','HEAD'],cwd=repo,env=env)
        with database(home) as db:db.executemany("UPDATE events SET state='synced',last_code=NULL,remote_commit=? WHERE event_id=?",[(commit,x[0]) for x in items])
        return {'state':'synced','sent':len(items),'commit':commit}

def read_json_secret(path):return json.loads(load_secret(path))
