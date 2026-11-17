"""核算诊断的加密、私有存储和受限 Git 传输。"""
from __future__ import annotations
import base64, contextlib, ctypes, hashlib, json, os, re, shlex, sqlite3, subprocess, sys, tempfile, time, uuid
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

def git_run(args,env=None,cwd=None,timeout=25):
    settings=dict(os.environ,GIT_TERMINAL_PROMPT='0',GCM_INTERACTIVE='Never')
    if env:settings.update(env)
    p=subprocess.run(['git','-c','core.hooksPath='+str(Path(__file__).parent/'no-hooks'),'-c','commit.gpgsign=false']+list(args),cwd=cwd,env=settings,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=timeout,creationflags=0x08000000 if os.name=='nt' else 0)
    if p.returncode:raise RuntimeError('GIT_TRANSPORT_FAILED')
    return p.stdout.strip()

@contextlib.contextmanager
def client_git_environment(home,config,org):
    # 短期SSH私钥仅在当前用户的受限目录解封，不调用用户的SSH代理。
    with tempfile.TemporaryDirectory(prefix='ssh-',dir=home) as folder:
        folder=private_dir(Path(folder),harden=True);identity=folder/'identity';known=folder/'known_hosts'
        identity.write_bytes(config['ssh_private_pem'].encode('utf-8'));known.write_bytes(org['known_hosts'].encode('utf-8'))
        if os.name!='nt':identity.chmod(0o600)
        command=shlex.join(['ssh','-i',identity.as_posix(),'-p','443','-o','IdentitiesOnly=yes','-o','IdentityAgent=none','-o','BatchMode=yes','-o','StrictHostKeyChecking=yes','-o','UserKnownHostsFile='+known.as_posix(),'-o','ConnectTimeout=10'])
        yield {'GIT_SSH_COMMAND':command,'GIT_SSH_VARIANT':'ssh'}

def cache(home,remote,env=None,transport_timeout=25):
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
    git_run(['fetch','origin','main','--depth','1'],cwd=target,env=env,timeout=transport_timeout)
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
        with client_git_environment(home,config,org) as env:
            remote='ssh://git@ssh.github.com:443/'+org['repository']+'.git';repo=cache(home,remote,env)
            for identifier,payload in items:
                identifier=str(uuid.UUID(identifier))
                destination=repo/'events'/identifier[:2]/(identifier+'.json');destination.parent.mkdir(parents=True,exist_ok=True)
                if destination.exists() and destination.read_text('utf-8')!=payload:raise ValueError('CLOUD_EVENT_COLLISION')
                destination.write_text(payload,encoding='utf-8')
            git_run(['add','--','events'],cwd=repo,env=env)
            if git_run(['status','--porcelain'],cwd=repo,env=env):
                git_run(['-c','user.name=诊断记录组件','-c','user.email=diagnostics@payroll.invalid','commit','-m','追加加密诊断事件'],cwd=repo,env=env)
                for attempt in range(3):
                    try:git_run(['push','origin','HEAD:main'],cwd=repo,env=env);break
                    except RuntimeError:
                        if attempt==2:raise
                        git_run(['-c','user.name=诊断记录组件','-c','user.email=diagnostics@payroll.invalid','pull','--rebase','origin','main'],cwd=repo,env=env)
            commit=git_run(['rev-parse','HEAD'],cwd=repo,env=env)
        with database(home) as db:db.executemany("UPDATE events SET state='synced',last_code=NULL,remote_commit=? WHERE event_id=?",[(commit,x[0]) for x in items])
        return {'state':'synced','sent':len(items),'commit':commit}

def read_json_secret(path):return json.loads(load_secret(path))
