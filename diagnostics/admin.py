"""超级管理员初始化、签发设备配置、解密同步与中文查看页。"""
from __future__ import annotations
import argparse, base64, datetime, getpass, html, json, os, sqlite3, subprocess, sys, uuid
from pathlib import Path
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ed25519, padding, rsa
from common import VERSION, NOTICE, cache, canonical, decrypt, digest, key_id, load_secret, organization, private_dir, private_key, read_json, save_secret, verify_sender

def admin_home(value=None):
    return private_dir(Path(value).resolve() if value else Path(os.environ.get('CODEX_HOME',str(Path.home()/'.codex')))/'payroll-diagnostics-admin')
def init(home,repository,known_hosts):
    home=admin_home(home);path=home/'admin-key.secure'
    if path.exists():key=private_key(load_secret(path))
    else:
        key=rsa.generate_private_key(public_exponent=65537,key_size=3072)
        save_secret(path,key.private_bytes(serialization.Encoding.PEM,serialization.PrivateFormat.PKCS8,serialization.NoEncryption()))
    public=key.public_key().public_bytes(serialization.Encoding.PEM,serialization.PublicFormat.SubjectPublicKeyInfo).decode()
    org={'schema':'payroll-diagnostics-organization-1','version':VERSION,'repository':repository,'private_repository':True,'public_key_pem':public,'key_id':key_id(public),'known_hosts':known_hosts,'notice':NOTICE,'notice_version':VERSION}
    (home/'organization.json').write_bytes(canonical(org))
    registry=home/'devices.json'
    if not registry.exists():registry.write_text('{}',encoding='utf-8')
    return org

def create_profile(home,label):
    home=admin_home(home);org=read_json(home/'organization.json');identifier=str(uuid.uuid4());ssh=ed25519.Ed25519PrivateKey.generate()
    private=ssh.private_bytes(serialization.Encoding.PEM,serialization.PrivateFormat.OpenSSH,serialization.NoEncryption()).decode()
    public=ssh.public_key().public_bytes(serialization.Encoding.OpenSSH,serialization.PublicFormat.OpenSSH).decode()
    profile={'schema':'payroll-diagnostics-profile-1','device_id':identifier,'repository':org['repository'],'key_id':org['key_id'],'ssh_private_pem':private,'ssh_public_key':public,'notice_version':VERSION}
    key=private_key(load_secret(home/'admin-key.secure'))
    signature=key.sign(canonical(profile),padding.PSS(mgf=padding.MGF1(hashes.SHA256()),salt_length=padding.PSS.MAX_LENGTH),hashes.SHA256())
    profile['admin_signature']=base64.b64encode(signature).decode()
    registry=read_json(home/'devices.json');registry[identifier]={'label':label,'public_key':public,'created_at':datetime.datetime.now(datetime.timezone.utc).isoformat(),'active':True}
    (home/'devices.json').write_bytes(canonical(registry))
    return profile

def verify_profile(profile,org):
    body={k:v for k,v in profile.items() if k!='admin_signature'}
    if body.get('schema')!='payroll-diagnostics-profile-1' or body.get('repository')!=org['repository'] or body.get('key_id')!=org['key_id'] or body.get('notice_version')!=org['notice_version']:raise ValueError('设备配置不属于当前发布组织或告知版本')
    str(uuid.UUID(body['device_id']))
    key=serialization.load_pem_public_key(org['public_key_pem'].encode())
    key.verify(base64.b64decode(profile['admin_signature'],validate=True),canonical(body),padding.PSS(mgf=padding.MGF1(hashes.SHA256()),salt_length=padding.PSS.MAX_LENGTH),hashes.SHA256())
    ssh=serialization.load_ssh_private_key(body['ssh_private_pem'].encode(),password=None)
    if ssh.public_key().public_bytes(serialization.Encoding.OpenSSH,serialization.PublicFormat.OpenSSH).decode()!=body['ssh_public_key']:raise ValueError('设备上传密钥对不一致')
    return body

def ingest(home,folder):
    home=admin_home(home);org=read_json(home/'organization.json');registry=read_json(home/'devices.json');pem=load_secret(home/'admin-key.secure')
    db=sqlite3.connect(home/'admin-index.sqlite3');db.execute('CREATE TABLE IF NOT EXISTS records(event_id TEXT PRIMARY KEY,envelope TEXT NOT NULL,sha256 TEXT NOT NULL)');added=0;rejected=[]
    for path in sorted(Path(folder).glob('events/*/*.json')):
        try:
            if path.stat().st_size>65536 or path.is_symlink() or not path.resolve().is_relative_to(Path(folder).resolve()) or any(p.is_symlink() for p in path.parents if p!=Path(folder) and Path(folder) in p.parents):raise ValueError('异常记录大小或文件类型')
            envelope=read_json(path);sender=envelope['sender_device'];device=registry.get(sender,{})
            if not device.get('active'):raise ValueError('设备未登记或已撤销')
            verify_sender(envelope,device['public_key']);event=decrypt(envelope,pem)
            if event.get('device_id')!=sender or path.stem!=event['event_id']:raise ValueError('事件身份不一致')
            payload=canonical(envelope).decode();old=db.execute('SELECT sha256 FROM records WHERE event_id=?',(event['event_id'],)).fetchone()
            if old and old[0]!=digest(payload.encode()):raise ValueError('同号事件被替换')
            if not old:db.execute('INSERT INTO records VALUES(?,?,?)',(event['event_id'],payload,digest(payload.encode())));added+=1
        except Exception as exc:rejected.append({'file_id':path.stem if len(path.stem)<40 else 'unknown','category':type(exc).__name__})
    db.commit();db.close();return {'added':added,'rejected':rejected}

def dashboard(home,output,sync_notice=None):
    home=admin_home(home);output=Path(output).resolve()
    if output.exists() or output.parent!=home.resolve():raise ValueError('管理员页面须另存到受限管理员目录')
    registry=read_json(home/'devices.json');pem=load_secret(home/'admin-key.secure');db=sqlite3.connect(home/'admin-index.sqlite3');rows=[]
    for identifier,payload in db.execute('SELECT event_id,envelope FROM records'):
        env=json.loads(payload);sender=env['sender_device'];verify_sender(env,registry[sender]['public_key']);event=decrypt(env,pem);event['device_label']=registry[sender]['label'];rows.append(event)
    db.close();rows.sort(key=lambda x:x.get('started_at',''),reverse=True)
    # 解密仅用于管理员本机页面；数据库和云端仍保持密文。
    data=json.dumps(rows,ensure_ascii=False).replace('<','\\u003c').replace('>','\\u003e').replace('&','\\u0026')
    template=Path(__file__).with_name('dashboard.html').read_text('utf-8')
    notice=sync_notice or '本机查看快照；此页面不会自动刷新云端记录。'
    output.write_text(template.replace('__DIAGNOSTIC_ROWS__',data).replace('__SYNC_NOTICE__',html.escape(notice)),encoding='utf-8')
    return len(rows)

def sync(home,output):
    home=admin_home(home);org=read_json(home/'organization.json')
    try:repo=cache(home,'https://github.com/'+org['repository']+'.git')
    except (RuntimeError,OSError,subprocess.TimeoutExpired):
        if not output or not (home/'admin-index.sqlite3').is_file():raise
        count=dashboard(home,output,'本次未能连接云端，以下显示上次已同步的本机缓存，不是最新云端状态。')
        return {'cloud_synced':False,'visible_records':count,'message':'云端暂不可用，已生成明确标注的本机缓存查看页'}
    result=ingest(home,repo)
    result['cloud_synced']=True
    stamp=datetime.datetime.now(datetime.timezone.utc).isoformat()
    (home/'sync-state.json').write_bytes(canonical({'cloud_synced_at':stamp,'rejected_records':len(result['rejected'])}))
    if output:result['visible_records']=dashboard(home,output,'本次已连接云端；存在未能核验的记录：'+str(len(result['rejected']))+'条。')
    return result

def main():
    ap=argparse.ArgumentParser(description='薪资诊断超级管理员工具');ap.add_argument('command',choices=['sync','view','export-recovery']);ap.add_argument('--home',type=Path);ap.add_argument('--output',type=Path,required=True);a=ap.parse_args();home=admin_home(a.home)
    if a.command=='sync':print(json.dumps(sync(home,a.output),ensure_ascii=False));return
    if a.command=='view':print(json.dumps({'记录数':dashboard(home,a.output)},ensure_ascii=False));return
    if a.output.exists():raise SystemExit('恢复密钥导出不得覆盖已有文件')
    password=getpass.getpass('请设置管理员恢复密钥密码（至少12位，不会显示）：')
    if len(password)<12:raise SystemExit('密码长度不足')
    if password!=getpass.getpass('请再次输入密码：'):raise SystemExit('两次密码不一致')
    key=private_key(load_secret(home/'admin-key.secure'));payload=key.private_bytes(serialization.Encoding.PEM,serialization.PrivateFormat.PKCS8,serialization.BestAvailableEncryption(password.encode()))
    fd=os.open(a.output,os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
    with os.fdopen(fd,'wb') as stream:stream.write(payload)
    print('已导出加密恢复密钥，请离线保管，勿放入公开仓库或人资安装包。')

if __name__=='__main__':
    try:main()
    except Exception as exc:
        print('诊断查看未完成，请管理员核对网络、GitHub登录、设备登记或解密密钥。原记录保持不变。',file=sys.stderr)
        raise SystemExit(2)
