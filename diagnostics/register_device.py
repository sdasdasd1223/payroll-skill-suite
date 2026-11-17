"""管理员批准设备登记或撤销上传密钥；不把GitHub账户凭据交给客户端。"""
from pathlib import Path
import argparse, datetime, json, os, subprocess, sys, urllib.request, uuid
from admin import admin_home
from common import canonical, read_json
from cryptography.hazmat.primitives import serialization

def api(path,method='GET',data=None):
    env=dict(os.environ,GIT_TERMINAL_PROMPT='0',GCM_INTERACTIVE='Never')
    result=subprocess.run(['git','credential','fill'],input='protocol=https\nhost=github.com\n\n',text=True,capture_output=True,timeout=30,env=env)
    token=dict(x.split('=',1) for x in result.stdout.splitlines() if '=' in x).get('password')
    if result.returncode or not token:raise ValueError('管理员本机尚未连接GitHub账户')
    payload=canonical(data) if data is not None else None
    request=urllib.request.Request('https://api.github.com'+path,data=payload,method=method,headers={'Authorization':'Bearer '+token,'Accept':'application/vnd.github+json','X-GitHub-Api-Version':'2026-03-10','Content-Type':'application/json','User-Agent':'payroll-diagnostics-admin'})
    with urllib.request.urlopen(request,timeout=30) as response:
        content=response.read();return json.loads(content) if content else None

def main():
    ap=argparse.ArgumentParser(description='管理员批准或撤销诊断设备');ap.add_argument('command',choices=['approve','revoke']);ap.add_argument('--home',type=Path);ap.add_argument('--request',type=Path);ap.add_argument('--label');ap.add_argument('--device-id');a=ap.parse_args()
    home=admin_home(a.home);org=read_json(home/'organization.json');registry=read_json(home/'devices.json');repo=org['repository'];owner=repo.split('/')[0]
    if api('/user').get('login')!=owner or api('/repos/'+repo).get('private') is not True:raise ValueError('当前账号或私有仓库不符')
    if a.command=='approve':
        if not a.request or not a.label:raise ValueError('需要明确登记申请文件及设备名称')
        request=read_json(a.request);device=str(uuid.UUID(request['device_id']))
        if request['repository']!=repo or request['key_id']!=org['key_id']:raise ValueError('申请不属于当前公司')
        serialization.load_ssh_public_key(request['ssh_public_key'].encode())
        if device in registry and registry[device]['public_key']!=request['ssh_public_key']:raise ValueError('设备编号冲突')
        keys=api('/repos/'+repo+'/keys');match=next((x for x in keys if x['key']==request['ssh_public_key']),None)
        if not match:match=api('/repos/'+repo+'/keys','POST',{'title':'核算诊断 '+a.label+' '+device[:8],'key':request['ssh_public_key'],'read_only':False})
        registry[device]={'label':a.label,'public_key':request['ssh_public_key'],'created_at':datetime.datetime.now(datetime.timezone.utc).isoformat(),'active':True,'deploy_key_id':match['id']}
        print('设备已登记，网络可用时后台将自动重试同步。')
    else:
        device=str(uuid.UUID(a.device_id));known=registry[device]
        for key in api('/repos/'+repo+'/keys'):
            if key['key']==known['public_key']:api('/repos/'+repo+'/keys/'+str(key['id']),'DELETE')
        known['active']=False;print('已撤销该设备的云端上传密钥，旧密文记录保留。')
    (home/'devices.json').write_bytes(canonical(registry))

if __name__=='__main__':
    try:main()
    except Exception as exc:
        print('设备登记操作未完成：'+type(exc).__name__+'；请管理员核对账户、私有仓库及申请文件。',file=sys.stderr);raise SystemExit(2)
