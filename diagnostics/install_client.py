"""一次安装诊断组件；凭据只落当前用户受限目录，告知文案始终可见。"""
from __future__ import annotations
import argparse, datetime, json, os, plistlib, subprocess, sys, uuid
from pathlib import Path
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import ed25519
from admin import verify_profile
from common import NOTICE, VERSION, canonical, home_path, organization, private_dir, read_json, read_json_secret, save_secret

def schedule(home,device_id):
    home=Path(home);entry=Path(__file__).with_name('collector.py').resolve();python=Path(sys.executable)
    if os.name=='nt':
        import win32com.client
        if python.with_name('pythonw.exe').is_file():python=python.with_name('pythonw.exe')
        service=win32com.client.Dispatch('Schedule.Service');service.Connect();folder=service.GetFolder('\\');name='PayrollDiagnostics-'+device_id[:8]
        args=subprocess.list2cmdline([str(entry),'flush','--home',str(home)])
        try:old=folder.GetTask(name)
        except Exception:old=None
        if old:
            action=old.Definition.Actions.Item(1)
            if Path(action.Path).resolve()!=python.resolve() or action.Arguments!=args:raise ValueError('同名定时任务不属于当前组件，拒绝替换')
        task=service.NewTask(0);task.RegistrationInfo.Description=NOTICE;task.Settings.Enabled=True;task.Settings.StartWhenAvailable=True;task.Settings.ExecutionTimeLimit='PT10M';task.Settings.MultipleInstances=2
        user=subprocess.run(['whoami'],capture_output=True,text=True,check=True,timeout=10).stdout.strip()
        task.Principal.UserId=user;task.Principal.LogonType=3;task.Principal.RunLevel=0
        trigger=task.Triggers.Create(2);trigger.StartBoundary=datetime.datetime.now().replace(microsecond=0).isoformat();trigger.DaysInterval=1;trigger.Repetition.Interval='PT5M';trigger.Repetition.Duration='P1D'
        logon=task.Triggers.Create(9);logon.UserId=user
        action=task.Actions.Create(0);action.Path=str(python);action.Arguments=args;action.WorkingDirectory=str(entry.parent)
        folder.RegisterTaskDefinition(name,task,6,user,None,3)
        return {'enabled':True,'kind':'windows_task','name':name}
    if sys.platform=='darwin':
        label='com.payroll.diagnostics.'+device_id[:8];directory=Path.home()/'Library/LaunchAgents';directory.mkdir(parents=True,exist_ok=True);path=directory/(label+'.plist')
        data={'Label':label,'ProgramArguments':[str(python),str(entry),'flush','--home',str(home)],'RunAtLoad':True,'StartInterval':300,'ProcessType':'Background'}
        if path.exists() and plistlib.loads(path.read_bytes()).get('Label')!=label:raise ValueError('已有启动项不属于本组件')
        path.write_bytes(plistlib.dumps(data));path.chmod(0o600)
        domain='gui/'+str(os.getuid());subprocess.run(['launchctl','bootout',domain,str(path)],capture_output=True,timeout=10)
        subprocess.run(['launchctl','bootstrap',domain,str(path)],capture_output=True,check=True,timeout=10)
        return {'enabled':True,'kind':'macos_launchagent','name':label}
    return {'enabled':False,'kind':'unsupported','reason':'每次核算后仍尝试同步；当前系统未配置定时重试'}

def install(home=None,profile_path=None,org=None,skip_schedule=False):
    org=org or organization();home=private_dir(home_path(home));secure=home/'client.secure'
    if secure.exists():
        config=read_json_secret(secure)
        if config['repository']!=org['repository'] or config['key_id']!=org['key_id']:raise ValueError('已有诊断凭据属于其他组织或密钥，停止覆盖')
        if profile_path and verify_profile(read_json(profile_path),org)['device_id']!=config['device_id']:raise ValueError('此电脑已有另一设备身份，不静默替换')
    elif profile_path:
        config=verify_profile(read_json(profile_path),org);save_secret(secure,canonical(config))
    else:
        ssh=ed25519.Ed25519PrivateKey.generate();config={'device_id':str(uuid.uuid4()),'repository':org['repository'],'key_id':org['key_id'],'notice_version':org['notice_version'],
          'ssh_private_pem':ssh.private_bytes(serialization.Encoding.PEM,serialization.PrivateFormat.OpenSSH,serialization.NoEncryption()).decode(),
          'ssh_public_key':ssh.public_key().public_bytes(serialization.Encoding.OpenSSH,serialization.PublicFormat.OpenSSH).decode(),'pending_enrollment':True}
        save_secret(secure,canonical(config))
        request={k:config[k] for k in ('device_id','repository','key_id','notice_version','ssh_public_key')};request['schema']='payroll-diagnostics-enrollment-1';(home/'管理员登记请求.json').write_bytes(canonical(request))
    result={'version':VERSION,'device_id':config['device_id'],'notice':NOTICE,'notice_version':org['notice_version'],'installed_at':datetime.datetime.now(datetime.timezone.utc).isoformat(),'pending_enrollment':config.get('pending_enrollment',False)}
    if skip_schedule:result['schedule']={'enabled':False,'kind':'isolated_test'}
    else:
        try:result['schedule']=schedule(home,config['device_id'])
        except Exception as exc:result['schedule']={'enabled':False,'kind':'registration_failed','category':type(exc).__name__}
    (home/'installation.json').write_bytes(canonical(result))
    return result

def main():
    ap=argparse.ArgumentParser(description='安装公司核算诊断记录组件');ap.add_argument('--home',type=Path);ap.add_argument('--profile',type=Path);ap.add_argument('--skip-schedule',action='store_true');ap.add_argument('--check',action='store_true');a=ap.parse_args();org=organization()
    if a.profile:verify_profile(read_json(a.profile),org)
    if a.check:print('组织公钥与设备配置检查通过');return
    result=install(a.home,a.profile,org,a.skip_schedule)
    print(json.dumps({'诊断组件版本':result['version'],'待管理员登记':result['pending_enrollment'],'五分钟定时同步':result['schedule']['enabled'],'设备编号':result['device_id']},ensure_ascii=False))
    if result['pending_enrollment']:print('云端同步等待管理员登记；登记请求在本机诊断记录目录，记录不会因此被称作已上传。')
    if not result['schedule']['enabled'] and not a.skip_schedule:print('定时同步未启用；每次核算后仍会尝试同步，管理员可查看安装状态排查。')
if __name__=='__main__':main()
