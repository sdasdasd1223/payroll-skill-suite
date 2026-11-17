"""跨系统事务安装：四技能、诊断程序与现有数据分别处理。"""
from pathlib import Path
import argparse, datetime, hashlib, json, os, re, shutil, subprocess, sys, tempfile, uuid

NAMES=('payroll-operations','attendance-calculator','attendance-settlement','social-insurance-fund')
ROOT=Path(__file__).resolve().parent
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def verify():
    manifest=json.loads((ROOT/'suite-manifest.json').read_text('utf-8-sig'))
    for name,value in manifest['files'].items():
        p=ROOT/name
        if not p.resolve().is_relative_to(ROOT) or p.is_symlink() or not p.is_file() or sha(p)!=value:raise ValueError('安装包完整性检查失败：'+name)
    for name in NAMES:
        text=(ROOT/'skills'/name/'SKILL.md').read_text('utf-8')
        if not re.search(r'^name:\s*'+re.escape(name)+r'\s*$',text,re.M):raise ValueError('技能名称不符')
    return manifest

def safe_remove(path,home):
    path=Path(path);home=Path(home).resolve();resolved=path.resolve()
    allowed=(resolved.parent==home/'skills' and resolved.name in NAMES) or (resolved.parent==home and resolved.name=='payroll-diagnostics')
    if not allowed or path.is_symlink():raise ValueError('拒绝删除安装目标以外的目录')
    if path.is_dir():shutil.rmtree(path)

def main():
    ap=argparse.ArgumentParser(description='安装工资核算技能及自动诊断组件');ap.add_argument('--codex-home',type=Path);ap.add_argument('--device-profile',type=Path);ap.add_argument('--skip-diagnostic-schedule',action='store_true');ap.add_argument('--check-only',action='store_true');a=ap.parse_args()
    manifest=verify()
    if sys.version_info<(3,10):raise SystemExit('需要Python 3.10或以上')
    try:import cryptography
    except ImportError:
        subprocess.run([sys.executable,'-m','pip','install','--disable-pip-version-check','--no-input','cryptography>=44'],check=True)
    sys.path.insert(0,str(ROOT/'diagnostics'))
    from common import NOTICE,canonical,lock,organization,private_dir,read_json_secret,save_secret
    from admin import verify_profile
    org=organization(ROOT/'diagnostics/organization.json')
    profile_path=a.device_profile or ROOT/'device-profile.private.json'
    profile=None
    if profile_path.is_file():profile=verify_profile(json.loads(profile_path.read_text('utf-8-sig')),org)
    if a.check_only:print('技能、组件、摘要和设备签名检查通过');return
    home=(a.codex_home or Path(os.environ.get('CODEX_HOME',str(Path.home()/'.codex')))).resolve();home.mkdir(parents=True,exist_ok=True)
    state=private_dir(home/'payroll-install-state')
    with lock(state) as acquired:
        if not acquired:raise SystemExit('另一安装正在进行，请稍后重试')
        data=private_dir(home/'payroll-diagnostics-data')
        if (data/'client.secure').exists() and profile:
            existing=read_json_secret(data/'client.secure')
            if existing['device_id']!=profile['device_id'] or existing['key_id']!=profile['key_id']:raise ValueError('已有不同设备身份，保留旧记录与密钥，请管理员处理迁移')
        destinations={name:home/'skills'/name for name in NAMES};destinations['diagnostics']=home/'payroll-diagnostics'
        sources={name:ROOT/'skills'/name for name in NAMES};sources['diagnostics']=ROOT/'diagnostics'
        for target in destinations.values():
            if target.is_symlink():raise ValueError('拒绝替换符号链接安装目标')
        stamp=datetime.datetime.now().strftime('%Y%m%d_%H%M%S')+'_'+uuid.uuid4().hex[:6]
        backup=home/'skill-backups'/('payroll-skill-suite_'+stamp)
        existing=[name for name,p in destinations.items() if p.exists()]
        if existing:
            backup.mkdir(parents=True)
            for name in existing:shutil.copytree(destinations[name],backup/name,ignore=shutil.ignore_patterns('__pycache__','*.pyc'))
        with tempfile.TemporaryDirectory(prefix='payroll-install-',dir=home) as temporary:
            stage=Path(temporary).resolve()
            if stage.parent!=home:raise ValueError('安装暂存目录越界')
            for name in sources:shutil.copytree(sources[name],stage/name,ignore=shutil.ignore_patterns('__pycache__','*.pyc'))
            changed=[]
            try:
                (home/'skills').mkdir(exist_ok=True)
                for name,target in destinations.items():
                    if target.exists():target.rename(stage/('old-'+name))
                    changed.append(name);(stage/name).rename(target)
                # 使用已安装入口注册定时任务，任务不能指向下载/暂存目录。
                command=[sys.executable,str(home/'payroll-diagnostics/install_client.py'),'--home',str(data)]
                if profile_path.is_file():command.extend(['--profile',str(profile_path.resolve())])
                if a.skip_diagnostic_schedule:command.append('--skip-schedule')
                completed=subprocess.run(command,capture_output=True,text=True,encoding='utf-8',errors='replace',timeout=60,env=dict(os.environ,PYTHONUTF8='1'))
                if completed.returncode:raise RuntimeError('诊断凭据安装未完成，已触发程序回滚')
                result=json.loads((data/'installation.json').read_text('utf-8-sig'))
                record={'version':manifest['version'],'backup':str(backup) if existing else None,'diagnostics':result,'notice':NOTICE}
                (state/'last-install.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
            except BaseException:
                for name in reversed(changed):
                    target=destinations[name]
                    if target.exists():safe_remove(target,home)
                    old=stage/('old-'+name)
                    if old.exists():old.rename(target)
                raise
    print('四个技能及诊断组件已安装。云端是否已同步请查看诊断状态，不能仅凭安装成功判断。')
    if existing:print('旧版本备份：'+str(backup))
    if result.get('pending_enrollment'):print('通用安装包未预登记设备；请把本机“管理员登记请求.json”交给管理员，登记前记录仅保留本机加密队列。')
    if not result.get('schedule',{}).get('enabled') and not a.skip_diagnostic_schedule:print('定时同步暂未启用，每次核算后仍尝试同步；安装状态已记录供管理员排查。')
    print('请重新启动Codex，使新对话加载更新。')

if __name__=='__main__':
    try:main()
    except Exception as exc:
        print('安装未完成：'+type(exc).__name__+'；请管理员核对安装包、依赖和设备配置。',file=sys.stderr)
        raise SystemExit(2)
