"""只停用本套件的舊自動記錄元件，既有資料保留。"""
from pathlib import Path
import argparse,datetime,json,os,plistlib,shutil,subprocess,sys,uuid
def main():
 p=argparse.ArgumentParser();p.add_argument('--codex-home',type=Path,required=True);a=p.parse_args();original=a.codex_home.absolute();home=original.resolve()
 if original.is_symlink() or home==Path(home.anchor):raise ValueError('安裝目錄不能是連結或磁碟根目錄')
 component=home/'payroll-diagnostics';backup=home/'skill-backups'/('disabled-recording-'+datetime.datetime.now().strftime('%Y%m%d-%H%M%S')+'-'+uuid.uuid4().hex[:6]);actions=[]
 if os.name=='nt':
  collector=str(component/'collector.py').replace("'","''")
  code="$ErrorActionPreference='Stop'; $entry='"+collector+"'; foreach($task in (Get-ScheduledTask -TaskName 'PayrollDiagnostics-*' -ErrorAction SilentlyContinue)) { if(@($task.Actions | Where-Object { $_.Arguments -like ('*'+$entry+'*') }).Count -gt 0) { Disable-ScheduledTask -InputObject $task | Out-Null; Stop-ScheduledTask -InputObject $task; Unregister-ScheduledTask -InputObject $task -Confirm:$false } }"
  result=subprocess.run(['powershell.exe','-NoProfile','-NonInteractive','-Command',code],capture_output=True,timeout=45)
  if result.returncode:raise RuntimeError('Windows 舊同步任務尚未停用')
 elif sys.platform=='darwin':
  for path in (Path.home()/'Library/LaunchAgents').glob('com.payroll.diagnostics.*.plist'):
   data=plistlib.loads(path.read_bytes())
   if str(component/'collector.py') not in data.get('ProgramArguments',[]):continue
   subprocess.run(['launchctl','bootout','gui/'+str(os.getuid()),str(path)],capture_output=True,timeout=15)
   check=subprocess.run(['launchctl','print','gui/'+str(os.getuid())+'/'+data['Label']],capture_output=True,timeout=15)
   if check.returncode==0:raise RuntimeError('macOS 舊背景任務尚未停用')
   backup.mkdir(parents=True,exist_ok=True);shutil.move(str(path),str(backup/path.name));actions.append('舊背景任務已停用')
 if component.exists():
  if component.is_symlink() or component.resolve().parent!=home:raise ValueError('舊元件路徑不符')
  backup.mkdir(parents=True,exist_ok=True);shutil.move(str(component),str(backup/component.name));actions.append('舊記錄程式已移至備份')
 shortcut=Path.home()/'Desktop/薪資案件管理員.lnk'
 if os.name=='nt' and home==Path.home()/'.codex' and shortcut.is_file():
  backup.mkdir(parents=True,exist_ok=True);shutil.move(str(shortcut),str(backup/shortcut.name))
 print(json.dumps({'已停用舊記錄功能':True,'既有資料保留':True,'處理':actions},ensure_ascii=True))
if __name__=='__main__':main()
