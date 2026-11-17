"""Portable input regressions. Synthetic fixtures only; no employee data or network calls."""
import argparse,importlib.util,json,tempfile,zipfile,sys
from pathlib import Path
sys.dont_write_bytecode=True
def load(name,path):
 s=importlib.util.spec_from_file_location(name,path);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);return m
def main():
 ap=argparse.ArgumentParser();ap.add_argument('--skills',required=True,type=Path);ap.add_argument('--output',required=True,type=Path);args=ap.parse_args()
 if args.output.exists():raise SystemExit('请用新结果文件名，不覆盖旧证据')
 a=load('input_att',args.skills/'attendance-calculator/scripts/process_attendance.py');r=load('input_reward',args.skills/'attendance-calculator/scripts/reward_policy.py')
 cases=[]
 def check(id,expected,fn):
  try:actual=fn();error=None;passed=actual==expected
  except Exception as e:actual=None;error=type(e).__name__+':'+str(e);passed=False
  cases.append(dict(id=id,expected=expected,actual=actual,error=error,passed=passed))
 base=dict(name='虚构甲',date='2026-09-14',tier=200,checkout='次日00:02',next_arrival='08:50',next_workday=True,application_status='approved',application_end='2026-09-15T00:00:00+08:00')
 for id,expected,changes in [
  ('matched_registration',200,dict(application_status='matched',application_source='source.json',application_record_ids=['record1'])),
  ('untraced_match_pending',None,dict(application_status='matched')),
  ('timezone_not_clock',0,dict(application_end='2026-09-14T23:00:00+08:00')),
  ('wrong_application_day',None,dict(application_end='2026-09-20T00:00:00+08:00')),
  ('same_day_midnight_with_evening_start',200,dict(application_start='2026-09-14T18:30:00+08:00',application_end='2026-09-14T00:00:00+08:00')),
  ('same_day_midnight_without_start',None,dict(application_end='2026-09-14T00:00:00+08:00')),
  ('legacy_plain_midnight',200,dict(application_end='00:00')),
  ('explicit_morning_not_night',0,dict(checkout='2026-09-14 08:00'))]:
  check(id,expected,lambda changes=changes:r.evaluate_claim(dict(base,**changes),'2026-09')['amount'])
 with tempfile.TemporaryDirectory() as folder:
  import openpyxl
  p=Path(folder)/'fixture.xlsx';b=openpyxl.Workbook();ws=b.active;ws.append(['UserID','name']);ws.append(['1000000000000000001','虚构甲']);ws.append(['1000000000000000002','虚构乙']);b.save(p);b.close()
  with zipfile.ZipFile(p) as z:parts={name:z.read(name) for name in z.namelist()}
  parts['xl/worksheets/sheet1.xml']=parts['xl/worksheets/sheet1.xml'].replace(b'dimension ref="A1:B3"',b'dimension ref="A1"')
  with zipfile.ZipFile(p,'w') as z:
   for name,value in parts.items():z.writestr(name,value)
  check('invalid_dimension_does_not_drop_rows',2,lambda:len(a.read_xlsx(p)))
 out=dict(case_count=len(cases),passed=sum(x['passed'] for x in cases),results=cases)
 args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps({k:v for k,v in out.items() if k!='results'}))
if __name__=='__main__':main()
