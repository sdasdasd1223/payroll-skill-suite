"""Additional scenarios, separate from the initial comparison suite."""
import argparse,json,sys,importlib.util,tempfile,copy,subprocess,csv
from pathlib import Path
from datetime import date
sys.dont_write_bytecode=True
def load(name,p):
 s=importlib.util.spec_from_file_location(name,p);m=importlib.util.module_from_spec(s);sys.modules[name]=m;s.loader.exec_module(m);return m
def run(root):
 a=load('hard_att',root/'attendance-calculator/scripts/process_attendance.py')
 r=load('hard_reward',root/'attendance-calculator/scripts/reward_policy.py')
 p=load('hard_pay',root/'payroll-operations/scripts/payroll_core.py')
 w=load('hard_workflow',root/'payroll-operations/scripts/workflow_audit.py')
 n=load('hard_norm',root/'social-insurance-fund/scripts/normalized_social.py')
 results=[]
 def check(id,title,expected,fn):
  try:actual=fn();passed=actual==expected;error=None
  except Exception as e:actual=None;passed=False;error=type(e).__name__+':'+str(e)
  results.append(dict(id=id,title=title,expected=expected,actual=actual,passed=passed,error=error,source_type='synthetic_additional'))
 base=dict(name='测试甲',date='2026-09-14',tier=50,checkout='22:06',next_arrival='08:50',next_workday=True)
 check('H01','明确当日上午不视为次日凌晨',0,lambda:r.evaluate_claim(dict(base,checkout='2026-09-14 08:00'),'2026-09')['amount'])
 check('H02','明确次日凌晨仍核准50',50,lambda:r.evaluate_claim(dict(base,checkout='2026-09-15 00:02'),'2026-09')['amount'])
 check('H03','次日到岗完整日期不匹配须待核',None,lambda:r.evaluate_claim(dict(base,next_arrival='2026-09-14 08:50'),'2026-09')['amount'])
 check('H04','同档重复申报只计一次',50,lambda:r.audit_claims([base,dict(base)],'2026-09')['totals']['测试甲']['amount'])
 check('H05','不同档申报冲突不自动选高档',None,lambda:r.audit_claims([dict(base,checkout='次日00:02'),dict(base,tier=200,checkout='次日00:02',application_end='00:00',application_status='approved')],'2026-09')['totals']['测试甲']['amount'])
 check('H06','月份最后一天缺次月工作记录待核',None,lambda:r.evaluate_claim(dict(base,date='2026-09-30',next_workday=None),'2026-09')['amount'])
 check('H07','不可解析最低工资不崩溃','pending',lambda:p.sick_leave_deduction(5800,8,184,80,'待核')['status'])
 check('H08','月底递增周期无同日待核','pending',lambda:p.salary_progression(4500,6000,6,'2026-01-31','2026-02',['2026-02-02'])['status'])
 check('H09','已完整生效档位进薪资本体',4800,lambda:p.salary_progression(4500,6000,6,'2026-06-01','2026-07',['2026-07-01'])['salary'])
 check('H10','离职未完成新周期不递增',4500,lambda:p.salary_progression(4500,6000,6,'2026-06-15','2026-07',['2026-07-15'],'2026-07-20')['salary'])
 check('H11','月中新增档位只折算生效计薪日',150,lambda:p.salary_progression(4500,6000,6,'2026-06-15','2026-07',['2026-07-01','2026-07-15'])['allowance'])
 check('H12','标准化社保空表不能通过',False,lambda:n.build([],'2026-09')['passed'])
 check('H13','社保错月不并入本月',0,lambda:len(n.build([dict(company='测试',name='甲',certificate='ID',month='2026-08',insurance='pension',employee=100,employer=200,source='f',source_row=4)],'2026-09')['people']))
 rec=dict(company='广思',name='甲',certificate='ID',month='2026-09',insurance='pension',employee=100,employer=190,source='f',source_row=4)
 check('H14','广思单位金额采用源数190',190,lambda:n.build([rec],'2026-09')['people'][0]['items']['pension']['employer'])
 check('H15','广思重复同险种不取后一笔',None,lambda:n.build([rec,dict(rec,employer=300)],'2026-09')['people'][0]['items']['pension']['employer'])
 with tempfile.TemporaryDirectory() as tmp:
  temp=Path(tmp);source=temp/'source.json';source.write_text('{}')
  sources={'source':dict(path='source.json',sha256=w.sha256(source),month='2026-09',role='attendance_confirmed')}
  record=dict(month='2026-09',policy_version='2026.11.16.1',stage='payroll',expected_employee_keys=['id1'],employees=[dict(employee_key='id1',items={i:dict(status='confirmed',source_ids=['source']) for i in w.REQUIRED['payroll']})],sources=sources,confirmations=[dict(source_id='source',sha256=w.sha256(source),confirmed_by='测试人资',confirmed_at='2026-11-16')])
  # V4记录新增源件假别观察；同一22案例仅更新完整输入样本。
  record['employees'][0]['observations']={'maternity':{'record_count':0,'coverage_complete':True,'source_ids':['source']}}
  check('H16','旧自报完成不能代替实际文件复核',False,lambda:w.validate_run(record,temp)['passed'])
  changed=copy.deepcopy(record);changed['sources']['source']['month']='2026-08'
  check('H17','上月确认文件不能冒充本月',False,lambda:w.validate_run(changed,temp)['passed'])
  changed=copy.deepcopy(record);changed['employees'][0]['items'].pop('reward')
  check('H18','漏奖励状态阻止交接',False,lambda:w.validate_run(changed,temp)['passed'])
  changed=copy.deepcopy(record);changed['expected_employee_keys'].append('id2')
  check('H19','人员漏一人阻止交接',False,lambda:w.validate_run(changed,temp)['passed'])
  source.write_text('{"changed":true}')
  check('H20','确认后文件内容变化阻止交接',False,lambda:w.validate_run(record,temp)['passed'])
 check('H21','已离开组织并缺手机完整列原因',2,lambda:len(w.payslip_eligibility(dict(account_query_complete=True,account_count=1,is_activated=True,is_cross_tenant=False,has_left=True,phone=None,net=10),'2026-10-10')['reasons']))
 check('H22','合法零实发不当缺失',True,lambda:w.payslip_eligibility(dict(account_query_complete=True,account_count=1,is_activated=True,is_cross_tenant=False,has_left=False,phone='13800000000',net=0),'2026-10-10')['included'])
 return results
def main():
 ap=argparse.ArgumentParser();ap.add_argument('--skills',type=Path,required=True);ap.add_argument('--output',type=Path,required=True);args=ap.parse_args()
 if args.output.exists():raise SystemExit('不覆盖证据')
 results=run(args.skills);out=dict(case_count=len(results),passed=sum(x['passed'] for x in results),results=results)
 args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_text(json.dumps(out,ensure_ascii=False,indent=2,default=str),encoding='utf-8');print(json.dumps({k:v for k,v in out.items() if k!='results'}))
if __name__=='__main__':main()
