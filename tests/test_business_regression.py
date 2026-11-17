"""Same observable business assertions against frozen baseline and candidate.
Synthetic cases have literal expectations; historical checks are reported separately.
"""
import argparse, importlib.util, json, sys, hashlib, traceback
from pathlib import Path
from datetime import date
from argparse import Namespace
from decimal import Decimal
sys.dont_write_bytecode = True

def module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    m = importlib.util.module_from_spec(spec)
    sys.modules[name] = m
    spec.loader.exec_module(m)
    return m

def run(root):
    a = module('attendance_test', root/'attendance-calculator/scripts/process_attendance.py')
    s = module('social_test', root/'social-insurance-fund/scripts/build_social_insurance_fund.py')
    results = []
    def check(id, title, expected, fn, rule):
        try:
            actual = fn()
            ok = actual == expected
            error = None
        except Exception as exc:
            actual, ok, error = None, False, type(exc).__name__+': '+str(exc)
        results.append(dict(id=id, title=title, source_type='synthetic', rule=rule,
                            expected=expected, actual=actual, passed=ok, error=error))
    def summary(overrides=None, history=True, overtime=True, roster=True):
        row=dict(name='测试甲', UserID='1000000000000000001', attendance_type='需要打卡',
                 late_minutes=2,missed_punch_count=3,personal_leave_hours=1,sick_leave_hours=1,
                 annual_leave_days=0,childcare_leave_days=0,required_work_hours=184,
                 reward_status='confirmed_zero')
        row.update(overrides or {})
        person=dict(name='测试甲',UserID=row['UserID'],job_level='组员',hire_date='2020-01-01')
        if overrides and 'job_level' in overrides: person['job_level']=overrides['job_level']
        h=dict(person,prior_annual_leave_entitlement_days=7,prior_annual_leave_remaining_days=2,
               prior_sick_leave_hours_year=3,history_month='2026-08')
        o=dict(person,overtime_balance_hours=8,prior_leave_balance_hours=2)
        data={'a':[row],'r':[person] if roster else [],'h':[h] if history else [],'o':[o] if overtime else []}
        old=a.read_table
        a.read_table=lambda path:data.get(path,[])
        try:
            return a.build_summary(Namespace(month='2026-09',attendance='a',roster='r',history='h',overtime='o',paid_days=23,hours_per_day=8))[0][0]
        finally:a.read_table=old
    annual='business-rules.md:周年月起享首年额度，向上取整至0.5天'
    check('A01','周年月中入职者当月新增年假',1,lambda:a.annual_leave_entitlement(date(2025,8,15),2026,date(2026,8,1)),annual)
    check('A02','周年月之前不提前新增',0,lambda:a.annual_leave_entitlement(date(2025,8,15),2026,date(2026,7,1)),annual)
    check('A03','周年月之后保持额度',1,lambda:a.annual_leave_entitlement(date(2025,8,15),2026,date(2026,9,1)),annual)
    check('A04','满五年年假7天',7,lambda:a.annual_leave_entitlement(date(2020,1,1),2026,date(2026,9,1)),annual)
    check('A05','斜线日期不能被拆开',['2024-09-11'],lambda:[str(x) for x in a.split_birthdates('2024/09/11')],'input-output-schema.md:出生日期格式')
    check('A06','育儿假跨周期清零',0,lambda:a.childcare_note(0,[date(2024,9,1)],date(2026,10,1),10,'截至2026年9月（含），有10天育儿假，已使用10天')[1],'business-rules.md:育儿假新周期清零')
    check('A07','育儿假超额仍显示11天',11,lambda:a.childcare_note(1,[date(2024,9,1)],date(2026,9,1),10,'截至2026年9月（含），有10天育儿假，已使用10天')[1],'business-rules.md:超额保留事实')
    check('A08','迟到按组员每分钟3元',-6,lambda:summary()['迟到扣款情况'],'business-rules.md:迟到')
    check('A09','第三次缺卡扣50元',-50,lambda:summary()['漏打卡扣款情况'],'business-rules.md:缺卡')
    check('A10','缺职级不能把总扣款伪装成完整金额','',lambda:summary({'job_level':'未知'})['考勤扣款情况（合计）'],'职级不明留空待核')
    check('A11','缺上月年假余额不得默认完整额度','',lambda:summary(history=False)['上月年假剩余天数'],'禁止把缺失资料视为已确认金额/余额')
    check('A12','缺病假累计不得默认为零','',lambda:summary(history=False)['上月病假时长'],'缺历史数据待核')
    check('A13','缺加班结余不得默认为零','',lambda:summary(overtime=False)['累计加班结余（小时）'],'缺加班资料待核')
    check('A14','事假累计不重复加本月',3,lambda:summary()['累计事假时长（小时）'],'上月请假结余2+本月1')
    check('A15','年假扣本月使用后余额',1,lambda:summary({'annual_leave_days':1})[next(k for k in summary() if '在职结余时间' in k)],'上月确认2-本月1')
    check('A16','病假与年假明细不可遗漏',True,lambda:all(x in summary({'sick_leave_detail':'9日1小时','annual_leave_detail':'10日1天'})['考勤情况'] for x in ['病假','年假']),'考勤情况含迟到/事假/病假/年假')
    check('A17','未核验奖励金必须产生异常',True,lambda:'奖励金' in summary({'reward_status':''})['异常提示'],'奖励金处理状态不得遗漏')
    check('A18','核准奖励金进入考勤表',50,lambda:summary({'reward_status':'calculated','reward_amount':50})['奖励金'],'奖励核准金额落表')
    check('A19','未知考勤类型不能静默视为免打卡',True,lambda:'考勤类型' in summary({'attendance_type':'未知'})['异常提示'],'未知类型待核')
    check('A20','历史余额不会被重置',2,lambda:summary()['上月年假剩余天数'],'人工确认余额优先')
    def social(rows):
        old=s.read_table;s.read_table=lambda path:rows
        try:
            records,detail,issues=s.build_records({'测试公司':{'pension':Path('fake.xlsx')}})
            return records,detail,issues
        finally:s.read_table=old
    headers=['姓名','证件号码','缴费基数','应缴费额','费率']
    normal=[headers,['测试甲','TEST-A','10000','800','8%']]
    check('S01','普通养老单位金额1600',Decimal('1600.00'),lambda:next(iter(social(normal)[0].values())).pension_employer,'现行公司规则16%')
    check('S02','缺公积金主表不能填零',None,lambda:next(iter(social(normal)[0].values())).housing_employee,'social-insurance-fund:缺险种不填零')
    check('S03','缺公积金仍有异常',True,lambda:any(x[2]=='housing_fund' for x in social(normal)[2]),'缺项必须留异常')
    duplicate=[headers,['测试甲','TEST-A','10000','800','8%'],['测试甲','TEST-B','12000','960','8%']]
    check('S04','同名不同证件不能后行覆盖为960',True,lambda:len(social(duplicate)[0])==2 or next(iter(social(duplicate)[0].values())).pension_employee is None,'重名待核不得自动匹配')
    check('S05','金额无法解析不是有效零值',None,lambda:s.parse_decimal('待确认'),'无法解析必须待核')
    reward_path=root/'attendance-calculator/scripts/reward_policy.py'
    if reward_path.exists():r=module('reward_test',reward_path)
    else:r=None
    def reward(**kw):
        if r is None:raise NotImplementedError('旧通用技能未提供奖励金核验程序')
        base=dict(name='测试甲',date='2026-09-14',tier=50,checkout='22:06',next_arrival='08:50',next_workday=True)
        base.update(kw)
        return r.evaluate_claim(base,'2026-09')
    for id,title,expected,kwargs in [
        ('R01','22点档核准50元',50,{}),('R02','未到9点半无现金奖励',0,{'next_arrival':'09:20'}),
        ('R03','午夜档缺申请不自动发200元',None,{'tier':200,'checkout':'次日00:02'}),
        ('R04','午夜档申请到00点可核准',200,{'tier':200,'checkout':'次日00:02','application_end':'00:00','application_status':'approved'}),
        ('R05','非目标月申报排除',0,{'date':'2026-07-02'}),('R06','次日休息不发奖励',0,{'next_workday':False}),
        ('R07','缺次日打卡属于待核',None,{'next_arrival':''})]:
        check(id,title,expected,lambda kw=kwargs:reward(**kw)['amount'],'工资总流程原阶段7')
    corepath=root/'payroll-operations/scripts/payroll_core.py'
    p=module('payroll_test',corepath) if corepath.exists() else None
    def core(fn,*args,**kw):
        if p is None:raise NotImplementedError('旧套件未提供统一工资计算核心')
        return getattr(p,fn)(*args,**kw)
    check('P01','非实习离职累计事假按174',-100,lambda:core('personal_leave_deduction',5800,3,False,184),'5800/174*3，负数')
    check('P02','实习离职仍按当月工时',-100,lambda:core('personal_leave_deduction',4600,4,True,184),'4600/184*4')
    check('P03','产假差额正数进入个人扣款',4192.57,lambda:core('maternity_difference',10004,2000,21,11),'原技能已确认算例')
    check('P04','病假跨9天时拒绝自选拆分',True,lambda:core('sick_leave_deduction',5800,24,184,64,2000)['status']=='pending','现行制度未定义跨阈值算法，不推断')
    check('P05','三个月试用期不套差额除五',0,lambda:core('increment_amount',4500,6000,3),'三个月不逐月递增')
    check('P06','六个月已确认方案月增额300',300,lambda:core('increment_amount',4500,6000,6),'(6000-4500)/5')
    return results

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--skills',type=Path,required=True);ap.add_argument('--output',type=Path,required=True);args=ap.parse_args()
    if args.output.exists():raise SystemExit('拒绝覆盖已有测试证据，请使用新轮次路径')
    results=run(args.skills)
    serial=json.loads(json.dumps(results,ensure_ascii=False,default=lambda x:float(x) if isinstance(x,Decimal) else str(x)))
    payload=dict(skills=str(args.skills.resolve()),case_count=len(serial),passed=sum(r['passed'] for r in serial),results=serial)
    args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_text(json.dumps(payload,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:v for k,v in payload.items() if k!='results'},ensure_ascii=True))
if __name__=='__main__':main()
