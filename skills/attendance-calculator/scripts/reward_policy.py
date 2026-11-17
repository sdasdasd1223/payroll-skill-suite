"""Normalize and audit declared rewards; missing evidence stays pending, never zero.
CLI: python reward_policy.py --month YYYY-MM --claims normalized.json --output audit.json
"""
import argparse, json, re
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

def minutes(value, overnight=False):
    if value is None or str(value).strip()=='':return None
    text=str(value)
    matches=re.findall(r'(\d{1,2}):(\d{2})',text)
    if not matches:return None
    h,m=map(int,matches[0])
    if h>24 or m>59 or (h==24 and m):return None
    n=h*60+m
    if overnight and (h<12 or '次日' in text):n+=1440 if n<1440 else 0
    return n

def evaluate_claim(claim, month):
    result=dict(claim,amount=None,status='pending',reason='')
    def done(amount,status,reason):
        result.update(amount=amount,status=status,reason=reason);return result
    try:d=date.fromisoformat(str(claim.get('date',''))[:10])
    except ValueError:return done(None,'pending','申报日期无法解析')
    if d.strftime('%Y-%m')!=month:return done(0,'excluded','申报日期不属于核算月')
    if not claim.get('name'):return done(None,'pending','员工身份缺失')
    tier=claim.get('tier')
    if tier not in (0,50,200):return done(None,'pending','申报档位不明，不能从实际打卡擅自升级')
    if d.weekday()==4:return done(0,'excluded','现行制度排除周五晚上')
    work=claim.get('next_workday')
    if work is False:return done(0,'excluded','次日非工作日')
    if work is not True:return done(None,'pending','缺少次日工作日依据（跨月须补次月记录）')
    def absolute_stamp(value):
        text=str(value or '')
        found=re.search(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})[ T](\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:\d{2})?)',text)
        if not found:return False,None
        try:
            y,m,day=map(int,re.split(r'[-/]',found[1]))
            clock=found[2];hour,tail=clock.split(':',1)
            stamp=datetime.fromisoformat(f'{y:04d}-{m:02d}-{day:02d}T{int(hour):02d}:{tail}'.replace('Z','+00:00'))
            if stamp.tzinfo is not None:stamp=stamp.astimezone(timezone(timedelta(hours=8))).replace(tzinfo=None)
            return True,stamp
        except ValueError:return True,None
    def dated_minutes(value, anchor, overnight=False):
        explicit,stamp=absolute_stamp(value)
        if explicit:
            if stamp is None:return None
            delta=(stamp.date()-anchor).days
            if delta not in (0,1) or (not overnight and delta!=0):return None
            return delta*1440+stamp.hour*60+stamp.minute
        return minutes(value,overnight)
    end=dated_minutes(claim.get('checkout'),d,True)
    arrival=dated_minutes(claim.get('next_arrival'),d+timedelta(days=1))
    if end is None or arrival is None:return done(None,'pending','缺少有效上下班打卡')
    threshold=1440 if tier==200 or claim.get('midnight_tier') else 1320
    if end<threshold:return done(0,'rejected','下班未达到申报档位')
    if arrival>540 or tier==0:return done(0,'confirmed_zero','不符合现金奖励到岗条件；报销资格另核')
    if threshold==1440:
        if claim.get('application_status') in ('rejected','cancelled'):
            return done(0,'rejected','加班申请未批准或已撤销')
        has_start,start_stamp=absolute_stamp(claim.get('application_start'))
        has_end,end_stamp=absolute_stamp(claim.get('application_end'))
        if has_start and (start_stamp is None or start_stamp.date()!=d):
            return done(None,'pending','申请实际开始日期与所申报当晚不一致，须核对原登记；不能移动绝对日期')
        app=dated_minutes(claim.get('application_end'),d,True)
        start_clock=(start_stamp.hour*60+start_stamp.minute) if has_start and start_stamp is not None else minutes(claim.get('application_start'))
        # 既有制度允许“本业务日晚上开始、同业务日00点结束”的登记简写。
        # 已有绝对开始日时必须等于申报日；前一晚的正常跨日申请不能再次推一天。
        same_day_midnight=(app==0 and start_clock is not None and start_clock>=720
                           and (not has_start or start_stamp.date()==d))
        if has_end and (end_stamp is None or has_start and end_stamp<start_stamp and not same_day_midnight):
            return done(None,'pending','申请绝对起止日期无效或先后矛盾，须核对原登记')
        if same_day_midnight:
            app=1440
            result['application_midnight_interpretation']='沿用原制度的同业务日午夜登记简写；原始起止不改'
        if app==0:
            return done(None,'pending','申请明确结束于申报日凌晨，未证明当晚；不能将明确日期自动推到次日')
        evidence=claim.get('application_record_ids') or claim.get('application_record_id')
        valid_status=claim.get('application_status')=='approved' or (claim.get('application_status')=='matched' and evidence and claim.get('application_source'))
        if not valid_status or app is None:
            return done(None,'pending','午夜档缺少有效加班登记记录或可核对的申请结束时间')
        if app<1440:return done(0,'rejected','申请未证明加班至午夜')
    return done(tier,'calculated','问卷、打卡、次日工作日及必要申请核验通过')

def audit_claims(claims, month):
    if not claims:
        return {'month':month,'rows':[],'totals':{},'pending_count':1,'issues':['奖励问卷为空，缺少完整收取或人工确认为零的依据']}
    rows=[evaluate_claim(c,month) for c in claims]
    groups={}
    for i,r in enumerate(rows):groups.setdefault((r.get('name'),r.get('date')),[]).append(i)
    for indexes in groups.values():
        valid=[i for i in indexes if rows[i]['status']=='calculated']
        # Conflicting positive declarations are unresolved; do not invent a tie-break policy.
        amounts={rows[i]['amount'] for i in valid}
        if len(amounts)>1:
            for i in valid:rows[i].update(amount=None,status='pending',reason='同人同日申报档位冲突，需确认采用记录')
        elif valid:
            for i in valid[1:]:rows[i].update(amount=0,status='duplicate',reason='同人同日重复申报，仅计一条')
    names={r['name'] for r in rows if r.get('name')}
    totals={}
    for name in names:
        own=[r for r in rows if r.get('name')==name]
        pending=any(r['status']=='pending' for r in own)
        totals[name]={'amount':None if pending else sum(r['amount'] or 0 for r in own),'status':'pending' if pending else 'calculated'}
    return {'month':month,'rows':rows,'totals':totals,'pending_count':sum(r['status']=='pending' for r in rows)}

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--month',required=True);ap.add_argument('--claims',type=Path,required=True);ap.add_argument('--output',type=Path,required=True)
    args=ap.parse_args();out=audit_claims(json.loads(args.claims.read_text(encoding='utf-8-sig')),args.month)
    if args.output.resolve()==args.claims.resolve():raise SystemExit('不得覆盖源文件')
    if args.output.exists():raise SystemExit('输出已存在，请另存新版本')
    args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
    raise SystemExit(2 if out['pending_count'] else 0)
if __name__=='__main__':main()
