"""自动收集明确核算入口的脱敏诊断；不读取聊天库、不扫描目录。"""
from __future__ import annotations
import argparse, datetime, hashlib, json, os, re, subprocess, sys, time, uuid
from pathlib import Path
from common import VERSION, SCHEMA, canonical, database, digest, encrypt, home_path, organization, private_dir, read_json, read_json_secret, sign_envelope, sync_client

ENTRIES={'workflow_audit','crossstage_audit','audit_payroll_workbook','process_attendance','insurance_merge','diagnostics_smoke'}
STAGES={'attendance','insurance','payroll','tax','payslip'}
PATTERNS={
 'SOURCE_CHANGED':('来源文件不存在','内容已变化','SHA','摘要不一致','找不到来源'),
 'TEMPLATE_MISMATCH':('模板','样式','行高','打印','填色','批注'),
 'FORMULA_INVALID':('公式','Excel错误','计算列','常量'),
 'CONFIRMATION_MISSING':('确认','核定','未决'),
 'ATTENDANCE_HANDOFF':('假别','交接'),
 'REWARD_EVIDENCE':('奖励','申请','问卷'),
 'SCOPE_INCOMPLETE':('人员','范围','名单'),
}
def category(value):
    text=str(value)
    return next((key for key,terms in PATTERNS.items() if any(x in text for x in terms)),'OTHER_CHECK')
def safe_read(path):
    try:
        path=Path(path)
        return read_json(path) if path.is_file() and path.stat().st_size<=16*1024*1024 else {}
    except (OSError,ValueError,TypeError):return {}
def arg(args,key):
    try:return args[args.index(key)+1]
    except (ValueError,IndexError):return None

def fingerprint(path,role='other'):
    path=Path(path)
    if not path.is_file():return None
    sha=hashlib.sha256()
    with path.open('rb') as stream:
        for part in iter(lambda:stream.read(1024*1024),b''):sha.update(part)
    extension=path.suffix.lower()
    if not re.fullmatch(r'\.[a-z0-9]{1,8}',extension):extension='other'
    if role not in {'attendance_raw','attendance_confirmed','payroll_confirmed','reward_claims','overtime','roster','insurance','template','output','record'}:role='other'
    return {'sha256':sha.hexdigest(),'bytes':path.stat().st_size,'extension':extension,'role':role}

def event_for(entry,args,started,exit_code,exception_kind=None,client=None):
    if entry not in ENTRIES:raise ValueError('未登记的诊断入口')
    record_path=arg(args,'--record');output_path=arg(args,'--output');record=safe_read(record_path) if record_path else {};result=safe_read(output_path) if output_path else {}
    if not result and output_path:result=safe_read(str(output_path)+'.audit.json')
    month=record.get('month') or arg(args,'--month')
    if not isinstance(month,str) or not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',month):month=None
    errors=result.get('errors',[]);pending=result.get('pending',result.get('pending_totals',result.get('issues',[])))
    if not isinstance(errors,list):errors=[]
    if not isinstance(pending,list):pending=[]
    counts={}
    for item in errors+pending:code=category(item);counts[code]=counts.get(code,0)+1
    files=[];sources=record.get('sources',{})
    if isinstance(sources,dict) and record_path:
        for source in list(sources.values())[:50]:
            if not isinstance(source,dict) or not source.get('path'):continue
            try:
                item=fingerprint(Path(record_path).resolve().parent/source['path'],source.get('role','other'))
                if item:files.append(item)
            except OSError:pass
    for flag in ('--workbook','--attendance','--roster','--history','--overtime','--plan','--input','--pension','--unemployment','--medical','--housing-fund','--source-periods'):
        value=arg(args,flag)
        if value:
            try:
                item=fingerprint(value)
                if item:files.append(item)
            except OSError:pass
    output=[]
    if output_path:
        try:
            item=fingerprint(output_path,'output')
            if item:output.append(item)
        except OSError:pass
    version=str(record.get('policy_version') or VERSION)
    if not re.fullmatch(r'[0-9.]{1,40}',version):version='unknown'
    people=record.get('employees',[])
    count=len(people) if isinstance(people,list) and record else result.get('employee_count',result.get('rows'))
    if isinstance(count,bool) or not isinstance(count,int) or not 0<=count<=100000:count=None
    technical=result.get('technical_delivery',{})
    if not isinstance(technical,dict):technical={}
    result_state='exception' if exception_kind else ('failed' if errors else ('passed' if result.get('passed') is True else ('pending' if exit_code==2 else ('failed' if exit_code else 'completed_unverified'))))
    identifier=str(uuid.uuid4());device=(client or {}).get('device_id','unregistered')
    return {'schema':SCHEMA,'event_id':identifier,'device_id':device,'run_id':digest((device+str(record_path or output_path or entry)).encode())[:24],
      'component':entry,'policy_version':version,'collector_version':VERSION,'month':month,
      'started_at':datetime.datetime.fromtimestamp(started,datetime.timezone.utc).isoformat(),'duration_ms':max(0,int((time.time()-started)*1000)),
      'exit_code':exit_code,'state':result_state,'error_count':len(errors) if result else None,'pending_count':len(pending) if result else None,'category_counts':counts,
      'exception_kind':exception_kind if exception_kind in {'ValueError','OSError','RuntimeError','FileNotFoundError','KeyboardInterrupt','TypeError','JSONDecodeError','ImportError'} else ('OTHER_EXCEPTION' if exception_kind else None),
      'employee_count':count,'source_versions':files,'source_file_count':len(sources) if isinstance(sources,dict) else 0,
      'source_versions_limit':50,'output_versions':output,
      'required_stages':[x for x in technical.get('required_template_stages',[]) if x in STAGES],
      'verified_stages':[x for x in technical.get('verified_template_stages',[]) if x in STAGES],
      'formal_ready':result.get('formal_ready') is True,'data_scope':'redacted_diagnostics_only'}

def capture_case(entry,args,home,org=None):
    try:
        from case_store import record_entry
        return record_entry(entry,args,home,org)
    except Exception as exc:
        try:(Path(home)/'case-health.json').write_bytes(canonical({'at':datetime.datetime.now(datetime.timezone.utc).isoformat(),'state':'capture_pending','category':type(exc).__name__}))
        except OSError:pass
        return {'state':'capture_pending'}

def health(home,state):
    # 无错误原文、路径、令牌、员工姓名。
    p=Path(home)/'health.json';temporary=p.with_name('health-'+uuid.uuid4().hex+'.tmp')
    try:
        temporary.write_text(json.dumps({'at':datetime.datetime.now(datetime.timezone.utc).isoformat(),'state':state},ensure_ascii=False),encoding='utf-8')
        for attempt in range(3):
            try:os.replace(temporary,p);break
            except PermissionError:
                if attempt==2:return
                time.sleep(.02)
    finally:temporary.unlink(missing_ok=True)
def emit(event,home=None,org=None):
    home=private_dir(home_path(home));org=org or organization()
    config=read_json_secret(home/'client.secure')
    if config['key_id']!=org['key_id']:raise ValueError('诊断加密公钥已变更，请由管理员核对')
    event=dict(event,device_id=config['device_id'])
    envelope=sign_envelope(encrypt(event,org['public_key_pem']),config['device_id'],config['ssh_private_pem'].encode())
    with database(home) as db:db.execute('INSERT INTO events(event_id,envelope,created,state) VALUES(?,?,?,?)',(event['event_id'],canonical(envelope).decode(),time.time(),'queued'))
    health(home,'queued');return envelope

def observed(function,entry,argv=None,home=None,org=None,background=True):
    started=time.time();args=list(sys.argv[1:] if argv is None else argv);code=0;kind=None
    try:
        returned=function()
        if isinstance(returned,int):code=returned
        return returned
    except SystemExit as exc:
        code=exc.code if isinstance(exc.code,int) else (0 if exc.code is None else 2);raise
    except BaseException as exc:code=2;kind=type(exc).__name__;raise
    finally:
        target=home_path(home)
        if (target/'client.secure').is_file():
            try:
                config=read_json_secret(target/'client.secure');event=event_for(entry,args,started,code,kind,config);case_result=capture_case(entry,args,target,org) or {};event['case_archive_state']=case_result.get('state','capture_pending');event['case_conversation_state']=case_result.get('conversation_state','unknown');emit(event,target,org)
                if background:
                    subprocess.Popen([sys.executable,str(Path(__file__).resolve()),'flush','--home',str(target)],stdin=subprocess.DEVNULL,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,creationflags=0x08000000 if os.name=='nt' else 0,start_new_session=os.name!='nt')
            except Exception:
                try:health(target,'record_failed')
                except OSError:pass

def main():
    for stream in (sys.stdout,sys.stderr):
        if hasattr(stream,'reconfigure'):stream.reconfigure(encoding='utf-8')
    ap=argparse.ArgumentParser(description='公司薪资核算诊断组件');ap.add_argument('command',choices=['flush','status','smoke','record']);ap.add_argument('--home',type=Path);ap.add_argument('--entry',choices=sorted(ENTRIES));ap.add_argument('--started',type=float);ap.add_argument('--exit-code',type=int,default=0);ap.add_argument('--exception-kind');ap.add_argument('--argument-list',default='[]');a=ap.parse_args();home=home_path(a.home)
    if a.command=='record':
        if not (home/'client.secure').is_file():return
        config=read_json_secret(home/'client.secure');args=json.loads(a.argument_list)
        if not isinstance(args,list) or len(args)>40 or any(not isinstance(x,str) for x in args):raise ValueError('诊断参数格式错误')
        event=event_for(a.entry,args,a.started or time.time(),a.exit_code,a.exception_kind,config)
        case_result=capture_case(a.entry,args,home) or {};event['case_archive_state']=case_result.get('state','capture_pending');event['case_conversation_state']=case_result.get('conversation_state','unknown');emit(event,home)
        subprocess.Popen([sys.executable,str(Path(__file__).resolve()),'flush','--home',str(home)],stdin=subprocess.DEVNULL,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,creationflags=0x08000000 if os.name=='nt' else 0,start_new_session=os.name!='nt')
        return
    if a.command=='status':
        if not (home/'client.secure').is_file():print('诊断尚未安装');return
        with database(home) as db:counts=dict(db.execute('SELECT state,COUNT(*) FROM events GROUP BY state'))
        print(json.dumps({'待同步':counts.get('queued',0),'已同步':counts.get('synced',0),'最近状态':safe_read(home/'health.json').get('state')},ensure_ascii=False));return
    if a.command=='smoke':
        config=read_json_secret(home/'client.secure');event=event_for('diagnostics_smoke',[],time.time(),0,client=config);event['month']='2099-09';event['synthetic_only']=True;emit(event,home);print('已生成明确合成诊断事件');return
    from case_store import refresh,sync as sync_cases
    refresh(home)
    failures=False;result={'state':'pending','sent':0};case_result={'state':'pending','sent':0}
    try:result=sync_client(home,organization())
    except Exception:failures=True
    try:case_result=sync_cases(home)
    except Exception:failures=True
    health(home,'sync_pending' if failures else result['state'])
    print(json.dumps({'診斷同步狀態':result['state'],'診斷本次筆數':result['sent'],'案件同步狀態':case_result['state'],'案件本次筆數':case_result['sent']},ensure_ascii=False))
    if failures:
        print('未完成部分仍保留本機密文隊列，不能當成全部同步成功。');raise SystemExit(2)
if __name__=='__main__':main()
