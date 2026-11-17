"""Persisted payroll handoff checks. Audit JSON is evidence, not authorization.
CLI: python workflow_audit.py --record run.json --output audit.json
Exit 2 means unresolved or invalid, 0 means this specified stage passes.
crossstage-1另输出逐人原批注业务处置CSV；复制对象不代表事假清算已解决。
"""
import argparse, hashlib, json
from pathlib import Path

REQUIRED={
 'attendance':('scope','identity','calendar','attendance','history','annual_leave','childcare','overtime','reward'),
 'payroll':('scope','identity','salary','progression','attendance','insurance','performance','reward','manual_adjustments','leave','maternity','tax'),
 'tax':('scope','identity','payroll_confirmed','tax_mapping'),
 'payslip':('scope','payroll_confirmed','accounts','phone','net_amount'),
}
DONE={'calculated','confirmed_zero','not_applicable','confirmed'}
def sha256(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def validate_run(record, base_dir=Path('.')):
    if record.get('schema_version')=='crossstage-1':
        from crossstage_audit import audit_crossstage
        return audit_crossstage(record,base_dir)
    errors=[];stage=record.get('stage');required=REQUIRED.get(stage)
    if required is None:return {'passed':False,'errors':['未知阶段'],'stage':stage}
    if not record.get('month') or not record.get('policy_version'):errors.append('缺核算月或规则版本')
    people=record.get('employees',[]);ids=[p.get('employee_key') for p in people]
    expected=record.get('expected_employee_keys')
    if not people or any(not x for x in ids) or len(ids)!=len(set(ids)):errors.append('人员为空、缺主键或主键重复')
    if not isinstance(expected,list) or set(expected)!=set(ids) or len(expected)!=len(ids):errors.append('人员范围未覆盖预期核定名单')
    sources=record.get('sources',{});valid_sources=set()
    for key,source in sources.items():
        p=base_dir/str(source.get('path',''))
        if not p.is_file() or not source.get('sha256') or sha256(p)!=source['sha256']:
            errors.append(f'来源文件不存在或内容已变化:{key}');continue
        if not source.get('month') or not source.get('role'):errors.append(f'来源缺归属期或用途:{key}');continue
        if source.get('role') in ('attendance_confirmed','payroll_confirmed') and source.get('month')!=record.get('month'):
            errors.append(f'确认来源属于其他月份:{key}');continue
        valid_sources.add(key)
    for person in people:
        for item in required:
            decision=person.get('items',{}).get(item,{})
            label=f"{person.get('employee_key')}:{item}"
            if decision.get('status') not in DONE:errors.append(label+':未完成');continue
            if decision.get('status')=='not_applicable' and not decision.get('reason'):errors.append(label+':不适用缺依据')
            if not decision.get('source_ids') or not set(decision['source_ids']).issubset(valid_sources):errors.append(label+':缺有效来源')
        if stage=='payroll':
            # Source facts are carried forward separately from a manually selected status.
            observed=person.get('observations',{}).get('maternity',{})
            count=observed.get('record_count')
            if isinstance(count,bool) or not isinstance(count,int) or count<0 or observed.get('coverage_complete') is not True:
                errors.append(str(person.get('employee_key'))+':产假原始记录覆盖未核实')
            elif not observed.get('source_ids') or not set(observed['source_ids']).issubset(valid_sources):
                errors.append(str(person.get('employee_key'))+':产假观察缺有效来源')
            elif count>0 and person.get('items',{}).get('maternity',{}).get('status')=='not_applicable':
                errors.append(str(person.get('employee_key'))+':原件有产假但交接误记不适用')
    for item in record.get('confirmations',[]):
        source=sources.get(item.get('source_id'),{})
        if not item.get('confirmed_by') or not item.get('confirmed_at') or item.get('sha256')!=source.get('sha256'):errors.append('确认记录未绑定人、日期和文件内容')
    if stage in ('payroll','tax','payslip'):
        role='attendance_confirmed' if stage=='payroll' else 'payroll_confirmed'
        bound={c.get('source_id') for c in record.get('confirmations',[]) if c.get('confirmed_by') and c.get('confirmed_at')}
        if not any(k in bound and v.get('role')==role for k,v in sources.items()):errors.append('缺已绑定文件的人工确认:'+role)
    # 旧版自报状态不能证明来源与产出完整；保留诊断，正式输出须采用实际文件契约。
    pending=['旧记录未核实际全范围来源及跨阶段输出；请迁移crossstage-1']
    if record.get('mode')!='draft':errors.extend(pending)
    purpose=record.get('purpose','template_delivery')
    return {'passed':not errors and purpose=='calculation_details','calculation_passed':not errors,'formal_ready':False,'purpose':purpose,'delivery_status':'calculation_details_only' if purpose=='calculation_details' else 'technical_incomplete','delivery_message':'旧记录仅可诊断，未核验合规模板；迁移crossstage-1后按实际交付阶段核验','technical_delivery':{'ready':False,'required_template_stages':[],'verified_template_stages':[],'all_five_stages_verified':False,'pending':pending},'stage':stage,'month':record.get('month'),'employee_count':len(people),'errors':errors,'pending':pending}

def payslip_eligibility(person, pay_date):
    reasons=[]
    if person.get('account_query_complete') is not True:reasons.append('账号查询未完整完成')
    if person.get('account_count')!=1:reasons.append('账号不存在或不唯一')
    if person.get('is_activated') is not True:reasons.append('账号未激活或状态不明')
    if person.get('is_cross_tenant') is not False:reasons.append('跨租户或状态不明')
    if person.get('has_left') is not False:reasons.append('已离开组织或状态未核')
    if not person.get('phone'):reasons.append('缺核实手机号')
    if person.get('net') is None:reasons.append('缺实发金额')
    if person.get('departure_date') and person['departure_date']<=pay_date:reasons.append('离职日期不晚于发薪日')
    return {'included':not reasons,'reasons':reasons}

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--record',type=Path,required=True);ap.add_argument('--output',type=Path,required=True);a=ap.parse_args()
    record=json.loads(a.record.read_text(encoding='utf-8-sig'))
    protected=[a.record.resolve()]+[(a.record.parent/str(s.get('path',''))).resolve() for s in record.get('sources',{}).values()]
    if a.output.resolve() in protected:
        print('检查输出不得覆盖记录或任何引用的来源文件');raise SystemExit(2)
    if a.output.exists():print('检查输出已存在，请另存新版本');raise SystemExit(2)
    r=validate_run(record,a.record.parent)
    if record.get('schema_version')=='crossstage-1':
        from crossstage_audit import write_scope_review_csv,write_comment_review_csv
        scope_csv=a.output.with_suffix('.scope.csv')
        if scope_csv.resolve() in protected:raise SystemExit('范围复核CSV不得覆盖来源')
        scope_csv.parent.mkdir(parents=True,exist_ok=True);write_scope_review_csv(r,scope_csv);r['scope_review_csv']=str(scope_csv)
        comment_csv=a.output.with_suffix('.comments.csv')
        if comment_csv.resolve() in protected:raise SystemExit('批注复核CSV不得覆盖来源')
        write_comment_review_csv(r,comment_csv);r['comment_review_csv']=str(comment_csv)
    a.output.parent.mkdir(parents=True,exist_ok=True);a.output.write_text(json.dumps(r,ensure_ascii=False,indent=2,default=str),encoding='utf-8');print(json.dumps(r,ensure_ascii=False,default=str))
    raise SystemExit(0 if r['passed'] else 2)
if __name__=='__main__':
    from diagnostics_bridge import observed
    observed(main,'workflow_audit')
