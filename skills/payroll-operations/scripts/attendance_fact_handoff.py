"""从实际原件生成假别事实及最终交接绑定；不推定工资待遇，不扫描自由文字。"""
import argparse,copy,csv,hashlib,importlib.util,json,re
from datetime import date,datetime
from decimal import Decimal,InvalidOperation
from pathlib import Path

ROLES=('hr_handoff','run_explanation')
def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def number(value):
    if value is None or isinstance(value,bool) or str(value).strip()=='':return None
    try:
        value=Decimal(str(value));return value if value.is_finite() else None
    except InvalidOperation:return None
def text_number(value):
    numeric=number(value)
    return format(numeric.normalize(),'f') if numeric is not None else str(value)
def context_for(record,base):
    spec=importlib.util.spec_from_file_location('_handoff_crossstage',Path(__file__).with_name('crossstage_audit.py'));module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
    context=module.Audit(record,base);context.load_sources();context.load_tables();context.scope();return context
def day(value):
    if isinstance(value,(date,datetime)):return value.strftime('%Y-%m-%d')
    for fmt in ('%Y-%m-%d','%Y/%m/%d','%Y.%m.%d'):
        try:return datetime.strptime(str(value)[:10],fmt).strftime('%Y-%m-%d')
        except ValueError:pass
    raise ValueError('原始日期无法按已定义格式读取')

def collect_facts(record,base_dir,context=None):
    from openpyxl import load_workbook
    from openpyxl.utils.cell import column_index_from_string,get_column_letter
    ctx=context or context_for(record,base_dir);spec=record.get('raw_attendance',{});original_id=spec.get('original_table') or spec.get('table');table=ctx.tables.get(original_id)
    result={'facts':[],'errors':[],'pending':[]}
    if not table or table['spec']['format'] not in ('xlsx','xls') or not {'name','date'}<=set(table['spec']['columns']):result['pending'].append('假别事实缺实际xlsx/xls原件及已有姓名/日期映射');return result
    source=ctx.sources[table['spec']['source_id']];path=source['resolved'];ts=table['spec'];header=ts.get('header_row',spec.get('header_row') if original_id==spec.get('table') else None) or ts['start_row']-1
    if header<1:result['pending'].append('假别事实缺实际表头行');return result
    if sha(path)!=source['sha256']:raise ValueError('原假别源SHA失效')
    if ts['format']=='xlsx':
        book=load_workbook(path,read_only=True,data_only=True)
        try:ws=book[ts['sheet']];ws.reset_dimensions();rows=[list(row) for row in ws.values]
        finally:book.close()
    else:
        import xlrd
        book=xlrd.open_workbook(path);ws=book.sheet_by_name(ts['sheet']);rows=[[xlrd.xldate_as_datetime(ws.cell_value(r,c),book.datemode) if ws.cell_type(r,c)==xlrd.XL_CELL_DATE else ws.cell_value(r,c) for c in range(ws.ncols)] for r in range(ws.nrows)]
    names={}
    for key,person in ctx.people.items():names.setdefault(person['name'],[]).append(key)
    name_col=column_index_from_string(ts['columns']['name'])-1;date_col=column_index_from_string(ts['columns']['date'])-1
    candidates=[]
    for col,value in enumerate(rows[header-1]):
        label=str(value or '').strip();base=re.split(r'[（(]',label)[0].strip();rule=spec.get('leave_columns',{}).get(label)
        if not base.endswith('假') and not rule:continue
        match=re.search(r'[（(]\s*(分钟|小时|天)\s*[）)]',label)
        unit=match.group(1) if match else {'minutes':'分钟','hours':'小时','days':'天'}.get((rule or {}).get('unit'))
        kind=(rule or {}).get('kind') or {'事假':'personal','病假':'sick','年假':'annual','育儿假':'childcare','产假':'maternity'}.get(base,'other')
        candidates.append((col,label,unit,kind))
    for index,raw in enumerate(rows,1):
        if index<ts['start_row']:continue
        name=str(raw[name_col] or '').strip() if name_col<len(raw) else ''
        if name not in names:continue
        active=[(col,label,unit,kind,raw[col]) for col,label,unit,kind in candidates if col<len(raw) and raw[col] not in (None,'') and number(raw[col])!=0]
        if not active:continue
        if len(names[name])!=1:result['pending'].append('原假别姓名不唯一，不能猜同人:'+name);continue
        try:stamp=day(raw[date_col])
        except (ValueError,IndexError) as exc:result['pending'].append(name+':'+str(exc));continue
        if stamp[:7]!=record.get('month'):continue
        key=names[name][0]
        for col,label,unit,kind,value in active:
            if not unit:result['pending'].append(name+':原假别单位未能从实际表头/已有映射证实:'+label)
            item={'personal':'personal_leave','sick':'sick_leave','annual':'annual_leave','childcare':'childcare_leave','maternity':'maternity_leave','other':'other_leave'}[kind]
            treatment=next((f for f in record.get('facts',[]) if f.get('employee_key')==key and f.get('stage')=='payroll' and f.get('item')==item),{})
            payroll_requested=record.get('through_stage') in ('payroll','tax','payslip')
            treatment_status='not_requested' if not payroll_requested else ('reviewed' if treatment.get('status') in ('calculated','confirmed','confirmed_zero','not_applicable') and treatment.get('policy_refs') else 'pending')
            fact={'employee_key':key,'name':name,'date':stamp,'original_label':label,'raw_value':text_number(value),'raw_unit':unit or '单位待核','business_classification':kind,'treatment_status':treatment_status,'source':{'path':str(path.resolve()),'sha256':source['sha256'],'sheet':ts['sheet'],'row':index,'header_cell':get_column_letter(col+1)+str(header),'value_cell':get_column_letter(col+1)+str(index)}}
            canonical=json.dumps(fact,ensure_ascii=False,sort_keys=True,separators=(',',':'));fact['fact_id']=hashlib.sha256(canonical.encode()).hexdigest()[:20];result['facts'].append(fact)
    if sha(path)!=source['sha256']:raise ValueError('读取原假别期间源SHA变化')
    return result

def sentence(fact):
    treatment={'pending':'工资待遇待核','reviewed':'工资处理见已有核定记录','not_requested':'工资待遇不在本次阶段'}[fact['treatment_status']]
    return '〔事实:'+fact['fact_id']+'〕'+fact['name']+'在'+fact['date']+'记录原假别“'+fact['original_label']+'”，原值'+fact['raw_value']+fact['raw_unit']+'；'+treatment+'。'
def block_markers(block_id):
    if not re.fullmatch(r'[A-Za-z0-9_-]+',str(block_id)):raise ValueError('事实段标识只能用字母数字下划线/短横线')
    return '<!-- attendance-facts:'+block_id+':BEGIN -->','<!-- attendance-facts:'+block_id+':END -->'
def expected_text(facts,ids):
    mapping={f['fact_id']:f for f in facts}
    if len(ids)!=len(set(ids)) or not set(ids)<=set(mapping):raise ValueError('最终事实绑定重复或引用不存在的原事实')
    return '\n'.join(sentence(mapping[key]) for key in ids)

def verify_handoff(record,base_dir,context=None):
    from openpyxl import load_workbook
    collected=collect_facts(record,base_dir,context);facts=collected['facts'];errors=list(collected['errors']);pending=list(collected['pending']);data_pending=[];checks=[]
    if not facts:return {'passed':not errors and not pending,'errors':errors,'pending':pending,'data_pending':[],'facts':facts,'checks':checks}
    spec=record.get('attendance_handoff',{});targets=spec.get('targets',[]);final_ids=spec.get('final_source_ids',[])
    required=set(f['fact_id'] for f in facts);coverage={role:set() for role in ROLES};covered_sources=set()
    if not final_ids:pending.append('缺最终人资清单/本次说明的实际文件声明，旁边事实日志不能替代最终交接')
    for sid,source in record.get('sources',{}).items():
        if source.get('role') in ROLES and sid not in final_ids:errors.append('最终交接用途来源未纳入绑定:'+sid)
    for target in targets:
        sid=target.get('source_id');source=record.get('sources',{}).get(sid,{})
        try:
            if sid not in final_ids or source.get('role') not in ROLES:raise ValueError('绑定不是已声明的最终人资交接/说明实物')
            path=Path(base_dir)/source['path']
            if sha(path)!=source.get('sha256'):raise ValueError('最终交接文件SHA已变化')
            wanted=expected_text(facts,target.get('fact_ids',[]))
            if target.get('format')=='xlsx':
                book=load_workbook(path,read_only=True,data_only=False)
                try:actual=book[target['sheet']][target['cell']].value
                finally:book.close()
            elif target.get('format')=='text':
                start,end=block_markers(target['block_id']);text=path.read_text(encoding='utf-8-sig')
                if text.count(start)!=1 or text.count(end)!=1:raise ValueError('最终说明缺唯一程序事实段')
                actual=text.split(start,1)[1].split(end,1)[0].strip('\r\n')
            else:raise ValueError('最终交接绑定只支持明确xlsx格或UTF8说明事实段')
            if actual!=wanted:raise ValueError('最终绑定的日期/假别/原值/单位或处理状态与实际源确定性短句不符')
            coverage[source['role']].update(target['fact_ids']);covered_sources.add(sid);checks.append({'source_id':sid,'actual_sha256':source['sha256'],'fact_ids':target['fact_ids'],'passed':True})
        except (ValueError,KeyError,OSError,TypeError) as exc:errors.append(str(sid)+':'+str(exc))
    for role,seen in coverage.items():
        if seen!=required:pending.append(role+':最终交接没有完整覆盖本人原假别事实')
    if set(final_ids)!=covered_sources:pending.append('部分最终交接实物未实际读回验证')
    for fact in facts:
        if fact['treatment_status']=='pending':data_pending.append(fact['name']+':原假别事实已保存，工资待遇仍待核:'+fact['fact_id'])
    return {'passed':not errors and not pending,'errors':errors,'pending':pending,'data_pending':data_pending,'facts':facts,'checks':checks,'note':'只核已声明最终交接的确定性事实绑定；自由说明不做泛自然语言扫描，不修改源事实或待遇'}

def render_handoff(record,base_dir,output_dir):
    from openpyxl import load_workbook
    output=Path(output_dir).resolve()
    if output.exists():raise ValueError('输出目录必须尚不存在，不得覆盖任何既有交接/源')
    collected=collect_facts(record,base_dir)
    if collected['errors'] or collected['pending']:raise ValueError('原假别尚未可确定性渲染:'+str(collected))
    output.mkdir(parents=True);result=copy.deepcopy(record);facts=collected['facts'];ids=[f['fact_id'] for f in facts]
    for source in result.get('sources',{}).values():source['path']=str((Path(base_dir)/source['path']).resolve())
    spec=result.setdefault('attendance_handoff',{});targets=spec.get('targets',[])
    existing=[sid for sid,s in result['sources'].items() if s.get('role') in ROLES]
    if not targets and existing:raise ValueError('已有最终交接文件须先在record定位真实目标格/事实段，不能另放正确日志替代')
    if not targets:
        targets=[];spec['final_source_ids']=[]
        for role in ROLES:
            sid='generated_'+role;path=output/(('人资假别事实交接单' if role=='hr_handoff' else '本次核算说明')+'.md');start,end=block_markers('leave')
            path.write_text('# '+('人资假别事实交接单' if role=='hr_handoff' else '本次核算说明')+'\n\n'+start+'\n{{attendance_facts:leave}}\n'+end+'\n',encoding='utf8')
            result['sources'][sid]={'path':str(path),'month':record['month'],'role':role};spec['final_source_ids'].append(sid);targets.append({'source_id':sid,'format':'text','block_id':'leave','fact_ids':ids})
        spec['targets']=targets
    changed={}
    for target in targets:
        sid=target['source_id'];source=result['sources'][sid]
        if sid not in changed:
            original=Path(source['path']);path=output/('final-'+str(len(changed)+1)+original.suffix)
            if target['format']=='xlsx':changed[sid]=(path,load_workbook(original))
            else:changed[sid]=(path,original.read_text(encoding='utf-8-sig'))
        path,content=changed[sid];wanted=expected_text(facts,target.get('fact_ids',ids));target['fact_ids']=target.get('fact_ids',ids)
        if target['format']=='xlsx':content[target['sheet']][target['cell']]=wanted
        else:
            start,end=block_markers(target['block_id']);placeholder='{{attendance_facts:'+target['block_id']+'}}'
            if content.count(start)!=1 or content.count(end)!=1:raise ValueError('最终说明必须先使用唯一事实段/占位符，不静默追加到旧错误摘要旁边')
            head,rest=content.split(start,1);_,tail=rest.split(end,1);content=head+start+'\n'+wanted+'\n'+end+tail
            changed[sid]=(path,content)
    for sid,(path,content) in changed.items():
        if isinstance(content,str):path.write_text(content,encoding='utf8')
        else:content.save(path);content.close()
        result['sources'][sid]['path']=str(path);result['sources'][sid]['sha256']=sha(path)
    (output/'source-facts.json').write_text(json.dumps(collected,ensure_ascii=False,indent=2),encoding='utf8')
    with (output/'人资假别事实单.csv').open('w',encoding='utf-8-sig',newline='') as stream:
        fields={'fact_id':'事实编号','name':'姓名','date':'日期','original_label':'原始假别表头','raw_value':'原始数值','raw_unit':'原始单位','business_classification':'业务分类','treatment_status':'待遇处理状态','source_sha256':'来源文件指纹','sheet':'来源工作表','header_cell':'来源表头格','value_cell':'来源数值格','short_sentence':'原始事实及待核说明'}
        writer=csv.DictWriter(stream,fieldnames=list(fields.values()));writer.writeheader()
        units={'minutes':'分钟','hours':'小时','days':'天'}
        states={'pending':'待核','calculated':'已计算','confirmed':'已确认','not_applicable':'不适用','confirmed_zero':'已确认零'}
        for fact in facts:
            values={**{k:fact.get(k,'') for k in fields},'source_sha256':fact['source']['sha256'],'sheet':fact['source']['sheet'],'header_cell':fact['source']['header_cell'],'value_cell':fact['source']['value_cell'],'short_sentence':sentence(fact)}
            values['raw_unit']=units.get(values['raw_unit'],values['raw_unit']);values['treatment_status']=states.get(values['treatment_status'],values['treatment_status'])
            values['business_classification']={'personal':'事假','sick':'病假','annual':'年假','childcare':'育儿假','maternity':'产假','other':'其他（原标签保留）'}.get(values['business_classification'],values['business_classification'])
            writer.writerow({label:values[key] for key,label in fields.items()})
    report=verify_handoff(result,base_dir);(output/'handoff-verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8');(output/'run-with-handoff.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
    return {'record':str(output/'run-with-handoff.json'),'passed':report['passed'],'data_pending':report['data_pending'],'output_dir':str(output)}

def main():
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('action',choices=('render','verify'));parser.add_argument('--record',type=Path,required=True);parser.add_argument('--output',type=Path,required=True);args=parser.parse_args();record=json.loads(args.record.read_text(encoding='utf-8-sig'))
    if args.action=='render':result=render_handoff(record,args.record.parent,args.output)
    else:
        if args.output.exists() or args.output.resolve()==args.record.resolve() or args.output.resolve() in {(args.record.parent/str(s['path'])).resolve() for s in record.get('sources',{}).values()}:raise ValueError('不得覆盖源/记录/既有报告')
        result=verify_handoff(record,args.record.parent);args.output.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps(result,ensure_ascii=False));raise SystemExit(0 if result['passed'] else 2)
if __name__=='__main__':main()
