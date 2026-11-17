"""工资考勤扣款只引用同人真实金额快照；只读核源SHA、身份、金额与实际公式。"""
import argparse
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re

def sha256(path): return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def blank(value): return value is None or value == ''
def numeric(value): return isinstance(value,(int,float)) and not isinstance(value,bool) and math.isfinite(value)
def equal(a,b): return abs(a-b)<.000001 if numeric(a) and numeric(b) else a==b or blank(a) and blank(b)

def formula_for_value(view_sheet, value_cell):
    spec=importlib.util.spec_from_file_location('_attendance_source_views',Path(__file__).with_name('template_source_views.py'))
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
    return module.source_view_formula(view_sheet,value_cell)

def same_formula(actual, expected):
    from openpyxl.formula.tokenizer import Tokenizer
    def tokens(expression):
        if not isinstance(expression,str) or not expression.startswith('='):return None
        result=[]
        for token in Tokenizer(expression).items:
            if token.type=='WHITE-SPACE':continue
            value=token.value
            if token.type=='OPERAND' and token.subtype=='RANGE' and '!' in value:
                sheet,cell=value.rsplit('!',1)
                if sheet.startswith("'") and sheet.endswith("'"):sheet=sheet[1:-1].replace("''", "'")
                value=(sheet.upper(),cell.replace('$','').upper())
            elif token.type!='OPERAND' or token.subtype!='TEXT':value=value.upper()
            result.append((token.type,token.subtype,value))
        return result
    try:return tokens(actual) is not None and tokens(actual)==tokens(expected)
    except Exception:return False

def link_for_employee(mapping, employee_identity):
    rows=[row for row in mapping['rows'] if row['identity']==employee_identity]
    if len(rows)!=1:raise ValueError('同人来源不存在或不唯一，不能用他人/常量补扣款')
    row=rows[0];cell=mapping['columns']['deduction']+str(row['view_row'])
    return {'identity':employee_identity,'value_cell':cell,'formula':formula_for_value(mapping['sheet'],cell),'source':mapping['source'],'source_row':row['source_row'],'source_cells':row.get('source_cells',{})}

def audit_attendance_links(workbook, inventory, required_targets, base_dir=None):
    from openpyxl import load_workbook
    from openpyxl.utils.cell import column_index_from_string
    base=Path(base_dir) if base_dir is not None else Path(workbook).resolve().parent
    errors=[];technical_pending=[];data_pending=[];checked=[];source_books=[]
    wf=load_workbook(workbook,read_only=False,data_only=False);wv=load_workbook(workbook,read_only=False,data_only=True)
    sources={}
    try:
        for item in (inventory or {}).get('sources',[]):
            source=item.get('source',{});path=base/str(source.get('path',''))
            if not path.is_file() or not source.get('sha256') or sha256(path)!=source['sha256']:
                errors.append('考勤原件SHA失效，必须重新导入及重核，不静默刷新:'+str(source.get('path')));continue
            try:
                book=load_workbook(path,read_only=True,data_only=True);source_books.append(book);ws=book[item['source_sheet']];ws.reset_dimensions();rows=list(ws.values)
                fields=item['source_columns'];name_col=column_index_from_string(fields['name'])-1;amount_col=column_index_from_string(fields['deduction'])-1
                names=[]
                for row in item['rows']:
                    raw=rows[row['source_row']-1];name=raw[name_col] if name_col<len(raw) else None;amount=raw[amount_col] if amount_col<len(raw) else None
                    if name!=row['identity'] or name!=row['values']['name'] or not equal(amount,row['values']['deduction']):errors.append('登记快照与实际同人原件不符:'+str(row['source_row']))
                    if not blank(amount) and not numeric(amount):errors.append('考勤源金额未核，不能将文本转换0')
                    names.append(name)
                if len(names)!=len(set(names)):errors.append('考勤源身份重复，不能仅按姓名取金额')
                if sha256(path)!=source['sha256']:errors.append('读取期间考勤源SHA变化，必须重新导入重核')
                sources[item['id']]=item
            except (KeyError,ValueError,IndexError,OSError) as exc:technical_pending.append('考勤源定位未完成:'+str(exc))
        links=(inventory or {}).get('links',[])
        for target in required_targets:
            sheet,cell=target['sheet'],target['cell'];label=sheet+'!'+cell
            own=[link for link in links if link.get('target_sheet')==sheet and link.get('target_cell')==cell]
            if len(own)!=1:technical_pending.append('缺唯一同人业务金额来源绑定:'+label);continue
            link=own[0];source=sources.get(link.get('view_id'))
            if source is None:technical_pending.append('来源未实际核实:'+label);continue
            rows=[row for row in source['rows'] if row['source_row']==link['source_row'] and row['view_row']==link['view_row']]
            if len(rows)!=1:errors.append('金额映射错行:'+label);continue
            row=rows[0]
            try:
                view=wf[link['view_sheet']];snapshot=view[link['value_cell']]
                actual_identity=wv[sheet][target['identity_cell']].value
                snapshot_identity=view[source['columns']['name']+str(row['view_row'])].value
                if actual_identity!=row['identity'] or snapshot_identity!=row['identity'] or link['identity']!=row['identity']:errors.append('工资/快照/原件不是同人:'+label)
                expected_cell=source['columns']['deduction']+str(row['view_row'])
                if link['view_sheet']!=source['sheet'] or link['value_cell']!=expected_cell:errors.append('引用未定位实际快照金额格:'+label)
                if snapshot.data_type=='f' or not(blank(snapshot.value) or numeric(snapshot.value)) or view.protection.sheet:errors.append('快照金额不是可编辑实际数字/空白:'+label)
                expected=formula_for_value(source['sheet'],expected_cell);actual=wf[sheet][cell].value
                if not same_formula(actual,expected):errors.append('考勤扣款不是同人真实业务金额式，身份/常量数组不算联动:'+label)
                if not equal(snapshot.value,row['values']['deduction']):data_pending.append('快照保留了人工金额差异，需按既有确认核定，不能静默重导覆盖:'+label)
                cached=wv[sheet][cell].value
                if blank(snapshot.value):
                    data_pending.append('本人考勤金额待核，空白不当0:'+label)
                    if not blank(cached):errors.append('空考勤金额被显示为数值:'+label)
                elif blank(cached):technical_pending.append('考勤公式尚缺实际重算缓存:'+label)
                elif not equal(cached,snapshot.value):errors.append('考勤扣款缓存未随实际金额更新:'+label)
                checked.append({'target':label,'identity':actual_identity,'source':source['source'],'source_sheet':source['source_sheet'],'source_row':row['source_row'],'source_cells':row.get('source_cells'),'view_sheet':source['sheet'],'value_cell':expected_cell,'snapshot_amount':snapshot.value,'target_cached':cached,'formula':actual})
            except (KeyError,ValueError,TypeError) as exc:technical_pending.append('业务金额引用实物未完成:'+label+':'+str(exc))
        return {'passed':not errors,'technical_passed':not errors and not technical_pending,'errors':errors,'technical_pending':technical_pending,'data_pending':data_pending,'checks':checked,'note':'只读实际内页与源指纹；不覆盖原确认金额，不承诺外部文件静默联动'}
    finally:
        wf.close();wv.close()
        for book in source_books:book.close()

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--workbook',type=Path,required=True);ap.add_argument('--plan',type=Path,required=True);ap.add_argument('--output',type=Path,required=True);args=ap.parse_args()
    plan=json.loads(args.plan.read_text(encoding='utf-8-sig'))
    protected={args.workbook.resolve(),args.plan.resolve()}|{(args.plan.parent/str(item.get('source',{}).get('path',''))).resolve() for item in (plan.get('source_view_inventory') or {}).get('sources',[])}
    if args.output.exists() or args.output.resolve() in protected:raise SystemExit('不得覆盖原件或既有报告')
    result=audit_attendance_links(args.workbook,plan.get('source_view_inventory'),plan['required_targets'],args.plan.parent)
    args.output.parent.mkdir(parents=True,exist_ok=True);args.output.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(result,ensure_ascii=False));raise SystemExit(0 if result['technical_passed'] else 2)

if __name__=='__main__':main()
