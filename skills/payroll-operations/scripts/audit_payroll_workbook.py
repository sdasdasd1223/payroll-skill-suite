"""只读核实际工资公式、缓存及未决合计，不代替业务确认。
计划包含sheet/header_row/field_columns/first_row/last_row/total_row/pending_rows、
formula_fields/complete_total_fields/calculated_cells。
标准I/Q/R/V工资模板另给leave_formula_parameters逐行绑定已审阅请假参数。
"""
import argparse,json,math,importlib.util
from pathlib import Path

def numeric(v):
    return isinstance(v,(int,float)) and not isinstance(v,bool) and math.isfinite(v)

def blank(v):
    return v is None or v==''

def audit(workbook,plan,base_dir=None):
    from openpyxl import load_workbook
    wf=load_workbook(workbook,data_only=False,read_only=True)
    wv=load_workbook(workbook,data_only=True,read_only=True)
    try:
        f=wf[plan['sheet']];v=wv[plan['sheet']]
        f.reset_dimensions();v.reset_dimensions()
        # Materialize once: repeated read-only coordinate access can reread XML.
        fr=list(f.values);vr=list(v.values)
        def get(rows,r,c):return rows[r-1][c-1] if r<=len(rows) and c<=len(rows[r-1]) else None
        headers={};errors=[];checked=[];pending=[];leave_checks=[];technical_pending=[];attendance_targets=[]
        for c,label in enumerate(fr[plan['header_row']-1],1):
            if blank(label):continue
            if label in headers:errors.append(f'重复字段:{label}')
            headers[label]=c
        # Grouped headings need a reviewed mapping, e.g. company/social vs personal/social.
        from openpyxl.utils.cell import column_index_from_string,coordinate_to_tuple
        for field,location in plan.get('field_columns',{}).items():
            c=column_index_from_string(location['column'])
            if not location.get('header_cells'):raise ValueError('多级表头映射缺定位依据')
            for coord,expected in location['header_cells'].items():
                r,hc=coordinate_to_tuple(coord)
                if r>=plan['first_row'] or get(fr,r,hc)!=expected:
                    errors.append(f'字段定位依据不符:{field}:{coord}')
            headers[field]=c
        first,last=plan['first_row'],plan['last_row']
        if not isinstance(first,int) or not isinstance(last,int) or not 1<=first<=last:
            raise ValueError('人员行范围无效')
        pending_rows=set(plan.get('pending_rows',[]))
        if not pending_rows.issubset(set(range(first,last+1))):raise ValueError('待核人员行超出范围')
        from openpyxl.utils import get_column_letter
        def require_formula(r,c,reason):
            cell=f'{get_column_letter(c)}{r}';value=get(fr,r,c)
            ok=isinstance(value,str) and value.startswith('=')
            checked.append({'cell':cell,'reason':reason,'formula_preserved':ok})
            if not ok:errors.append(f'计算列未保留公式:{cell}:{reason}')
        # 指定工资表实际存在的核心计算列不能靠漏填计划跳过。
        mandatory_formula_fields=('请假扣款','考勤扣款','入离职差额','应发工资','财务核算薪资','实发金额','原应税金额','招行')
        formula_fields=list(dict.fromkeys(list(plan['formula_fields'])+[x for x in mandatory_formula_fields if x in headers]))
        for field in formula_fields:
            if field not in headers:errors.append(f'缺计算字段:{field}');continue
            for r in range(first,last+1):
                if field=='考勤扣款':
                    identity=plan.get('attendance_identity_column') or (get_column_letter(headers['姓名']) if '姓名' in headers else None)
                    if identity:attendance_targets.append({'sheet':f.title,'cell':f'{get_column_letter(headers[field])}{r}','identity_cell':f'{identity}{r}'})
                    else:technical_pending.append(f'考勤扣款缺实际人员身份列:{r}')
                # 待核金额与是否保留既定计算公式是两件事，不能整行豁免。
                # 规则确实尚未明确时，逐格说明并保留待核；不能用常数代替。
                loc=f'{get_column_letter(headers[field])}{r}'
                exemption=plan.get('pending_formula_exemptions',{}).get(loc)
                if exemption:
                    if r not in pending_rows or not exemption.get('reason') or not exemption.get('source_refs'):
                        errors.append(f'公式待核豁免缺实际依据:{loc}')
                    if not blank(get(fr,r,headers[field])):
                        errors.append(f'公式待核格不得改为常数或未核定公式:{loc}')
                    pending.append({'field':field,'cell':loc,'reason':exemption.get('reason'),'source_refs':exemption.get('source_refs')})
                else:
                    require_formula(r,headers[field],field)
                    if field=='请假扣款' and all(k in headers for k in ('事假时间','病假时间')):
                        import re
                        from openpyxl.formula.tokenizer import Tokenizer
                        expression=get(fr,r,headers[field])
                        refs=set()
                        if isinstance(expression,str) and expression.startswith('='):
                            for token in Tokenizer(expression).items:
                                if token.type!='OPERAND' or token.subtype!='RANGE':continue
                                location=token.value
                                if '!' in location:
                                    origin,location=location.rsplit('!',1)
                                    if origin.strip("'").replace("''", "'")!=f.title:continue
                                refs.update(re.findall(r'(?<![A-Za-z0-9_])\$?([A-Z]{1,3})\$?([1-9][0-9]*)(?![A-Za-z0-9_])',location.upper()))
                        for input_field in ('事假时间','病假时间'):
                            if (get_column_letter(headers[input_field]),str(r)) not in refs:
                                errors.append(f'请假扣款缺本行可编辑输入引用:{loc}:{input_field}')
                        parameters=plan.get('leave_formula_parameters',{}).get(str(r))
                        canonical=(headers[field],headers['事假时间'],headers['病假时间'])==(19,17,18)
                        if parameters is None:
                            pending.append({'field':field,'cell':loc,'reason':'缺本行已核请假参数，未证明病假门槛及人工补录安全'})
                            if canonical:errors.append(f'请假扣款缺核定参数或逐格空白待核豁免，不能沿用固定半薪式:{loc}')
                        else:
                            module_spec=importlib.util.spec_from_file_location('_leave_formula_audit',Path(__file__).with_name('payroll_formulas.py'))
                            formulas=importlib.util.module_from_spec(module_spec);module_spec.loader.exec_module(formulas)
                            base=Path(base_dir) if base_dir is not None else Path(workbook).resolve().parent
                            generated=formulas.leave_deduction_formula(r,parameters,base)
                            normalize=lambda text:re.sub(r'\s+','',str(text)).replace('$','').upper()
                            if not canonical or normalize(expression)!=normalize(generated['formula']):errors.append(f'请假扣款与本行已核参数生成式不符:{loc}')
                            actual=formulas.evaluate_leave_deduction(parameters,get(vr,r,9),get(vr,r,17),get(vr,r,18),get(vr,r,22),base)
                            leave_checks.append({'cell':loc,'formula_matches':canonical and normalize(expression)==normalize(generated['formula']),'calculation':actual,'parameters':generated['parameters'],'source_checks':generated['source_checks']})
                            cached=get(vr,r,headers[field])
                            if actual['status']=='pending':
                                pending.append({'field':field,'cell':loc,'reason':'请假组成或年度档位未核，辅助独立项目可保留','components':actual})
                                if not blank(cached):errors.append(f'未核请假扣款被显示为数值:{loc}')
                            elif blank(cached):technical_pending.append(f'请假公式尚缺实际重算缓存:{loc}')
                            elif not numeric(cached) or abs(cached-actual['amount'])>.005:errors.append(f'请假扣款缓存与已核参数计算不符:{loc}')
        for cell in plan.get('calculated_cells',[]):
            from openpyxl.utils.cell import coordinate_to_tuple
            r,c=coordinate_to_tuple(cell)
            if not first<=r<=last:raise ValueError('计算明细坐标超出人员范围')
            require_formula(r,c,'本次自动计算明细')
        total_fields=list(dict.fromkeys(list(plan['complete_total_fields'])+[x for x in ('应发工资','财务核算薪资','实发金额','招行') if x in headers]))
        for field in total_fields:
            if field not in headers:errors.append(f'缺合计字段:{field}');continue
            c=headers[field];values=[get(vr,r,c) for r in range(first,last+1)]
            actual=get(vr,plan['total_row'],c);loc=f'{get_column_letter(c)}{plan["total_row"]}'
            if any(not numeric(x) for x in values):
                pending.append({'field':field,'cell':loc,'missing_people':sum(not numeric(x) for x in values),'actual':actual})
                if not blank(actual):errors.append(f'待核金额被显示为全员合计:{loc}:{field}:{actual}')
            elif not numeric(actual) or abs(actual-sum(values))>.005:
                errors.append(f'已核金额合计不符:{loc}:{field}')
        attendance_audit=None
        if attendance_targets:
            spec=importlib.util.spec_from_file_location('_attendance_link_audit',Path(__file__).with_name('attendance_link.py'))
            linker=importlib.util.module_from_spec(spec);spec.loader.exec_module(linker)
            attendance_audit=linker.audit_attendance_links(workbook,plan.get('source_view_inventory'),attendance_targets,base_dir)
            errors.extend(attendance_audit['errors']);technical_pending.extend(attendance_audit['technical_pending'])
            pending.extend({'field':'考勤扣款','reason':text} for text in attendance_audit['data_pending'])
        error_values={'#REF!','#DIV/0!','#VALUE!','#NAME?','#NUM!','#N/A','#SPILL!'}
        for r,row in enumerate(vr,1):
            for c,x in enumerate(row,1):
                if isinstance(x,str) and x in error_values:errors.append(f'Excel错误:{get_column_letter(c)}{r}:{x}')
        return {'passed':not errors,'technical_passed':not errors and not technical_pending,'status':'needs_correction' if errors else ('technical_incomplete' if technical_pending else 'draft_checked'),'technical_pending':technical_pending,'errors':errors,'formula_checks':checked,'leave_formula_checks':leave_checks,'attendance_link_audit':attendance_audit,'pending_totals':pending,'pending_rows':sorted(pending_rows),'business_confirmation':False,'note':'passed只表示本组件未发现错误；技术完成另看technical_passed，交付仍须实际模板gate。'}
    finally:
        wf.close();wv.close()

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--workbook',type=Path,required=True);ap.add_argument('--plan',type=Path,required=True);ap.add_argument('--output',type=Path,required=True);a=ap.parse_args()
    if a.output.exists() or a.output.resolve() in (a.workbook.resolve(),a.plan.resolve()):
        raise SystemExit('检查输出不得覆盖原表、计划或既有结果')
    result=audit(a.workbook,json.loads(a.plan.read_text(encoding='utf-8-sig')),a.plan.parent)
    a.output.parent.mkdir(parents=True,exist_ok=True);a.output.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'passed':result['passed'],'technical_passed':result['technical_passed'],'technical_pending':result['technical_pending'],'errors':len(result['errors']),'formula_checks':len(result['formula_checks']),'pending_totals':len(result['pending_totals'])}))
    raise SystemExit(0 if result['technical_passed'] else 2)
if __name__=='__main__':
    from diagnostics_bridge import observed
    observed(main,'audit_payroll_workbook')
