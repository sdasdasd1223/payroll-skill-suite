"""生成既有I/Q/R/V工资模板的请假扣款式；不写工资源表、不猜缺失参数。
CLI: python payroll_formulas.py --plan 已审阅参数.json --output 新公式计划.json
"""
import argparse
from decimal import Decimal, InvalidOperation
import hashlib
import importlib.util
import json
from pathlib import Path

def number(value):
    if value is None or isinstance(value, bool) or str(value).strip() == '': return None
    try:
        result = Decimal(str(value)); return result if result.is_finite() else None
    except InvalidOperation: return None

def sha256(path): return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def core():
    spec = importlib.util.spec_from_file_location('_payroll_formula_core', Path(__file__).with_name('payroll_core.py'))
    value = importlib.util.module_from_spec(spec); spec.loader.exec_module(value); return value

def reviewed_parameters(parameters, base_dir=Path('.')):
    """核已审阅参数的实际来源指纹；不把缺源或未核数字补成0。"""
    evidence = parameters.get('source_evidence', {}); checks = {}
    for key in ('salary','is_intern','personal_input','sick_input','required_hours','annual_before','minimum_wage'):
        source = evidence.get(key, {}); path = Path(base_dir) / str(source.get('path', ''))
        valid = bool(source.get('locator') and source.get('sha256') and path.is_file() and sha256(path) == source['sha256'])
        checks[key] = {'verified':valid, 'source':source}
    prior, minimum = number(parameters.get('annual_before')), number(parameters.get('minimum_wage'))
    values = {'is_intern':parameters.get('is_intern') if isinstance(parameters.get('is_intern'), bool) and checks['is_intern']['verified'] else None,
        'annual_before':prior if prior is not None and prior >= 0 and checks['annual_before']['verified'] else None,
        'minimum_wage':minimum if minimum is not None and minimum >= 0 and checks['minimum_wage']['verified'] else None,
        'personal_input_complete':parameters.get('personal_input_complete') is True and checks['personal_input']['verified'],
        'sick_input_complete':parameters.get('sick_input_complete') is True and checks['sick_input']['verified'],
        'salary_verified':checks['salary']['verified'], 'required_hours_verified':checks['required_hours']['verified']}
    return values, checks

def leave_deduction_formula(row, parameters, base_dir=Path('.')):
    if isinstance(row, bool) or not isinstance(row, int) or not 1 <= row <= 1048576: raise ValueError('工资行号无效')
    values, checks = reviewed_parameters(parameters, base_dir)
    I,Q,R,V = (f'{column}{row}' for column in ('I','Q','R','V'))
    wage_guard = f'NOT(ISNUMBER({I}))' if values['salary_verified'] else f'OR(NOT(ISNUMBER({I})),TRUE)'
    hour_guard = f'OR(NOT(ISNUMBER({V})),{V}<=0)' if values['required_hours_verified'] else f'OR(NOT(ISNUMBER({V})),{V}<=0,TRUE)'
    blank_q = '0' if values['personal_input_complete'] else '""'
    blank_r = '0' if values['sick_input_complete'] else '""'
    if values['is_intern'] is None: personal_positive = '""'
    else:
        denominator = V if values['is_intern'] else '174'
        guard = f'OR({wage_guard},{hour_guard})' if values['is_intern'] else wage_guard
        personal_positive = f'IF({guard},"",ROUND(-{I}/{denominator}*{Q},2))'
    personal = f'IF({Q}="",{blank_q},IF(OR(NOT(ISNUMBER({Q})),{Q}<0),"",IF({Q}=0,0,{personal_positive})))'
    prior, minimum = values['annual_before'], values['minimum_wage']
    if prior is None: policy = '""'
    elif prior < 72:
        policy = f'IF({format(prior,"f")}+{R}>72,"",ROUND(-{I}/{V}*{R}/2,2))'
    else:
        policy = '""' if minimum is None else f'ROUND(-({I}-{format(minimum,"f")}*0.8)/{V}*{R},2)'
    sick_positive = f'IF(OR({wage_guard},{hour_guard}),"",{policy})'
    sick = f'IF({R}="",{blank_r},IF(OR(NOT(ISNUMBER({R})),{R}<0),"",IF({R}=0,0,{sick_positive})))'
    formula = f'=IF(AND(ISNUMBER({personal}),ISNUMBER({sick})),ROUND({personal}+{sick},2),"")'
    return {'row':row,'cell':f'S{row}','formula':formula,'component_formulas':{'personal':'='+personal,'sick':'='+sick},
        'parameters':{k:str(v) if isinstance(v,Decimal) else v for k,v in values.items()},'source_checks':checks,
        'note':'普通在职Q可按制度留白；原始月输入未知的Q/R空白仍待核。跨72小时留空，未新增分段制度。'}

def evaluate_leave_deduction(parameters, salary, personal_hours, sick_hours, required_hours, base_dir=Path('.')):
    values, _ = reviewed_parameters(parameters, base_dir); arithmetic = core()
    def resolve(value, complete):
        if value in (None,''): return Decimal(0) if complete else None
        n=number(value); return n if n is not None and n>=0 else None
    q=resolve(personal_hours,values['personal_input_complete']); r=resolve(sick_hours,values['sick_input_complete'])
    personal={'status':'pending','amount':None,'reason':'事假输入、本人类型或薪资来源未核'}
    sick={'status':'pending','amount':None,'reason':'病假输入或年度/薪资/计薪来源未核'}
    if q==0: personal={'status':'calculated','amount':0}
    elif q is not None and values['is_intern'] is not None and values['salary_verified'] and (not values['is_intern'] or values['required_hours_verified']):
        try: personal={'status':'calculated','amount':arithmetic.personal_leave_deduction(salary,q,values['is_intern'],required_hours)}
        except (ValueError,TypeError): pass
    if r==0: sick={'status':'calculated','amount':0}
    elif r is not None and values['annual_before'] is not None and values['salary_verified'] and values['required_hours_verified']:
        sick=arithmetic.sick_leave_deduction(salary,r,required_hours,values['annual_before'],values['minimum_wage'])
    pending=personal['status']=='pending' or sick['status']=='pending'
    return {'status':'pending' if pending else 'calculated','amount':None if pending else arithmetic.money(Decimal(str(personal['amount']))+Decimal(str(sick['amount']))),'personal':personal,'sick':sick}

def main():
    ap=argparse.ArgumentParser(); ap.add_argument('--plan',required=True,type=Path); ap.add_argument('--output',required=True,type=Path); args=ap.parse_args()
    plan=json.loads(args.plan.read_text(encoding='utf-8-sig'))
    for row in plan.get('rows',[]):
        for source in row.get('parameters',{}).get('source_evidence',{}).values():
            if source.get('path'): source['path']=str((args.plan.parent/source['path']).resolve())
    protected={args.plan.resolve()} | {(args.plan.parent/str(s.get('path',''))).resolve() for row in plan.get('rows',[]) for s in row.get('parameters',{}).get('source_evidence',{}).values()}
    if args.output.exists() or args.output.resolve() in protected: raise SystemExit('新公式计划不得覆盖原件或既有结果')
    rows=[leave_deduction_formula(row['row'],row['parameters'],args.plan.parent) for row in plan['rows']]
    args.output.parent.mkdir(parents=True,exist_ok=True); args.output.write_text(json.dumps({'sheet':plan['sheet'],'rows':rows,'leave_formula_parameters':{str(row['row']):row['parameters'] for row in plan['rows']},'business_confirmation':False},ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'rows':len(rows),'output':str(args.output),'business_confirmation':False},ensure_ascii=False))

if __name__=='__main__': main()
