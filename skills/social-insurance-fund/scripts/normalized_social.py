"""Company-specific social amounts, including Guangsi, from auditable normalized rows.
Never applies another company's default rates. Missing amounts propagate as None.
CLI input JSON rows: company,name,certificate,month,insurance,employee,employer,source,source_row.
"""
import argparse,json
from decimal import Decimal,InvalidOperation,ROUND_HALF_UP
from pathlib import Path

def read_guangsi_xlsx(path):
    """Read explicit project/side headers. Returns normalized rows plus retained extras.
    Missing dates are never inferred from filename or export time.
    """
    from openpyxl import load_workbook
    from datetime import date, datetime, timedelta
    b=load_workbook(path,read_only=False,data_only=True);ws=b.active
    header=None
    for n in range(1,min(ws.max_row,20)+1):
        vals=[str(c.value or '').strip() for c in ws[n]]
        if '姓名' in vals and '证件号码' in vals and '费款所属期起' in vals:header=n;break
    if header is None:raise ValueError('广思缺已识别的身份及费用所属期表头')
    cols={str(c.value).strip():c.column for c in ws[header] if c.value}
    projects={};last=''
    for col in range(1,ws.max_column+1):
        group=str(ws.cell(header-1,col).value or '').strip()
        if group:last=group
        if str(ws.cell(header,col).value or '').strip()=='应缴费额':projects[col]=last
    def month(v):
        if isinstance(v,(date,datetime)):return v.strftime('%Y-%m')
        if isinstance(v,(int,float)):return (datetime(1899,12,30)+timedelta(days=v)).strftime('%Y-%m')
        text=str(v or '').strip().replace('/','-')
        import re
        m=re.match(r'(\d{4})-?(\d{2})',text)
        return f'{m[1]}-{m[2]}' if m else None
    rows=[];extras=[]
    for n in range(header+1,ws.max_row+1):
        name=ws.cell(n,cols['姓名']).value;cert=ws.cell(n,cols['证件号码']).value
        if not name or str(name).strip() in ('合计','总计') or not cert:continue
        period=month(ws.cell(n,cols['费款所属期起']).value)
        end=month(ws.cell(n,cols.get('费款所属期止',cols['费款所属期起'])).value)
        if period!=end:raise ValueError('广思来源行跨费用所属月，需确认如何拆分')
        collected={}
        for col,project in projects.items():
            value=ws.cell(n,col).value
            kind=next((kind for text,kind in [('基本养老保险','pension'),('基本医疗保险','medical'),('失业保险','unemployment'),('生育保险','maternity'),('工伤保险','injury')] if text in project),None)
            if kind is None:
                extras.append(dict(name=name,project=project,amount=value,source=str(path),source_row=n));continue
            side='employee' if '个人' in project else 'employer'
            rec=collected.setdefault(kind,dict(company='广思',name=str(name).strip(),certificate=str(cert),month=period,insurance=kind,employee=0 if kind in ('maternity','injury') else None,employer=None,source=str(path),source_row=n))
            rec[side]=value
        rows.extend(collected.values())
    b.close()
    if not rows:raise ValueError('广思来源无有效员工明细，不能确认为零')
    return {'rows':rows,'excluded_projects':extras}

def amount(value):
    if value is None or str(value).strip()=='':return None
    try:
        d=Decimal(str(value).replace(',',''))
        return float(d.quantize(Decimal('.01'),rounding=ROUND_HALF_UP)) if d.is_finite() else None
    except InvalidOperation:return None
def build(rows,month):
    people={};issues=[];detail=[];seen=set();names={}
    if not rows:return {'month':month,'people':[],'details':[],'issues':[{'reason':'原始明细为空，不证明本月无需缴纳'}],'passed':False}
    for i,row in enumerate(rows):
        detail.append(dict(row))
        key=(row.get('company'),row.get('certificate') or row.get('name'))
        namekey=(row.get('company'),row.get('name'))
        if not all(row.get(x) for x in ('company','name','source','source_row','month')):
            issues.append({'row':i,'reason':'缺身份、归属期或来源行号'});continue
        if row['month']!=month:issues.append({'row':i,'reason':'归属期与核算月不符'});continue
        insurance=row.get('insurance')
        if insurance not in ('pension','unemployment','medical','maternity','injury','housing_fund'):
            issues.append({'row':i,'reason':'额外项目未获确认纳入'});continue
        p=people.setdefault(key,{'company':row['company'],'name':row['name'],'certificate':row.get('certificate'),'items':{}})
        if namekey in names and names[namekey]!=key:issues.append({'row':i,'reason':'同名不同证件，保留独立身份，工资匹配仍须用证件'})
        names[namekey]=key
        dup=(key,insurance)
        if dup in seen:
            p['items'][insurance]={'employee':None,'employer':None};issues.append({'row':i,'reason':'同一身份同险种多笔，补退费或重复须确认'});continue
        seen.add(dup)
        emp=amount(row.get('employee'));co=amount(row.get('employer'))
        if insurance in ('maternity','injury'):
            if emp not in (None,0):issues.append({'row':i,'reason':'单位险种出现个人金额，待核'});emp=None
            else:emp=0
        if emp is None or co is None:issues.append({'row':i,'reason':'金额缺失或不可解析'})
        p['items'][insurance]={'employee':emp,'employer':co}
    for p in people.values():
        for kind in ('pension','unemployment','medical','housing_fund'):
            if kind not in p['items']:
                p['items'][kind]={'employee':None,'employer':None};issues.append({'name':p['name'],'reason':'缺少'+kind})
    return {'month':month,'people':list(people.values()),'details':detail,'issues':issues,'passed':not issues}
def main():
    ap=argparse.ArgumentParser();ap.add_argument('--input',type=Path,required=True);ap.add_argument('--month',required=True);ap.add_argument('--output',type=Path,required=True);a=ap.parse_args()
    result=build(json.loads(a.input.read_text(encoding='utf-8-sig')),a.month)
    if a.input.resolve()==a.output.resolve():raise SystemExit('不得覆盖源文件')
    a.output.parent.mkdir(parents=True,exist_ok=True);a.output.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8');raise SystemExit(0 if result['passed'] else 2)
if __name__=='__main__':
    from diagnostics_bridge import observed
    observed(main,'insurance_merge')
