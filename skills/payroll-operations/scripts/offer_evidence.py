"""只读提取录用函正文报到日期，核对预核记录；文件名不作为日期依据。
约定报到日不等于实际入职日，不在此覆盖花名册或裁定两者冲突。
"""
import argparse,hashlib,json,re,unicodedata
from datetime import date
from pathlib import Path

LABEL=re.compile(r'入职报到日期|入职日期|报到日期|报到时间')
DATE=re.compile(r'(?P<y>\d{4})\s*(?:年|[-/])\s*(?P<m>\d{1,2})\s*(?:月|[-/])\s*(?P<d>\d{1,2})\s*日?')
def normalize(text):return re.sub(r'\s+','',unicodedata.normalize('NFKC',text or ''))
def parse_pages(pages,employee=None):
    candidates=[];problems=[]
    identity=bool(employee and any(normalize(employee) in normalize(text) for text in pages))
    for page,text in enumerate(pages,1):
        compact=normalize(text)
        for label in LABEL.finditer(compact):
            tail=compact[label.end():label.end()+80]
            # 日期必须紧接字段，不能跳过其他条款去抓签订/打印日期。
            found=DATE.match(re.sub(r'^[：:，,\-—]+','',tail))
            if not found:continue
            try:value=date(int(found['y']),int(found['m']),int(found['d'])).isoformat()
            except ValueError:
                problems.append({'page':page,'label':label.group(),'reason':'正文日期不合法','quote':compact[label.start():label.end()+len(found.group())+2]});continue
            raw_quote=None
            lines=text.splitlines()
            for n,line in enumerate(lines):
                if label.group() in normalize(line):
                    excerpt=[]
                    for following in lines[n:n+4]:
                        excerpt.append(following)
                        if DATE.search(unicodedata.normalize('NFKC','\n'.join(excerpt))):break
                    raw_quote='\n'.join(excerpt);break
            candidates.append({'page':page,'label':label.group(),'reporting_date':value,'quote':compact[label.start():label.end()+len(found.group())+2],'raw_quote':raw_quote})
    dates=sorted({x['reporting_date'] for x in candidates})
    if not identity:problems.append({'reason':'正文未能核实本人员姓名；不能仅用文件名确认归属'})
    if not candidates:problems.append({'reason':'未取得可核验的正文报到日期；扫描件或特殊格式待人工提取，不回退文件名'})
    elif len(dates)!=1:problems.append({'reason':'正文存在多个不同报到日期，待人工核定'})
    return {'status':'pending' if problems else 'extracted','employee':employee,'reporting_date':dates[0] if len(dates)==1 and not problems else None,'candidates':candidates,'pending':problems,'note':'仅核约定报到日正文事实，不证明实际入职日、薪资标准或人工确认'}
def inspect(pdf,employee,expected_sha256,reported_date=None):
    from pypdf import PdfReader
    pdf=Path(pdf).resolve();digest=hashlib.sha256(pdf.read_bytes()).hexdigest()
    if digest!=expected_sha256:raise ValueError('录用函SHA与事先固定原件不一致')
    pages=[p.extract_text() or '' for p in PdfReader(pdf).pages]
    result=parse_pages(pages,employee);result.update(schema='payroll-offer-evidence/1',source={'path':str(pdf),'sha256':digest},reported_date=reported_date,errors=[])
    if reported_date is not None and result['reporting_date'] is not None and reported_date!=result['reporting_date']:
        result['errors'].append('预核记录日期与正文日期不一致；不得采用文件名日期或虚构冲突')
    result['passed']=not result['errors'] and result['status']=='extracted'
    return result
def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--pdf',type=Path,required=True);p.add_argument('--employee',required=True);p.add_argument('--expected-sha256',required=True);p.add_argument('--reported-date');p.add_argument('--output',type=Path,required=True)
    a=p.parse_args()
    if a.output.exists() or a.output.resolve()==a.pdf.resolve():raise SystemExit('证据必须另存，不能覆盖录用函或既有结果')
    try:r=inspect(a.pdf,a.employee,a.expected_sha256,a.reported_date)
    except Exception as exc:r={'passed':False,'errors':[str(exc)],'status':'pending'}
    a.output.parent.mkdir(parents=True,exist_ok=True);a.output.write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8');print(json.dumps(r,ensure_ascii=False));raise SystemExit(0 if r['passed'] else 2)
if __name__=='__main__':main()
