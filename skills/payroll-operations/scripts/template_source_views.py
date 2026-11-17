"""固定SHA业务原件的简单值快照和同人引用；不重算源、不复制历史模板制度。"""
from pathlib import Path
import re
from zipfile import ZipFile
from lxml import etree as ET
from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter


def source_view_formula(sheet, cell):
    """生成器与审计共用；合法标识页名不用可被应用删掉的多余引号。"""
    if not isinstance(sheet,str) or not sheet or len(sheet)>31 or re.search(r'[\\/*?:\[\]]',sheet):
        raise ValueError('来源页名不合法')
    if not isinstance(cell,str) or not re.fullmatch(r'\$?[A-Z]{1,3}\$?[1-9][0-9]*',cell):
        raise ValueError('来源值坐标须为实际单格引用')
    plain=bool(re.fullmatch(r'[^\W\d][\w.]*',sheet))
    name=sheet if plain else "'"+sheet.replace("'","''")+"'"
    reference=name+'!'+cell
    return f'=IF({reference}="","",{reference})'


def inventory_sources(configs, base):
    from template_guard import evidence, sha256, scalar
    result=[];ids=set();names=set()
    for config in configs:
        key=config.get('id');name=config.get('sheet')
        if not key or key in ids or not name or name in names or len(name)>31 or re.search(r'[\\/*?:\[\]]',name):
            raise ValueError('来源值页id/页名缺失、重复或不合法')
        ids.add(key);names.add(name)
        source=evidence(config.get('source'),base);path=Path(source['path'])
        if path.suffix.lower()!='.xlsx':raise ValueError('来源值页目前只支持简单xlsx区域')
        formulas=load_workbook(path,data_only=False);cached=load_workbook(path,data_only=True)
        sheet=config['source_sheet'];ws=formulas[sheet];values=cached[sheet]
        if ws._charts or ws._images or ws.tables or ws.merged_cells.ranges:
            raise ValueError('来源值页不支持图表/图片/表格对象/合并区域；须选择简单实际业务区域')
        fields=config.get('fields',{})
        if list(fields)[:2]!=['name','deduction'] or any(not re.fullmatch(r'[A-Z]{1,3}',str(v)) for v in fields.values()):
            raise ValueError('来源值页fields须依次含name、deduction的实际列字母')
        first,last=config['first_row'],config['last_row']
        if isinstance(first,bool) or isinstance(last,bool) or not 1<=first<=last<=ws.max_row:
            raise ValueError('来源值页实际行范围无效')
        # 公式源必须有真实缓存，空字符串缓存允许；不把缺失缓存伪装为业务空白。
        formula_cache={}
        with ZipFile(path) as archive:
            main='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
            rel='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
            wb=ET.fromstring(archive.read('xl/workbook.xml'))
            rels=ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
            targets={x.get('Id'):x.get('Target') for x in rels}
            node=next(x for x in wb.find(main+'sheets') if x.get('name')==sheet)
            target=targets[node.get(rel)];part=target.lstrip('/') if target.startswith('/') else 'xl/'+target
            for cell in ET.fromstring(archive.read(part)).iter(main+'c'):
                if cell.find(main+'f') is not None:
                    v=cell.find(main+'v')
                    formula_cache[cell.get('r')]=v is not None and (v.text is not None or cell.get('t')=='str')
        rows=[];identities=set()
        for row in range(first,last+1):
            actual={};refs={}
            for field,column in fields.items():
                c=ws[f'{column}{row}'];v=values[c.coordinate].value
                if c.data_type=='f' and not formula_cache.get(c.coordinate):
                    raise ValueError(f'来源公式缺真实缓存：{sheet}!{c.coordinate}')
                if values[c.coordinate].data_type=='e':raise ValueError('来源值页有公式错误缓存')
                actual[field]=scalar(v);refs[field]=c.coordinate
            identity=actual['name']
            if identity in (None,''):
                if any(v not in (None,'') for k,v in actual.items() if k!='name'):
                    raise ValueError(f'来源行有金额却无人员身份：{sheet}!{row}')
                continue
            if not isinstance(identity,str) or identity!=identity.strip() or identity in identities:
                raise ValueError('来源真实姓名须非空、无首尾空格且唯一；同名须先人工核定主键')
            if actual['deduction'] not in (None,'') and (isinstance(actual['deduction'],bool) or not isinstance(actual['deduction'],(int,float))):
                raise ValueError('已算扣款只接受实读数值或空白，不把待核文字转换成0')
            if any(isinstance(v,dict) or isinstance(v,str) and v.startswith('=') for v in actual.values()):
                raise ValueError('简单来源快照不支持日期对象或可执行公式文本')
            identities.add(identity)
            rows.append({'source_row':row,'view_row':len(rows)+2,'identity':identity,'values':actual,'source_cells':refs})
        if sha256(path)!=source['sha256']:raise ValueError('读取期间业务源SHA发生变化')
        result.append({'id':key,'source':source,'source_sheet':sheet,'sheet':name,
            'first_row':first,'last_row':last,'source_columns':dict(fields),
            'columns':{f:get_column_letter(i) for i,f in enumerate(fields,1)},'rows':rows})
    return result


def apply_source_views(book, policy, plan, mappings, base, registered):
    from template_guard import evidence
    configs=policy.get('sheet_policy',{}).get('source_views',[])
    sources=inventory_sources(configs,base)
    if sources!=registered:raise ValueError('业务来源值页实读内容与预先登记盘点不符')
    links=[]
    for config,source in zip(configs,sources):
        name=source['sheet']
        if name in book.sheetnames:raise ValueError('来源值页不得替换指定主模板或既有附页')
        ws=book.create_sheet(name)
        ws.append(list(source['columns']))
        for row in source['rows']:ws.append([row['values'][field] for field in source['columns']])
        link=config.get('link')
        if not link:continue
        target=book[link['target_sheet']];identity_column=link['identity_column'];column=link['target_column']
        personnel=policy.get('sheets',{}).get(target.title,{}).get('personnel')
        if not personnel or identity_column not in personnel['identity_columns']:
            raise ValueError('来源值页自动引用须登记目标人员区和实际身份列')
        lookup={x['identity']:x for x in source['rows']}
        for n,mapping in enumerate(mappings[target.title],1):
            if not personnel['start_row']<=mapping['source_row']<=personnel['end_row']:continue
            identity=target[f'{identity_column}{n}'].value
            if not identity:continue
            if identity not in lookup:raise ValueError(f'来源值页无该实际员工，拒绝错源/错人：{identity}')
            cell=target[f'{column}{n}']
            if any(x['sheet']==target.title and x['cell']==cell.coordinate for x in plan.get('writes',[])):
                raise ValueError('自动来源引用格不得同时手工写入或姓名查常量')
            row=lookup[identity];previous_manual=None;authorization=None
            if cell.value not in (None,'') and cell.data_type!='f':
                previous_manual=cell.value
                if isinstance(cell.value,bool) or cell.value!=row['values']['deduction']:
                    if not link.get('reason'):raise ValueError('来源金额与目标非空人工值不同，须停止核定；例外缺实际授权理由')
                    authorization=evidence(link.get('authorization'),base)
            value_cell=source['columns']['deduction']+str(row['view_row'])
            formula=source_view_formula(name,value_cell)
            cell.value=formula
            links.append({'view_id':source['id'],'target_sheet':target.title,'target_cell':cell.coordinate,
                'identity':identity,'view_sheet':name,'view_row':row['view_row'],'value_cell':value_cell,
                'source_row':row['source_row'],'formula':formula,'previous_manual_value':previous_manual,
                'authorization':authorization,'authorization_reason':link.get('reason') if authorization else None})
    return {'sources':sources,'links':links}
