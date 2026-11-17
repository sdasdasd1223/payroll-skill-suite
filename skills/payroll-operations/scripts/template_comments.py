"""工资原件批注全量盘点与同人对象复制；由template_guard登记并重放。

source_comments在登记policy固定，comment_copies只覆盖明确坐标，不改变来源。
未匹配到目标人员只记录not_in_target，不证明该人应被排除；人员去向另由流程审计。
"""
from copy import copy
from pathlib import Path
import argparse
import hashlib
import json
import re

from openpyxl import load_workbook
from openpyxl.comments import Comment
from openpyxl.cell.cell import MergedCell
from openpyxl.utils.cell import column_index_from_string, get_column_letter, coordinate_to_tuple


def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read_source(config, base_dir):
    source=config.get('source',{})
    if not source.get('path') or not source.get('sha256') or not source.get('locator'):
        raise ValueError('批注源须有source.path/sha256/locator，不能用自报已保留代替')
    path=Path(source['path'])
    if not path.is_absolute():path=Path(base_dir)/path
    path=path.resolve()
    if path.suffix.lower() not in ('.xlsx','.xlsm'):
        raise ValueError('工资批注源仅支持xlsx/xlsm；不能把无法读取的xls批注当成0条')
    if sha256(path)!=source['sha256']:
        raise ValueError(f'工资批注源SHA不符：{path}')
    book=load_workbook(path,data_only=False)
    items=[]
    for ws in book:
        for row in ws:
            for cell in row:
                if cell.comment is not None:
                    items.append({'sheet':ws.title,'cell':cell.coordinate,'row':cell.row,'column':cell.column,
                        'text':cell.comment.text,'author':cell.comment.author})
    if sha256(path)!=source['sha256']:
        raise ValueError('读取工资批注期间源文件发生变化')
    return book,{'source_id':config['id'],'source':{'path':str(path),'sha256':source['sha256'],'locator':source['locator']},
        'comment_count':len(items),'sheet_titles':{ws.title:str(ws['A1'].value or '') for ws in book},'items':items}


def inventory_sources(configs, base_dir):
    """只读，返回真实工资原件全部工作表的批注；原SHA不符即拒绝。"""
    if not isinstance(configs,list):raise ValueError('policy.source_comments必须为数组')
    result=[];ids=set()
    for config in configs:
        if not config.get('id') or config['id'] in ids:raise ValueError('批注源id为空或重复')
        ids.add(config['id']);book,inventory=read_source(config,base_dir)
        regions=config.get('regions',[]);ignored=config.get('ignored_sheets',{})
        if inventory['comment_count'] and not regions:
            raise ValueError('有批注的工资原件必须登记人员区，不能全部笼统忽略')
        for sheet,reason in ignored.items():
            if sheet not in book.sheetnames or not str(reason).strip():
                raise ValueError('忽略历史附页须为真实工作表并说明原因')
        for region in regions:
            if region.get('sheet') not in book.sheetnames or region['sheet'] in ignored:
                raise ValueError('批注人员区工作表缺失或同时被忽略')
            first,last=region.get('first_row'),region.get('last_row')
            if not isinstance(first,int) or not isinstance(last,int) or not 1<=first<=last<=book[region['sheet']].max_row:
                raise ValueError('批注源人员区边界无效')
            if not region.get('identity_columns') or not region.get('target_identity_columns') or not region.get('target_sheet'):
                raise ValueError('批注源/目标须登记实际身份列和目标页')
        result.append(inventory)
    return result


def column_number(value):
    if isinstance(value,int) and not isinstance(value,bool) and 1<=value<=16384:return value
    text=str(value).upper()
    if not re.fullmatch('[A-Z]{1,3}',text):raise ValueError('批注字段列必须是列字母或正整数')
    number=column_index_from_string(text)
    if number>16384:raise ValueError('批注字段列超出范围')
    return number


def owner(ws,row,col):
    cell=ws.cell(row,col)
    if isinstance(cell,MergedCell):
        for merged in ws.merged_cells.ranges:
            if cell.coordinate in merged:return ws.cell(merged.min_row,merged.min_col)
    return cell


def field_key(ws,col,header_rows):
    values=[]
    for row in header_rows:
        value=owner(ws,row,col).value
        if value in (None,''):continue
        value=re.sub(r'\s+','',str(value))
        if value=='备注1':value='备注'
        if not values or values[-1]!=value:values.append(value)
    return '|'.join(values)


def identity(ws,row,columns):
    return tuple(str(ws.cell(row,column_number(col)).value or '').strip() for col in columns)


def coordinate(value):
    text=str(value).upper().replace('$','')
    if not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]*',text):raise ValueError('批注目标坐标无效')
    row,col=coordinate_to_tuple(text)
    if row>1048576 or col>16384:raise ValueError('批注目标坐标越界')
    return row,col,text


def append_original_comment(cell,item,source_file,field,mode):
    """不改业务值；保留已有人工注正文/作者，只追加带原作者的来源段。"""
    existing=cell.comment
    original_author=item['author'] if item['author'] is not None else ''
    if mode=='direct' and existing is None:
        cell.comment=Comment(item['text'],original_author)
        return 'copied_original_object'
    if mode=='direct' and existing is not None and existing.text==item['text'] and existing.author==original_author:
        return 'existing_identical_object'
    section=(f'原工资批注来源：{Path(source_file).name} / {item["sheet"]}!{item["cell"]}\n'
             f'原字段：{field or "未命名字段"}\n原作者：{original_author}\n原文：\n{item["text"]}')
    if existing is None:
        cell.comment=Comment(section,original_author)
    else:
        # 合并不删除或改写原人工注；新批注的作者另原样写入来源段。
        merged=Comment(existing.text+'\n\n--- 保留的原工资批注 ---\n'+section,existing.author)
        merged.width=existing.width;merged.height=existing.height;cell.comment=merged
    return 'appended_with_original_author' if existing else 'aggregated_original_object'


def apply_source_comments(book,policy,plan,mappings,base_dir,frozen_inventory):
    configs=policy.get('source_comments',[])
    if policy.get('require_source_comments') and not configs:
        raise ValueError('工资登记require_source_comments=true，必须登记实际工资原件，0条批注也不能省略')
    if 'source_comments' in plan and plan['source_comments']!=configs:
        raise ValueError('批注源须在登记政策固定，计划不得事后替换/清空source_comments')
    if not configs:
        if plan.get('comment_copies'):raise ValueError('comment_copies没有预先登记source_comments')
        return None
    actual=inventory_sources(configs,base_dir)
    if actual!=frozen_inventory:
        raise ValueError('工资原件全量批注与登记盘点不符')
    overrides={}
    for item in plan.get('comment_copies',[]):
        key=(item.get('source_id'),item.get('source_sheet'),str(item.get('source_cell','')).upper().replace('$',''))
        if key in overrides:raise ValueError('同一原批注出现重复comment_copies')
        overrides[key]=item
    used=set();reports=[]
    for config,inventory in zip(configs,actual):
        source_book,_=read_source(config,base_dir)
        copies=[]
        for item in inventory['items']:
            key=(config['id'],item['sheet'],item['cell'])
            source_ws=source_book[item['sheet']]
            record=dict(item)
            if item['sheet'] in config.get('ignored_sheets',{}):
                if key in overrides:raise ValueError('历史附页已忽略，不能同时指派复制')
                record.update(status='ignored_historical_sheet',reason=config['ignored_sheets'][item['sheet']])
                copies.append(record);continue
            candidates=[r for r in config.get('regions',[]) if r['sheet']==item['sheet']]
            regions=[r for r in candidates if r['first_row']<=item['row']<=r['last_row']]
            is_person=bool(regions)
            if len(regions)>1:raise ValueError('批注人员区重叠，无法确定身份')
            if not regions:
                regions=[r for r in candidates if item['row']<r['first_row']]
            if len(regions)!=1:
                raise ValueError(f'原工资全量盘点有未归属批注，请补人员区/历史附页说明：{item["sheet"]}!{item["cell"]}')
            region=regions[0];target_sheet=region['target_sheet']
            if target_sheet not in book.sheetnames:raise ValueError('批注目标页未保留')
            target_ws=book[target_sheet]
            target_policy=policy.get('sheets',{}).get(target_sheet,{}).get('personnel',{})
            first=region.get('target_first_row',target_policy.get('start_row'))
            if not isinstance(first,int) or first<1:raise ValueError('批注目标须登记人员起始行')
            source_headers=region.get('header_rows',list(range(max(1,region['first_row']-2),region['first_row'])))
            target_headers=region.get('target_header_rows',list(range(max(1,first-2),first)))
            source_identity=identity(source_ws,item['row'],region['identity_columns']) if is_person else None
            target_names={}
            for row in range(first,target_ws.max_row+1):
                value=identity(target_ws,row,region['target_identity_columns'])
                if not any(value):continue
                if value in target_names:raise ValueError(f'批注目标身份重复，不能猜测同名：{value}')
                target_names[value]=row
            if is_person and not any(source_identity):raise ValueError('原工资有批注的人员行缺身份')
            override=overrides.get(key)
            if is_person and source_identity not in target_names:
                if override:raise ValueError('原批注员工不在目标表，不能把批注贴给别人')
                record.update(status='not_in_target',identity=list(source_identity),reason='实际目标表没有该源身份；这不证明业务上应排除')
                copies.append(record);continue
            field=field_key(source_ws,item['column'],source_headers)
            target_by_field={}
            for col in range(1,target_ws.max_column+1):
                label=field_key(target_ws,col,target_headers)
                if label:target_by_field.setdefault(label,[]).append(col)
            field_override=region.get('field_map',{}).get(get_column_letter(item['column']),'__auto__')
            target_col=None
            if field_override!='__auto__':
                target_col=column_number(field_override) if field_override is not None else None
            elif field and len(target_by_field.get(field,[]))==1:
                target_col=target_by_field[field][0]
            fallback=False
            if override:
                used.add(key)
                if override.get('target_sheet',target_sheet)!=target_sheet:raise ValueError('批注目标页与核定身份页不同')
                row,col,_=coordinate(override['target_cell'])
            elif is_person:
                row=target_names[source_identity]
                if target_col is None:fallback=True;col=column_number(region.get('fallback_column','AG'))
                else:col=target_col
            else:
                if item['row'] not in source_headers or target_col is None:
                    raise ValueError('全局/表头批注无法自动定位，请提供明确comment_copies')
                header_index=source_headers.index(item['row'])
                row=target_headers[min(header_index,len(target_headers)-1)];col=target_col
            if col>target_ws.max_column:raise ValueError('批注不得新建业务列')
            target=owner(target_ws,row,col)
            if is_person:
                if identity(target_ws,row,region['target_identity_columns'])!=source_identity:
                    raise ValueError(f'批注错人：源{source_identity}，目标{identity(target_ws,row,region["target_identity_columns"])}')
                # 部门纵向合并的主人可能在另一员工行，不能把批注误归给组首。
                if target.row!=row:fallback=True
                if target.comment is not None and (target.comment.text!=item['text'] or target.comment.author!=(item['author'] or '')):
                    fallback=True
                if fallback:
                    col=column_number(region.get('fallback_column','AG'))
                    if col>target_ws.max_column or '备注' not in field_key(target_ws,col,target_headers):
                        raise ValueError('缺对应字段/已有人工注，须用同人原模板备注栏汇总，不能新建列')
                    target=target_ws.cell(row,col)
                    if isinstance(target,MergedCell):raise ValueError('同人备注格被合并，无法保留独立批注对象')
            elif row>=first:
                raise ValueError('非人员批注不能未经核定贴到员工行')
            action=append_original_comment(target,item,inventory['source']['path'],field,'aggregate' if fallback else 'direct')
            record.update(status='copied_employee' if is_person else 'copied_header',identity=list(source_identity) if source_identity else None,
                target_sheet=target_sheet,target_cell=target.coordinate,action=action,source_field=field,
                original_text_preserved=item['text'] in target.comment.text,
                original_author_preserved=(target.comment.author==(item['author'] or '') or (item['author'] or '') in target.comment.text))
            copies.append(record)
        reports.append({'source_id':config['id'],'source':inventory['source'],'comment_count':inventory['comment_count'],
            'copied_employee_count':sum(r['status']=='copied_employee' for r in copies),
            'copied_header_count':sum(r['status']=='copied_header' for r in copies),'items':copies})
    if set(overrides)-used:raise ValueError('comment_copies指向不存在/未采用的源批注')
    return {'sources':reports,'all_source_comment_count':sum(r['comment_count'] for r in reports),
            'copied_employee_count':sum(r['copied_employee_count'] for r in reports),
            'copied_header_count':sum(r['copied_header_count'] for r in reports)}


def rewrite_package_comments(original,generated,source_part,output_part,generated_part,
                             source_xml,generated_xml,content_types,replacements,index):
    """把实际Comment对象序列化的XML/VML及关系装入原模板包，其他关系保留。"""
    from lxml import etree as ET
    import posixpath
    rel_ns='http://schemas.openxmlformats.org/package/2006/relationships'
    ct_ns='http://schemas.openxmlformats.org/package/2006/content-types'
    sheet_ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rid='{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    def relpath(part):
        directory,filename=part.rsplit('/',1)
        return directory+'/_rels/'+filename+'.rels'
    def target_part(part,target):
        return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(posixpath.dirname(part),target))
    source_rels=relpath(source_part);generated_rels=relpath(generated_part)
    root=ET.fromstring(original.read(source_rels)) if source_rels in original.namelist() else ET.Element('{'+rel_ns+'}Relationships',nsmap={None:rel_ns})
    for relation in list(root):
        if relation.get('Type','').endswith(('/comments','/vmlDrawing')):
            if relation.get('Type','').endswith('/vmlDrawing'):
                vml=ET.fromstring(original.read(target_part(source_part,relation.get('Target'))))
                if any(node.get('ObjectType') not in (None,'Note') for node in vml.iter()):
                    raise ValueError('原VML含非批注控件，不能替换；须原生保留核对')
            root.remove(relation)
    for node in list(source_xml):
        if ET.QName(node).localname=='legacyDrawing':source_xml.remove(node)
    if generated_rels in generated.namelist():
        generated_root=ET.fromstring(generated.read(generated_rels))
        for relation in generated_root:
            kind=relation.get('Type','')
            if not kind.endswith(('/comments','/vmlDrawing')):continue
            is_vml=kind.endswith('/vmlDrawing')
            target=(f'xl/drawings/templateGuardComments{index}.vml' if is_vml else
                    f'xl/comments/templateGuardComments{index}.xml')
            replacements[target]=generated.read(target_part(generated_part,relation.get('Target')))
            new_id=f'templateGuardCommentVML{index}' if is_vml else f'templateGuardComment{index}'
            if any(r.get('Id')==new_id for r in root):raise ValueError('批注关系标识冲突')
            new=ET.SubElement(root,'{'+rel_ns+'}Relationship',Id=new_id,Type=kind,Target='/'+target)
            if is_vml:
                if not any(n.get('Extension')=='vml' for n in content_types):
                    ET.SubElement(content_types,'{'+ct_ns+'}Default',Extension='vml',
                        ContentType='application/vnd.openxmlformats-officedocument.vmlDrawing')
                drawing=ET.Element('{'+sheet_ns+'}legacyDrawing');drawing.set(rid,new_id)
                before=next((i for i,n in enumerate(source_xml) if ET.QName(n).localname=='extLst'),len(source_xml))
                source_xml.insert(before,drawing)
            elif not any(n.get('PartName')=='/'+target for n in content_types):
                ET.SubElement(content_types,'{'+ct_ns+'}Override',PartName='/'+target,
                    ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml')
    replacements[relpath(output_part)]=ET.tostring(root,encoding='UTF-8',xml_declaration=True,standalone=True)


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source',required=True,type=Path);parser.add_argument('--sha256',required=True)
    parser.add_argument('--output',required=True,type=Path)
    a=parser.parse_args()
    config={'id':'inspection','source':{'path':str(a.source.resolve()),'sha256':a.sha256,'locator':'完整工作簿全部批注'}}
    try:
        _,result=read_source(config,Path.cwd())
        if a.output.exists():raise ValueError('批注盘点输出已存在，拒绝覆盖')
        a.output.parent.mkdir(parents=True,exist_ok=True)
        a.output.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf8')
        print(json.dumps({'source':result['source'],'comment_count':result['comment_count'],'output':str(a.output)},ensure_ascii=False))
        return 0
    except (ValueError,KeyError,OSError) as exc:
        print(json.dumps({'status':'rejected','error':str(exc)},ensure_ascii=False));return 2


if __name__=='__main__':raise SystemExit(main())
