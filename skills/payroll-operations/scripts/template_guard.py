"""指定模板登记、复制、来源行映射与交付实物复核；不计算工资、不声称原生重算。

所有写操作拒绝覆盖。登记 SHA 必须来自用户指定原件的预先盘点；登记文件 SHA
须固定在运行记录中。verify 不采信 JSON 的 passed，而从封存模板重放计划并读回成品。
"""
from __future__ import annotations

import argparse
from collections import defaultdict
from copy import copy, deepcopy
from datetime import date, datetime
import hashlib
import io
import json
from pathlib import Path
import re
import shutil
import sys
from xml.etree.ElementTree import tostring
from zipfile import ZipFile
from lxml import etree as ET

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import get_column_letter, range_boundaries, coordinate_to_tuple

SCHEMA = 'payroll-template-guard/1'
CELL_REF = re.compile(r'^(\$?)([A-Z]{1,3})(\$?)([1-9][0-9]*)$')


class RowMappings(dict):
    """JSON仍为普通逐表数组，内部另记原始行数以正确映射空白预留范围。"""
    source_sizes = None


def sha256(path):
    h = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(block)
    return h.hexdigest()


def digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                    separators=(',', ':')).encode('utf-8')).hexdigest()


def read_json(path):
    return json.loads(Path(path).read_text(encoding='utf-8-sig'))


def write_json(path, value):
    path = Path(path)
    if path.exists():
        raise ValueError(f'拒绝覆盖已有文件：{path}')
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('x', encoding='utf-8') as out:
        json.dump(value, out, ensure_ascii=False, indent=2)
        out.write('\n')


def xml(obj):
    return None if obj is None else tostring(obj.to_tree(), encoding='unicode')


def scalar(value):
    if isinstance(value, (datetime, date)):
        return {'date': value.isoformat()}
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def style(cell):
    return {key: xml(copy(getattr(cell, key))) for key in
            ('font', 'fill', 'border', 'alignment', 'protection')} | {
        'number_format': cell.number_format,
        'quotePrefix': bool(cell.quotePrefix), 'pivotButton': bool(cell.pivotButton)}


def dimension(value, default_height=None):
    attributes = {k: v for k, v in dict(value).items() if k not in ('s', 'style')}
    if default_height is not None and attributes.get('ht') is not None and float(attributes['ht']) == float(default_height):
        attributes.pop('ht'); attributes.pop('customHeight', None)
    if value.has_style:
        attributes['actual_style'] = style(value)
    return attributes


def fingerprint(book):
    """读取实际值/公式、逐格样式和表级参数；样式不以 style_id 代替。"""
    result = {'sheetnames': book.sheetnames, 'sheets': {},
              'defined_names': sorted(xml(v) for v in book.defined_names.values()),
              'security': xml(book.security), 'views': [xml(v) for v in book.views]}
    for ws in book:
        properties = copy(ws.sheet_properties)
        properties.codeName = None  # .xlsx无宏；应用自动分配的模块名不影响工资格式。
        cells = {}
        for row in ws:
            for c in row:
                if isinstance(c, MergedCell):
                    continue
                # 显式空字符串经过 Excel/openpyxl 保存会成为空白；不把它当人工金额。
                value = None if c.value == '' else scalar(c.value)
                if value is not None or c.has_style or c.comment or c.hyperlink:
                    cells[c.coordinate] = {'value': value, 'formula': c.data_type == 'f',
                        'style': style(c), 'comment': None if c.comment is None else
                        [c.comment.text, c.comment.author, c.comment.width, c.comment.height],
                        'hyperlink': None if c.hyperlink is None else xml(c.hyperlink)}
        result['sheets'][ws.title] = {
            'size': [ws.max_row, ws.max_column], 'state': ws.sheet_state,
            'cells': cells, 'merges': sorted(str(v) for v in ws.merged_cells.ranges),
            'rows': {str(k): dimension(v, ws.sheet_format.defaultRowHeight) for k, v in ws.row_dimensions.items()
                     if dimension(v, ws.sheet_format.defaultRowHeight)},
            'columns': {str(k): dimension(v) for k, v in ws.column_dimensions.items() if dict(v)},
            'sheet_format': xml(ws.sheet_format), 'properties': xml(properties),
            'views': xml(ws.views), 'protection': xml(ws.protection),
            'page_setup': xml(ws.page_setup), 'page_margins': xml(ws.page_margins),
            'print_options': xml(ws.print_options), 'header_footer': xml(ws.HeaderFooter),
            'print_area': str(ws.print_area), 'print_title_rows': ws.print_title_rows,
            'print_title_cols': ws.print_title_cols, 'row_breaks': xml(ws.row_breaks),
            'col_breaks': xml(ws.col_breaks), 'auto_filter': xml(ws.auto_filter),
            'data_validations': xml(ws.data_validations),
            'conditional_formatting': [(str(k), [xml(r) for r in v])
                                      for k, v in ws.conditional_formatting._cf_rules.items()],
            'tables': sorted(xml(t) for t in ws.tables.values())}
    return result


def normalized_book(book):
    buffer = io.BytesIO()
    book.save(buffer)
    buffer.seek(0)
    return load_workbook(buffer, data_only=False, keep_links=True)


def package_fingerprint(path):
    def node(element):
        return [element.tag, sorted(element.attrib.items()), element.text or '',
                [node(child) for child in element]]
    result = {}
    with ZipFile(path) as package:
        for name in package.namelist():
            data = package.read(name)
            if name.endswith('.xml'):
                tree = ET.fromstring(data)
                if name.startswith('xl/worksheets/'):
                    # 单元格/样式/人工值/行尺寸已经逐格实读检查；省略非业务的 spans
                    # 与重复样式索引序列化差异，其余扩展节点仍在封装检查范围。
                    for child in list(tree):
                        if ET.QName(child).localname == 'sheetData':
                            tree.remove(child)
                result[name] = digest(node(tree))
            else:
                result[name] = hashlib.sha256(data).hexdigest()
    return result


def package_invariants(path):
    """Excel重算可重写缓存/文档时间；原扩展、附件及非数据部件仍须实际保留。"""
    def node(element):
        return [element.tag, sorted(element.attrib.items()), (element.text or '').strip(),
                [node(child) for child in element]]
    result = {}
    with ZipFile(path) as package:
        ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        rid = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        workbook = ET.fromstring(package.read('xl/workbook.xml'))
        rels = ET.fromstring(package.read('xl/_rels/workbook.xml.rels'))
        targets = {r.get('Id'): r.get('Target') for r in rels}
        logical = {}
        for sheet in workbook.find(ns + 'sheets'):
            target = targets[sheet.get(rid)]
            part = target.lstrip('/') if target.startswith('/') else 'xl/' + target
            logical[part] = 'worksheet:' + sheet.get('name')
        attachments = []
        for name in package.namelist():
            data = package.read(name)
            if name.endswith('.xml'):
                tree = ET.fromstring(data)
                extensions = [node(v) for v in tree.iter() if ET.QName(v).localname == 'extLst' and
                              (len(v) or v.attrib or (v.text or '').strip())]
                if extensions:
                    result[logical.get(name, name) + '#extensions'] = digest(extensions)
                if name.startswith(('customXml/', 'xl/charts/', 'xl/drawings/')):
                    attachments.append(['xml', digest(node(tree))])
            elif not name.endswith(('.rels', '.vml')):
                attachments.append([Path(name).suffix, hashlib.sha256(data).hexdigest()])
        result['attachments'] = sorted(attachments)
    return result


def check_xlsx(path):
    if Path(path).suffix.lower() != '.xlsx':
        raise ValueError('本工具只处理 .xlsx；旧 .xls 不转换、不冒充完成，须用 Excel 原生保留格式另行核对')
    with ZipFile(path) as package:
        names = package.namelist()
        unsupported = [n for n in names if any(v in n for v in
                       ('/activeX/', '/ctrlProps/', '/embeddings/', '/slicer', '/pivot', 'vbaProject'))]
        if unsupported:
            raise ValueError('模板含本工具无法可靠保留的扩展，须改用 Excel 原生处理：' + ','.join(unsupported))
        # openpyxl 会移除某些扩展；不静默另存并宣称格式保留。
        # extLst 保留在原 ZIP 部件中；不让 openpyxl 重建这些部件。


def save_preserving_package(template, expected, output):
    """只替换必要工作表节点，其他 ZIP 部件逐字节保留，字体/格式表也不重建。"""
    buffer = io.BytesIO(); expected.save(buffer); buffer.seek(0)
    tags = {'sheetData', 'dimension', 'mergeCells', 'autoFilter', 'rowBreaks',
            'printOptions', 'pageMargins', 'pageSetup', 'headerFooter'}
    namespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    def local(element):
        return ET.QName(element).localname
    replacements = {}
    removed_parts = set()
    comments_rewritten = False
    with ZipFile(template) as original, ZipFile(buffer) as generated:
        # 新样式索引不能指向原样式表以外；现有样式直接复用，禁止重建近似样式。
        old_styles = ET.fromstring(original.read('xl/styles.xml'))
        new_styles = ET.fromstring(generated.read('xl/styles.xml'))
        baseline = io.BytesIO(); normalized_book(load_workbook(template)).save(baseline); baseline.seek(0)
        with ZipFile(baseline) as baseline_zip:
            baseline_styles = ET.fromstring(baseline_zip.read('xl/styles.xml'))
        for kind in ('fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'numFmts'):
            original_list = old_styles.find(f'{{{namespace}}}{kind}')
            generated_list = new_styles.find(f'{{{namespace}}}{kind}')
            permitted_list = baseline_styles.find(f'{{{namespace}}}{kind}')
            if generated_list is None:
                continue
            if len(generated_list) > len(permitted_list):
                raise ValueError(f'行/合并计划产生模板外新样式 {kind}，须改用模板既有样式或 Excel 核对')
            permitted = {ET.tostring(v, method='c14n') for v in permitted_list}
            if any(ET.tostring(v, method='c14n') not in permitted for v in generated_list):
                raise ValueError(f'样式 {kind} 不等于原模板规范化副本，拒绝重建样式')
            if original_list is None and len(generated_list):
                raise ValueError('不能新增模板不存在的样式类别')
            if original_list is not None:
                # Excel 可使用不连续自定义数字格式编号；整组复制来源模板规范化的
                # 索引才能保持日期/数字含义，不能只追加 xf 却保留旧编号。
                old_styles.replace(original_list, deepcopy(generated_list))
        replacements['xl/styles.xml'] = ET.tostring(old_styles, xml_declaration=True, encoding='UTF-8', standalone=True)
        # 用工作簿关系定位真实部件名，不假定 sheet1 是第一张表。
        wb = ET.fromstring(original.read('xl/workbook.xml'))
        rels = ET.fromstring(original.read('xl/_rels/workbook.xml.rels'))
        targets = {r.get('Id'): r.get('Target') for r in rels}
        rid_ns = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
        original_sheets = list(wb.find(f'{{{namespace}}}sheets'))
        sheet_nodes = {s.get('name'): s for s in original_sheets}
        content_types = ET.fromstring(original.read('[Content_Types].xml'))
        original_book = load_workbook(template)
        source_view_names = set(getattr(expected, '_source_view_names', []))
        for sheet in original_sheets:
            if sheet.get('name') not in expected.sheetnames:
                rid = sheet.get(rid_ns)
                target = targets[rid]
                removed_parts.add(target.lstrip('/') if target.startswith('/') else 'xl/' + target)
                wb.find(f'{{{namespace}}}sheets').remove(sheet)
                for relation in list(rels):
                    if relation.get('Id') == rid:
                        rels.remove(relation)
        sources = getattr(expected, '_template_sheet_sources', {})
        for index, name in enumerate(expected.sheetnames, 1):
            if name in source_view_names:
                path=f'xl/worksheets/templateGuardSourceView{index}.xml'
                new_rid=f'templateGuardSourceView{index}'
                ET.SubElement(wb.find(f'{{{namespace}}}sheets'), f'{{{namespace}}}sheet',
                    name=name,sheetId=str(max(int(s.get('sheetId')) for s in original_sheets)+index),
                    attrib={rid_ns:new_rid})
                ET.SubElement(rels,'{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',
                    Id=new_rid,Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                    Target=path.removeprefix('xl/'))
                ET.SubElement(content_types,'{http://schemas.openxmlformats.org/package/2006/content-types}Override',
                    PartName='/'+path,ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
                replacements[path]=generated.read(f'xl/worksheets/sheet{index}.xml')
                continue
            source_name = sources.get(name, name)
            sheet = sheet_nodes[source_name]
            target = targets[sheet.get(rid_ns)]
            path = target.lstrip('/') if target.startswith('/') else 'xl/' + target
            src = ET.fromstring(original.read(path))
            if name not in sheet_nodes:
                if original_book[source_name]._charts or original_book[source_name]._images or any(c.hyperlink for row in original_book[source_name] for c in row):
                    raise ValueError('新增页模板含图片/链接，须Excel原生复制并核对关系')
                path = f'xl/worksheets/templateGuardCopy{index}.xml'
                new_rid = f'templateGuardCopy{index}'
                new_sheet = deepcopy(sheet); new_sheet.set('name', name)
                new_sheet.set('sheetId', str(max(int(s.get('sheetId')) for s in original_sheets) + index))
                new_sheet.set(rid_ns, new_rid)
                wb.find(f'{{{namespace}}}sheets').append(new_sheet)
                new_rel = deepcopy(next(r for r in ET.fromstring(original.read('xl/_rels/workbook.xml.rels')) if r.get('Id') == sheet.get(rid_ns)))
                new_rel.set('Id', new_rid); new_rel.set('Target', path.removeprefix('xl/'))
                rels.append(new_rel)
                ET.SubElement(content_types, '{http://schemas.openxmlformats.org/package/2006/content-types}Override',
                    PartName='/' + path, ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml')
            gen = ET.fromstring(generated.read(f'xl/worksheets/sheet{index}.xml'))
            if name not in sheet_nodes:
                view = gen.find(f'{{{namespace}}}sheetViews')
                old_view = src.find(f'{{{namespace}}}sheetViews')
                if old_view is not None and view is not None:
                    src.replace(old_view, deepcopy(view))
            merge_node = gen.find(f'{{{namespace}}}mergeCells')
            if merge_node is not None:
                merge_node[:] = sorted(merge_node, key=lambda child: child.get('ref'))
            # 批注对象重建独立关系，超链接移动仍须另行原生核对。
            orig_ws, dest_ws = original_book[source_name], expected[name]
            if any(c.hyperlink for row in orig_ws for c in row) and any(
                    (c.hyperlink.target if c.hyperlink else None) !=
                    (dest_ws[c.coordinate].hyperlink.target if dest_ws[c.coordinate].hyperlink else None)
                    for row in orig_ws for c in row if not isinstance(c, MergedCell)):
                raise ValueError('超链接位置变化须 Excel 原生保留核对')
            for child in list(src):
                if local(child) in tags:
                    replacement = gen.find(child.tag)
                    if replacement is None:
                        src.remove(child)
                    else:
                        src.replace(child, deepcopy(replacement))
            order = [local(c) for c in gen]
            for child in gen:
                if local(child) not in tags or src.find(child.tag) is not None:
                    continue
                following = set(order[order.index(local(child))+1:])
                at = next((i for i, c in enumerate(src) if local(c) in following), len(src))
                src.insert(at, deepcopy(child))
            def comments(ws):
                return {c.coordinate:(c.comment.text,c.comment.author) for row in ws for c in row if c.comment is not None}
            if comments(orig_ws)!=comments(dest_ws) or name not in sheet_nodes and comments(dest_ws):
                from template_comments import rewrite_package_comments
                original_path = target.lstrip('/') if target.startswith('/') else 'xl/' + target
                rewrite_package_comments(original,generated,original_path,path,
                    f'xl/worksheets/sheet{index}.xml',src,gen,content_types,replacements,index)
                comments_rewritten = True
            replacements[path] = ET.tostring(src, xml_declaration=True, encoding='UTF-8', standalone=True)
        generated_wb = ET.fromstring(generated.read('xl/workbook.xml'))
        old_names, new_names = wb.find(f'{{{namespace}}}definedNames'), generated_wb.find(f'{{{namespace}}}definedNames')
        if old_names is not None and new_names is not None:
            wb.replace(old_names, deepcopy(new_names))
        old_views, new_views = wb.find(f'{{{namespace}}}bookViews'), generated_wb.find(f'{{{namespace}}}bookViews')
        if old_views is not None and new_views is not None:
            wb.replace(old_views, deepcopy(new_views))
        replacements['xl/workbook.xml'] = ET.tostring(wb, xml_declaration=True, encoding='UTF-8', standalone=True)
        sheet_container = wb.find(f'{{{namespace}}}sheets')
        sheet_container[:] = sorted(sheet_container, key=lambda item: expected.sheetnames.index(item.get('name')))
        replacements['xl/workbook.xml'] = ET.tostring(wb, xml_declaration=True, encoding='UTF-8', standalone=True)
        if removed_parts or comments_rewritten or any(name not in sheet_nodes for name in expected.sheetnames):
            for item in list(content_types):
                if item.get('PartName', '').lstrip('/') in removed_parts:
                    content_types.remove(item)
            replacements['[Content_Types].xml'] = ET.tostring(content_types, xml_declaration=True, encoding='UTF-8', standalone=True)
            replacements['xl/_rels/workbook.xml.rels'] = ET.tostring(rels, xml_declaration=True, encoding='UTF-8', standalone=True)
            for path in list(removed_parts):
                directory, filename = path.rsplit('/', 1)
                removed_parts.add(directory + '/_rels/' + filename + '.rels')
        with ZipFile(output, 'w') as result:
            for info in original.infolist():
                if info.filename not in removed_parts:
                    result.writestr(info, replacements.pop(info.filename, original.read(info.filename)))
            for name, data in replacements.items():
                result.writestr(name, data)


def select_sheets(book, policy):
    rule = policy.get('sheet_policy', {})
    keep = rule.get('keep', book.sheetnames)
    copies = rule.get('copies', {})
    if not keep or len(keep) != len(set(keep)) or set(keep) - set(book.sheetnames):
        raise ValueError('工作表保留政策为空、重复或引用模板外工作表')
    dropped = set(book.sheetnames) - set(keep)
    for name in keep:
        for row in book[name]:
            for cell in row:
                if cell.data_type != 'f':
                    continue
                for token in Tokenizer(cell.value).items:
                    if token.subtype == 'RANGE' and '!' in token.value:
                        target = token.value.rsplit('!', 1)[0].strip("'").replace("''", "'")
                        if target in dropped or dropped and ':' in target:
                            raise ValueError(f'删除页仍被公式引用，须保留必要内部页：{name}!{cell.coordinate}')
    for value in book.defined_names.values():
        if any(re.search(rf"(?:'{re.escape(name)}'|(?<![\w]){re.escape(name)})!", value.attr_text or '') for name in dropped):
            raise ValueError('删除页仍被命名范围引用')
    for name in list(book.sheetnames):
        if name in dropped:
            book.remove(book[name])
    # 保留页相对顺序也须来自登记；允许明确指定业务页次序。
    book._sheets = [book[name] for name in keep]
    if dropped:
        for view in book.views:
            view.firstSheet = min(view.firstSheet or 0, len(keep)-1)
            view.activeTab = min(view.activeTab or 0, len(keep)-1)
    sources = {name: name for name in keep}
    for target, source in copies.items():
        if source not in keep or target in book.sheetnames or not target or len(target) > 31 or re.search(r'[\\/*?:\[\]]', target):
            raise ValueError('新增页须在登记政策copies中指定合法新页名及保留的模板来源页')
        copied = book.copy_worksheet(book[source]); copied.title = target
        sources[target] = source
    return sources


def evidence(item, base):
    """只证明来源文件与定位可追溯；不把自填 approved/passed 当授权真实性。"""
    if not isinstance(item, dict) or not item.get('path') or not item.get('sha256') or not item.get('locator'):
        raise ValueError('来源须含 path、预先固定的 sha256 和非空 locator')
    path = Path(item['path'])
    if not path.is_absolute():
        path = Path(base) / path
    if sha256(path) != item['sha256']:
        raise ValueError(f'来源 SHA 不符：{path}')
    return {'path': str(path.resolve()), 'sha256': item['sha256'], 'locator': item['locator']}


def instruction_identity(value):
    """仅识别明确指引文字；不把普通样例姓名或正常空人员行判成说明。"""
    if not isinstance(value,str) or value.startswith('='):return False
    text=re.sub(r'\s+','',value).replace('：',':')
    if text.casefold() in {'说明','填写指引','填写说明','示例','姓名示例','姓名','员工姓名','人员姓名','员工编号','工号','userid'}:
        return True
    return bool(re.match(r'^(?:从.{1,80}(?:提取|获取|取数|导入|读取)|取自.{1,80}|'
        r'请(?:在此)?(?:填写|填入|输入|录入|选择)|(?:在此|此处)(?:填写|填入|输入|录入)|'
        r'根据.{1,80}(?:填写|填入|提取|获取)|《.{1,60}》(?:统计|提取|获取|取数)|'
        r'(?:示例|例如|填写示例|填写指引|填写说明):)',text))


def personnel_instruction_rows(book, policy):
    """从原件实际登记身份列读取，坐标/行号不写死；登记范围不豁免说明行。"""
    result={}
    for name,rule in policy.get('sheets',{}).items():
        region=rule.get('personnel')
        if not region or name not in book.sheetnames:continue
        rows=[];ws=book[name]
        for row in range(1,ws.max_row+1):
            cells=[ws[f'{column}{row}'] for column in region.get('identity_columns',[])]
            notes=[{'cell':c.coordinate,'value':c.value} for c in cells if instruction_identity(c.value)]
            if notes:rows.append({'source_row':row,'instruction_identity_cells':notes})
        result[name]=rows
    return result


def register(template, expected_sha256, policy, registry):
    template, registry = Path(template).resolve(), Path(registry).resolve()
    if sha256(template) != expected_sha256:
        raise ValueError('指定模板 SHA 与预先固定值不同，停止登记')
    check_xlsx(template)
    book = load_workbook(template)
    before = fingerprint(book)
    after = fingerprint(normalized_book(book))
    if before != after:
        raise ValueError('内存规范化已改变模板结构/样式；须 Excel 原生处理，不能放宽为仅查表头')
    select_sheets(load_workbook(template), policy)
    from template_comments import inventory_sources
    if policy.get('require_source_comments') and not policy.get('source_comments'):
        raise ValueError('工资必须登记source_comments原件清单；无批注也要实际盘点')
    comment_inventories = inventory_sources(policy.get('source_comments',[]),registry.parent)
    from template_source_views import inventory_sources as inventory_views
    source_view_inventories=inventory_views(policy.get('sheet_policy',{}).get('source_views',[]),registry.parent)
    selected=select_sheets(load_workbook(template),policy)
    if any(item['sheet'] in selected for item in source_view_inventories):
        raise ValueError('业务来源值页不得替换模板工作表')
    for name, rule in policy.get('sheets', {}).items():
        if name not in book.sheetnames:
            raise ValueError(f'政策工作表不存在：{name}')
        for coord, spec in rule.get('formula_rules', {}).items():
            if not str(spec.get('expression', '')).startswith('='):
                raise ValueError(f'核定公式缺 expression：{name}!{coord}')
            evidence(spec.get('source'), registry.parent)
    sealed = registry.parent / (registry.stem + '.template.xlsx')
    if sealed.exists() or registry.exists() or sealed == template:
        raise ValueError('登记文件/封存模板已存在或指向原件')
    sealed.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(template, sealed)
    if sha256(sealed) != expected_sha256 or sha256(template) != expected_sha256:
        raise ValueError('封存过程中源件发生变化，未完成登记；保留失败副本供核对')
    record = {'schema': SCHEMA, 'template': {'path': str(template), 'sha256': expected_sha256},
              'sealed_template': {'path': str(sealed), 'sha256': sha256(sealed)},
              'policy': policy, 'template_fingerprint': before,
              'comment_source_inventories': comment_inventories,
              'source_view_inventories': source_view_inventories,
              'personnel_instruction_rows':personnel_instruction_rows(book,policy),
              'template_fingerprint_sha256': digest(before),
              'created_at': datetime.now().astimezone().isoformat()}
    write_json(registry, record)
    return {'registry': str(registry), 'registry_sha256': sha256(registry),
            'template_sha256': expected_sha256, 'sheets': book.sheetnames,
            'status': 'registered', 'native_recalculation': False}


def load_registry(registry, expected_registry_sha256):
    if sha256(registry) != expected_registry_sha256:
        raise ValueError('登记文件 SHA 不符；不能事后换政策或把重建表重新登记为本次模板')
    record = read_json(registry)
    if record.get('schema') != SCHEMA:
        raise ValueError('不支持的模板登记版本')
    for key in ('template', 'sealed_template'):
        if sha256(record[key]['path']) != record[key]['sha256']:
            raise ValueError(f'{key} 原件或封存文件 SHA 不符')
    book = load_workbook(record['sealed_template']['path'])
    if fingerprint(book) != record['template_fingerprint']:
        raise ValueError('封存模板指纹不符')
    return record, book


def mapped_range(ref, source_sheet, mappings):
    c1, r1, c2, r2 = range_boundaries(ref.replace('$', ''))
    rows = [n for n, m in enumerate(mappings[source_sheet], 1) if r1 <= m['source_row'] <= r2]
    source_max = (getattr(mappings, 'source_sizes', None) or {}).get(source_sheet, max(i['source_row'] for i in mappings[source_sheet]))
    delta = len(mappings[source_sheet]) - source_max
    if r2 > source_max:
        rows.extend(range(max(r1, source_max+1)+delta, r2+delta+1))
    if not rows or rows != list(range(min(rows), max(rows) + 1)):
        raise ValueError(f'引用/合并范围删除或变成不连续，须核定明确公式/合并映射：{source_sheet}!{ref}')
    if (r1, r2) == (min(rows), max(rows)):
        return ref
    return f'{get_column_letter(c1)}{min(rows)}:{get_column_letter(c2)}{max(rows)}'


def remap_formula(formula, sheet, source_row, output_row, mappings):
    """本行相对引用随人员行移动，跨行/跨表引用依来源行映射，合计范围含新增行。"""
    tokens = Tokenizer(formula).items
    for token in tokens:
        if token.type != 'OPERAND' or token.subtype != 'RANGE':
            continue
        target_sheet, ref = sheet, token.value
        prefix = ''
        if '!' in ref:
            prefix, ref = ref.rsplit('!', 1)
            target_sheet = prefix.strip("'").replace("''", "'")
            prefix += '!'
        if target_sheet not in mappings:
            # 外部工作簿链接、命名公式不能猜测重排含义。
            if any(any(i['source_row'] != n or i['mode'] != 'copy' for n, i in enumerate(v, 1))
                   for v in mappings.values()):
                raise ValueError(f'重排行存在无法追踪的外部/三维引用：{token.value}')
            continue
        if ':' in ref:
            endpoints = ref.split(':')
            if not all(CELL_REF.fullmatch(v) for v in endpoints):
                raise ValueError(f'暂不支持重排整列/命名范围公式：{token.value}')
            start, end = [CELL_REF.fullmatch(v) for v in endpoints]
            if target_sheet == sheet and int(start[4]) == source_row == int(end[4]) and not start[3] and not end[3]:
                mapped = f'{start[2]}{output_row}:{end[2]}{output_row}'
            else:
                mapped = mapped_range(ref, target_sheet, mappings)
            # 保留各端绝对列/行标志；行坐标按实际插删映射。
            target = mapped.split(':')
            token.value = prefix + ':'.join(
                ('$' if a.startswith('$') else '') + CELL_REF.fullmatch(b)[2] +
                ('$' if CELL_REF.fullmatch(a)[3] else '') + CELL_REF.fullmatch(b)[4]
                for a, b in zip(endpoints, target))
            continue
        match = CELL_REF.fullmatch(ref)
        if not match:
            if any(i['source_row'] != n for n, i in enumerate(mappings[target_sheet], 1)):
                raise ValueError(f'命名引用涉及重排行，须 Excel 原生核对：{ref}')
            continue
        row = int(match[4])
        if target_sheet == sheet and row == source_row and not match[3]:
            dest = output_row
        else:
            candidates = [n for n, item in enumerate(mappings[target_sheet], 1)
                          if item['source_row'] == row and item['mode'] == 'copy']
            if len(candidates) != 1:
                raise ValueError(f'公式引用行删除或不唯一：{token.value}')
            dest = candidates[0]
        token.value = prefix + match[1] + match[2] + match[3] + str(dest)
    return '=' + ''.join(t.value for t in tokens)


def in_ranges(coord, ranges):
    from openpyxl.worksheet.cell_range import CellRange
    return any(coord in CellRange(r) for r in ranges)


def build_expected(registry, registry_sha256, plan_path):
    record, book = load_registry(registry, registry_sha256)
    sheet_sources = select_sheets(book, record['policy'])
    plan = read_json(plan_path)
    base = Path(plan_path).resolve().parent
    mappings, provenance = RowMappings(), []
    mappings.source_sizes = {ws.title: ws.max_row for ws in book}
    policies = record['policy'].get('sheets', {})
    if set(plan.get('row_maps', {})) - set(book.sheetnames):
        raise ValueError('行映射含模板外工作表')
    for ws in book:
        rows = plan.get('row_maps', {}).get(ws.title)
        if rows is None:
            rows = [{'source_row': n, 'mode': 'copy'} for n in range(1, ws.max_row + 1)]
        rows = [dict(source_row=i, mode='copy') if isinstance(i, int) else dict(i) for i in rows]
        if not rows:
            raise ValueError(f'不能删除整个工作表：{ws.title}')
        for item in rows:
            item.setdefault('mode', 'copy')
            if item['mode'] not in ('copy', 'new') or isinstance(item.get('source_row'), bool) or not isinstance(item.get('source_row'), int) or not 1 <= item['source_row'] <= ws.max_row:
                raise ValueError('来源行或模式无效')
        changed = rows != [{'source_row': n, 'mode': 'copy'} for n in range(1, ws.max_row + 1)]
        if changed:
            policy = policies.get(ws.title, {})
            region = policy.get('personnel')
            if not region or not region.get('identity_columns'):
                raise ValueError(f'未登记可变人员区和身份列：{ws.title}')
            start, end = region['start_row'], region['end_row']
            if ws.max_row > end and [r['source_row'] for r in rows[-(ws.max_row-end):]] != list(range(end+1, ws.max_row+1)):
                raise ValueError('只能在登记人员区内增删/重排，表头与尾部必须沿原序保留')
            # 独立判断，避免尾区为空时忽略头部。
            if [r['source_row'] for r in rows[:start-1]] != list(range(1, start)):
                raise ValueError('表头行必须沿原序保留')
            mid_end = len(rows) - (ws.max_row - end)
            if any(not start <= r['source_row'] <= end for r in rows[start-1:mid_end]):
                raise ValueError('人员区行不得取自表头或合计区')
            copied = [r['source_row'] for r in rows if r['mode'] == 'copy']
            if len(copied) != len(set(copied)):
                raise ValueError('不能复制同一既有员工为两名员工；新增须 mode:new')
            provenance.append(evidence(plan.get('row_change_sources', {}).get(ws.title), base))
            if ws.tables or ws.conditional_formatting or ws.data_validations.count or ws._charts or ws._images:
                raise ValueError('此表含表格/条件格式/数据验证/图表，重排行须 Excel 原生处理')
        mappings[ws.title] = rows
    instruction_rows=personnel_instruction_rows(book,record['policy'])
    for name,items in mappings.items():
        notes={x['source_row']:x for x in instruction_rows.get(name,[])}
        for dest,item in enumerate(items,1):
            if item['source_row'] not in notes:continue
            writes=[w for w in plan.get('writes',[]) if w.get('sheet')==name and coordinate_to_tuple(w['cell'])[0]==dest]
            if item['mode']=='new' or writes:
                raise ValueError(f'明确说明行不得作为人员样式原型或改写员工身份：{name}!原行{item["source_row"]}→行{dest}；'
                                 '请选正常空人员行/合法样例原型，说明仅可原样保留')
    originals = {s.title: {key: copy(c) for key, c in s._cells.items()} for s in book}
    any_row_changes = any(items != [{'source_row': n, 'mode': 'copy'} for n in range(1, book[name].max_row+1)]
                          for name, items in mappings.items())
    for ws in book:
        rows = mappings[ws.title]
        sheet_row_changes = rows != [{'source_row': n, 'mode': 'copy'} for n in range(1, ws.max_row+1)]
        old_dimensions = {key: copy(value) for key, value in ws.row_dimensions.items()}
        old_merges = [str(m) for m in ws.merged_cells.ranges]
        for merge in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(merge))
        ws._cells = {}
        ws.row_dimensions.clear()
        for dest, item in enumerate(rows, 1):
            src, mode = item['source_row'], item['mode']
            if src in old_dimensions:
                dim = copy(old_dimensions[src]); dim.index = dest
                ws.row_dimensions[dest] = dim
            for (r, col), original in originals[ws.title].items():
                if r != src or isinstance(original, MergedCell):
                    continue
                cell = ws.cell(dest, col)
                cell._style = copy(original._style)
                cell.protection = copy(original.protection)
                if mode == 'copy':
                    cell.value = original.value
                    cell.comment = copy(original.comment)
                    cell._hyperlink = copy(original._hyperlink)
                    if cell._hyperlink:
                        cell._hyperlink.ref = cell.coordinate
                elif original.data_type == 'f':
                    cell.value = original.value
                if original.data_type == 'f' and any_row_changes:
                    cell.value = remap_formula(original.value, ws.title, src, dest, mappings)
        merges = plan.get('merge_maps', {}).get(ws.title)
        if merges is not None:
            provenance.append(evidence(plan.get('row_change_sources', {}).get(ws.title), base))
            for item in merges:
                if item['source'] not in old_merges:
                    raise ValueError('合并来源不是模板原合并范围')
                a, _, c, _ = range_boundaries(item['source'])
                b, _, d, _ = range_boundaries(item['target'])
                if (a, c) != (b, d):
                    raise ValueError('不得改变模板合并列结构')
                _, source_r, _, _ = range_boundaries(item['source'])
                _, target_r, _, _ = range_boundaries(item['target'])
                source_title = originals[ws.title].get((source_r, a))
                if source_title is not None:
                    ws.cell(target_r, b)._style = copy(source_title._style)
                    if ws.cell(target_r, b).value is None:
                        ws.cell(target_r, b).value = source_title.value
                ws.merge_cells(item['target'])
        else:
            for merge in old_merges:
                c1, r1, c2, r2 = range_boundaries(merge)
                if r1 == r2:
                    for n, item in enumerate(rows, 1):
                        if item['source_row'] == r1:
                            ws.merge_cells(start_row=n, end_row=n, start_column=c1, end_column=c2)
                else:
                    if not any(r1 <= item['source_row'] <= r2 for item in rows):
                        continue  # 整个核定人员组已删除；其组标题合并一并删除。
                    try:
                        target = mapped_range(merge, ws.title, mappings)
                        _, target_r, _, _ = range_boundaries(target)
                        source_title = originals[ws.title].get((r1, c1))
                        if source_title is not None:
                            ws.cell(target_r, c1)._style = copy(source_title._style)
                            if ws.cell(target_r, c1).value is None:
                                ws.cell(target_r, c1).value = source_title.value
                        ws.merge_cells(target)
                    except ValueError:
                        raise ValueError(f'合并组边界已变，请提供核定 merge_maps：{ws.title}!{merge}')
        if ws.print_area and sheet_row_changes:
            ws.print_area = [mapped_range(str(v), ws.title, mappings) for v in ws.print_area.ranges]
        if ws.print_title_rows and sheet_row_changes:
            first, last = map(int, ws.print_title_rows.replace('$', '').split(':'))
            ref = mapped_range(f'A{first}:A{last}', ws.title, mappings)
            _, first, _, last = range_boundaries(ref)
            ws.print_title_rows = f'{first}:{last}'
        if ws.auto_filter.ref and sheet_row_changes:
            ws.auto_filter.ref = mapped_range(ws.auto_filter.ref, ws.title, mappings)
        for br in ws.row_breaks.brk:
            cand = [n for n, item in enumerate(rows, 1) if item['source_row'] == br.id]
            if len(cand) != 1:
                raise ValueError('分页符来源行不唯一，须 Excel 核对')
            br.id = cand[0]
    seen = set()
    for item in plan.get('writes', []):
        name, coord = item['sheet'], item['cell']
        if (name, coord) in seen:
            raise ValueError(f'重复写入：{name}!{coord}')
        seen.add((name, coord))
        ws = book[name]; target = ws[coord]
        if target.row > len(mappings[name]):
            raise ValueError('写入超出核定行映射')
        src = mappings[name][target.row-1]['source_row']
        original_coord = f'{target.column_letter}{src}'
        rule = policies.get(name, {})
        if not in_ranges(original_coord, rule.get('writable_ranges', [])):
            raise ValueError(f'目标不在登记可写区：{name}!{coord}')
        if 'expected_old' not in item or scalar(target.value) != item['expected_old']:
            raise ValueError(f'原值与写入计划不符：{name}!{coord}')
        if not item.get('reason'):
            raise ValueError('修改须说明理由')
        provenance.append(evidence(item.get('source'), base))
        if target.value not in (None, ''):
            provenance.append(evidence(item.get('authorization'), base))
        value = item.get('value')
        formula_rule = rule.get('formula_rules', {}).get(original_coord)
        if target.data_type == 'f' or isinstance(value, str) and value.startswith('='):
            if not formula_rule:
                raise ValueError(f'公式替换/补入没有预先登记规则：{name}!{coord}')
            approved = remap_formula(formula_rule['expression'], name, src, target.row, mappings)
            if value != approved:
                raise ValueError(f'公式必须与核定来源表达式及行映射一致：{name}!{coord}')
            refs = [t.value for t in Tokenizer(value).items if t.subtype == 'RANGE']
            if not refs:
                raise ValueError('公式仅加等号且没有单元格依赖，拒绝伪联动')
            for column in formula_rule.get('required_row_columns', []):
                if not any(re.search(rf'(?<![A-Z0-9_])\$?{column}\$?{target.row}(?![0-9])', ref) for ref in refs):
                    raise ValueError(f'公式缺本行依赖：{column}{target.row}')
            provenance.append(evidence(formula_rule['source'], Path(registry).resolve().parent))
        target.value = value
    for name, rows in mappings.items():
        region = policies.get(name, {}).get('personnel')
        if not region:
            continue
        identities = set()
        for dest, item in enumerate(rows, 1):
            if not region['start_row'] <= item['source_row'] <= region['end_row']:
                continue
            identity = tuple(scalar(book[name][f'{col}{dest}'].value) for col in region['identity_columns'])
            if item['mode'] == 'new' and not any(identity):
                raise ValueError('新增人员行必须从来源填写身份，不能复制原员工姓名')
            if any(identity) and identity in identities:
                raise ValueError('人员身份重复，请先核对身份主键/同名')
            identities.add(identity)
            if item['mode'] == 'copy':
                old_identity = tuple(scalar(originals[name].get((item['source_row'], range_boundaries(f'{col}1:{col}1')[0])).value)
                                     if originals[name].get((item['source_row'], range_boundaries(f'{col}1:{col}1')[0])) else None
                                     for col in region['identity_columns'])
                if any(old_identity) and identity != old_identity:
                    raise ValueError('既有员工身份不得改成另一人；新员工应使用 mode:new 并清空旧人工值')
    from template_comments import apply_source_comments
    comment_report = apply_source_comments(book,record['policy'],plan,mappings,Path(registry).resolve().parent,
        record.get('comment_source_inventories',[]))
    from template_source_views import apply_source_views
    source_view_report=apply_source_views(book,record['policy'],plan,mappings,Path(registry).resolve().parent,
        record.get('source_view_inventories',[]))
    result = normalized_book(book)
    result._template_sheet_sources = sheet_sources
    result._source_comment_report = comment_report
    result._source_view_report = source_view_report
    result._source_view_names = [item['sheet'] for item in source_view_report['sources']]
    return record, result, mappings, provenance


def compare(expected, actual):
    errors = []
    if expected['sheetnames'] != actual['sheetnames']:
        return ['工作表名称或顺序改变']
    for key in ('defined_names', 'security', 'views'):
        if expected[key] != actual[key]:
            errors.append(f'工作簿参数漂移：{key}')
    for name, reference in expected['sheets'].items():
        current = actual['sheets'][name]
        for key in reference:
            if reference[key] == current[key]:
                continue
            if key == 'cells':
                for coord in sorted(set(reference[key]) | set(current[key])):
                    if reference[key].get(coord) != current[key].get(coord):
                        errors.append(f'值/公式/人工值/样式/批注与来源行不符：{name}!{coord}')
                        if len(errors) >= 40:
                            return errors
            else:
                errors.append(f'表结构/样式/打印参数漂移：{name}.{key}')
    return errors


def source_view_mapping(registry, registry_sha256, view_id=None):
    """只读实读固定来源；不手填几十个员工行坐标。"""
    record,_=load_registry(registry,registry_sha256)
    from template_source_views import inventory_sources
    actual=inventory_sources(record['policy'].get('sheet_policy',{}).get('source_views',[]),Path(registry).resolve().parent)
    if actual!=record.get('source_view_inventories',[]):
        raise ValueError('业务来源值页盘点与登记不符')
    selected=[x for x in actual if view_id is None or x['id']==view_id]
    if view_id is not None and not selected:raise ValueError('来源值页id不存在')
    return {'sources':selected}


def build(registry, registry_sha256, plan, output, receipt):
    output = Path(output).resolve()
    if output.exists() or Path(receipt).exists():
        raise ValueError('输出或来源凭据已存在，拒绝覆盖')
    record, expected, mappings, provenance = build_expected(registry, registry_sha256, plan)
    output.parent.mkdir(parents=True, exist_ok=True)
    # 输出首先逐字节复制封存模板；随后在该副本上保存核定计划形成的新工作簿状态。
    shutil.copyfile(record['sealed_template']['path'], output)
    copied_sha = sha256(output)
    save_preserving_package(record['sealed_template']['path'], expected, output)
    errors = compare(fingerprint(expected), fingerprint(load_workbook(output)))
    if errors:
        raise ValueError('; '.join(errors))
    result = {'schema': SCHEMA, 'status': 'verified_structure_pending_recalculation',
              'output': {'path': str(output), 'sha256': sha256(output)},
              'template_sha256': record['template']['sha256'], 'copied_template_sha256': copied_sha,
              'registry': {'path': str(Path(registry).resolve()), 'sha256': registry_sha256},
              'plan': {'path': str(Path(plan).resolve()), 'sha256': sha256(plan)},
              'row_mapping': mappings, 'source_evidence': provenance,
              'source_comment_inventory': getattr(expected,'_source_comment_report',None),
              'source_view_inventory': getattr(expected,'_source_view_report',None),
              'fingerprint_sha256': digest(fingerprint(expected)),
              'native_recalculation': False, 'formal_ready': False,
              'limitations': ['只证明结构、公式来源和样式；不证明业务金额正确或人工授权真实',
                              '未执行 Excel 原生重算或改单元格联动测试']}
    write_json(receipt, result)
    return result


def verify(registry, registry_sha256, plan, output, receipt, report=None,
           recalculation_log=None, recalculation_log_sha256=None):
    actual_sha = sha256(output)
    errors = []
    comment_report = None
    source_view_report = None
    try:
        record, expected, mappings, provenance = build_expected(registry, registry_sha256, plan)
        comment_report = getattr(expected,'_source_comment_report',None)
        source_view_report = getattr(expected,'_source_view_report',None)
        origin = read_json(receipt)
        log_evidence = None
        if recalculation_log is not None:
            log_evidence = evidence({'path': str(Path(recalculation_log).resolve()),
                'sha256': recalculation_log_sha256, 'locator': '实际Excel重算运行日志全文'}, Path.cwd())
            log = read_json(recalculation_log)
            if (log.get('source_sha256') != origin.get('output', {}).get('sha256') or
                log.get('output_sha256') != actual_sha or not log.get('engine') or
                not log.get('events') or not log.get('started_at') or not log.get('finished_at')):
                raise ValueError('重算日志须绑定原生成SHA及本次输出SHA，包含引擎、起止时间及实际事件；不采信passed参数')
        for label, left, right in (
            ('登记', origin.get('registry', {}).get('sha256'), registry_sha256),
            ('计划', origin.get('plan', {}).get('sha256'), sha256(plan)),
            ('模板', origin.get('template_sha256'), record['template']['sha256']),
            ('复制来源', origin.get('copied_template_sha256'), record['template']['sha256']),
            ('输出SHA', origin.get('output', {}).get('sha256'),
             origin.get('output', {}).get('sha256') if recalculation_log is not None else actual_sha),
            ('人员来源行', origin.get('row_mapping'), mappings)):
            if left != right:
                errors.append(f'{label}与生成凭据不符')
        if origin.get('source_comment_inventory') != comment_report:
            errors.append('工资原件全量批注盘点及对象去向与生成凭据不符')
        if origin.get('source_view_inventory') != source_view_report:
            errors.append('固定业务原件值快照及同人引用与生成凭据不符')
        check_xlsx(output)
        errors += compare(fingerprint(expected), fingerprint(load_workbook(output)))
        replay = io.BytesIO()
        save_preserving_package(record['sealed_template']['path'], expected, replay)
        if recalculation_log is not None:
            expected_parts, actual_parts = package_invariants(replay), package_invariants(output)
        else:
            expected_parts, actual_parts = package_fingerprint(replay), package_fingerprint(output)
        for part in sorted(set(expected_parts) | set(actual_parts)):
            if expected_parts.get(part) != actual_parts.get(part):
                errors.append(f'原模板封装/扩展或核定内容不符：{part}')
    except (ValueError, KeyError, TypeError, OSError) as exc:
        errors.append(str(exc)); mappings = {}; record = {}; log_evidence = None
    result = {'schema': SCHEMA, 'passed': not errors, 'errors': errors,
              'output': {'path': str(Path(output).resolve()), 'sha256': actual_sha},
              'registry_sha256': registry_sha256, 'plan_sha256': sha256(plan),
              'receipt_sha256': sha256(receipt),
              'template_sha256': record.get('template', {}).get('sha256'),
              'row_mapping': mappings, 'native_recalculation': False, 'formal_ready': False,
              'source_comment_inventory': comment_report,
              'source_view_inventory': source_view_report,
              'finalized_output_sha256': actual_sha, 'recalculation_log': log_evidence,
              'status': 'verified_structure_pending_recalculation' if not errors else 'rejected'}
    if report is not None:
        write_json(report, result)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    reg = commands.add_parser('register', help='以预先固定 SHA 登记用户指定模板')
    reg.add_argument('--template', required=True, type=Path)
    reg.add_argument('--expected-sha256', required=True)
    reg.add_argument('--policy', required=True, type=Path)
    reg.add_argument('--registry', required=True, type=Path)
    view=commands.add_parser('source-views',help='实读固定业务来源并列出自动来源页/员工行映射')
    view.add_argument('--registry',required=True,type=Path)
    view.add_argument('--registry-sha256',required=True)
    view.add_argument('--view-id')
    for command in ('build', 'verify', 'finalize'):
        sub = commands.add_parser(command)
        for name in ('registry', 'plan', 'output', 'receipt'):
            sub.add_argument('--' + name, required=True, type=Path)
        sub.add_argument('--registry-sha256', required=True)
        if command in ('verify', 'finalize'):
            sub.add_argument('--report', required=True, type=Path)
        if command == 'finalize':
            sub.add_argument('--recalculation-log', required=True, type=Path)
            sub.add_argument('--recalculation-log-sha256', required=True)
    args = vars(parser.parse_args()); command = args.pop('command')
    try:
        if command == 'register':
            args['policy'] = read_json(args['policy'])
            result = register(**args)
        elif command=='source-views':
            result=source_view_mapping(**args)
        else:
            result = globals()['verify' if command == 'finalize' else command](**args)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 2 if result.get('passed') is False else 0
    except (ValueError, KeyError, TypeError, OSError) as exc:
        print(json.dumps({'status': 'rejected', 'error': str(exc)}, ensure_ascii=False))
        return 2


if __name__ == '__main__':
    sys.exit(main())
