"""仅从同公式/同输入的实际计算副本搬运缓存，完整保留原模板样式封装。

计算应用可能舍入行列尺寸，因此不交付该另存副本。此工具不执行计算，也不证明
缓存正确；独立原生/兼容应用审计仍须核对。不得把搬运缓存称为Microsoft Excel重算。
"""
import argparse
from copy import deepcopy
import json
from pathlib import Path
from zipfile import ZipFile
from lxml import etree as ET
from openpyxl import load_workbook
from template_guard import sha256, scalar, write_json

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


def sheet_parts(package):
    book = ET.fromstring(package.read('xl/workbook.xml'))
    rels = ET.fromstring(package.read('xl/_rels/workbook.xml.rels'))
    targets = {r.get('Id'): r.get('Target') for r in rels}
    rid = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    return {s.get('name'): target.lstrip('/') if target.startswith('/') else 'xl/' + target
            for s in book.find(NS+'sheets') for target in [targets[s.get(rid)]]}


def carry(base, calculated, output, report=None):
    base, calculated, output = map(lambda p: Path(p).resolve(), (base, calculated, output))
    if output.exists() or report and Path(report).exists():
        raise ValueError('缓存目标/报告必须另存新文件，不覆盖源件')
    original, computed = load_workbook(base), load_workbook(calculated)
    cache = load_workbook(calculated, data_only=True)
    if original.sheetnames != computed.sheetnames:
        raise ValueError('计算副本工作表与来源不一致')
    for ws in original:
        actual = computed[ws.title]
        for row, col in set(ws._cells) | set(actual._cells):
            left, right = ws.cell(row, col), actual.cell(row, col)
            if scalar(left.value) != scalar(right.value) or (left.data_type == 'f') != (right.data_type == 'f'):
                raise ValueError(f'计算副本输入/公式不同，拒绝拿其他人的缓存：{ws.title}!{left.coordinate}')
    replacements, carried = {}, []
    with ZipFile(base) as source_zip, ZipFile(calculated) as calculated_zip:
        originals, computed_parts = sheet_parts(source_zip), sheet_parts(calculated_zip)
        for name, part in originals.items():
            tree = ET.fromstring(source_zip.read(part))
            computed_tree = ET.fromstring(calculated_zip.read(computed_parts[name]))
            computed_cells = {c.get('r'): c for c in computed_tree.iter(NS+'c')}
            for cell in tree.iter(NS+'c'):
                if cell.find(NS+'f') is None:
                    continue
                coordinate = cell.get('r'); actual = computed_cells.get(coordinate)
                value = None if actual is None else actual.find(NS+'v')
                if actual is None or value is None or cache[name][coordinate].data_type == 'e':
                    raise ValueError(f'计算副本公式无缓存或错误：{name}!{coordinate}')
                for old in list(cell):
                    if old.tag in (NS+'v', NS+'is'):
                        cell.remove(old)
                kind = actual.get('t')
                if kind not in (None, 'n', 'b', 'str', 's'):
                    raise ValueError(f'不支持的公式缓存类型：{name}!{coordinate} {kind}')
                if value.text in (None, '') and kind != 'str':
                    raise ValueError(f'公式只有空v节点，没有实际缓存：{name}!{coordinate}')
                if kind == 's':
                    value = ET.Element(NS+'v'); value.text = str(cache[name][coordinate].value or '')
                    kind = 'str'
                if kind:
                    cell.set('t', kind)
                else:
                    cell.attrib.pop('t', None)
                cell.append(deepcopy(value)); carried.append(name+'!'+coordinate)
            replacements[part] = ET.tostring(tree, encoding='UTF-8', xml_declaration=True, standalone=True)
        output.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(output, 'w') as target:
            for info in source_zip.infolist():
                target.writestr(info, replacements.get(info.filename, source_zip.read(info.filename)))
    result = {'schema': 'payroll-template-cache/1', 'base': {'path':str(base),'sha256':sha256(base)},
              'calculated': {'path':str(calculated),'sha256':sha256(calculated)},
              'output': {'path':str(output),'sha256':sha256(output)}, 'cached_cells': carried,
              'native_recalculation': False, 'formal_ready': False,
              'status': 'cache_carried_pending_independent_calculation_audit'}
    if report is not None:
        write_json(report, result)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for key in ('base','calculated','output','report'):
        parser.add_argument('--'+key, type=Path, required=True)
    try:
        print(json.dumps(carry(**vars(parser.parse_args())), ensure_ascii=False, indent=2)); return 0
    except (ValueError, KeyError, OSError) as exc:
        print(json.dumps({'status':'rejected','error':str(exc)}, ensure_ascii=False)); return 2


if __name__ == '__main__':
    raise SystemExit(main())
