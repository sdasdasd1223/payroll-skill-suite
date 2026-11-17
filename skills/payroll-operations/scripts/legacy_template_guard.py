"""只读核对政府原始 .xls 模板，不转换格式、不制作工资、不冒称公式已复核。

按用户明确例外可允许原始黄色字段清底色；此例外不允许其他字体/边框/列宽漂移。
"""
import argparse
from copy import deepcopy
import json
from pathlib import Path
import struct
import sys

import xlrd
from template_guard import sha256, digest, write_json


def plain(obj):
    if isinstance(obj, (str, int, float, bool)) or obj is None:
        return obj
    if isinstance(obj, (tuple, list)):
        return [plain(v) for v in obj]
    if isinstance(obj, dict):
        return {str(k): plain(v) for k, v in obj.items()}
    return {k: plain(v) for k, v in vars(obj).items() if not k.startswith('_')}


def cell_style(book, sheet, row, col):
    xf = book.xf_list[sheet.cell_xf_index(row, col)]
    return {key: plain(getattr(xf, key)) for key in ('alignment', 'border', 'background', 'protection')} | {
        'font': plain(book.font_list[xf.font_index]),
        'number_format': book.format_map[xf.format_key].format_str}


def yellow(style, book):
    background = style['background']
    rgb = book.colour_map.get(background['pattern_colour_index'])
    return bool(background['fill_pattern'] and rgb and rgb[0] >= 200 and rgb[1] >= 170 and rgb[2] < 170)


def print_records(path):
    """保留 BIFF 打印/页边距/页眉页脚/分页记录，不用表头代替打印核对。"""
    from xlrd.compdoc import CompDoc
    document = CompDoc(Path(path).read_bytes())
    stream = document.get_named_stream('Workbook') or document.get_named_stream('Book')
    codes = {0xA1, 0x26, 0x27, 0x28, 0x29, 0x14, 0x15, 0x2B, 0x2A, 0x83, 0x84, 0x1A, 0x1B}
    records, offset = [], 0
    while offset + 4 <= len(stream):
        code, length = struct.unpack_from('<HH', stream, offset)
        data = stream[offset+4:offset+4+length]
        if code in codes:
            records.append([code, data.hex()])
        offset += 4 + length
    return records


def snapshot(book):
    result = {}
    for sheet in book.sheets():
        result[sheet.name] = {
            'shape': [sheet.nrows, sheet.ncols], 'visibility': sheet.visibility,
            'merges': plain(sheet.merged_cells),
            'row_dimensions': plain(sheet.rowinfo_map), 'column_dimensions': plain(sheet.colinfo_map),
            'default_col_width': sheet.defcolwidth, 'default_row_height': sheet.default_row_height,
            'cells': {f'{r+1},{c+1}': {'value': sheet.cell_value(r, c),
                'type': sheet.cell_type(r, c), 'style': cell_style(book, sheet, r, c)}
                for r in range(sheet.nrows) for c in range(sheet.ncols)}}
    return result


def row_settings(sheet,row):
    info=sheet.rowinfo_map.get(row)
    return {'height':info.height if info is not None and not info.has_default_height else sheet.default_row_height,
        'hidden':bool(info.hidden) if info else False,'outline_level':info.outline_level if info else 0}


def inspect(template, expected_sha256, report=None, output=None, clear_original_yellow=False,
            clear_all_fills=False, personnel_prototype_rows=None):
    template = Path(template).resolve()
    if template.suffix.lower() != '.xls' or sha256(template) != expected_sha256:
        raise ValueError('须为用户指定原始 .xls，且 SHA 与预先固定值一致')
    source = xlrd.open_workbook(template, formatting_info=True)
    original = snapshot(source)
    yellow_fields, yellow_cells = {}, {}
    for sheet in source.sheets():
        yellow_cells[sheet.name] = [f'{r+1},{c+1}' for r in range(sheet.nrows) for c in range(sheet.ncols)
                                    if yellow(cell_style(source, sheet, r, c), source)]
        yellow_fields[sheet.name] = [{'column': c+1, 'field': sheet.cell_value(0, c)}
            for c in range(sheet.ncols) if sheet.nrows and yellow(cell_style(source, sheet, 0, c), source)]
    errors, checks = [], []
    row_mapping=[]
    body=source.sheet_by_index(0)
    prototypes=list(range(2,body.nrows+1)) if personnel_prototype_rows is None else list(personnel_prototype_rows)
    if not prototypes or any(isinstance(r,bool) or not isinstance(r,int) or not 2<=r<=body.nrows for r in prototypes):
        raise ValueError('合法人员原型行须为原业务页第2行起的实际行，不得使用表头/说明页')
    if source.sheet_names() != ['正常工资薪金收入', '填表说明'] or source.sheet_by_index(0).ncols != 22:
        errors.append('本案例政府模板须为22列且保留正常工资薪金收入/填表说明两张表')
    if output is not None:
        output = Path(output).resolve()
        if output == template or output.suffix.lower() != '.xls':
            raise ValueError('输出须另存原始 .xls，不得覆盖模板或转换 .xlsx')
        book = xlrd.open_workbook(output, formatting_info=True)
        actual = snapshot(book)
        if book.sheet_names() != source.sheet_names():
            errors.append('政府模板工作表名称/说明页缺失或改变')
        else:
            original_notes, actual_notes = deepcopy(original['填表说明']), deepcopy(actual['填表说明'])
            if clear_all_fills:
                for note in (original_notes, actual_notes):
                    for cell in note['cells'].values():
                        cell['style'].pop('background')
                for cell in actual['填表说明']['cells'].values():
                    if cell['style']['background']['fill_pattern']:
                        errors.append('填表说明仍有底色，未完成全部清底色')
            if actual_notes != original_notes:
                errors.append('政府模板填表说明的文字、样式或结构改变')
            name = source.sheet_names()[0]
            src, dest = source.sheet_by_name(name), book.sheet_by_name(name)
            if dest.ncols != 22 or dest.row_values(0) != src.row_values(0):
                errors.append('政府模板22列或字段顺序改变')
            for key in ('column_dimensions', 'merges', 'visibility', 'default_col_width', 'default_row_height'):
                if original[name][key] != actual[name][key]:
                    errors.append(f'政府模板参数改变：{key}')
            yellow_columns = {item['column']-1 for item in yellow_fields[name]}
            for r in range(dest.nrows):
                def comparable_style(b,s,row,col):
                    value=cell_style(b,s,row,col)
                    if clear_all_fills or clear_original_yellow and col in yellow_columns:
                        value=deepcopy(value);value.pop('background')
                    return value
                candidates=[r] if r<src.nrows else [n-1 for n in prototypes]
                def row_score(sr):
                    return sum(comparable_style(source,src,sr,c)!=comparable_style(book,dest,r,c)
                               for c in range(min(dest.ncols,src.ncols)))
                exact=[sr for sr in candidates if row_score(sr)==0]
                height_exact=[sr for sr in exact if row_settings(src,sr)==row_settings(dest,r)]
                source_r=height_exact[0] if height_exact else exact[0] if exact else min(candidates,key=row_score)
                if row_settings(src,source_r)!=row_settings(dest,r):
                    errors.append(f'政府模板人员行高度/隐藏/分组与完整来源原型不符：{name}!行{r+1}，'
                        f'原型行{source_r+1}={row_settings(src,source_r)}，成品={row_settings(dest,r)}')
                row_mapping.append({'output_row':r+1,'template_row':source_r+1,
                    'added_personnel_row':r>=src.nrows,'template_row_settings':row_settings(src,source_r),
                    'actual_row_settings':row_settings(dest,r)})
                for c in range(min(dest.ncols, src.ncols)):
                    expected, current = cell_style(source, src, source_r, c), cell_style(book, dest, r, c)
                    allowed = clear_all_fills or clear_original_yellow and c in yellow_columns
                    if allowed:
                        if current['background']['fill_pattern'] != 0:
                            errors.append(f'核定清底色范围仍有填充：{name}!{r+1},{c+1}')
                        expected, current = deepcopy(expected), deepcopy(current)
                        expected.pop('background'); current.pop('background')
                    if expected != current:
                        errors.append(f'政府模板非核定样式变化：{name}!{r+1},{c+1}')
                    if len(errors) >= 50:
                        break
                if len(errors) >= 50:
                    break
            if print_records(template) != print_records(output):
                errors.append('政府模板打印/页边距/页眉页脚/分页记录改变')
            checks = ['原始22列与顺序', '完整说明页', '逐格字体/边框/数字格式/对齐/填充', '列尺寸/合并/打印BIFF记录',
                      '新增人员整行样式与有效行高绑定合法原型（排除表头说明）']
    result = {'schema': 'payroll-legacy-template-guard/1', 'passed': not errors, 'errors': errors,
              'template': {'path': str(template), 'sha256': expected_sha256},
              'template_fingerprint_sha256': digest(original), 'original_yellow_fields': yellow_fields,
              'original_yellow_cells': yellow_cells, 'clear_original_yellow_allowed': clear_original_yellow,
              'clear_all_fills_allowed': clear_all_fills,
              'personnel_prototype_rows':prototypes,'personnel_row_mapping':row_mapping,
              'checks': checks, 'native_recalculation': False, 'formal_ready': False,
              'status': 'legacy_structure_pending_native_formula_audit' if output else 'legacy_template_read_only_inspected',
              'limitations': ['xlrd仅取得公式缓存，不能证明公式保留或原生重算；须另用Excel只读公式核对',
                              '数据行金额和人员是否正确须由核定来源及跨环节审计检查，本工具不证明这些业务事实']}
    if output:
        result['output'] = {'path': str(output), 'sha256': sha256(output)}
        result['finalized_output_sha256'] = result['output']['sha256']
    if report is not None:
        write_json(report, result)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--template', type=Path, required=True)
    parser.add_argument('--expected-sha256', required=True)
    parser.add_argument('--report', type=Path, required=True)
    parser.add_argument('--output', type=Path)
    parser.add_argument('--clear-original-yellow', action='store_true')
    parser.add_argument('--clear-all-fills', action='store_true', help='先识别原始黄色必填字段，再检查所有工作表全部单元格无底色')
    parser.add_argument('--personnel-prototype-rows',type=lambda value:[int(x) for x in value.split(',')],
        help='预先核定的1起算人员原型行，如2,3；默认原业务页第2行起，排除表头/说明')
    try:
        result = inspect(**vars(parser.parse_args()))
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result['passed'] else 2
    except (ValueError, KeyError, OSError, xlrd.XLRDError) as exc:
        print(json.dumps({'status': 'rejected', 'error': str(exc)}, ensure_ascii=False)); return 2


if __name__ == '__main__':
    sys.exit(main())
