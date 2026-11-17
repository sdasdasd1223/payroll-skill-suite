"""指定模板填写入口；必须经 template_guard 登记与来源计划，不能只查表头。"""
import argparse
import json
from pathlib import Path
import re


def complete_total_formula(column, first_row, last_row):
    """完整核定才显示合计；空白或部分待核保持空白。"""
    from openpyxl.utils.cell import column_index_from_string
    column = str(column).upper()
    if not re.fullmatch(r'[A-Z]+', column) or column_index_from_string(column) > 16384:
        raise ValueError('合计列坐标无效')
    if isinstance(first_row, bool) or isinstance(last_row, bool) or not isinstance(first_row, int) or not isinstance(last_row, int):
        raise ValueError('合计行坐标必须为整数')
    if not 1 <= first_row <= last_row <= 1048576:
        raise ValueError('合计行范围无效')
    cells = f'{column}{first_row}:{column}{last_row}'
    return f'=IF(COUNT({cells})={last_row-first_row+1},SUM({cells}),"")'


def fill(template, plan, output, registry=None, registry_sha256=None, receipt=None, plan_path=None):
    from template_guard import build, load_registry, write_json
    if registry is None or registry_sha256 is None or receipt is None:
        raise ValueError('必须提供预先登记的registry、固定registry_sha256和新receipt；不得绕过模板保护')
    record, _ = load_registry(registry, registry_sha256)
    if Path(template).resolve() != Path(record['template']['path']).resolve():
        raise ValueError('填写入口模板不等于登记的用户指定原件')
    if 'cells' in plan or 'sheet' in plan:
        raise ValueError('旧版表头/姓名/source_id计划不足，请按template-guard.md改用writes与真实来源SHA')
    if plan_path is None:
        plan_path = Path(output).with_suffix('.guard-plan.json')
        write_json(plan_path, plan)
    return build(registry, registry_sha256, plan_path, output, receipt)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for key in ('template', 'plan', 'output', 'registry', 'receipt'):
        parser.add_argument('--' + key, type=Path, required=True)
    parser.add_argument('--registry-sha256', required=True)
    args = vars(parser.parse_args())
    try:
        plan_path = args['plan']
        args['plan'] = json.loads(plan_path.read_text(encoding='utf-8-sig'))
        result = fill(**args, plan_path=plan_path)
        print(json.dumps(result, ensure_ascii=False, indent=2)); return 0
    except (ValueError, KeyError, OSError) as exc:
        print(json.dumps({'status': 'rejected', 'error': str(exc)}, ensure_ascii=False)); return 2


if __name__ == '__main__':
    raise SystemExit(main())
