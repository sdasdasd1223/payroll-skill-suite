"""生成可执行的跨阶段记录；真实输入只计算能独立证实的考勤事实。
演示：python crossstage_scaffold.py --demo --output-dir 全新目录
真实：--scope 实际名单.json --month YYYY-MM [--attendance-config 映射.json] --output-dir 全新目录
"""
import argparse
import calendar
import json
from pathlib import Path
from crossstage_audit import audit_crossstage, sha256, REQUIRED, LEAVE_ITEMS, write_scope_review_csv

def save(path, value):
    if path.exists(): raise ValueError('不得覆盖既有输出:' + str(path))
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')

def add_table(record, directory, table_id, rows, role='calculation'):
    path = directory / (table_id + '.json'); save(path, rows)
    record['sources'][table_id] = {'path': str(path.resolve()), 'sha256': sha256(path), 'month': record['month'], 'role': role}
    columns = {k: k for row in rows for k in row}
    record['tables'][table_id] = {'source_id': table_id, 'format': 'json', 'start_row': 1, 'end_row': max(1, len(rows)), 'columns': columns}

def make_record(directory, month, policy_version):
    directory = Path(directory).resolve()
    if directory.exists(): raise ValueError('请指定尚不存在的新目录')
    directory.mkdir(parents=True)
    return directory, {'schema_version': 'crossstage-1', 'month': month, 'policy_version': policy_version, 'mode': 'draft', 'purpose': 'calculation_details', 'through_stage': 'attendance', 'scope_table': 'scope', 'sources': {}, 'tables': {}, 'stages': {}, 'facts': [], 'confirmations': []}

def finish(directory, record):
    save(directory / 'run.json', record)
    report = audit_crossstage(record, directory)
    write_scope_review_csv(report, directory / 'scope-review.csv')
    report['scope_review_csv'] = str(directory / 'scope-review.csv')
    save(directory / 'audit.json', report)
    return {'record': str(directory / 'run.json'), 'report': str(directory / 'audit.json'), 'passed': report['passed'], 'formal_ready': report['formal_ready'], 'purpose': report['purpose'], 'delivery_status': report['delivery_status'], 'delivery_message': report['delivery_message'], 'technical_delivery': report['technical_delivery'], 'errors': report['errors'], 'pending_count': len(report['pending'])}

def demo(output_dir):
    """全部为明确合成数据，无任何真实确认。"""
    directory, record = make_record(output_dir, '2026-09', 'synthetic-demonstration-only')
    person = {'employee_key': 'SYNTHETIC-1', 'name': '明确合成人员', 'company': '明确合成公司'}
    add_table(record, directory, 'scope', [person], 'scope_approved')
    add_table(record, directory, 'daily', [dict(person, date=f'2026-09-{d:02}', attendance_type='需要打卡', **{k: 0 for k in LEAVE_ITEMS}) for d in range(1, 31)], 'raw')
    record['raw_attendance'] = {'table': 'daily', 'leave_columns': {k: {'kind': k, 'unit': 'hours'} for k in LEAVE_ITEMS}}
    save(directory / 'attendance-config.json', {'path': str((directory / 'daily.json').resolve()), 'table': record['tables']['daily'], 'raw_attendance': record['raw_attendance']})
    add_table(record, directory, 'claims', [dict(person, date='2026-09-01', tier=50, next_workday=True, checkout='22:01', next_arrival='08:59')], 'raw')
    record['raw_rewards'] = {'table': 'claims'}
    insurance = dict(person, month='2026-09', medical_employer_mode='combined')
    for k, amount in zip(REQUIRED['insurance'], (100, 10, 20, 30)):
        insurance[k + '_employee'] = amount; insurance[k + '_employer'] = amount * 2
    add_table(record, directory, 'insurance_raw', [insurance], 'raw')
    record['raw_insurance'] = {'table': 'insurance_raw', 'components': {k: {'employee': k + '_employee', 'employer': k + '_employer'} for k in REQUIRED['insurance']}}
    add_table(record, directory, 'salary_inputs', [dict(person, salary=10000, salary_adjustment=0, progression=0, performance=0, manual_adjustments=0)], 'manual')
    record['through_stage'] = 'payslip'
    for stage, required_items in REQUIRED.items():
        items = required_items + ('finance', 'bank', 'third_party', 'cash_other') if stage == 'payroll' else required_items
        row = dict(person, **{item: 0 for item in items})
        if 'reward' in items: row['reward'] = 50
        if stage == 'insurance': row.update({k: insurance[k + '_employee'] for k in items})
        if stage == 'payroll': row.update(salary=10000, personal_social=110, personal_medical=20, personal_fund=30, employer_social=220, employer_medical=40, employer_fund=60, gross=10050, finance=10050, net=9890, bank=9890)
        if stage == 'tax': row.update(taxable=10050, personal_social=110, personal_medical=20, personal_fund=30)
        if stage == 'payslip': row.update(gross=10050, net=9890)
        add_table(record, directory, stage, [row])
        record['stages'][stage] = {'table': stage, 'dispositions': [{'employee_key': person['employee_key'], 'route': 'output', 'row': 1}]}
        for item in items:
            ref = {'table': stage, 'row': 1, 'column': item}
            fact = {'employee_key': person['employee_key'], 'stage': stage, 'item': item, 'status': 'calculated', 'amount': row[item], 'source_refs': [ref], 'value_refs': [ref]}
            if item in LEAVE_ITEMS.values(): fact.update(observed_rows=[], unit='hours')
            if stage == 'insurance': fact['employer_amount'] = insurance[item + '_employer']
            if stage == 'reward': fact['claim_rows'] = [1]
            if stage == 'payroll' and item in ('salary', 'salary_adjustment', 'progression', 'performance', 'manual_adjustments'):
                ref = {'table': 'salary_inputs', 'row': 1, 'column': item}
                fact['source_refs'] = [ref]; fact['derivation'] = {'operation': 'copy', 'refs': [ref]}
            record['facts'].append(fact)
    return finish(directory, record)

def scaffold(scope, month, output_dir, attendance_config=None, policy_version='待核定规则版本'):
    directory, record = make_record(output_dir, month, policy_version)
    scope = Path(scope).resolve()
    rows = json.loads(scope.read_text(encoding='utf-8-sig'))
    if not isinstance(rows, list) or not rows: raise ValueError('名单须为非空JSON对象数组')
    if any(not all(p.get(k) for k in ('employee_key', 'name', 'company')) for p in rows): raise ValueError('名单须提供employee_key/name/company')
    record['sources']['scope'] = {'path': str(scope), 'sha256': sha256(scope), 'month': month, 'role': 'scope_approved'}
    record['tables']['scope'] = {'source_id': 'scope', 'format': 'json', 'start_row': 1, 'end_row': len(rows), 'columns': {k: k for p in rows for k in p}}
    record['stages']['attendance'] = {'dispositions': [{'employee_key': p['employee_key'], 'route': 'pending', 'reason': '待逐项核算；不得仅因缺资料删除本人'} for p in rows]}
    if attendance_config:
        config_path = Path(attendance_config).resolve()
        config = json.loads(config_path.read_text(encoding='utf-8-sig'))
        source = Path(config['path'])
        if not source.is_absolute(): source = config_path.parent / source
        record['sources']['daily'] = {'path': str(source.resolve()), 'sha256': sha256(source), 'month': month, 'role': 'raw'}
        record['tables']['daily'] = dict(config['table'], source_id='daily')
        record['raw_attendance'] = dict(config['raw_attendance'], table='daily')
    # 先读取实际全范围观察；尚无事实时不会捏造金额，身份/映射错误仍保留错误。
    observed = audit_crossstage(record, directory).get('evidence', {}).get('raw_attendance', {})
    calculated = []
    for p in rows:
        obs = observed.get(p['employee_key'], {})
        output = dict(p, **{item: None for item in REQUIRED['attendance']})
        for item in REQUIRED['attendance']:
            fact = {'employee_key': p['employee_key'], 'stage': 'attendance', 'item': item, 'status': 'pending', 'amount': None, 'reason': '缺对应原始依据，其他可算项目继续'}
            kind = next((k for k, v in LEAVE_ITEMS.items() if v == item), None)
            if kind:
                leaves = obs.get('leave', {}).get(kind, [])
                fact.update(unit='hours', observed_rows=[{'row': x['row'], 'column': x['column']} for x in leaves])
                complete = obs.get('coverage_complete') and obs.get('types') in (['需要打卡'], ['不需要打卡']) and not any(x['row'] in obs.get('rows', []) for x in obs.get('unknown', []))
                if kind in ('personal', 'sick') and obs.get('types') == ['不需要打卡']: complete = False
                if complete:
                    amount = float(obs['business_hours'][kind]); output[item] = amount
                    first = obs['rows'][0]
                    fact.update(status='calculated', amount=amount, reason='从完整实际日明细独立计算', source_refs=[{'table': 'daily', 'row': x['row'], 'column': x['column']} for x in leaves] or [{'table': 'daily', 'row': first, 'column': 'date'}])
            record['facts'].append(fact)
        calculated.append(output)
    add_table(record, directory, 'attendance_calculation', calculated)
    record['stages']['attendance']['table'] = 'attendance_calculation'
    for n, disposition in enumerate(record['stages']['attendance']['dispositions'], 1): disposition['row'] = n
    for f in record['facts']:
        n = next(i for i, p in enumerate(rows, 1) if p['employee_key'] == f['employee_key'])
        f['value_refs'] = [{'table': 'attendance_calculation', 'row': n, 'column': f['item']}]
    return finish(directory, record)

def main():
    ap = argparse.ArgumentParser(); group = ap.add_mutually_exclusive_group(required=True)
    group.add_argument('--demo', action='store_true'); group.add_argument('--scope', type=Path)
    ap.add_argument('--month'); ap.add_argument('--output-dir', type=Path, required=True); ap.add_argument('--attendance-config', type=Path); ap.add_argument('--policy-version', default='待核定规则版本')
    a = ap.parse_args()
    if a.scope and not a.month: ap.error('真实名单模式必须指定--month')
    try: result = demo(a.output_dir) if a.demo else scaffold(a.scope, a.month, a.output_dir, a.attendance_config, a.policy_version)
    except (ValueError, KeyError, OSError) as exc: ap.error(str(exc))
    print(json.dumps(result, ensure_ascii=False)); raise SystemExit(0 if result['passed'] else 2)

if __name__ == '__main__': main()
