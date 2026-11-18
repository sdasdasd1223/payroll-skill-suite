"""按私有固定名单台账读取实际来源与产出；详细契约见run-contract.md。
草稿保留独立可算项目，正式输出要求前置完整。本程序不生成确认。
"""
import argparse
import csv
import hashlib
import importlib.util
import json
from collections import Counter
from datetime import date, datetime
import calendar
import re
from decimal import Decimal, InvalidOperation, ROUND_CEILING
from pathlib import Path

STAGES = ('attendance', 'reward', 'insurance', 'payroll', 'tax', 'payslip')
DONE = {'calculated', 'confirmed', 'confirmed_zero', 'not_applicable'}
REQUIRED = {
    'attendance': ('attendance', 'personal_leave', 'sick_leave', 'annual_leave', 'childcare_leave', 'maternity_leave', 'other_leave', 'overtime', 'reward'),
    'reward': ('reward',),
    'insurance': ('pension', 'unemployment', 'medical', 'housing_fund'),
    'payroll': ('salary', 'salary_adjustment', 'progression', 'performance', 'manual_adjustments', 'other_allowance', 'leave_deduction', 'attendance_deduction', 'hire_exit_difference', 'other_income', 'other_deduction', 'personal_leave', 'sick_leave', 'annual_leave', 'childcare_leave', 'maternity_leave', 'other_leave', 'overtime', 'reward', 'personal_social', 'personal_medical', 'personal_fund', 'employer_social', 'employer_medical', 'employer_fund', 'gross', 'net'),
    'tax': ('taxable', 'personal_social', 'personal_medical', 'personal_fund', 'combined_tax', 'tax'),
    'payslip': ('gross', 'net', 'tax'),
}
LEAVE_ITEMS = {'personal': 'personal_leave', 'sick': 'sick_leave', 'annual': 'annual_leave', 'childcare': 'childcare_leave', 'maternity': 'maternity_leave', 'other': 'other_leave'}

class MissingEvidence(ValueError):
    pass

def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def number(value):
    if isinstance(value, bool) or value is None or str(value).strip() == '':
        return None
    try:
        n = Decimal(str(value).replace(',', ''))
        return n if n.is_finite() else None
    except InvalidOperation:
        return None

def equal(a, b):
    x, y = number(a), number(b)
    return x == y if x is not None and y is not None else a == b

def _module(skill, filename):
    path = Path(__file__).resolve().parents[2] / skill / 'scripts' / filename
    spec = importlib.util.spec_from_file_location('_crossstage_' + path.stem, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module

class Audit:
    def __init__(self, record, base_dir):
        self.r, self.base = record, Path(base_dir)
        self.errors, self.pending, self.evidence = [], [], {}
        self.sources, self.tables, self.people, self.facts = {}, {}, {}, {}
        self.unresolved_performance_display = set()
        self.technical_errors, self.technical_pending = [], []

    def error(self, message):
        self.errors.append(message)

    def wait(self, message):
        self.pending.append(message)

    def tech_error(self, message):
        self.technical_errors.append(message); self.error(message)

    def tech_wait(self, message):
        self.technical_pending.append(message); self.wait(message)

    def required_delivery_stages(self):
        return [x for x in getattr(self,'stages',()) if x in ('attendance','insurance','payroll','tax','payslip')]

    def no_punch_evidence(self, key, first_day, last_day, normalize):
        refs=self.r.get('raw_attendance',{}).get('no_punch_refs',{}).get(key,[])
        for ref in refs:
            try:
                if normalize(self.ref(ref,self.people[key]))!='不需要打卡':raise ValueError('原记录并非免打卡')
                table=self.tables[ref['table']];row=table['rows'][ref['row']]
                if ref['column']!='attendance_type':raise ValueError('须定位实际免打卡字段')
                if row.get('month'):
                    if str(row['month'])[:7]!=self.r['month']:raise ValueError('不是本月确认')
                else:
                    start=str(row.get('effective_start') or '')[:10];end=str(row.get('effective_end') or '9999-12-31')[:10]
                    date.fromisoformat(start);date.fromisoformat(end)
                    if start>first_day or end<last_day:raise ValueError('长期名单有效期未覆盖本人本月在职范围')
                self.evidence.setdefault('no_punch',{})[key]={'source_ref':ref,'source_id':table['spec']['source_id'],'effective_first':first_day,'effective_last':last_day}
                return True
            except (KeyError,ValueError,TypeError) as exc:self.wait(key+':免打卡替代证据待核:'+str(exc))
        return False

    def load_sources(self):
        for key, spec in self.r.get('sources', {}).items():
            path = self.base / str(spec.get('path', ''))
            if not path.is_file():
                self.error(f'来源不存在:{key}'); continue
            digest = sha256(path)
            if digest != spec.get('sha256'):
                self.error(f'来源变化，旧确认失效:{key}'); continue
            if not spec.get('role') or not spec.get('month'):
                self.error(f'来源缺用途或归属期:{key}'); continue
            self.sources[key] = dict(spec, resolved=path, actual_sha256=digest)
        self.evidence['sources'] = {k: {'path': str(v['resolved']), 'sha256': v['actual_sha256']} for k, v in self.sources.items()}

    def load_tables(self):
        """Excel/CSV使用物理行号，JSON数组位置从1开始。
        从start_row扫描全部实际行，end_row仅作范围断言，不能用来隐藏新增人员。
        """
        for key, spec in self.r.get('tables', {}).items():
            source = self.sources.get(spec.get('source_id'))
            if source is None: continue
            try:
                fmt, path = spec.get('format'), source['resolved']
                if fmt == 'xlsx':
                    from openpyxl import load_workbook
                    book = load_workbook(path, read_only=True, data_only=True)
                    try:
                        ws = book[spec['sheet']]
                        # 一些有效导出声明只有A1，内部却有数千行，不能信任尺寸声明。
                        ws.reset_dimensions()
                        raw = [list(row) for row in ws.values]
                    finally: book.close()
                    from openpyxl.utils.cell import column_index_from_string
                    columns = {k: column_index_from_string(v) - 1 for k, v in spec['columns'].items()}
                    rows = {i + 1: {k: row[c] if c < len(row) else None for k, c in columns.items()} for i, row in enumerate(raw)}
                elif fmt == 'xls':
                    import xlrd
                    from openpyxl.utils.cell import column_index_from_string
                    book = xlrd.open_workbook(path)
                    ws = book.sheet_by_name(spec['sheet'])
                    columns = {k: column_index_from_string(v) - 1 for k, v in spec['columns'].items()}
                    def legacy_value(i, c):
                        if c >= ws.ncols: return None
                        value = ws.cell_value(i, c)
                        return xlrd.xldate_as_datetime(value, book.datemode) if ws.cell_type(i, c) == xlrd.XL_CELL_DATE else value
                    rows = {i+1: {k: legacy_value(i, c) for k, c in columns.items()} for i in range(ws.nrows)}
                elif fmt == 'csv':
                    with path.open(encoding='utf-8-sig', newline='') as f: raw = list(csv.reader(f))
                    columns = {k: int(v) - 1 for k, v in spec['columns'].items()}
                    rows = {i + 1: {k: row[c] if c < len(row) else None for k, c in columns.items()} for i, row in enumerate(raw)}
                elif fmt == 'json':
                    raw = json.loads(path.read_text(encoding='utf-8-sig'))
                    for part in spec.get('list_path', []): raw = raw[part]
                    if not isinstance(raw, list): raise ValueError('list_path不是数组')
                    rows = {i + 1: {k: row.get(v) for k, v in spec['columns'].items()} for i, row in enumerate(raw)}
                else: raise ValueError('仅支持显式xlsx/xls/csv/json表')
                start, end = spec.get('start_row'), spec.get('end_row')
                if isinstance(start, bool) or not isinstance(start, int) or not isinstance(end, int) or not 1 <= start <= end:
                    raise ValueError('缺有效完整扫描范围')
                excluded = spec.get('excluded_rows', {})
                for n, reason in excluded.items():
                    if not reason or not str(n).isdigit(): raise ValueError('排除行必须逐行说明依据')
                constants = spec.get('constants', {})
                if set(constants) & set(spec['columns']): raise ValueError('常量不得覆盖实际字段')
                for row in rows.values(): row.update(constants)
                active = {n: row for n, row in rows.items() if n >= start and any(row.get(c) not in (None, '') for c in spec['columns']) and str(n) not in excluded}
                if any(n > end for n in active): self.error(f'声明范围之后仍有实际数据:{key}')
                if not active and not spec.get('empty_evidence_refs'):
                    self.wait(f'实际明细为空，不能证明零或覆盖:{key}')
                self.tables[key] = dict(spec=spec, rows=rows, active=active)
            except (ValueError, KeyError, TypeError, OSError) as exc:
                self.error(f'实际表无法读取:{key}:{exc}')

    def ref(self, ref, person=None):
        if not isinstance(ref, dict): raise ValueError('引用不是对象')
        table = self.tables.get(ref.get('table'))
        if table is None: raise ValueError('引用无有效实际表')
        row, column = ref.get('row'), ref.get('column')
        if row not in table['active'] or column not in table['spec']['columns']:
            raise ValueError('引用不在完整数据行或声明字段中')
        values = table['rows'][row]
        if person is not None and not self.matches(values, person): raise ValueError('引用行身份不符')
        return values[column]

    def refs(self, refs, label, person=None):
        if not isinstance(refs, list) or not refs:
            self.wait(label + ':缺实际来源定位'); return False
        valid = True
        for ref in refs:
            try: self.ref(ref, person if ref.get('person_bound', True) else None)
            except (ValueError, KeyError, TypeError) as exc: self.error(label + ':' + str(exc)); valid = False
        return valid

    def matches(self, row, person):
        # 有可靠主键时必须一致；无主键时只接受公司+唯一姓名的弱匹配。
        if row.get('employee_key') not in (None, ''):
            return str(row['employee_key']) == person['employee_key'] and all(row.get(k) in (None, '', person.get(k)) for k in ('name', 'company'))
        return row.get('name') == person.get('name') and row.get('company') == person.get('company')

    def owner(self, row):
        owners = [k for k, p in self.people.items() if self.matches(row, p)]
        return owners[0] if len(owners) == 1 else None

    def scope(self):
        table = self.tables.get(self.r.get('scope_table'))
        if not table or not table['active']:
            self.error('缺实际核定人员名单'); return
        for n, row in table['active'].items():
            key = str(row.get('employee_key') or '')
            if not key or not row.get('name') or not row.get('company') or key in self.people:
                self.error(f'核定名单缺身份或重复:行{n}'); continue
            self.people[key] = dict(row, employee_key=key)
        names = [(p['company'], p['name']) for p in self.people.values()]
        self.ambiguous_names = {n for n, count in Counter(names).items() if count > 1}
        expected = self.r.get('expected_employee_keys')
        if expected is not None and (len(expected) != len(self.people) or set(expected) != set(self.people)):
            self.error('自报名单与实际核定名单不同')
        self.evidence['employee_keys'] = list(self.people)

    def stage_rows(self):
        through = self.r.get('through_stage')
        if through not in STAGES:
            self.error('缺有效through_stage'); return
        self.stages = STAGES[:STAGES.index(through) + 1]
        for stage in self.stages:
            spec = self.r.get('stages', {}).get(stage)
            if not isinstance(spec, dict): self.error('阶段缺唯一人员去向:' + stage); continue
            dispositions = spec.get('dispositions', [])
            keys = [d.get('employee_key') for d in dispositions]
            if len(keys) != len(set(keys)) or set(keys) != set(self.people):
                self.error('阶段人员范围缺漏/重复/越界:' + stage)
            table = self.tables.get(spec.get('table'))
            located = []
            for d in dispositions:
                key, route = d.get('employee_key'), d.get('route')
                label = str(key) + ':' + stage
                if route not in ('output', 'pending', 'excluded'): self.error(label + ':去向无效'); continue
                if route in ('pending', 'excluded') and not d.get('reason'): self.error(label + ':去向缺原因')
                if route == 'pending': self.wait(label + ':人员待核')
                if route == 'excluded':
                    self.refs(d.get('source_refs'), label + ':排除依据', self.people.get(key))
                    if d.get('row') is not None: self.error(label + ':排除人员仍有输出行')
                elif d.get('row') is not None:
                    if not table or d['row'] not in table['active'] or key not in self.people or not self.matches(table['rows'][d['row']], self.people[key]):
                        self.error(label + ':实际输出行或身份不符')
                    located.append(d['row'])
                elif route == 'output': self.error(label + ':缺实际输出行')
                if stage in ('tax', 'payslip') and route == 'output':
                    earlier = self.r.get('stages', {}).get('payroll', {}).get('dispositions', [])
                    if not any(x.get('employee_key') == key and x.get('route') == 'output' for x in earlier): self.error(label + ':工资未核人员进入输出')
            if len(located) != len(set(located)): self.error(stage + ':同一实际输出行映射多人')
            if table and set(located) != set(table['active']): self.error(stage + ':实际输出有未映射/多余人员行')
            if not table and any(d.get('route') == 'output' for d in dispositions): self.error(stage + ':无实际输出文件')
            self.evidence[stage] = {'source_id': table['spec']['source_id'] if table else None, 'rows': located}

    def check_facts(self):
        for fact in self.r.get('facts', []):
            key = (fact.get('employee_key'), fact.get('stage'), fact.get('item'))
            label = ':'.join(str(x) for x in key)
            if key in self.facts: self.error(label + ':事实重复'); continue
            self.facts[key] = fact
            person = self.people.get(key[0])
            if person is None or key[1] not in STAGES: self.error(label + ':事实身份或阶段无效'); continue
            status, amount = fact.get('status'), fact.get('amount')
            performance_display = key[1:] == ('payroll', 'performance') and fact.get('display_exception', {}).get('kind') == 'confirmed_zero_performance'
            display_render = self.performance_zero_display(fact, person) if performance_display else None
            if status == 'pending':
                self.wait(label + ':' + str(fact.get('reason', '缺资料')))
                if amount is not None: self.error(label + ':待核不得补零或保留可发金额')
            elif status not in DONE: self.error(label + ':未知状态')
            else:
                if not performance_display: self.refs(fact.get('source_refs'), label, person)
                if status == 'not_applicable' and not fact.get('reason'): self.error(label + ':不适用缺依据')
                if status == 'confirmed_zero' and number(amount) != 0: self.error(label + ':确认零与金额冲突')
                documented_nonrecording = status == 'not_applicable' and key[2] in LEAVE_ITEMS.values() and fact.get('treatment') in ('not_recorded_no_punch', 'no_auto_deduction', 'policy_confirmed') and fact.get('policy_refs')
                if key[1] in ('reward', 'insurance', 'payroll', 'tax', 'payslip') and key[2] != 'overtime' and number(amount) is None and not documented_nonrecording:
                    self.wait(label + ':完成项目仍缺明确金额，不能证明合计完整')
                if fact.get('policy_refs') and not performance_display: self.refs(fact['policy_refs'], label + ':适用口径', person)
                if key[1] == 'payroll' and key[2] in ('salary', 'salary_adjustment', 'progression', 'performance', 'manual_adjustments') and status != 'not_applicable' and not performance_display:
                    derivation = fact.get('derivation', {})
                    refs = derivation.get('refs', [])
                    if derivation.get('operation') not in ('copy', 'sum', 'negate') or not refs:
                        self.wait(label + ':缺实际核定金额或逐笔调整金额推导')
                    elif self.refs(refs, label + ':金额依据', person):
                        if any(ref.get('table') == self.r.get('stages', {}).get('payroll', {}).get('table') for ref in refs): self.error(label + ':工资输出不能自证本人原始核定金额')
                        values = [number(self.ref(ref, person if ref.get('person_bound', True) else None)) for ref in refs]
                        if any(v is None for v in values): self.error(label + ':源金额为空或非数字')
                        else:
                            expected = sum(values) if derivation['operation'] == 'sum' else (-values[0] if derivation['operation'] == 'negate' else values[0])
                            if number(amount) != expected: self.error(label + ':与实际核定金额不符')
            outputs = fact.get('value_refs', [])
            if status in DONE and not outputs: self.wait(label + ':缺实际计算或输出金额定位')
            for ref in outputs:
                try:
                    value = self.ref(ref, person)
                    primary_performance_display = performance_display and ref.get('table') == self.r.get('stages', {}).get('payroll', {}).get('table') and ref.get('column') == 'performance'
                    if status == 'pending' and value not in (None, ''): self.error(label + ':实际输出将待核变为值')
                    elif primary_performance_display and value in (None, ''):
                        if display_render == 'zero': self.error(label + ':明确要求显示0，但实际主表仍空白')
                        # 未证实的显示例外已逐项待核，不能靠空格推断可计入的0。
                        elif display_render == 'blank': self.evidence['performance_display'][key[0]]['actual_display'] = 'blank'
                    elif status != 'pending' and not equal(value, amount): self.error(label + ':实际金额与事实台账不一致')
                    elif primary_performance_display and display_render == 'blank': self.error(label + ':本次政策要求绩效留白，不能自动填0')
                except (ValueError, KeyError, TypeError) as exc: self.error(label + ':' + str(exc))
            stage_spec = self.r.get('stages', {}).get(key[1], {})
            own = [d for d in stage_spec.get('dispositions', []) if d.get('employee_key') == key[0] and d.get('row') is not None]
            table = self.tables.get(stage_spec.get('table'))
            if own and table and key[2] in table['spec']['columns']:
                expected_ref = {'table': stage_spec['table'], 'row': own[0]['row'], 'column': key[2]}
                if not any(all(r.get(k) == v for k, v in expected_ref.items()) for r in outputs): self.error(label + ':漏绑定实际阶段金额列')
        for stage in getattr(self, 'stages', ()):
            for person in self.people:
                route = next((d.get('route') for d in self.r.get('stages', {}).get(stage, {}).get('dispositions', []) if d.get('employee_key') == person), None)
                if route == 'excluded': continue
                for item in REQUIRED[stage]:
                    if (person, stage, item) not in self.facts: self.wait(f'{person}:{stage}:{item}:缺项')
        # 固定交接关系不能通过删掉自报依赖列表而绕开。
        links = [('reward', 'reward', 'attendance', 'reward'), ('attendance', 'reward', 'payroll', 'reward'), ('attendance', 'overtime', 'payroll', 'overtime'), ('insurance', 'medical', 'payroll', 'personal_medical'), ('insurance', 'housing_fund', 'payroll', 'personal_fund'), ('payroll', 'personal_social', 'tax', 'personal_social'), ('payroll', 'personal_medical', 'tax', 'personal_medical'), ('payroll', 'personal_fund', 'tax', 'personal_fund'), ('payroll', 'gross', 'payslip', 'gross'), ('payroll', 'net', 'payslip', 'net'), ('tax', 'tax', 'payslip', 'tax')]
        # 假别的观察量和工资适用性分别在原始事实检查中核验，不把时数/金额混为同一项。
        for person in self.people:
            for a, item, b, dest in links:
                x, y = self.facts.get((person, a, item)), self.facts.get((person, b, dest))
                if not x or not y: continue
                if x.get('status') == 'pending' and y.get('status') != 'pending': self.error(f'{person}:{a}.{item}待核在{b}丢失')
                elif not equal(x.get('amount'), y.get('amount')): self.error(f'{person}:{a}.{item}至{b}.{dest}金额丢失或改变')
            self.check_social_sum(person)
            # 沿用payroll_core既有应发桶；调薪、递增和人工补扣是组成说明，不能再相加。
            for item in ('salary_adjustment', 'progression', 'manual_adjustments'):
                adjustment = self.facts.get((person, 'payroll', item))
                if adjustment and adjustment.get('status') in DONE and number(adjustment.get('amount')) not in (None, 0):
                    if adjustment.get('included_in') not in ('salary', 'other_allowance', 'performance', 'reward', 'leave_deduction', 'attendance_deduction', 'hire_exit_difference', 'other_income', 'other_deduction'):
                        self.wait(person + ':' + item + ':缺已包含的实际工资桶，禁止独立重复相加')
            self.check_payroll_totals(person)
        for key, fact in self.facts.items():
            deps = fact.get('depends_on', [])
            items = [self.facts.get((key[0], d.get('stage'), d.get('item'))) for d in deps]
            if deps and any(not self.fact_complete(x) for x in items):
                if fact.get('status') != 'pending': self.error(':'.join(key) + ':依赖待核却已完成')
            if fact.get('operation') in ('copy', 'sum', 'negate') and items and all(self.fact_complete(x) and number(x.get('amount')) is not None for x in items):
                values = [number(x['amount']) for x in items]
                expected = sum(values) if fact['operation'] == 'sum' else (-values[0] if fact['operation'] == 'negate' else values[0])
                if number(fact.get('amount')) != expected: self.error(':'.join(key) + ':依赖金额计算不符')

    def fact_complete(self, fact):
        return bool(fact and fact.get('status') in DONE and (fact.get('employee_key'), fact.get('stage'), fact.get('item')) not in self.unresolved_performance_display)

    def performance_zero_display(self, fact, person):
        """只复用本人当月原有无绩效确认；规范化0不等于用户重新确认了新JSON。"""
        key = (fact['employee_key'], fact['stage'], fact['item'])
        label = ':'.join(key)
        rule = fact['display_exception']
        def pending(reason):
            self.unresolved_performance_display.add(key)
            self.wait(label + ':绩效零值显示例外未核实:' + reason)
            self.evidence.setdefault('performance_display', {})[key[0]] = {'effective_status': 'pending', 'reason': reason}
            return None
        if fact.get('status') != 'confirmed_zero' or number(fact.get('amount')) != 0:
            return pending('仅支持confirmed_zero及明确数值0，不能把缺资料当0')
        if rule.get('render') not in ('blank', 'zero'): return pending('缺明确留白或显示0的政策选择')
        source_id = rule.get('confirmation_source_id'); source = self.sources.get(source_id)
        if not source or source.get('month') != self.r.get('month'): return pending('缺本核算月实际原确认来源')
        if source.get('role') not in ('performance_confirmation', 'manual_approval'): return pending('原来源未声明本人绩效确认用途')
        confirmations = [c for c in self.r.get('confirmations', []) if c.get('source_id') == source_id]
        if len(confirmations) != 1: return pending('缺唯一原确认留痕，不能由程序或AI新造确认')
        confirmation = confirmations[0]
        if not confirmation.get('confirmed_by') or not confirmation.get('confirmed_at') or confirmation.get('sha256') != source['sha256']:
            return pending('原确认未绑定实际文件内容、确认人及时间')
        if confirmation.get('employee_key', key[0]) != key[0] or confirmation.get('month', self.r['month']) != self.r['month']:
            return pending('原确认属于他人或其他月份')
        dependencies = confirmation.get('dependencies', {})
        if not dependencies or any(k not in self.sources or self.sources[k]['sha256'] != digest for k, digest in dependencies.items()):
            return pending('原确认的实际依赖缺失或已变化')
        scope_source = self.r.get('tables', {}).get(self.r.get('scope_table'), {}).get('source_id')
        if scope_source not in dependencies: return pending('原确认未绑定实际本人范围来源')
        amount_ref, month_ref = rule.get('amount_ref'), rule.get('month_ref')
        if not amount_ref or not month_ref or (amount_ref.get('table'), amount_ref.get('row')) != (month_ref.get('table'), month_ref.get('row')):
            return pending('数值0和核算月须来自同一本人规范化记录行')
        try:
            if number(self.ref(amount_ref, person)) != 0 or str(self.ref(month_ref, person))[:7] != self.r['month']:
                return pending('实际本人记录不是本月确认无绩效0')
            if not fact.get('source_refs') or not fact.get('policy_refs'): return pending('缺实际金额依据或既有显示政策引用')
            for ref in fact['source_refs']: self.ref(ref, person if ref.get('person_bound', True) else None)
            for ref in fact['policy_refs']: self.ref(ref, person if ref.get('person_bound', True) else None)
            normalized_source = self.r['tables'][amount_ref['table']]['source_id']
            if normalized_source != source_id:
                evidence = rule.get('original_evidence', {})
                if evidence.get('source_id') != source_id: return pending('规范化0未指回原聊天或文件确认')
                quote = evidence.get('quote')
                if not isinstance(quote, str) or not quote: return pending('缺原确认文字定位')
                if evidence.get('ref'):
                    raw_ref = evidence['ref']
                    if self.r.get('tables', {}).get(raw_ref.get('table'), {}).get('source_id') != source_id: return pending('原文定位指向其他来源')
                    text = str(self.ref(raw_ref))
                else:
                    line = evidence.get('line')
                    lines = source['resolved'].read_text(encoding='utf-8-sig').splitlines()
                    if isinstance(line, bool) or not isinstance(line, int) or not 1 <= line <= len(lines): return pending('原确认行定位无效')
                    text = lines[line-1]
                if quote not in text: return pending('规范化原文与实际原确认内容不符')
                if not any(token in quote for token in ('无绩效', '無績效', '没有绩效', '沒有績效', 'no_performance')) or any(token in quote for token in ('待核', '未确认', '未確認', '尚未', '待提供', '未提供', '?', '？')):
                    return pending('原文未明确确认无绩效，不能把尚未提供规范化成0')
                identifier = re.escape(str(person['employee_key']))
                named = bool(re.search(r'(?<![A-Za-z0-9_-])' + identifier + r'(?![A-Za-z0-9_-])', quote))
                name = str(person.get('name') or '')
                if name and re.search(re.escape(name) + r'(?=\s|[:：,，。/、]|本月|当月|當月)', quote): named = True
                collective = set(confirmation.get('employee_keys', [])) == set(self.people) and any(token in quote for token in ('全部', '所有', '全员', '全員', '这些人员', '以上人员'))
                if not named and not collective: return pending('原确认文字未定位本人或已核定整份人员范围，不能套用他人确认')
                if (confirmation.get('employee_key') != key[0] and not collective) or confirmation.get('month') != self.r['month']:
                    return pending('规范化原确认缺本人及核算月上下文')
            if rule['render'] == 'zero':
                ref = rule.get('zero_required_ref')
                required = self.ref(ref, person if ref.get('person_bound', True) else None) if ref else None
                if not ref or isinstance(required, bool) or required not in (0, '0', 'zero_required', '填0', '要求填0'):
                    return pending('模板或人资没有明确要求显示0，不能自动改主表')
        except (ValueError, KeyError, TypeError, OSError, UnicodeError) as exc:
            return pending('实际本人确认/政策引用无效:' + str(exc))
        self.evidence.setdefault('performance_display', {})[key[0]] = {'effective_status': 'confirmed_zero', 'calculation_amount': 0, 'render': rule['render'], 'confirmation_source_id': source_id, 'confirmation_sha256': source['sha256'], 'month': self.r['month'], 'amount_ref': amount_ref, 'month_ref': month_ref, 'original_evidence': rule.get('original_evidence'), 'note': '主表展示与辅助计算分开；未修改主表或原确认'}
        return rule['render']

    def check_social_sum(self, person):
        parts = [self.facts.get((person, 'insurance', item)) for item in ('pension', 'unemployment')]
        out = self.facts.get((person, 'payroll', 'personal_social'))
        if out and all(parts):
            if any(p.get('status') == 'pending' for p in parts):
                if out.get('status') != 'pending': self.error(person + ':保险缺项在工资合并时丢失')
            elif all(number(p.get('amount')) is not None for p in parts) and number(out.get('amount')) != sum(number(p['amount']) for p in parts): self.error(person + ':个人养老失业合并金额不符')

    def check_payroll_totals(self, person):
        """复用既有工资核心；辅助已知小计不写入任何正式合计事实。"""
        if not any(k[0] == person and k[1] == 'payroll' for k in self.facts): return
        mapping = {k: ('payroll', v) for k, v in {'fixed_salary': 'salary', 'other_allowance': 'other_allowance', 'performance': 'performance', 'reward': 'reward', 'leave_deduction': 'leave_deduction', 'attendance_deduction': 'attendance_deduction', 'hire_exit_difference': 'hire_exit_difference', 'other': 'other_income', 'personal_deduction': 'other_deduction', 'personal_social': 'personal_social', 'personal_medical': 'personal_medical', 'personal_fund': 'personal_fund', 'third_party': 'third_party', 'cash_other': 'cash_other'}.items()}
        mapping.update(tax=('tax', 'tax'), combined_tax=('tax', 'combined_tax'))
        known = {}
        for key, (stage, item) in mapping.items():
            f = self.facts.get((person, stage, item))
            if self.fact_complete(f) and number(f.get('amount')) is not None: known[key] = number(f['amount'])
        gross_deps = ('fixed_salary', 'other_allowance', 'performance', 'reward', 'leave_deduction', 'attendance_deduction', 'hire_exit_difference')
        finance_deps = gross_deps + ('other',)
        net_deps = finance_deps + ('personal_deduction', 'personal_social', 'personal_medical', 'personal_fund', 'tax')
        dependencies = {'gross': gross_deps, 'finance': finance_deps, 'net': net_deps, 'taxable': finance_deps + ('personal_deduction', 'combined_tax'), 'bank': net_deps + ('third_party', 'cash_other')}
        summary = {'label': '仅已知应发项目小计，不是正式应发/财务/实发', 'known_gross_subtotal': str(sum(known[k] for k in gross_deps if k in known)), 'unknown_gross_items': [k for k in gross_deps if k not in known]}
        self.evidence.setdefault('known_subtotals', {})[person] = summary
        payroll_totals = _module('payroll-operations', 'payroll_core.py').payroll_totals
        for output, deps in dependencies.items():
            f = self.facts.get((person, 'tax' if output == 'taxable' else 'payroll', output))
            if not f: continue
            missing = [k for k in deps if k not in known]
            if missing:
                if f.get('status') != 'pending': self.error(person + ':' + output + ':实际工资桶缺项/待核却给出正式合计:' + ','.join(missing))
                continue
            # 仅未参与本项输出的参数给中性值，以调用同一个核心函数。
            # 缺失且会影响本项的输入已在上方阻止；不把中性值写回事实、来源或输出。
            inputs = {k: known[k] if k in known else Decimal(0) for k in mapping}
            expected = payroll_totals(inputs)[output]
            if f.get('status') in DONE and not equal(f.get('amount'), expected): self.error(person + ':' + output + ':与payroll_core既有工资桶计算不符')

    def raw_attendance(self):
        spec = self.r.get('raw_attendance')
        if not spec:
            self.wait('缺原始全部假别与考勤类型覆盖资料'); return
        table = self.tables.get(spec.get('table'))
        if not table: self.wait('原始考勤表不可用'); return
        columns = table['spec']['columns']
        needed = {'name', 'date', 'attendance_type'}
        leave_cols = spec.get('leave_columns', {})
        if not needed <= set(columns) or not leave_cols or any(not isinstance(v, dict) or v.get('kind') not in LEAVE_ITEMS or v.get('unit') not in ('minutes', 'hours', 'days') for v in leave_cols.values()) or not set(leave_cols) <= set(columns):
            self.error('原始考勤映射缺日期/假别/类型或假别口径未知'); return
        # 实际表头中全部含“假”的列都要分类；其他假别不能静默视为工资不适用。
        if table['spec']['format'] == 'xlsx':
            from openpyxl import load_workbook
            book = load_workbook(self.sources[table['spec']['source_id']]['resolved'], read_only=True, data_only=True)
            try:
                ws = book[table['spec']['sheet']]
                ws.reset_dimensions()
                header = spec.get('header_row')
                if not isinstance(header, int) or not 1 <= header < table['spec']['start_row']: raise ValueError('原始考勤缺有效表头行')
                mapped = {table['spec']['columns'][c] for c in leave_cols}
                for cell in ws[header]:
                    if '假' in str(cell.value or '') and cell.column_letter not in mapped:
                        self.error('原始假别列未完整映射:' + str(cell.value))
            finally: book.close()
        normalize = _module('attendance-calculator', 'process_attendance.py').normalize_attendance_type
        if spec.get('blank_means_zero') is True: self.refs(spec.get('blank_zero_refs'), '原始考勤空白表示零的实际依据')
        observations = {p: {'rows': [], 'raw_rows': [], 'dates': [], 'types': set(), 'raw_types': set(), 'excluded_employment_rows': [], 'leave': {k: [] for k in LEAVE_ITEMS}, 'raw_hours': {k: Decimal(0) for k in LEAVE_ITEMS}, 'business_hours': {k: Decimal(0) for k in LEAVE_ITEMS}, 'unknown': []} for p in self.people}
        unresolved = []
        for n, row in table['active'].items():
            d = row.get('date')
            text = d.strftime('%Y-%m') if isinstance(d, (date, datetime)) else str(d or '')[:7]
            if text != self.r.get('month'):
                if not text: self.error(f'原始考勤日期缺失:行{n}')
                continue
            owner = self.owner(row)
            if not owner:
                if any(d.get('name') == str(row.get('name') or '').strip() and d.get('route') == 'excluded' for d in self.evidence.get('scope_review', [])): continue
                unresolved.append(n); continue
            stamp = d.strftime('%Y-%m-%d') if isinstance(d, (date, datetime)) else str(d)[:10]
            day = date.fromisoformat(stamp)
            person_scope = self.people[owner]
            hire = date.fromisoformat(str(person_scope['hire_date'])[:10]) if person_scope.get('hire_date') else None
            departure = date.fromisoformat(str(person_scope['departure_date'])[:10]) if person_scope.get('departure_date') else None
            in_employment = (hire is None or day >= hire) and (departure is None or day <= departure)
            obs = observations[owner]; obs['raw_rows'].append(n); obs['raw_types'].add(normalize(row.get('attendance_type')))
            if in_employment:
                obs['rows'].append(n); obs['dates'].append(stamp); obs['types'].add(normalize(row.get('attendance_type')))
            else:
                obs['excluded_employment_rows'].append({'row': n, 'date': stamp, 'reason': '入职前' if hire and day < hire else '离职后'})
            daily = {}
            for column, rule in leave_cols.items():
                kind, unit = rule['kind'], rule['unit']
                value = row.get(column)
                if value in (0, '0'): continue
                if value in (None, '') and spec.get('blank_means_zero') is True:
                    if not spec.get('blank_zero_refs'): self.wait('原始空白当零缺口径依据')
                    continue
                obs['leave'][kind].append({'row': n, 'column': column, 'raw': str(value)})
                numeric = number(value)
                if numeric is None or numeric < 0:
                    obs['unknown'].append({'row': n, 'column': column}); continue
                hours = numeric / 60 if unit == 'minutes' else (numeric * 8 if unit == 'days' else numeric)
                daily[kind] = daily.get(kind, Decimal(0)) + hours
            for kind, hours in daily.items():
                obs['raw_hours'][kind] += hours
                # 仅事病假按现有制度自然日向上取整至1小时并封顶8小时。
                # 其他假别保留原始单位换算，不在跨阶段检查中另创舍入规则。
                if in_employment: obs['business_hours'][kind] += min(Decimal(8), hours.to_integral_value(rounding=ROUND_CEILING)) if kind in ('personal', 'sick') else hours
        if unresolved: self.wait('原始考勤存在范围或身份待核人员，未自动纳入工资:' + ','.join(map(str, unresolved)))
        for person, obs in observations.items():
            year, month = map(int, self.r['month'].split('-'))
            expected_dates = {date(year, month, day).isoformat() for day in range(1, calendar.monthrange(year, month)[1] + 1)}
            # 入离职日期只能取自实际核定花名册，用于缩小日明细覆盖范围。
            hire, exit_date = self.people[person].get('hire_date'), self.people[person].get('departure_date')
            if hire: expected_dates = {d for d in expected_dates if d >= str(hire)[:10]}
            if exit_date: expected_dates = {d for d in expected_dates if d <= str(exit_date)[:10]}
            coverage = bool(expected_dates) and set(obs['dates']) == expected_dates
            if len(obs['dates']) != len(set(obs['dates'])): self.error(person + ':原始日明细同日重复，须先核实')
            verified_no_punch = bool(expected_dates) and self.no_punch_evidence(person,min(expected_dates),max(expected_dates),normalize)
            exempt = (coverage and obs['types'] == {'不需要打卡'}) or (verified_no_punch and obs['types'] <= {'不需要打卡'})
            if exempt:
                # 原始观察保留；制度不记录的月度项不是“计入0小时”。
                obs['business_hours']['personal'] = None
                obs['business_hours']['sick'] = None
            if not exempt and (not coverage or any(t not in ('需要打卡', '不需要打卡') for t in obs['types']) or len(obs['types']) != 1):
                self.wait(person + ':原始日明细缺失/类型未知或混合，不得自动豁免')
                fact = self.facts.get((person, 'attendance', 'personal_leave'))
                if fact and fact.get('status') == 'not_applicable': self.error(person + ':无完整本人原始明细却豁免事假')
            for kind, item in LEAVE_ITEMS.items():
                observed = obs['leave'][kind]
                applicable_observed = [x for x in observed if x['row'] in obs['rows']]
                a = self.facts.get((person, 'attendance', item))
                p = self.facts.get((person, 'payroll', item))
                allowed_payroll_nonrecording = False
                if p and p.get('status') == 'not_applicable' and p.get('policy_refs'):
                    treatment = p.get('treatment')
                    if treatment == 'not_recorded_no_punch' and exempt and kind in ('personal', 'sick'):
                        allowed_payroll_nonrecording = True
                    elif treatment == 'no_auto_deduction' and kind == 'personal':
                        employee_type = self.people[person].get('employee_type')
                        departure = str(self.people[person].get('departure_date') or '')[:7]
                        allowed_payroll_nonrecording = employee_type in ('全职', '全职员工', '正式员工', '试用期员工', '正式', '试用') and 'departure_date' in self.people[person] and departure != self.r['month']
                    elif treatment == 'policy_confirmed':
                        allowed_payroll_nonrecording = True
                if a and p and a.get('status') == 'pending' and p.get('status') != 'pending' and not allowed_payroll_nonrecording:
                    self.error(f'{person}:attendance.{item}待核在payroll丢失')
                for stage in ('attendance', 'payroll'):
                    fact = self.facts.get((person, stage, item))
                    if not fact: continue
                    reported = fact.get('observed_rows')
                    actual = sorted((x['row'], x['column']) for x in observed)
                    if reported is None or sorted((x.get('row'), x.get('column')) for x in reported) != actual:
                        self.error(f'{person}:{stage}:{item}:原始全部假别观察漏项/错行')
                    not_recorded = exempt and kind in ('personal', 'sick') and fact.get('treatment') == 'not_recorded_no_punch' and fact.get('policy_refs')
                    if applicable_observed and fact.get('status') == 'not_applicable' and not (not_recorded or (stage == 'payroll' and allowed_payroll_nonrecording)):
                        self.error(f'{person}:{stage}:{item}:原件有记录却标不适用')
                    if applicable_observed and stage == 'payroll' and not fact.get('policy_refs') and fact.get('status') != 'pending':
                        self.error(f'{person}:{item}:有假别但工资适用口径无来源')
                    if kind in ('personal', 'sick') and exempt and stage == 'attendance' and fact.get('status') != 'pending':
                        if fact.get('status') not in ('not_applicable', 'confirmed_zero') or number(fact.get('amount')) not in (None, 0) or not not_recorded:
                            self.error(person + ':不需要打卡者事病假月度项不记录，须保留原始观察与政策依据')
                    if stage == 'attendance' and fact.get('status') in DONE and not (exempt and kind in ('personal', 'sick')):
                        if any(leave_cols[x['column']]['kind'] == kind and x['row'] in obs['rows'] for x in obs['unknown']):
                            self.error(f'{person}:{item}:原始假别数值未知却已完成')
                        elif fact.get('unit') not in ('hours', 'days'):
                            self.error(f'{person}:{item}:缺明确时数/天数单位')
                        else:
                            expected = obs['business_hours'][kind] / 8 if fact['unit'] == 'days' else obs['business_hours'][kind]
                            if number(fact.get('amount')) != expected: self.error(f'{person}:{item}:从原始日明细提取的总量不符')
            obs['types'] = sorted(obs['types']); obs['raw_types'] = sorted(obs['raw_types']); obs['personal_leave_exempt'] = exempt; obs['coverage_complete'] = coverage
            obs['raw_hours'] = {k: str(v) for k, v in obs['raw_hours'].items()}
            obs['business_hours'] = {k: str(v) if v is not None else None for k, v in obs['business_hours'].items()}
        self.evidence['raw_attendance'] = observations

    def raw_scope_dispositions(self):
        spec = self.r.get('raw_attendance', {}); normalized = self.tables.get(spec.get('table'))
        original_id = spec.get('original_table') or (spec.get('table') if normalized and normalized['spec']['format'] in ('xlsx', 'xls') else None)
        original = self.tables.get(original_id); review = []
        self.evidence['scope_review'] = review
        if not original or original['spec']['format'] not in ('xlsx', 'xls') or not {'name', 'date'} <= set(original['spec']['columns']):
            reason = '原始姓名前置未绑定：标准化数据须另绑实际xlsx/xls的original_table及姓名/日期映射'
            self.wait(reason); review.append({'name':'','route':'pending','reason':reason,'original_rows':[]})
            self.evidence['scope_reverse_status'] = 'pending_original'; return
        names = {}
        # 不使用active；excluded_rows也不能把原始姓名从反查中删掉。
        for row_number, row in original['rows'].items():
            value = row.get('date'); month = value.strftime('%Y-%m') if isinstance(value, (date, datetime)) else str(value or '')[:7]
            name = str(row.get('name') or '').strip()
            if name and month == self.r['month']: names.setdefault(name, []).append(row_number)
        fixed = {str(p.get('name') or '').strip() for p in self.people.values()}
        dispositions = self.r.get('scope_dispositions', [])
        for name, rows in names.items():
            if name in fixed: continue
            own = [d for d in dispositions if str(d.get('name') or '').strip() == name]
            decision = own[0] if len(own) == 1 else {}
            route, reason = decision.get('route'), str(decision.get('reason') or '')
            valid_refs = bool(decision.get('source_refs'))
            for ref in decision.get('source_refs', []):
                table = self.tables.get(ref.get('table')); row = table['rows'].get(ref.get('row'), {}) if table else {}
                if not table or ref.get('column') not in table['spec']['columns'] or str(row.get('name') or '').strip() != name or row.get(ref.get('column')) in (None, ''): valid_refs = False
            circular = any(token in reason.replace(' ', '').lower() for token in ('已选名单', '所选名单', 'selectedscope'))
            if route != 'excluded' or not reason or circular or not valid_refs:
                route = 'pending'
                reason = reason if decision.get('route') == 'pending' and reason else ('排除缺实际同人依据或仅以已选名单为由' if decision else '原始本月有记录但固定scope未收录，负责人范围待核')
                self.wait('范围待核:' + name + ':' + reason)
            review.append({'name':name,'route':route,'reason':reason,'original_table':original_id,'original_source_id':original['spec']['source_id'],'original_rows':rows,'source_refs':decision.get('source_refs',[])})
        self.evidence['scope_reverse_status'] = 'pending_candidates' if any(r['route']=='pending' for r in review) else 'complete'
        self.evidence['original_month_name_count'] = len(names)

    def raw_insurance(self):
        """原标准四类，按个人/单位分别核验；额外单位险种保留原件观察。"""
        spec = self.r.get('raw_insurance')
        if not spec:
            if self.r.get('through_stage') in ('insurance', 'payroll', 'tax', 'payslip'): self.wait('缺保险实际分项与缺项覆盖资料')
            self.insurance_employer_handoff({})
            return
        table = self.tables.get(spec.get('table'))
        if not table:
            self.wait('保险实际来源不可用'); self.insurance_employer_handoff({}); return
        components = spec.get('components', {})
        if not set(REQUIRED['insurance']) <= set(components): self.error('保险须声明养老/失业/医保/公积金四类列映射，缺列以null明确保留')
        if set(components) - set(REQUIRED['insurance']) - {'maternity', 'injury'}: self.error('保险含未识别的额外项目映射')
        observed = {}
        for n, row in table['active'].items():
            person = self.owner(row)
            if not person: self.error(f'保险原件身份未唯一匹配:行{n}'); continue
            if person in observed: self.error(person + ':保险多行须先核实补退费，不能擅自去重或相加'); continue
            mode_column = spec.get('medical_employer_mode_column', 'medical_employer_mode')
            obs = observed[person] = {'row': n, 'parts': {}, 'source_id': table['spec']['source_id'], 'medical_employer_mode': row.get(mode_column) if mode_column in table['spec']['columns'] else None}
            period = row.get('month')
            if str(period)[:7] != self.r.get('month'):
                self.error(person + ':保险来源归属期不符或未知'); continue
            for kind in tuple(REQUIRED['insurance']) + tuple(k for k in ('maternity', 'injury') if k in components):
                mapping = components.get(kind) or {}
                emp, employer = row.get(mapping.get('employee')), row.get(mapping.get('employer'))
                fact = self.facts.get((person, 'insurance', kind))
                part = obs['parts'][kind] = {'employee': emp, 'employer': employer, 'employee_column': mapping.get('employee'), 'employer_column': mapping.get('employer')}
                if kind in ('maternity', 'injury'):
                    # 原制度无个人缴纳；没有个人字段也不伪造一个原件0。
                    if number(emp) not in (None, Decimal(0)): self.error(person + ':单位险种出现个人金额:' + kind)
                    part['detail_status'] = 'observed' if number(employer) is not None else 'detail_pending'
                    part['employer_status'] = 'calculated' if number(employer) is not None else 'pending'
                    if fact:
                        part['employer_status'] = fact.get('employer_status', 'calculated' if number(fact.get('employer_amount')) is not None else 'pending')
                        if part['employer_status'] not in DONE | {'pending'}: self.error(person + ':保险分侧状态无效:' + kind + ':employer')
                        if part['employer_status'] == 'pending' and number(fact.get('employer_amount')) is not None: self.error(person + ':单位保险待核不得保留金额:' + kind)
                        elif part['employer_status'] in DONE and not equal(fact.get('employer_amount'), employer):
                            part['employer_status'] = 'pending'
                            self.error(person + ':保险单位分项金额不符:' + kind)
                    continue
                for side, actual, field in (('employee', emp, 'amount'), ('employer', employer, 'employer_amount')):
                    status = fact.get('status') if fact else 'pending'
                    if side == 'employer': status = fact.get('employer_status', 'calculated' if number(fact.get(field)) is not None else 'pending') if fact else 'pending'
                    part[side + '_status'] = status
                    if status not in DONE | {'pending'}: self.error(person + ':保险分侧状态无效:' + kind + ':' + side)
                    if side == 'employer' and status == 'pending' and fact and number(fact.get(field)) is not None: self.error(person + ':单位保险待核不得保留金额:' + kind)
                    if number(actual) is None:
                        part[side + '_status'] = 'pending'
                        self.wait(person + ':保险分项缺' + ('个人' if side == 'employee' else '单位') + '金额:' + kind)
                        if fact and number(fact.get(field)) is not None: self.error(person + ':保险分侧缺额却标完成:' + kind + ':' + side)
                    elif not fact or status == 'pending':
                        self.wait(person + ':保险分侧尚未完成交接:' + kind + ':' + side)
                    elif not equal(fact.get(field), actual):
                        part[side + '_status'] = 'pending'
                        self.error(person + ':保险分项金额不符:' + kind + ':' + side)
        for person in self.people:
            if person not in observed:
                self.wait(person + ':保险实际来源缺本人行')
                for kind in REQUIRED['insurance']:
                    fact = self.facts.get((person, 'insurance', kind))
                    if fact and fact.get('status') in DONE and fact.get('status') != 'not_applicable': self.error(person + ':缺保险原件行却完成:' + kind)
        self.evidence['raw_insurance'] = observed
        self.insurance_employer_handoff(observed)

    def insurance_employer_handoff(self, observed):
        """原单位社保/医保/基金三栏；不进入个人扣款或payroll_core工资桶。"""
        if 'payroll' not in getattr(self, 'stages', ()) and not any(k[1] == 'payroll' and k[2] in ('employer_social', 'employer_medical', 'employer_fund') for k in self.facts): return
        stage = self.r.get('stages', {}).get('payroll', {})
        actual_table = self.tables.get(stage.get('table'))
        evidence = self.evidence.setdefault('employer_insurance_handoff', {})
        for person in self.people:
            own = next((d for d in stage.get('dispositions', []) if d.get('employee_key') == person), {})
            if own.get('route') == 'excluded': continue
            obs = observed.get(person, {}); parts = obs.get('parts', {})
            def employer(kind):
                p = parts.get(kind, {})
                return number(p.get('employer')) if p.get('employer_status') in DONE else None
            pension, unemployment = employer('pension'), employer('unemployment')
            mode = obs.get('medical_employer_mode'); medical = employer('medical')
            if mode == 'split':
                maternity = employer('maternity')
                medical = medical + maternity if medical is not None and maternity is not None else None
            elif mode != 'combined':
                medical = None
                self.wait(person + ':单位医保缺实际源模式combined/split，不能猜测生育是否已包含')
            expected = {'employer_social': pension + unemployment if pension is not None and unemployment is not None else None, 'employer_medical': medical, 'employer_fund': employer('housing_fund')}
            evidence[person] = {'source_id': obs.get('source_id'), 'source_row': obs.get('row'), 'medical_employer_mode': mode, 'expected': {k: str(v) if v is not None else None for k, v in expected.items()}, 'items': {}}
            for item, value in expected.items():
                fact = self.facts.get((person, 'payroll', item)); result = evidence[person]['items'][item] = {'expected': str(value) if value is not None else None}
                if not actual_table or item not in actual_table['spec']['columns'] or own.get('row') not in actual_table['active']:
                    self.wait(person + ':单位三栏未绑定实际工资列:' + item)
                    if fact and fact.get('status') in DONE: self.error(person + ':单位三栏完成但缺实际工资列:' + item)
                    continue
                actual = actual_table['rows'][own['row']].get(item); result['actual'] = actual
                if not fact:
                    self.wait(person + ':工资单位三栏缺事实:' + item)
                    if actual not in (None, ''): self.error(person + ':工资单位三栏有值但无交接依据:' + item)
                elif value is None:
                    self.wait(person + ':单位保险前置待核:' + item)
                    if fact.get('status') != 'pending' or actual not in (None, ''): self.error(person + ':单位保险待核在工资丢失:' + item)
                elif fact.get('status') == 'pending':
                    self.wait(person + ':工资单位保险仍待核:' + item)
                    if actual not in (None, ''): self.error(person + ':待核单位保险主表不得填值:' + item)
                elif not equal(fact.get('amount'), value) or not equal(actual, value): self.error(person + ':单位保险至工资金额不符:' + item)

    def raw_rewards(self):
        spec = self.r.get('raw_rewards')
        if not spec: self.wait('缺完整原始奖励逐笔核验资料'); return
        table = self.tables.get(spec.get('table'))
        if not table: self.wait('原始奖励表不可用'); return
        fields = ('date', 'tier', 'next_workday', 'checkout', 'next_arrival')
        if not all(f in table['spec']['columns'] for f in fields): self.error('原始奖励缺固定字段映射'); return
        claims = []
        for n, row in table['active'].items():
            person = self.owner(row)
            if not person: self.error(f'奖励原件身份未唯一匹配:行{n}'); continue
            # 补充审批/日历字段必须绑定实际来源行，不能直接填布尔结论。
            claim = dict(row); claim['name'] = person
            for field, ref in spec.get('enrichment', {}).get(str(n), {}).items():
                if field not in ('next_workday', 'checkout', 'next_arrival', 'application_status', 'application_end', 'application_start', 'application_record_id', 'application_source'): self.error('奖励补充字段不在契约中:' + field); continue
                try: claim[field] = self.ref(ref, self.people[person] if ref.get('person_bound', True) else None)
                except ValueError as exc: self.error(f'奖励行{n}:{exc}')
            if isinstance(claim.get('date'), (date, datetime)): claim['date'] = claim['date'].date().isoformat() if isinstance(claim['date'], datetime) else claim['date'].isoformat()
            claim['_source_row'] = n; claims.append(claim)
        result = _module('attendance-calculator', 'reward_policy.py').audit_claims(claims, self.r.get('month'))
        for person in self.people:
            own = [r for r in result['rows'] if r.get('name') == person]
            total = result['totals'].get(person)
            fact = self.facts.get((person, 'reward', 'reward'))
            if not own:
                if not spec.get('collection_complete_refs'): self.wait(person + ':奖励无申报记录且收取覆盖未证实')
                if fact and fact.get('status') != 'pending' and not spec.get('collection_complete_refs'): self.error(person + ':无申报不能自动确认零')
                continue
            if not fact: self.wait(person + ':奖励逐笔结果未交接'); continue
            if total['status'] == 'pending' and fact.get('status') != 'pending': self.error(person + ':奖励逐笔待核在个人汇总丢失')
            if not equal(total['amount'], fact.get('amount')): self.error(person + ':奖励逐笔至个人汇总金额不符')
            if sorted(fact.get('claim_rows', [])) != sorted(r['_source_row'] for r in own): self.error(person + ':奖励逐笔未完整交接')
        if spec.get('collection_complete_refs'): self.refs(spec['collection_complete_refs'], '奖励收取覆盖')
        self.evidence['rewards'] = result

    def confirmations(self):
        confirmed = set()
        for c in self.r.get('confirmations', []):
            source = self.sources.get(c.get('source_id'))
            if not source or not c.get('confirmed_by') or not c.get('confirmed_at') or c.get('sha256') != source['sha256']:
                self.error('确认未绑定实际文件内容、人员与时间'); continue
            dependencies = c.get('dependencies', {})
            if not dependencies:
                self.wait('确认缺前置来源快照:' + str(c.get('source_id'))); continue
            if any(k not in self.sources or self.sources[k]['sha256'] != digest for k, digest in dependencies.items()):
                self.error('确认后前置来源已变化，必须重新确认:' + c['source_id']); continue
            confirmed.add(c['source_id'])
        self.confirmed_sources = confirmed
        through = self.r.get('through_stage')
        required = ['scope_approved']
        if through in ('payroll', 'tax', 'payslip'): required += ['attendance_confirmed']
        if through in ('tax', 'payslip'): required += ['payroll_confirmed']
        for role in required:
            matches = [k for k, s in self.sources.items() if s['role'] == role and s['month'] == self.r.get('month')]
            if not any(k in confirmed for k in matches): self.wait('缺已绑定前置快照的人工确认:' + role)
            actual_table = self.r.get('scope_table') if role == 'scope_approved' else self.r.get('stages', {}).get('attendance' if role == 'attendance_confirmed' else 'payroll', {}).get('table')
            actual_source = self.r.get('tables', {}).get(actual_table, {}).get('source_id')
            if not any(c.get('source_id') in confirmed and c.get('source_id') in matches and (actual_source == c.get('source_id') or actual_source in c.get('dependencies', {})) for c in self.r.get('confirmations', [])):
                self.wait('人工确认未绑定实际核定名单/阶段输出:' + role)
        # 所有实际前置来源必须在确认快照内，不能略掉变化的来源沿用旧确认。
        for c in self.r.get('confirmations', []):
            if c.get('source_id') not in confirmed: continue
            role = self.sources[c['source_id']]['role']
            if role not in ('attendance_confirmed', 'payroll_confirmed'): continue
            before = ('attendance', 'reward') if role == 'attendance_confirmed' else ('attendance', 'reward', 'insurance', 'payroll')
            references = []
            display_confirmation_sources = set()
            for fact in self.facts.values():
                if fact.get('stage') not in before: continue
                for field in ('source_refs', 'value_refs', 'policy_refs'): references += fact.get(field, [])
                references += fact.get('derivation', {}).get('refs', [])
                display = fact.get('display_exception', {})
                if fact.get('item') == 'performance' and display.get('kind') == 'confirmed_zero_performance':
                    if display.get('confirmation_source_id'): display_confirmation_sources.add(display['confirmation_source_id'])
                    references += [display[k] for k in ('amount_ref', 'month_ref', 'zero_required_ref') if display.get(k)]
                    if display.get('original_evidence', {}).get('ref'): references.append(display['original_evidence']['ref'])
            for entry in self.r.get('raw_rewards', {}).get('enrichment', {}).values(): references += list(entry.values())
            references += self.r.get('raw_rewards', {}).get('collection_complete_refs', [])
            references += self.r.get('raw_attendance', {}).get('blank_zero_refs', [])
            for refs in self.r.get('raw_attendance', {}).get('no_punch_refs', {}).values(): references += refs
            if role=='payroll_confirmed':
                for disposition in self.r.get('comment_dispositions',[]):
                    references+=disposition.get('evidence_refs',[])
                    if disposition.get('amount_ref'):references.append(disposition['amount_ref'])
                    references+=disposition.get('component_refs',[])
                    if disposition.get('total_leave_ref'):references.append(disposition['total_leave_ref'])
            used = {self.r['tables'][ref['table']]['source_id'] for ref in references if ref.get('table') in self.r.get('tables', {})}
            used |= display_confirmation_sources
            if role=='payroll_confirmed':used|=set(self.r.get('payroll_original_source_ids',[]))
            raw_names = ('raw_attendance', 'raw_rewards') if role == 'attendance_confirmed' else ('raw_attendance', 'raw_rewards', 'raw_insurance')
            used |= {self.r['tables'][self.r[x]['table']]['source_id'] for x in raw_names if self.r.get(x, {}).get('table') in self.r.get('tables', {})}
            scope_table = self.r.get('tables', {}).get(self.r.get('scope_table'), {})
            if scope_table.get('source_id'): used.add(scope_table['source_id'])
            original = self.r.get('raw_attendance', {}).get('original_table')
            if original in self.r.get('tables', {}): used.add(self.r['tables'][original]['source_id'])
            for stage in before:
                used |= {value for key, value in self.r.get('template_bindings', {}).get(stage, {}).items() if key.endswith('_source_id') and isinstance(value, str) and value in self.sources}
            used.discard(c['source_id'])
            if not used <= set(c.get('dependencies', {})): self.error('确认快照遗漏实际前置来源:' + c['source_id'])

    def template_receipts(self):
        linked = {}
        required_stages = self.required_delivery_stages()
        for stage in getattr(self, 'stages', ()):
            if stage not in required_stages: continue
            stage_spec = self.r.get('stages', {}).get(stage, {})
            table = self.tables.get(stage_spec.get('table'))
            binding = self.r.get('template_bindings', {}).get(stage)
            if not binding:
                self.tech_wait(stage + ':缺指定正式模板的实际复核绑定；省略template_required不能豁免')
                linked[stage] = {'verified': False, 'reason': '缺实际模板绑定'}; continue
            if not table:
                self.tech_wait(stage + ':尚无正式模板实际产出，保留独立明细'); continue
            try:
                actual = self.sources[table['spec']['source_id']]
                def pinned(source_field, digest_field):
                    source = self.sources.get(binding.get(source_field))
                    if source is None: raise MissingEvidence('缺有效来源:' + source_field)
                    if binding.get(digest_field) != source['sha256']: raise ValueError('固定SHA与实际来源不符:' + source_field)
                    return source
                if binding.get('kind', 'xlsx') == 'xlsx':
                    if table['spec']['format'] != 'xlsx': raise MissingEvidence('当前仅有计算辅助表，尚无实际正式xlsx模板成品')
                    registry = pinned('registry_source_id', 'registry_sha256')
                    plan = pinned('plan_source_id', 'plan_sha256')
                    receipt = pinned('receipt_source_id', 'receipt_sha256')
                    kwargs = {}
                    if binding.get('recalculation_log_source_id'):
                        log = pinned('recalculation_log_source_id', 'recalculation_log_sha256')
                        kwargs.update(recalculation_log=log['resolved'], recalculation_log_sha256=log['sha256'])
                    guard = _module('payroll-operations', 'template_guard.py')
                    checked = guard.verify(registry['resolved'], registry['sha256'], plan['resolved'], actual['resolved'], receipt['resolved'], None, **kwargs)
                elif binding.get('kind') == 'legacy_xls':
                    if stage != 'tax' or table['spec']['format'] != 'xls': raise ValueError('legacy_xls仅用于实际政府税表原始xls')
                    template = pinned('template_source_id', 'template_sha256')
                    policy_source = pinned('policy_source_id', 'policy_sha256')
                    policy = json.loads(policy_source['resolved'].read_text(encoding='utf-8-sig'))
                    guard = _module('payroll-operations', 'legacy_template_guard.py')
                    checked = guard.inspect(template['resolved'], template['sha256'], None, actual['resolved'], clear_original_yellow=policy.get('clear_original_yellow') is True, clear_all_fills=policy.get('clear_all_fills') is True)
                else: raise ValueError('未知模板复核路径')
                linked[stage] = {'structure': checked, 'actual_output_sha256': actual['sha256']}
                if not checked.get('passed'):
                    for e in checked.get('errors', ['未通过']): self.tech_error(stage + ':实际模板重放拒绝:' + str(e))
                    continue
                if checked.get('finalized_output_sha256') != actual['sha256']: raise ValueError('实际模板复核未绑定本次最终SHA')
                if stage == 'payroll':
                    registration = json.loads(registry['resolved'].read_text(encoding='utf-8-sig'))
                    policy = registration.get('policy', {})
                    originals = set(self.r.get('payroll_original_source_ids', [])) | {key for key, value in self.sources.items() if value.get('role') in ('payroll_original', 'wage_original')}
                    if policy.get('require_source_comments') is not True or not originals:
                        self.tech_wait('payroll:缺本次工资原件批注强制保护或实际原件清单')
                    registered = set()
                    for item in policy.get('source_comments', []):
                        value = item.get('source', {}); path = Path(str(value.get('path', '')))
                        if not path.is_absolute(): path = registry['resolved'].parent / path
                        registered.add((str(path.resolve()), value.get('sha256')))
                    for original_id in originals:
                        value = self.sources.get(original_id)
                        if not value or (str(value['resolved'].resolve()), value['sha256']) not in registered or original_id == table['spec']['source_id']:
                            self.tech_wait('payroll:工资原件未完整登记批注盘点（0条也须登记）:' + str(original_id))
                    if not registration.get('comment_source_inventories') or 'source_comment_inventory' not in checked:
                        self.tech_wait('payroll:当前模板能力或回执缺实际原件批注盘点，不能自称保真')
                    formula_plan = pinned('formula_audit_plan_source_id', 'formula_audit_plan_sha256')
                    plan_data = json.loads(formula_plan['resolved'].read_text(encoding='utf-8-sig'))
                    plan_data['source_view_inventory'] = checked.get('source_view_inventory')
                    plan_data['attendance_identity_column'] = table['spec']['columns'].get('name',plan_data.get('attendance_identity_column'))
                    if plan_data.get('sheet') != table['spec'].get('sheet') or plan_data.get('first_row') != table['spec']['start_row'] or plan_data.get('last_row') != table['spec']['end_row']: raise ValueError('公式审计计划未覆盖实际工资人员范围')
                    for exemption in plan_data.get('pending_formula_exemptions', {}).values(): self.refs(exemption.get('source_refs'), '工资公式未核口径来源')
                    checked_formulas = _module('payroll-operations', 'audit_payroll_workbook.py').audit(actual['resolved'], plan_data)
                    linked[stage]['formula_audit'] = checked_formulas
                    for e in checked_formulas.get('errors', []): self.tech_error(stage + ':实际公式/缓存拒绝:' + str(e))
                    for e in checked_formulas.get('technical_pending', []): self.tech_wait(stage + ':公式技术未完成:' + str(e))
                    if checked_formulas.get('pending_totals') or checked_formulas.get('pending_rows'): self.wait(stage + ':实际公式/合计仍有待核事项')
                native = binding.get('native_recalculation', {})
                if native.get('run') is not True:
                    self.tech_wait(stage + ':未执行隔离副本的实际原生重算与缓存核对；字符串passed无效'); continue
                checks = native.get('linkage_checks', [])
                if stage == 'payroll' and not checks:
                    self.tech_wait(stage + ':缺实际可编辑事假输入到请假扣款/应发/实发的联动核对'); continue
                if stage == 'payroll':
                    from openpyxl.utils.cell import coordinate_to_tuple
                    source_links=(checked.get('source_view_inventory') or {}).get('links',[])
                    if not any(c.get('kind')=='attendance_snapshot' for c in checks): self.tech_wait(stage + ':缺工资内考勤快照业务金额到T的实际联动核对')
                    if not any(c.get('kind')!='attendance_snapshot' for c in checks): self.tech_wait(stage + ':缺原有请假输入联动核对')
                    for check in checks:
                        if check.get('kind')=='attendance_snapshot':
                            links=[x for x in source_links if x['view_sheet']==check.get('sheet') and x['value_cell']==check.get('input_cell') and x['target_sheet']==table['spec']['sheet']]
                            if len(links)!=1: raise ValueError('考勤联动输入未绑定实际同人可编辑快照金额')
                            link=links[0]; target_row=coordinate_to_tuple(link['target_cell'])[0]
                            if target_row not in table['active']: raise ValueError('考勤联动目标不是本次人员行')
                            outputs=check.get('outputs',[])
                            if not any(x.get('item')=='attendance_deduction' and x.get('sheet')==link['target_sheet'] and x.get('cell')==link['target_cell'] for x in outputs): raise ValueError('考勤快照联动缺实际工资T输出')
                            for item in outputs:
                                if item.get('sheet')!=table['spec']['sheet'] or item.get('cell')!=str(table['spec']['columns'].get(item.get('item')))+str(target_row): raise ValueError('考勤跨页联动输出未绑定实际工资字段')
                            continue
                        if check.get('sheet') != table['spec']['sheet'] or coordinate_to_tuple(check['input_cell'])[0] not in table['active']: raise ValueError('联动测试未绑定实际工资人员行')
                        expected_items = {'leave_deduction', 'gross', 'net'}
                        if not expected_items <= {x.get('item') for x in check.get('outputs', [])}: raise ValueError('联动测试缺请假扣款/应发/实发')
                        for item in check['outputs']:
                            if item.get('item') in expected_items and item.get('cell') != str(table['spec']['columns'].get(item['item'])) + str(coordinate_to_tuple(check['input_cell'])[0]): raise ValueError('联动输出未绑定实际工资字段')
                native_checked = _module('payroll-operations', 'native_recalc_verify.py').verify_native_copy(actual['resolved'], actual['sha256'], checks)
                linked[stage]['native_recalculation'] = native_checked
                for e in native_checked.get('errors', []): self.tech_error(stage + ':原生复核:' + str(e))
                unresolved=native_checked.get('pending',[])
                data_only=bool(native_checked.get('actual_engine') and unresolved and not native_checked.get('errors') and all('联动输出仍为空白' in e or '缺核定数值预期' in e for e in unresolved))
                linked[stage]['native_linkage_complete']=native_checked.get('verified') is True
                linked[stage]['native_technical_verified']=native_checked.get('verified') is True or data_only
                if data_only:self.wait(stage + ':原生已执行，金额依赖/预期仍待核:' + ';'.join(unresolved))
                elif native_checked.get('verified') is not True:self.tech_wait(stage + ':原生重算/缓存未证实:' + ';'.join(unresolved))
            except MissingEvidence as exc: self.tech_wait(stage + ':模板实物复核资料不足:' + str(exc))
            except (ValueError, TypeError, OSError, KeyError) as exc: self.tech_error(stage + ':模板实物复核拒绝:' + str(exc))
        for stage,item in linked.items():
            item['technical_verified']=bool(item.get('structure',{}).get('passed') and item.get('native_technical_verified') and not any(x.startswith(stage+':') for x in self.technical_errors+self.technical_pending))
        self.evidence['template_receipts'] = linked

    def maternity_handoff(self):
        for person in self.people:
            fact = self.facts.get((person, 'payroll', 'maternity_leave'))
            if not fact or fact.get('status') == 'pending' or number(fact.get('amount')) in (None, 0): continue
            if number(fact.get('amount')) < 0 or fact.get('included_in') != 'other_deduction':
                self.error(person + ':产假月中差额须正数计入其他个人扣款，不能另入请假/考勤/入离职差额'); continue
            original_ref = fact.get('manual_original_ref')
            if not original_ref:
                self.wait(person + ':产假差额缺原其他个人扣款格定位，不能判断是否覆盖人工值'); continue
            try: original = self.ref(original_ref, self.people[person])
            except ValueError as exc: self.error(person + ':产假原人工值:' + str(exc)); continue
            bucket = self.facts.get((person, 'payroll', 'other_deduction'))
            if original not in (None, ''):
                approval_id = fact.get('merge_confirmation_source_id')
                approval = self.sources.get(approval_id, {})
                confirmation = next((c for c in self.r.get('confirmations', []) if c.get('source_id') == approval_id), {})
                ref = fact.get('merge_amount_ref')
                source_ids = {self.r['tables'][original_ref['table']]['source_id']} | {self.r['tables'][r['table']]['source_id'] for r in fact.get('source_refs', []) if r.get('table') in self.r.get('tables', {})}
                if approval_id not in getattr(self, 'confirmed_sources', set()) or approval.get('role') != 'manual_approval' or not source_ids <= set(confirmation.get('dependencies', {})) or not ref or self.r.get('tables', {}).get(ref.get('table'), {}).get('source_id') != approval_id:
                    self.wait(person + ':原其他个人扣款已有人工值，产假是否已包含及合并金额未实际确认'); continue
                try: approved_amount = self.ref(ref, self.people[person])
                except ValueError as exc: self.error(person + ':产假合并依据:' + str(exc)); continue
                if bucket and bucket.get('status') in DONE and not equal(bucket.get('amount'), approved_amount): self.error(person + ':产假合并后实际其他个人扣款与本人确认金额不符')
            elif bucket and bucket.get('status') in DONE:
                if not any(d.get('stage') == 'payroll' and d.get('item') == 'maternity_leave' for d in bucket.get('depends_on', [])):
                    self.wait(person + ':缺产假专项差额到实际其他个人扣款桶的组成核对')

    def comment_business_review(self):
        inventory=self.evidence.get('template_receipts',{}).get('payroll',{}).get('structure',{}).get('source_comment_inventory')
        if inventory is None:return
        def resolve(ref,key):
            value=self.ref(ref,self.people[key]);table=self.tables[ref['table']]
            return value,table['rows'][ref['row']],self.sources[table['spec']['source_id']]
        def components(table_id,key):
            table=self.tables[table_id]
            if not {'component_id','month','amount'}<=set(table['spec']['columns']):raise ValueError('完整组成表缺实际项目标识/月/金额映射')
            refs=[]
            for n,actual in table['rows'].items():
                if n<table['spec']['start_row'] or not self.matches(actual,self.people[key]):continue
                if not actual.get('month'):raise ValueError('本人组成明细有未定月份，不能宣称覆盖完整')
                if actual['month']==self.r['month']:refs.append({'table':table_id,'row':n,'column':'amount'})
            return refs
        rows=_module('payroll-operations','payroll_comment_review.py').review_comments(inventory,self.people,self.r['month'],self.r.get('comment_dispositions',[]),resolve,components)
        states=[];affected=set()
        for row in rows:
            key=row['employee_key']
            if row['effective_status']=='resolved':
                if row.get('adopted_total_leave') is not None:
                    fact=self.facts.get((key,'payroll','leave_deduction'))
                    if not fact or not equal(fact.get('amount'),row['adopted_total_leave']):self.error(str(key)+':本人已核完整请假扣款总额与实际S事实不符')
                continue
            self.wait(str(key)+':原工资事假清算批注业务待核:'+row['source_sheet']+'!'+row['source_cell']+':'+row['reason'])
            if key not in self.people:continue
            affected.add(key)
            targets=[('payroll',x) for x in ('leave_deduction','gross','finance','net','bank')]+[('tax','taxable'),('payslip','gross'),('payslip','net')]
            for stage,item in targets:
                fact=self.facts.get((key,stage,item))
                stage_spec=self.r.get('stages',{}).get(stage,{})
                table=self.tables.get(stage_spec.get('table'))
                own=[d for d in stage_spec.get('dispositions',[]) if d.get('employee_key')==key and d.get('row') is not None]
                refs=list((fact or {}).get('value_refs',[]))
                if own and table and item in table['spec']['columns']:
                    actual_ref={'table':stage_spec['table'],'row':own[0]['row'],'column':item}
                    if actual_ref not in refs:refs.append(actual_ref)
                if not fact and not refs:continue
                fact=fact or {'employee_key':key,'stage':stage,'item':item,'status':'pending','amount':None}
                values=[]
                if fact.get('status')!='pending' or fact.get('amount') is not None:self.error(key+':原清算批注未决但'+stage+'.'+item+'仍声明已算/不适用或金额')
                for ref in refs:
                    try:
                        value=self.ref(ref,self.people[key]);values.append({'ref':ref,'actual_value':value})
                        if number(value) is not None:self.tech_error('payroll:'+key+':原清算批注未决，实际'+stage+'.'+item+'仍显示数值；S及依赖应发须真实待核，不能仅改审计状态')
                    except ValueError as exc:self.error(key+':清算待核实物定位:'+str(exc))
                states.append({'employee_key':key,'stage':stage,'item':item,'effective_status':'pending','effective_amount':None,'declared_status':fact.get('status'),'declared_amount':fact.get('amount'),'actual_values':values,'reason':'原工资清算批注尚未逐人处置'})
                self.facts[(key,stage,item)]=dict(fact,status='pending',amount=None)
        self.evidence['comment_business_review']={'items':rows,'effective_facts':states,'note':'批注对象保留与业务处置分别核验；不修改源或成品，不推定清算小时或扣款'}
        for key in affected:self.check_payroll_totals(key)
        if any(x.startswith('payroll:') for x in self.technical_errors):self.evidence['template_receipts']['payroll']['technical_verified']=False

    def attendance_handoff_review(self):
        checked=_module('payroll-operations','attendance_fact_handoff.py').verify_handoff(self.r,self.base,self)
        self.evidence['attendance_fact_handoff']=checked
        for message in checked['errors']:self.tech_error('attendance:最终人资事实交接拒绝:'+message)
        for message in checked['pending']:self.tech_wait('attendance:最终人资事实交接未完成:'+message)
        for message in checked['data_pending']:self.wait(message)
        if not checked['passed'] and 'attendance' in self.evidence.get('template_receipts',{}):self.evidence['template_receipts']['attendance']['technical_verified']=False

    def run(self):
        if self.r.get('schema_version') != 'crossstage-1': self.error('不是crossstage-1契约')
        if not self.r.get('month') or not self.r.get('policy_version'): self.error('缺月份或规则版本')
        self.load_sources(); self.load_tables(); self.scope(); self.stage_rows(); self.check_facts()
        self.raw_scope_dispositions(); self.raw_attendance(); self.raw_rewards(); self.raw_insurance(); self.template_receipts(); self.comment_business_review(); self.attendance_handoff_review(); self.confirmations(); self.maternity_handoff()
        self.errors = list(dict.fromkeys(self.errors)); self.pending = list(dict.fromkeys(self.pending))
        mode = self.r.get('mode', 'draft')
        if mode not in ('draft', 'formal'): self.error('模式须为draft/formal')
        purpose=self.r.get('purpose','template_delivery')
        if purpose not in ('template_delivery','calculation_details'):self.error('purpose须为template_delivery/calculation_details')
        required=self.required_delivery_stages()
        verified=[x for x in required if self.evidence.get('template_receipts',{}).get(x,{}).get('technical_verified')]
        technical_ready=bool(required) and set(required)<=set(verified) and not self.technical_errors and not self.technical_pending
        data_errors=[x for x in self.errors if x not in self.technical_errors]
        data_pending=[x for x in self.pending if x not in self.technical_pending]
        calculation_passed=not data_errors; business_ready=calculation_passed and not data_pending
        formal_ready=purpose=='template_delivery' and business_ready and technical_ready
        passed=calculation_passed and (mode=='draft' or business_ready) and (purpose=='calculation_details' or technical_ready)
        status='calculation_details_only' if purpose=='calculation_details' else ('technical_incomplete' if not technical_ready else ('formal_template_verified' if mode=='formal' and formal_ready else 'template_draft_technically_verified'))
        message='仅计算明细，未交付合规模板' if purpose=='calculation_details' else ('模板技术交付未完成；可算独立明细已保留' if not technical_ready else '本次实际交付范围技术已核验，金额/政策待核另列')
        return {'passed':passed,'calculation_passed':calculation_passed,'business_ready':business_ready,'formal_ready':formal_ready,'purpose':purpose,'delivery_status':status,'delivery_message':message,'technical_delivery':{'ready':technical_ready,'required_template_stages':required,'verified_template_stages':verified,'all_five_stages_verified':set(('attendance','insurance','payroll','tax','payslip'))<=set(verified),'input_only_stages':[x for x in getattr(self,'stages',()) if x in ('attendance','insurance','payroll','tax','payslip') and x not in required],'errors':self.technical_errors,'pending':self.technical_pending},'mode':mode,'month':self.r.get('month'),'employee_count':len(self.people),'errors':self.errors,'pending':self.pending,'evidence':self.evidence}

def audit_crossstage(record, base_dir=Path('.')):
    try: return Audit(record, base_dir).run()
    except (ValueError, TypeError, KeyError, AttributeError, IndexError) as exc:
        return {'passed': False, 'formal_ready': False, 'errors': ['输入契约无效:' + str(exc)], 'pending': []}

def write_scope_review_csv(report, path):
    path = Path(path)
    if path.exists(): raise ValueError('范围复核CSV已存在，不得覆盖')
    fields = ('name','route','reason','original_table','original_rows')
    with path.open('w', encoding='utf-8-sig', newline='') as stream:
        writer = csv.DictWriter(stream, fieldnames=fields); writer.writeheader()
        for row in report.get('evidence', {}).get('scope_review', []):
            writer.writerow({key: ';'.join(map(str,row.get(key,[]))) if key=='original_rows' else row.get(key,'') for key in fields})

def write_comment_review_csv(report,path):
    path=Path(path)
    if path.exists():raise ValueError('原工资批注业务复核CSV已存在，不得覆盖')
    fields=('name','employee_key','source_path','source_sheet','source_cell','original_text','object_status','action','effective_status','reason')
    with path.open('w',encoding='utf-8-sig',newline='') as stream:
        writer=csv.DictWriter(stream,fieldnames=fields);writer.writeheader()
        for row in report.get('evidence',{}).get('comment_business_review',{}).get('items',[]):writer.writerow({key:row.get(key,'') for key in fields})

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--record', type=Path, required=True); ap.add_argument('--output', type=Path, required=True)
    a = ap.parse_args(); record = json.loads(a.record.read_text(encoding='utf-8-sig'))
    protected = {a.record.resolve()} | {(a.record.parent / str(s.get('path', ''))).resolve() for s in record.get('sources', {}).values()}
    if a.output.resolve() in protected or a.output.exists(): raise SystemExit('输出不得覆盖记录、来源或既有结果')
    result = audit_crossstage(record, a.record.parent)
    a.output.parent.mkdir(parents=True, exist_ok=True)
    scope_csv = a.output.with_suffix('.scope.csv')
    if scope_csv.resolve() in protected: raise SystemExit('范围复核CSV不得覆盖来源')
    write_scope_review_csv(result, scope_csv); result['scope_review_csv'] = str(scope_csv)
    comment_csv=a.output.with_suffix('.comments.csv')
    if comment_csv.resolve() in protected:raise SystemExit('批注复核CSV不得覆盖来源')
    write_comment_review_csv(result,comment_csv);result['comment_review_csv']=str(comment_csv)
    a.output.write_text(json.dumps(result, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    print(json.dumps({k:result.get(k) for k in ('passed','calculation_passed','formal_ready','purpose','delivery_status','delivery_message','technical_delivery','errors','pending')},ensure_ascii=False))
    raise SystemExit(0 if result['passed'] else 2)

if __name__ == '__main__': main()
