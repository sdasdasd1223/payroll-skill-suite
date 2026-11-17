"""明确合成：已确认零绩效的主表留白，与缺资料保持不同状态。"""
import copy
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from openpyxl import Workbook, load_workbook

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts', 'payroll-operations/scripts', 'skills/payroll-operations/scripts', 'candidate/skills/payroll-operations/scripts'):
            path = parent / relative / 'crossstage_audit.py'
            if path.is_file(): return path
    raise RuntimeError('未找到相邻技能程序')

def load_program(path):
    spec = importlib.util.spec_from_file_location('_synthetic_performance_' + str(id(path)), path)
    program = importlib.util.module_from_spec(spec); spec.loader.exec_module(program); return program

PROGRAM = load_program(locate())

def build(root, render='blank'):
    root = Path(root).resolve(); root.mkdir(parents=True, exist_ok=True)
    person = {'employee_key': 'SYNTHETIC-KEY', 'name': 'SYNTHETIC', 'company': 'SYNTHETIC'}
    record = {'schema_version': 'crossstage-1', 'month': '2099-01', 'policy_version': 'SYNTHETIC-ONLY', 'mode': 'draft', 'sources': {}, 'tables': {}, 'scope_table': 'scope', 'stages': {'payroll': {'table': 'main', 'dispositions': [{'employee_key': person['employee_key'], 'route': 'output', 'row': 2}]}}, 'facts': [], 'confirmations': []}
    def table(key, rows, role='calculation'):
        path = root / (key + '.json'); path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding='utf-8')
        record['sources'][key] = {'path': str(path), 'sha256': PROGRAM.sha256(path), 'month': record['month'], 'role': role}
        record['tables'][key] = {'source_id': key, 'format': 'json', 'start_row': 1, 'end_row': len(rows), 'columns': {k:k for row in rows for k in row}}
    table('scope', [person])
    table('policy', [{'rule': '已确认本人本月无绩效时主表留空，模板或人资明确要求时显示0', 'zero_required': 'zero_required' if render == 'zero' else 'blank_allowed'}], 'policy')
    statement = root / 'original-confirmation.txt'
    quote = 'SYNTHETIC-KEY：已确认本人本月无绩效。'
    statement.write_text(quote + '\n', encoding='utf-8')
    record['sources']['original'] = {'path': str(statement), 'sha256': PROGRAM.sha256(statement), 'month': record['month'], 'role': 'performance_confirmation'}
    # 规范化辅助记录是程序留痕，用户确认仍绑定原文字，不冒充用户确认这个新JSON。
    table('decision', [dict(person, month=record['month'], amount=0, original_source_id='original')], 'normalized_decision')
    numbers = dict(person, salary=1, other_allowance=0, performance=0, reward=0, leave_deduction=0, attendance_deduction=0, hire_exit_difference=0, other_income=0, other_deduction=0, personal_social=0, personal_medical=0, personal_fund=0, tax=0, combined_tax=0, gross=1, net=1, taxable=1)
    table('numbers', [numbers])
    main = root / 'performance-main.xlsx'; book = Workbook(); ws = book.active; ws.title = 'SYNTHETIC'
    ws['A1'] = 'employee_key'; ws['B1'] = 'name'; ws['C1'] = 'company'; ws['O1'] = 'performance'
    ws['A2'] = person['employee_key']; ws['B2'] = person['name']; ws['C2'] = person['company']; ws['O2'] = 0 if render == 'zero' else None
    book.save(main)
    record['sources']['main'] = {'path': str(main), 'sha256': PROGRAM.sha256(main), 'month': record['month'], 'role': 'draft_output'}
    record['tables']['main'] = {'source_id': 'main', 'format': 'xlsx', 'sheet': 'SYNTHETIC', 'start_row': 2, 'end_row': 2, 'columns': {'employee_key': 'A', 'name': 'B', 'company': 'C', 'performance': 'O'}}
    for item, amount in numbers.items():
        if item in person: continue
        stage = 'tax' if item in ('tax', 'combined_tax', 'taxable') else 'payroll'
        ref = {'table': 'numbers', 'row': 1, 'column': item}
        fact = {'employee_key': person['employee_key'], 'stage': stage, 'item': item, 'status': 'calculated', 'amount': amount, 'source_refs': [ref], 'value_refs': [ref]}
        if item == 'salary': fact['derivation'] = {'operation': 'copy', 'refs': [ref]}
        if item == 'performance':
            source = {'table': 'decision', 'row': 1, 'column': 'amount'}
            fact.update(status='confirmed_zero', source_refs=[source], derivation={'operation':'copy', 'refs':[source]}, value_refs=[{'table': 'main', 'row': 2, 'column': 'performance'}], policy_refs=[{'table':'policy', 'row':1, 'column':'rule', 'person_bound':False}], display_exception={'kind':'confirmed_zero_performance', 'render':render, 'confirmation_source_id':'original', 'amount_ref':source, 'month_ref':{'table':'decision','row':1,'column':'month'}, 'original_evidence':{'source_id':'original','line':1,'quote':quote}})
            if render == 'zero': fact['display_exception']['zero_required_ref'] = {'table':'policy','row':1,'column':'zero_required','person_bound':False}
        record['facts'].append(fact)
    record['confirmations'] = [{'source_id':'original', 'sha256':record['sources']['original']['sha256'], 'confirmed_by':'SYNTHETIC-USER', 'confirmed_at':'2099-01-31T12:00:00+08:00', 'employee_key':person['employee_key'], 'month':record['month'], 'dependencies':{'scope':record['sources']['scope']['sha256']}}]
    return record

def performance(record): return next(f for f in record['facts'] if f['item'] == 'performance')

def pending_totals(record):
    path = Path(record['sources']['numbers']['path']); rows = json.loads(path.read_text(encoding='utf-8'))
    for f in record['facts']:
        if f['item'] in ('gross', 'net', 'taxable'):
            f.update(status='pending', amount=None, reason='绩效确认未证实，正式合计待核'); rows[0][f['item']] = None
    path.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding='utf-8'); record['sources']['numbers']['sha256'] = PROGRAM.sha256(path)

def check(record, program=PROGRAM):
    audit = program.Audit(record, Path('.')); audit.load_sources(); audit.load_tables(); audit.scope(); audit.stages = ()
    audit.check_facts()
    return audit

class PerformanceDisplay(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix='synthetic-performance-blank-'); self.root = Path(self.temporary.name).resolve(); self.record = build(self.root)
    def tearDown(self):
        if self.root.parent != Path(tempfile.gettempdir()).resolve(): raise RuntimeError('隔离目录越界')
        self.temporary.cleanup()
    def assert_pending(self):
        pending_totals(self.record); result = check(self.record)
        self.assertFalse(result.errors, result.errors)
        self.assertTrue(any('绩效' in p for p in result.pending), result.pending)
        self.assertIn('performance', result.evidence['known_subtotals']['SYNTHETIC-KEY']['unknown_gross_items'])
        self.assertEqual(next(f for f in self.record['facts'] if f['item']=='salary')['amount'], 1)
    def test_confirmed_person_current_month_blank_uses_zero_only_in_core(self):
        digest = PROGRAM.sha256(self.root/'performance-main.xlsx'); result = check(self.record)
        self.assertFalse(result.errors, result.errors); self.assertFalse(result.pending, result.pending)
        self.assertEqual(performance(self.record)['amount'], 0)
        book = load_workbook(self.root/'performance-main.xlsx', read_only=True, data_only=True)
        self.assertIsNone(book.active['O2'].value); book.close()
        self.assertEqual(digest, PROGRAM.sha256(self.root/'performance-main.xlsx'))
    def test_auxiliary_zero_is_not_confused_with_main_blank_display(self):
        performance(self.record)['value_refs'].append({'table':'numbers','row':1,'column':'performance'})
        result = check(self.record)
        self.assertFalse(result.errors, result.errors); self.assertFalse(result.pending, result.pending)
    def test_missing_confirmation_stays_pending(self):
        self.record['confirmations'] = []; self.assert_pending()
    def test_other_person_confirmation_stays_pending(self):
        self.record['confirmations'][0]['employee_key'] = 'SYNTHETIC-OTHER'; self.assert_pending()
    def test_other_person_original_text_cannot_be_relabelled(self):
        text = 'SYNTHETIC-OTHER：已确认本人本月无绩效。'
        path = Path(self.record['sources']['original']['path']); path.write_text(text, encoding='utf-8')
        self.record['sources']['original']['sha256'] = PROGRAM.sha256(path); self.record['confirmations'][0]['sha256'] = PROGRAM.sha256(path)
        performance(self.record)['display_exception']['original_evidence']['quote'] = text
        self.assert_pending()
    def test_wrong_month_confirmation_stays_pending(self):
        self.record['confirmations'][0]['month'] = '2098-12'; self.assert_pending()
    def test_changed_scope_dependency_stays_pending(self):
        self.record['confirmations'][0]['dependencies']['scope'] = 'CHANGED'; self.assert_pending()
    def test_unprovided_performance_is_not_default_zero(self):
        self.record['confirmations'] = []; performance(self.record).update(status='pending', amount=None, reason='未提供绩效资料'); performance(self.record).pop('display_exception'); self.assert_pending()
    def test_not_provided_original_text_cannot_be_normalized_as_zero(self):
        text = 'SYNTHETIC-KEY：本月绩效资料未提供。'
        path = Path(self.record['sources']['original']['path']); path.write_text(text, encoding='utf-8')
        self.record['sources']['original']['sha256'] = PROGRAM.sha256(path); self.record['confirmations'][0]['sha256'] = PROGRAM.sha256(path)
        performance(self.record)['display_exception']['original_evidence']['quote'] = text
        self.assert_pending()
    def test_missing_policy_stays_pending(self):
        performance(self.record)['policy_refs'] = []; self.assert_pending()
    def test_zero_display_requires_explicit_template_or_user_rule(self):
        other = self.root/'explicit-zero'; record = build(other, 'zero'); digest = PROGRAM.sha256(other/'performance-main.xlsx')
        result = check(record); self.assertFalse(result.errors, result.errors); self.assertFalse(result.pending, result.pending)
        self.assertEqual(digest, PROGRAM.sha256(other/'performance-main.xlsx'))
        performance(record)['display_exception'].pop('zero_required_ref'); self.record = record; self.assert_pending()
    def test_blank_zero_without_explicit_exception_is_rejected(self):
        performance(self.record).pop('display_exception'); self.assertTrue(any('实际金额' in e for e in check(self.record).errors))
    def test_exception_does_not_apply_to_another_payroll_bucket(self):
        other = next(f for f in self.record['facts'] if f['item']=='other_allowance'); other.update(amount=0, value_refs=[{'table':'main','row':2,'column':'performance'}], display_exception=copy.deepcopy(performance(self.record)['display_exception']))
        self.assertTrue(any('实际金额' in e for e in check(self.record).errors))

if __name__ == '__main__': unittest.main(verbosity=2)
