"""可发布的明确合成反例；无真实人员、工资或机器绝对路径。
运行：python -B test_crossstage_public.py；从相邻技能目录动态加载实际程序。
"""
import json
from pathlib import Path
import sys
import tempfile
import unittest
from zipfile import ZipFile, ZIP_DEFLATED

def locate():
    for parent in Path(__file__).resolve().parents:
        for candidate in (parent / 'scripts', parent / 'payroll-operations/scripts', parent / 'skills/payroll-operations/scripts', parent / 'candidate/skills/payroll-operations/scripts'):
            if (candidate / 'crossstage_audit.py').is_file() and (candidate / 'template_guard.py').is_file(): return candidate
    raise RuntimeError('须放在相邻技能目录或其测试目录运行')

sys.path.insert(0, str(locate()))
import template_guard as guard
from crossstage_audit import Audit, sha256
from openpyxl import Workbook, load_workbook

class TemplateBoundary(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix='synthetic-crossstage-public-')
        self.root = Path(self.temporary.name).resolve()
        self.template = self.root / 'synthetic-template.xlsx'
        b = Workbook(); s = b.active; s.title = 'SYNTHETIC'
        s.append(['employee_key', 'name', 'company', 'SYNTHETIC_LENGTH'])
        s.append(['SYNTHETIC-KEY', 'SYNTHETIC', 'SYNTHETIC', '=LEN(A2)']); b.save(self.template)
        self.registry, self.plan = self.root / 'registry.json', self.root / 'plan.json'
        reg = guard.register(self.template, sha256(self.template), {}, self.registry)
        self.plan.write_text('{}', encoding='utf-8')
        self.output, self.receipt = self.root / 'output.xlsx', self.root / 'receipt.json'
        guard.build(self.registry, reg['registry_sha256'], self.plan, self.output, self.receipt)
        self.record = {'schema_version': 'crossstage-1', 'month': '2099-01', 'sources': {}, 'tables': {}, 'stages': {'attendance': {'table': 'actual'}}, 'template_bindings': {}}
        for key, path in (('actual', self.output), ('registry', self.registry), ('plan', self.plan), ('receipt', self.receipt)): self.source(key, path)
        self.record['tables']['actual'] = {'source_id': 'actual', 'format': 'xlsx', 'sheet': 'SYNTHETIC', 'start_row': 2, 'end_row': 2, 'columns': {'employee_key': 'A', 'name': 'B', 'company': 'C'}}
        self.binding = {'kind': 'xlsx', 'registry_source_id': 'registry', 'registry_sha256': sha256(self.registry), 'plan_source_id': 'plan', 'plan_sha256': sha256(self.plan), 'receipt_source_id': 'receipt', 'receipt_sha256': sha256(self.receipt)}
    def source(self, key, path):
        self.record['sources'][key] = {'path': str(path), 'sha256': sha256(path), 'month': '2099-01', 'role': 'SYNTHETIC'}
    def tearDown(self):
        if self.root.parent != Path(tempfile.gettempdir()).resolve(): raise RuntimeError('测试隔离目录越界')
        self.temporary.cleanup()
    def check(self):
        a = Audit(self.record, self.root); a.stages = ('attendance',)
        a.load_sources(); a.load_tables(); a.template_receipts()
        return a
    def bind(self): self.record['template_bindings']['attendance'] = self.binding
    def test_omitting_template_required_does_not_bypass(self):
        a = self.check()
        self.assertTrue(any('省略template_required不能豁免' in p for p in a.pending))
    def test_real_replay_is_required_and_fake_native_passed_is_ignored(self):
        self.bind(); self.binding['native_recalculation'] = {'passed': True}
        a = self.check()
        self.assertFalse(a.errors, a.errors)
        self.assertTrue(a.evidence['template_receipts']['attendance']['structure']['passed'])
        self.assertTrue(any('字符串passed无效' in p for p in a.pending))
    def test_forged_matching_sha_receipt_does_not_prove_template(self):
        self.bind()
        self.receipt.write_text(json.dumps({'passed': True, 'output': {'path': str(self.output), 'sha256': sha256(self.output)}, 'finalized_output_sha256': sha256(self.output)}), encoding='utf-8')
        self.source('receipt', self.receipt); self.binding['receipt_sha256'] = sha256(self.receipt)
        self.assertTrue(any('实际模板重放拒绝' in e for e in self.check().errors))
    def test_changed_plan_rehashed_source_cannot_reuse_fixed_binding(self):
        self.bind(); self.plan.write_text('{"SYNTHETIC":"CHANGED"}', encoding='utf-8'); self.source('plan', self.plan)
        self.assertTrue(any('固定SHA' in e for e in self.check().errors))
    def test_same_output_path_new_sha_and_forged_receipt_cannot_hide_style_change(self):
        self.bind(); book = load_workbook(self.output); book.active['B2'].number_format = '0.000'; book.save(self.output)
        self.source('actual', self.output)
        receipt = json.loads(self.receipt.read_text(encoding='utf-8')); receipt['output']['sha256'] = sha256(self.output)
        self.receipt.write_text(json.dumps(receipt), encoding='utf-8'); self.source('receipt', self.receipt); self.binding['receipt_sha256'] = sha256(self.receipt)
        self.assertTrue(any('实际模板重放拒绝' in e for e in self.check().errors))
    def test_same_file_cache_refresh_log_is_not_actual_native_proof(self):
        self.bind(); old_sha = sha256(self.output)
        with ZipFile(self.output) as z: parts = {name: z.read(name) for name in z.namelist()}
        changed = False
        for name, value in parts.items():
            if name.startswith('xl/worksheets/'):
                from lxml import etree
                root = etree.fromstring(value)
                for cell in root.findall('.//{*}c'):
                    formula = cell.find('{*}f')
                    if formula is not None and formula.text == 'LEN(A2)':
                        cached = cell.find('{*}v')
                        if cached is None: cached = etree.SubElement(cell, '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')
                        cached.text = '13'; changed = True
                parts[name] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
        self.assertTrue(changed)
        with ZipFile(self.output, 'w', ZIP_DEFLATED) as z:
            for name, data in parts.items(): z.writestr(name, data)
        self.source('actual', self.output)
        log = self.root / 'synthetic-log.json'
        log.write_text(json.dumps({'engine': 'SYNTHETIC-NOT-REAL', 'events': ['SYNTHETIC'], 'started_at': 'SYNTHETIC', 'finished_at': 'SYNTHETIC', 'source_sha256': old_sha, 'output_sha256': sha256(self.output), 'passed': True}), encoding='utf-8')
        self.source('log', log); self.binding.update(recalculation_log_source_id='log', recalculation_log_sha256=sha256(log), native_recalculation={'passed': True})
        a = self.check()
        self.assertFalse(a.errors, a.errors)
        self.assertTrue(a.evidence['template_receipts']['attendance']['structure']['passed'])
        self.assertTrue(any('字符串passed无效' in p for p in a.pending))

if __name__ == '__main__': unittest.main(verbosity=2)
