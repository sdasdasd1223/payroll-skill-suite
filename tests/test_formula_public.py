"""明确合成工资表：待核不免除公式，审计计划不能删掉必查字段。"""
import importlib.util,tempfile,unittest
from pathlib import Path
from openpyxl import Workbook

ROOT=Path(__file__).resolve().parents[1]
if (ROOT/'skills').is_dir():SKILLS=ROOT/'skills'
elif (ROOT/'candidate/skills').is_dir():SKILLS=ROOT/'candidate/skills'
else:raise RuntimeError('请放在安装包tests内，或本轮formula-audit-tests内执行')
spec=importlib.util.spec_from_file_location('formula_audit_test',SKILLS/'payroll-operations/scripts/audit_payroll_workbook.py')
audit=importlib.util.module_from_spec(spec);spec.loader.exec_module(audit)

class FormulaTests(unittest.TestCase):
 def run_case(self,value,pending=True,exemption=None,formula_fields=None,full=False):
  with tempfile.TemporaryDirectory() as folder:
   w=Workbook();s=w.active;s.title='明确合成工资'
   s.append(['姓名','请假扣款']+(['事假时间','病假时间'] if full else []));s.append(['明确合成人员',value]+([None,None] if full else []));s.append(['合计',None])
   p=Path(folder)/'明确合成.xlsx';w.save(p)
   plan={'sheet':s.title,'header_row':1,'first_row':2,'last_row':2,'total_row':3,'pending_rows':[2] if pending else [],'formula_fields':formula_fields if formula_fields is not None else ['请假扣款'],'complete_total_fields':[]}
   if exemption:plan['pending_formula_exemptions']={'B2':exemption}
   return audit.audit(p,plan)
 def test_pending_constant_rejected(self):self.assertFalse(self.run_case(0)['passed'])
 def test_pending_formula_kept(self):self.assertTrue(self.run_case('=SUM(C2:D2)')['passed'])
 def test_plan_omission_cannot_skip(self):self.assertFalse(self.run_case(0,formula_fields=[])['passed'])
 def test_missing_rule_must_be_explicit(self):self.assertTrue(self.run_case(None,exemption={'reason':'明确合成制度待核','source_refs':[{'table':'明确合成源','row':1,'column':'制度'}]})['pending_totals'])
 def test_exemption_no_constant(self):self.assertFalse(self.run_case(0,exemption={'reason':'明确合成待核','source_refs':[{'table':'明确合成源','row':1,'column':'制度'}]})['passed'])
 def test_nonpending_cannot_exempt(self):self.assertFalse(self.run_case(None,pending=False,exemption={'reason':'明确合成待核','source_refs':[{'table':'明确合成源','row':1,'column':'制度'}]})['passed'])
 def test_wrong_person_row_rejected(self):self.assertFalse(self.run_case('=SUM(C3:D3)',full=True)['passed'])
 def test_literal_ref_cannot_spoof(self):self.assertFalse(self.run_case('=IF("C2 D2"="",0,1)',full=True)['passed'])
 def test_other_sheet_cannot_spoof(self):self.assertFalse(self.run_case('=SUM(其他表!C2:D2)',full=True)['passed'])
 def test_own_editable_inputs(self):self.assertTrue(self.run_case('=SUM(C2:D2)',full=True)['passed'])

if __name__=='__main__':unittest.main(verbosity=2)
