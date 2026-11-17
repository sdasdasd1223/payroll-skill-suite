"""十项明确合成请假公式规则测试；无真实人员或真实工资数据。"""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from openpyxl import Workbook

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts','payroll-operations/scripts','skills/payroll-operations/scripts','candidate/skills/payroll-operations/scripts'):
            path=parent/relative/'payroll_formulas.py'
            if path.is_file(): return path.parent
    raise RuntimeError('未找到相邻工资技能程序')

def load(name):
    spec=importlib.util.spec_from_file_location('_synthetic_'+name,locate()/(name+'.py'))
    value=importlib.util.module_from_spec(spec);spec.loader.exec_module(value);return value

FORMULAS=load('payroll_formulas'); AUDIT=load('audit_payroll_workbook')

class LeaveRules(unittest.TestCase):
    def setUp(self): self.temp=tempfile.TemporaryDirectory(prefix='synthetic-payroll-formulas-');self.root=Path(self.temp.name).resolve();self.counter=0
    def tearDown(self):
        if self.root.parent!=Path(tempfile.gettempdir()).resolve():raise RuntimeError('隔离目录越界')
        self.temp.cleanup()
    def parameters(self,**changes):
        p={'is_intern':False,'annual_before':0,'minimum_wage':None,'personal_input_complete':True,'sick_input_complete':True};p.update(changes)
        self.counter+=1;path=self.root/f'source-{self.counter}.json';path.write_text(json.dumps(p),encoding='utf-8')
        p['source_evidence']={key:{'path':str(path),'sha256':FORMULAS.sha256(path),'locator':'明确合成已审阅参数:'+key} for key in ('salary','is_intern','personal_input','sick_input','required_hours','annual_before','minimum_wage')}
        return p
    def test_regular_or_departing_nonintern_uses_174_and_editable_row_inputs(self):
        p=self.parameters(); f=FORMULAS.leave_deduction_formula(2,p)['formula']
        for ref in ('I2','Q2','R2','V2'):self.assertIn(ref,f)
        self.assertIn('/174*Q2',f);self.assertEqual(FORMULAS.evaluate_leave_deduction(p,174,None,0,160)['amount'],0)
        self.assertEqual(FORMULAS.evaluate_leave_deduction(p,174,1,0,160)['amount'],-1)
    def test_intern_uses_dynamic_v(self):
        p=self.parameters(is_intern=True);self.assertIn('/V2*Q2',FORMULAS.leave_deduction_formula(2,p)['formula'])
        self.assertEqual(FORMULAS.evaluate_leave_deduction(p,168,1,0,168)['amount'],-1)
        self.assertEqual(FORMULAS.evaluate_leave_deduction(p,168,1,0,160)['amount'],-1.05)
    def test_total_72_is_half_pay_without_minimum_wage(self):
        p=self.parameters();self.assertEqual(FORMULAS.evaluate_leave_deduction(p,1740,0,72,174)['amount'],-360)
    def test_current_80_crosses_from_zero_and_is_pending(self):
        p=self.parameters();r=FORMULAS.evaluate_leave_deduction(p,174,1,80,174)
        self.assertEqual(r['status'],'pending');self.assertIsNone(r['amount']);self.assertEqual(r['personal']['amount'],-1)
    def test_71_plus_two_crosses_threshold(self):
        p=self.parameters(annual_before=71);self.assertEqual(FORMULAS.evaluate_leave_deduction(p,174,0,2,174)['status'],'pending')
    def test_above_threshold_reuses_core_minimum_wage_branch(self):
        for before in (72,80):
            p=self.parameters(annual_before=before,minimum_wage=100)
            self.assertEqual(FORMULAS.evaluate_leave_deduction(p,1000,0,1,100)['amount'],-9.2)
    def test_missing_minimum_does_not_block_r_zero_other_deduction(self):
        p=self.parameters(annual_before=80)
        self.assertEqual(FORMULAS.evaluate_leave_deduction(p,174,1,1,174)['status'],'pending')
        self.assertEqual(FORMULAS.evaluate_leave_deduction(p,174,1,0,174)['amount'],-1)
    def test_unknown_annual_is_not_assumed_zero(self):
        p=self.parameters(annual_before=None)
        self.assertEqual(FORMULAS.evaluate_leave_deduction(p,174,0,1,174)['status'],'pending')
        self.assertEqual(FORMULAS.evaluate_leave_deduction(p,174,1,0,174)['amount'],-1)
    def test_complete_vs_unknown_blank_inputs_and_unknown_salary(self):
        p=self.parameters();self.assertEqual(FORMULAS.evaluate_leave_deduction(p,None,None,None,None)['amount'],0)
        self.assertEqual(FORMULAS.evaluate_leave_deduction(p,None,1,0,174)['status'],'pending')
        p=self.parameters(personal_input_complete=False,sick_input_complete=False)
        self.assertEqual(FORMULAS.evaluate_leave_deduction(p,174,None,None,174)['status'],'pending')
    def test_audit_accepts_generated_rule_rejects_fixed_half_or_wrong_row(self):
        p=self.parameters();book=Workbook();ws=book.active;ws.title='SYNTHETIC'
        for cell,label in {'I1':'薪资','Q1':'事假时间','R1':'病假时间','S1':'请假扣款','V1':'应出勤小时数'}.items():ws[cell]=label
        ws['I2']=174;ws['V2']=174;ws['S2']=FORMULAS.leave_deduction_formula(2,p)['formula']
        path=self.root/'synthetic.xlsx';book.save(path)
        plan={'sheet':ws.title,'header_row':1,'first_row':2,'last_row':2,'total_row':3,'pending_rows':[],'formula_fields':['请假扣款'],'complete_total_fields':[],'leave_formula_parameters':{'2':p}}
        self.assertTrue(AUDIT.audit(path,plan)['passed'])
        ws['S2']='=ROUND(-I2/174*N(Q2)-I2/V2*N(R2)/2,2)';book.save(path)
        self.assertFalse(AUDIT.audit(path,plan)['passed'])
        ws['S2']=FORMULAS.leave_deduction_formula(3,p)['formula'];book.save(path)
        self.assertFalse(AUDIT.audit(path,plan)['passed'])
        ws['S2']=FORMULAS.leave_deduction_formula(2,p)['formula'];book.save(path);plan.pop('leave_formula_parameters')
        self.assertFalse(AUDIT.audit(path,plan)['passed'])

if __name__=='__main__':unittest.main(verbosity=2)
