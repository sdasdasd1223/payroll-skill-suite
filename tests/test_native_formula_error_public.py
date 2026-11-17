"""可移植明确合成mock：未知金额预期不能掩盖真实公式错误。"""
import importlib.util
from pathlib import Path
import unittest

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts','payroll-operations/scripts','skills/payroll-operations/scripts','candidate/skills/payroll-operations/scripts'):
            path=parent/relative/'native_recalc_verify.py'
            if path.is_file():return path
    raise RuntimeError('未找到相邻技能程序')
spec=importlib.util.spec_from_file_location('_native_error_public',locate());NATIVE=importlib.util.module_from_spec(spec);spec.loader.exec_module(NATIVE)

class Cell:
    def __init__(self,value,text,is_error=False,has_formula=True):
        self.Value2=value;self.Text=text;self.HasFormula=has_formula;self.Parent=self;self.detected=is_error
    def Address(self,*args):return 'S2'
    def Evaluate(self,expression):
        if expression!='ISERROR(S2)':raise ValueError('须核实际同格错误类型')
        if self.detected is None:raise RuntimeError('明确模拟类型接口不可用')
        return self.detected

class NativeErrors(unittest.TestCase):
    def assess(self,before,after,expected=None):
        return NATIVE.assess_linkage_output(NATIVE.observe_cell(before),NATIVE.observe_cell(after),expected)
    def test_actual_value_error_with_unknown_expected_is_technical_failure(self):
        cell=Cell(-2146826273,'#VALUE!',True);observed=NATIVE.observe_cell(cell)
        self.assertEqual(observed['text'],'#VALUE!');self.assertTrue(observed['is_error'])
        status,reason=self.assess(Cell(1,'1'),cell);self.assertEqual(status,'failed');self.assertIn('公式错误',reason)
    def test_real_blank_with_unknown_expected_stays_pending(self):
        status,reason=self.assess(Cell(None,''),Cell('',''));self.assertEqual(status,'pending');self.assertIn('仍为空白',reason)
    def test_normal_numeric_with_unknown_expected_stays_pending(self):
        status,reason=self.assess(Cell(1,'1'),Cell(2,'2'));self.assertEqual(status,'pending');self.assertIn('缺核定数值预期',reason)
    def test_narrow_column_hash_display_is_not_formula_error(self):
        before=Cell(1,'####');after=Cell(2,'########')
        self.assertFalse(NATIVE.observe_cell(after)['is_error']);self.assertEqual(self.assess(before,after)[0],'pending');self.assertEqual(self.assess(before,after,2)[0],'verified')
    def test_com_error_code_with_error_display_is_reliable_fallback(self):
        after=Cell(-2146826273,'#VALUE!',None);observed=NATIVE.observe_cell(after)
        self.assertTrue(observed['is_error']);self.assertEqual(observed['error_type_source'],'COM_error_code_and_display');self.assertEqual(self.assess(Cell(1,'1'),after)[0],'failed')
    def test_literal_error_text_with_false_iserror_is_not_fabricated_failure(self):
        after=Cell('#VALUE!','#VALUE!',False);self.assertFalse(NATIVE.observe_cell(after)['is_error']);self.assertEqual(self.assess(Cell(1,'1'),after)[0],'pending')
    def test_real_error_before_edit_cannot_be_hidden_by_correct_after_value(self):
        self.assertEqual(self.assess(Cell(-2146826273,'#VALUE!',True),Cell(2,'2'),2)[0],'failed')
    def test_nonstandard_error_display_still_detected_by_actual_type(self):
        self.assertEqual(self.assess(Cell(None,''),Cell(-1,'本机错误显示',True))[0],'failed')

if __name__=='__main__':unittest.main(verbosity=2)
