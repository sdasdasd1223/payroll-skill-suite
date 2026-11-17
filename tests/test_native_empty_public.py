"""明确合成缓存类型测试，不含真实人员、工资或机器绝对路径。"""
import importlib.util
from pathlib import Path
import unittest

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts', 'payroll-operations/scripts', 'skills/payroll-operations/scripts', 'candidate/skills/payroll-operations/scripts'):
            path = parent / relative / 'native_recalc_verify.py'
            if path.is_file(): return path
    raise RuntimeError('未找到相邻技能的native_recalc_verify.py')

spec = importlib.util.spec_from_file_location('_synthetic_native_cache', locate())
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)

def metadata(cell):
    xml = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">' + cell + '</row></sheetData></worksheet>'
    return native.formula_cache_metadata(xml.encode('utf-8'))['A1']

class EmptyCache(unittest.TestCase):
    def test_typed_empty_string_with_v_is_present(self):
        value = metadata('<c r="A1" t="str"><f>IF(TRUE,"","")</f><v/></c>')
        self.assertTrue(value['cache_present'])
        self.assertTrue(value['empty_string_cache'])
    def test_typed_string_without_v_is_missing(self):
        value = metadata('<c r="A1" t="str"><f>IF(TRUE,"","")</f></c>')
        self.assertFalse(value['cache_present'])
        self.assertFalse(value['empty_string_cache'])
    def test_untyped_empty_v_is_missing_not_zero(self):
        value = metadata('<c r="A1"><f>1-1</f><v/></c>')
        self.assertFalse(value['cache_present'])
        self.assertFalse(value['empty_string_cache'])
    def test_explicit_numeric_zero_cache_is_present(self):
        value = metadata('<c r="A1"><f>1-1</f><v>0</v></c>')
        self.assertTrue(value['cache_present'])
        self.assertFalse(value['empty_string_cache'])

if __name__ == '__main__': unittest.main(verbosity=2)
