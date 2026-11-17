"""明确合成正文日期检查；不含任何真实录用函或员工。"""
import importlib.util,unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]
SKILLS=ROOT/'skills' if (ROOT/'skills').is_dir() else ROOT/'candidate/skills'
sp=importlib.util.spec_from_file_location('offer_evidence_test',SKILLS/'payroll-operations/scripts/offer_evidence.py');m=importlib.util.module_from_spec(sp);sp.loader.exec_module(m)
class OfferTests(unittest.TestCase):
 def test_body_date(self):self.assertEqual(m.parse_pages(['明确合成人员 入职报到日期：2026 年 6 月 15 日。打印日期2026/6/25'],'明确合成人员')['reporting_date'],'2026-06-15')
 def test_slash(self):self.assertEqual(m.parse_pages(['明确合成人员 报到日期:2026/7/1'],'明确合成人员')['reporting_date'],'2026-07-01')
 def test_filename_no_fallback(self):self.assertIsNone(m.parse_pages(['明确合成人员文件名_20260625；没有正文日期字段'],'明确合成人员')['reporting_date'])
 def test_conflict(self):self.assertIsNone(m.parse_pages(['明确合成人员 入职日期2026-06-15','报到日期2026-06-25'],'明确合成人员')['reporting_date'])
 def test_invalid(self):self.assertIsNone(m.parse_pages(['明确合成人员 入职日期2026年2月30日'],'明确合成人员')['reporting_date'])
 def test_other_person(self):self.assertIsNone(m.parse_pages(['明确合成乙 入职日期2026/6/15'],'明确合成甲')['reporting_date'])
 def test_no_skipping_other_clause(self):self.assertIsNone(m.parse_pages(['明确合成人员 报到时间待定。合同签订日期2026/6/25'],'明确合成人员')['reporting_date'])
if __name__=='__main__':unittest.main(verbosity=2)
