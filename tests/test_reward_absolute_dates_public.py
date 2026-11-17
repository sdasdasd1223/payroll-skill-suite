"""明确合成申请日期；不能用晚起点移动源件已写明的绝对结束日期。"""
import importlib.util
from pathlib import Path
import unittest

def load():
    for parent in Path(__file__).resolve().parents:
        for rel in ('skills/attendance-calculator/scripts/reward_policy.py','candidate/skills/attendance-calculator/scripts/reward_policy.py'):
            path=parent/rel
            if path.is_file():
                spec=importlib.util.spec_from_file_location('_reward_absolute',path);m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m
    raise RuntimeError('未找到相邻分配的奖励规则')
RULE=load()
class Dates(unittest.TestCase):
    def setUp(self):
        self.c={'name':'明确合成人员甲','date':'2026-07-23','tier':200,'checkout':'2026-07-24T00:03:00+08:00','next_arrival':'2026-07-24T08:58:00+08:00','next_workday':True,'application_status':'matched','application_source':'合成申请原件','application_record_id':'SYNTHETIC','application_start':'2026-07-23T18:30:00+08:00','application_end':'2026-07-24T00:00:00+08:00'}
    def evaluate(self,**change):return RULE.evaluate_claim(dict(self.c,**change),'2026-07')
    def test_actual_same_night_valid_registration_needs_no_approval_field(self):self.assertEqual(self.evaluate()['amount'],200)
    def test_previous_night_cannot_be_shifted_to_claimed_night(self):
        a=self.evaluate(application_start='2026-07-22T18:30:00+08:00',application_end='2026-07-23T00:00:00+08:00');self.assertEqual((a['status'],a['amount']),('pending',None))
    def test_existing_same_business_day_midnight_notation_is_preserved(self):
        a=self.evaluate(application_end='2026-07-23T00:00:00+08:00');self.assertEqual(a['amount'],200);self.assertIn('application_midnight_interpretation',a)
    def test_explicit_midnight_without_start_cannot_move_date(self):self.assertIsNone(self.evaluate(application_start=None,application_end='2026-07-23T00:00:00+08:00')['amount'])
    def test_existing_time_only_local_registration_stays_supported(self):self.assertEqual(self.evaluate(application_start='18:30',application_end='00:00')['amount'],200)
    def test_declared_absolute_utc_is_normalized_before_comparing_workday(self):self.assertEqual(self.evaluate(application_start='2026-07-23T10:30:00Z',application_end='2026-07-23T16:00:00Z')['amount'],200)
    def test_midnight_notation_uses_normalized_evening_start_clock(self):self.assertEqual(self.evaluate(application_start='2026-07-23T10:30:00Z',application_end='2026-07-22T16:00:00Z')['amount'],200)
    def test_foreign_evening_text_is_not_local_evening(self):self.assertIsNone(self.evaluate(application_start='2026-07-22T18:30:00-08:00',application_end='2026-07-22T08:00:00-08:00')['amount'])
    def test_missing_or_unparseable_end_stays_pending(self):self.assertIsNone(self.evaluate(application_end='2026-07-23T99:00:00+08:00')['amount'])
    def test_earlier_end_than_actual_start_is_not_valid_registration(self):self.assertIsNone(self.evaluate(application_end='2026-07-23T18:00:00+08:00')['amount'])
    def test_before_midnight_application_is_still_rejected(self):self.assertEqual(self.evaluate(application_end='2026-07-23T23:30:00+08:00')['amount'],0)
    def test_pending_claim_keeps_personal_total_unknown(self):
        c=dict(self.c,application_start='2026-07-22T18:30:00+08:00',application_end='2026-07-23T00:00:00+08:00');a=RULE.audit_claims([c],'2026-07');self.assertEqual(a['totals'][c['name']],{'amount':None,'status':'pending'})
if __name__=='__main__':unittest.main(verbosity=2)
