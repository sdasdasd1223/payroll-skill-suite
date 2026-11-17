"""明确合成免打卡既有来源替代日明细覆盖，不含真实人员。"""
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts','payroll-operations/scripts','skills/payroll-operations/scripts','candidate/skills/payroll-operations/scripts'):
            path=parent/relative/'crossstage_audit.py'
            if path.is_file():return path
    raise RuntimeError('未找到相邻技能')
spec=importlib.util.spec_from_file_location('_no_punch_crossstage',locate());CROSS=importlib.util.module_from_spec(spec);spec.loader.exec_module(CROSS)

class NoPunch(unittest.TestCase):
    def run_case(self, evidence=None, daily_type='不需要打卡'):
        with tempfile.TemporaryDirectory(prefix='synthetic-no-punch-') as temp:
            root=Path(temp);person={'employee_key':'SYNTHETIC','name':'SYNTHETIC','company':'SYNTHETIC'}
            record={'month':'2099-01','scope_table':'scope','sources':{},'tables':{},'raw_attendance':{'table':'daily','leave_columns':{k:{'kind':k,'unit':'hours'} for k in CROSS.LEAVE_ITEMS}},'facts':[]}
            rows={'scope':[person],'daily':[dict(person,date='2099-01-01',attendance_type=daily_type,**{k:0 for k in CROSS.LEAVE_ITEMS})]}
            if evidence is not None:
                rows['existing']= [dict(person,attendance_type='不需要打卡',**evidence)]
                record['raw_attendance']['no_punch_refs']={'SYNTHETIC':[{'table':'existing','row':1,'column':'attendance_type'}]}
            for key,values in rows.items():
                path=root/(key+'.json');path.write_text(json.dumps(values),encoding='utf-8')
                record['sources'][key]={'path':str(path),'sha256':CROSS.sha256(path),'month':'2099-01','role':'actual-source'}
                record['tables'][key]={'source_id':key,'format':'json','start_row':1,'end_row':1,'columns':{k:k for k in values[0]}}
            audit=CROSS.Audit(record,root);audit.load_sources();audit.load_tables();audit.scope();audit.raw_attendance()
            return audit.evidence['raw_attendance']['SYNTHETIC'],audit.pending
    def test_existing_same_person_current_month_no_punch_replaces_daily_proof(self):
        obs,pending=self.run_case({'month':'2099-01'})
        self.assertTrue(obs['personal_leave_exempt']);self.assertFalse(obs['coverage_complete']);self.assertIsNone(obs['business_hours']['personal']);self.assertFalse(pending)
    def test_existing_long_term_list_uses_actual_effective_period(self):
        obs,pending=self.run_case({'effective_start':'2098-01-01','effective_end':None})
        self.assertTrue(obs['personal_leave_exempt']);self.assertFalse(pending)
    def test_absence_alone_is_not_exemption(self):
        obs,pending=self.run_case();self.assertFalse(obs['personal_leave_exempt']);self.assertTrue(pending)
    def test_other_person_record_cannot_exempt(self):
        obs,pending=self.run_case({'employee_key':'OTHER','month':'2099-01'})
        self.assertFalse(obs['personal_leave_exempt']);self.assertTrue(pending)
    def test_wrong_month_or_expired_record_cannot_exempt(self):
        for evidence in ({'month':'2098-12'},{'effective_start':'2098-01-01','effective_end':'2099-01-15'}):
            obs,pending=self.run_case(evidence);self.assertFalse(obs['personal_leave_exempt']);self.assertTrue(pending)
    def test_conflicting_or_unknown_actual_daily_status_stays_pending(self):
        for value in ('需要打卡','未知类型'):
            obs,pending=self.run_case({'month':'2099-01'},value);self.assertFalse(obs['personal_leave_exempt']);self.assertTrue(pending)

if __name__=='__main__':unittest.main(verbosity=2)
