"""明确合成保险兼容反例，无真实姓名、原工资或私人路径。
原制度 social-insurance-fund/SKILL.md 标准四类、Current Calculation Rules、
Output Meaning：单位社保=养老+失业；单位医保=基本医保+生育；单位基金来自原件。
测试均创建实际JSON/XLSX并绑定当前SHA，不伪造工伤单位0。
"""
import copy
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from openpyxl import Workbook

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts', 'payroll-operations/scripts', 'skills/payroll-operations/scripts', 'candidate/skills/payroll-operations/scripts'):
            path=parent/relative/'crossstage_audit.py'
            if path.is_file(): return path
    raise RuntimeError('未找到相邻技能程序')

def load(path):
    spec=importlib.util.spec_from_file_location('synthetic_insurance_'+path.stem,path)
    module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module);return module

PROGRAM=load(locate())
CORE=load(locate().with_name('payroll_core.py'))

class InsuranceCompatibility(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(prefix='synthetic-insurance-public-')
        self.root=Path(self.temp.name);self.person={'employee_key':'SYNTHETIC-KEY','name':'SYNTHETIC-PERSON','company':'SYNTHETIC-COMPANY'}
        # 完全合成：基数10000，8%与0.7%仅用来展示原制度已有的两种表示。
        self.values={'pension':(800,1600),'unemployment':(50,50),'medical':(200,870),'housing_fund':(400,400)}
        self.raw=dict(self.person,month='2099-01',medical_employer_mode='combined')
        for kind,(employee,employer) in self.values.items():self.raw.update({kind+'_employee':employee,kind+'_employer':employer})
        self.components={k:{'employee':k+'_employee','employer':k+'_employer'} for k in self.values}
        self.pay=dict(self.person,personal_social=850,personal_medical=200,personal_fund=400,employer_social=1650,employer_medical=870,employer_fund=400)
        self.insurance_status={};self.employer_status={};self.pay_status={};self.optional_facts=[];self.extra_raw=[];self.omit_pay_columns=set()
    def tearDown(self):self.temp.cleanup()
    def run_fragment(self, mutate_record=None, write_sources=True):
        record={'month':'2099-01','through_stage':'payroll','sources':{},'tables':{},'stages':{'payroll':{'table':'payroll','dispositions':[{'employee_key':'SYNTHETIC-KEY','route':'output','row':2}]}},'facts':[],'raw_insurance':{'table':'insurance','components':copy.deepcopy(self.components),'medical_employer_mode_column':'medical_employer_mode'}}
        source=self.root/'synthetic-insurance.json';source.write_text(json.dumps([self.raw]+self.extra_raw,ensure_ascii=False),encoding='utf8')
        record['sources']['insurance']={'path':str(source),'sha256':PROGRAM.sha256(source),'month':'2099-01','role':'raw'}
        record['tables']['insurance']={'source_id':'insurance','format':'json','start_row':1,'end_row':1+len(self.extra_raw),'columns':{k:k for row in [self.raw]+self.extra_raw for k in row}}
        paypath=self.root/'synthetic-payroll.xlsx';book=Workbook();sheet=book.active;sheet.title='SYNTHETIC-PAYROLL'
        cols={'employee_key':'A','name':'B','company':'C','personal_social':'D','personal_medical':'E','personal_fund':'F','employer_social':'J','employer_medical':'K','employer_fund':'L'}
        cols={k:c for k,c in cols.items() if k not in self.omit_pay_columns}
        for key,col in cols.items():sheet[col+'1']=key;sheet[col+'2']=self.pay.get(key)
        book.save(paypath);book.close()
        record['sources']['payroll']={'path':str(paypath),'sha256':PROGRAM.sha256(paypath),'month':'2099-01','role':'calculation'}
        record['tables']['payroll']={'source_id':'payroll','format':'xlsx','sheet':sheet.title,'start_row':2,'end_row':2,'columns':cols}
        for kind in self.values:
            employee=self.raw.get(kind+'_employee');employer=self.raw.get(kind+'_employer')
            fact={'employee_key':'SYNTHETIC-KEY','stage':'insurance','item':kind,'status':self.insurance_status.get(kind,'calculated' if employee is not None else 'pending'),'amount':employee,'employer_amount':employer,'employer_status':self.employer_status.get(kind,'calculated' if employer is not None else 'pending'),'source_refs':[{'table':'insurance','row':1,'column':kind+'_employee'}],'value_refs':[{'table':'insurance','row':1,'column':kind+'_employee'}]}
            record['facts'].append(fact)
        for key in self.pay:
            if key in self.person:continue
            ref={'table':'payroll','row':2,'column':key}
            record['facts'].append({'employee_key':'SYNTHETIC-KEY','stage':'payroll','item':key,'status':self.pay_status.get(key,'calculated' if self.pay[key] is not None else 'pending'),'amount':self.pay[key],'source_refs':[ref],'value_refs':[ref] if key in cols else []})
        record['facts'].extend(copy.deepcopy(self.optional_facts))
        if mutate_record:mutate_record(record)
        audit=PROGRAM.Audit(record,self.root);audit.load_sources();audit.load_tables();audit.people={'SYNTHETIC-KEY':self.person};audit.stages=('insurance',)
        # 隔离保险必需项和固定交接，不制造考勤、薪资或税额已确认事实。
        audit.check_facts();audit.raw_insurance();return audit
    def assert_clean(self,audit):self.assertEqual([],audit.errors);self.assertEqual([],audit.pending)
    def test_standard_four_complete_without_injury_or_maternity_zero(self):
        a=self.run_fragment();self.assert_clean(a)
        self.assertEqual({'pension','unemployment','medical','housing_fund'},set(a.evidence['raw_insurance']['SYNTHETIC-KEY']['parts']))
        self.assertEqual('870',a.evidence['employer_insurance_handoff']['SYNTHETIC-KEY']['expected']['employer_medical'])
    def test_split_basic_and_maternity_is_870_and_real_injury_stays_detail(self):
        self.raw.update(medical_employer_mode='split',medical_employer=800,maternity_employer=70,injury_employer=31)
        self.components.update(maternity={'employee':None,'employer':'maternity_employer'},injury={'employee':None,'employer':'injury_employer'})
        a=self.run_fragment();self.assert_clean(a)
        self.assertEqual(31,a.evidence['raw_insurance']['SYNTHETIC-KEY']['parts']['injury']['employer'])
        self.assertIsNone(a.evidence['raw_insurance']['SYNTHETIC-KEY']['parts']['injury']['employee'])
        self.assertEqual('870',a.evidence['employer_insurance_handoff']['SYNTHETIC-KEY']['expected']['employer_medical'])
    def test_combined_870_does_not_add_retained_maternity_70_again(self):
        self.raw['maternity_employer']=70;self.components['maternity']={'employee':None,'employer':'maternity_employer'}
        self.assert_clean(self.run_fragment())
        self.pay['employer_medical']=940
        self.assertTrue(any('单位保险至工资金额不符:employer_medical' in e for e in self.run_fragment().errors))
    def test_corrupt_each_employer_column_is_rejected_with_fresh_source_sha(self):
        for key in ('employer_social','employer_medical','employer_fund'):
            with self.subTest(key=key):
                old=self.pay[key];self.pay[key]=old+17;a=self.run_fragment();self.pay[key]=old
                self.assertTrue(any('单位保险至工资金额不符:'+key in e for e in a.errors),a.errors)
    def test_missing_each_employer_column_cannot_hide_behind_auxiliary_fact(self):
        for key in ('employer_social','employer_medical','employer_fund'):
            with self.subTest(key=key):
                self.omit_pay_columns={key};a=self.run_fragment()
                self.assertTrue(any('单位三栏完成但缺实际工资列:'+key in e for e in a.errors),a.errors)
    def test_missing_employer_amount_preserves_known_employee_deduction(self):
        self.raw['pension_employer']=None;self.pay['employer_social']=None
        a=self.run_fragment();self.assertFalse(a.errors,a.errors)
        self.assertTrue(any('缺单位金额:pension' in p for p in a.pending))
        self.assertEqual(850,a.facts[('SYNTHETIC-KEY','payroll','personal_social')]['amount'])
        self.assertEqual('calculated',a.evidence['raw_insurance']['SYNTHETIC-KEY']['parts']['pension']['employee_status'])
        self.assertEqual('870',a.evidence['employer_insurance_handoff']['SYNTHETIC-KEY']['expected']['employer_medical'])
    def test_missing_employee_amount_preserves_known_employer_amount(self):
        self.raw['medical_employee']=None;self.pay['personal_medical']=None
        a=self.run_fragment();self.assertFalse(a.errors,a.errors)
        self.assertTrue(any('缺个人金额:medical' in p for p in a.pending))
        self.assertEqual('870',a.evidence['employer_insurance_handoff']['SYNTHETIC-KEY']['expected']['employer_medical'])
    def test_missing_unit_amount_cannot_become_completed_payroll_zero(self):
        self.raw['housing_fund_employer']=None;self.pay['employer_fund']=0
        self.assertTrue(any('单位保险待核在工资丢失:employer_fund' in e for e in self.run_fragment().errors))
    def test_unknown_medical_mode_only_blocks_unit_medical(self):
        self.raw['medical_employer_mode']=None;self.pay['employer_medical']=None
        a=self.run_fragment();self.assertFalse(a.errors,a.errors)
        self.assertTrue(any('缺实际源模式' in p for p in a.pending))
        self.assertEqual('1650',a.evidence['employer_insurance_handoff']['SYNTHETIC-KEY']['expected']['employer_social'])
        self.assertEqual(200,a.facts[('SYNTHETIC-KEY','payroll','personal_medical')]['amount'])
    def test_split_without_real_maternity_amount_remains_pending(self):
        self.raw.update(medical_employer_mode='split',medical_employer=800);self.pay['employer_medical']=None
        a=self.run_fragment();self.assertFalse(a.errors,a.errors)
        self.assertTrue(any('单位保险前置待核:employer_medical' in p for p in a.pending))
    def test_optional_unavailable_injury_is_detail_not_payroll_required(self):
        self.components['injury']={'employee':None,'employer':None}
        a=self.run_fragment();self.assert_clean(a)
        self.assertEqual('detail_pending',a.evidence['raw_insurance']['SYNTHETIC-KEY']['parts']['injury']['detail_status'])
    def test_wrong_month_company_duplicate_or_sha_still_rejected(self):
        old=copy.deepcopy(self.raw)
        self.raw['month']='2099-02';self.assertTrue(any('归属期' in e for e in self.run_fragment().errors));self.raw=copy.deepcopy(old)
        self.raw['company']='OTHER-SYNTHETIC';self.assertTrue(any('身份未唯一匹配' in e for e in self.run_fragment().errors));self.raw=copy.deepcopy(old)
        self.extra_raw=[copy.deepcopy(self.raw)];self.assertTrue(any('保险多行' in e for e in self.run_fragment().errors));self.extra_raw=[]
        a=self.run_fragment(lambda r:r['sources']['insurance'].update(sha256='not-current-sha'))
        self.assertTrue(any('来源变化' in e for e in a.errors))
    def test_employer_cost_does_not_change_existing_payroll_buckets(self):
        inputs=dict.fromkeys(('fixed_salary','other_allowance','performance','reward','leave_deduction','attendance_deduction','hire_exit_difference','other','personal_deduction','personal_social','personal_medical','personal_fund','tax','combined_tax','third_party','cash_other'),0)
        inputs.update(fixed_salary=4000,personal_social=85,personal_medical=20,personal_fund=40)
        before=CORE.payroll_totals(inputs);inputs.update(employer_social=1650,employer_medical=870,employer_fund=400)
        self.assertEqual(before,CORE.payroll_totals(inputs))

if __name__=='__main__':unittest.main()
