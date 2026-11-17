"""明确合成：原批注对象保留与事假清算业务处置分别核验。"""
import copy
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from openpyxl import Workbook,load_workbook
from openpyxl.comments import Comment

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts','payroll-operations/scripts','skills/payroll-operations/scripts','candidate/skills/payroll-operations/scripts'):
            path=parent/relative/'crossstage_audit.py'
            if path.is_file():return path.parent
    raise RuntimeError('未找到相邻技能程序')
def load(name):
    spec=importlib.util.spec_from_file_location('_comment_test_'+name,locate()/(name+'.py'));module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module);return module
CROSS=load('crossstage_audit');COMMENTS=load('template_comments')

class CommentBusiness(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(prefix='synthetic-comment-business-');self.root=Path(self.temp.name)
        self.person={'employee_key':'SYNTHETIC-ONE','name':'SYNTHETIC-ONE','company':'SYNTHETIC'}
        self.source=self.root/'source.xlsx';book=Workbook();sheet=book.active;sheet.title='原工资'
        for cell,value in {'A1':'主键','B1':'姓名','C1':'公司','U1':'请假扣款','A2':self.person['employee_key'],'B2':self.person['name'],'C2':self.person['company']}.items():sheet[cell]=value
        sheet['U2'].comment=Comment('含事假清算扣款，须逐人核本人适用月份。','SYNTHETIC');book.save(self.source)
        self.output=self.root/'output.xlsx';book=Workbook();sheet=book.active;sheet.title='工资表'
        for cell,value in {'A1':'主键','B1':'姓名','C1':'公司','I1':'薪资','S1':'请假扣款','W1':'应发工资','A2':self.person['employee_key'],'B2':self.person['name'],'C2':self.person['company'],'I2':1,'S2':0,'W2':1}.items():sheet[cell]=value
        config={'id':'synthetic-original','source':{'path':str(self.source),'sha256':CROSS.sha256(self.source),'locator':'明确合成原批注'},'regions':[{'sheet':'原工资','first_row':2,'last_row':2,'identity_columns':['B'],'target_sheet':'工资表','target_identity_columns':['B'],'header_rows':[1],'target_header_rows':[1]}]}
        policy={'source_comments':[config],'sheets':{'工资表':{'personnel':{'start_row':2}}}}
        inventory=COMMENTS.inventory_sources([config],self.root)
        self.inventory=COMMENTS.apply_source_comments(book,policy,{},None,self.root,inventory);book.save(self.output)
        self.record={'month':'2099-01','scope_table':'scope','sources':{},'tables':{},'stages':{'payroll':{'table':'main','dispositions':[{'employee_key':self.person['employee_key'],'route':'output','row':2}]}},'facts':[],'comment_dispositions':[]}
        self.table('scope',[self.person]);self.record['sources']['main']={'path':str(self.output),'sha256':CROSS.sha256(self.output),'month':'2099-01','role':'draft_output'}
        self.record['tables']['main']={'source_id':'main','format':'xlsx','sheet':'工资表','start_row':2,'end_row':2,'columns':{'employee_key':'A','name':'B','company':'C','salary':'I','leave_deduction':'S','gross':'W'}}
        for item,value in (('salary',1),('leave_deduction',0),('gross',1)):
            ref={'table':'main','row':2,'column':item};self.record['facts'].append({'employee_key':self.person['employee_key'],'stage':'payroll','item':item,'status':'calculated','amount':value,'source_refs':[ref],'value_refs':[ref]})
    def tearDown(self):self.temp.cleanup()
    def table(self,key,rows,month='2099-01'):
        path=self.root/(key+'.json');path.write_text(json.dumps(rows,ensure_ascii=False),encoding='utf-8')
        self.record['sources'][key]={'path':str(path),'sha256':CROSS.sha256(path),'month':month,'role':'existing_evidence'}
        self.record['tables'][key]={'source_id':key,'format':'json','start_row':1,'end_row':len(rows),'columns':{k:k for r in rows for k in r}}
    def decision(self,text='本人事假清算已结清，本月不重复扣款。',action='historical_not_applicable',person=None,month='2099-01'):
        self.table('existing',[dict(person or self.person,month=month,note=text,amount=0)],month)
        disposition={'employee_key':self.person['employee_key'],'source_sha256':CROSS.sha256(self.source),'source_sheet':'原工资','source_cell':'U2','action':action,'reason':'引用既有本人本月处置，不要求重复批准','evidence_refs':[{'table':'existing','row':1,'column':'note'}]}
        if action=='adopted':disposition['amount_ref']={'table':'existing','row':1,'column':'amount'}
        self.record['comment_dispositions']=[disposition]
    def audit(self):
        audit=CROSS.Audit(self.record,self.root);audit.load_sources();audit.load_tables();audit.scope();audit.facts={(f['employee_key'],f['stage'],f['item']):f for f in self.record['facts']}
        audit.evidence['template_receipts']={'payroll':{'structure':{'source_comment_inventory':self.inventory},'technical_verified':True}}
        audit.comment_business_review();return audit
    def test_copied_object_without_resolution_is_pending_and_numeric_output_rejected(self):
        self.assertEqual(self.inventory['sources'][0]['items'][0]['status'],'copied_employee')
        before=(CROSS.sha256(self.source),CROSS.sha256(self.output));audit=self.audit()
        self.assertTrue(audit.pending);self.assertTrue(audit.errors);self.assertFalse(audit.evidence['template_receipts']['payroll']['technical_verified'])
        self.assertEqual(audit.facts[('SYNTHETIC-ONE','payroll','leave_deduction')]['status'],'pending');self.assertEqual(audit.facts[('SYNTHETIC-ONE','payroll','gross')]['status'],'pending')
        self.assertEqual(audit.facts[('SYNTHETIC-ONE','payroll','salary')]['amount'],1)
        self.assertEqual(before,(CROSS.sha256(self.source),CROSS.sha256(self.output)))
    def test_pending_real_cells_are_blank_and_known_salary_is_preserved(self):
        book=load_workbook(self.output);book['工资表']['S2']=None;book['工资表']['W2']=None;book.save(self.output);self.record['sources']['main']['sha256']=CROSS.sha256(self.output)
        for fact in self.record['facts']:
            if fact['item']!='salary':fact.update(status='pending',amount=None)
        audit=self.audit();self.assertTrue(audit.pending);self.assertFalse(audit.errors);self.assertEqual(audit.evidence['known_subtotals']['SYNTHETIC-ONE']['known_gross_subtotal'],'1')
    def test_pending_flags_alone_cannot_hide_numeric_actual_cells(self):
        for fact in self.record['facts']:
            if fact['item']!='salary':fact.update(status='pending',amount=None)
        audit=self.audit();self.assertTrue(audit.technical_errors);self.assertFalse(audit.evidence['template_receipts']['payroll']['technical_verified'])
    def test_actual_same_person_current_month_existing_resolution_passes(self):
        self.decision();audit=self.audit();self.assertFalse(audit.pending);self.assertFalse(audit.errors)
    def test_existing_adopted_amount_must_match_actual_fact(self):
        self.decision('本人事假清算已计入，本月完整请假扣款核定为零。','adopted');self.record['comment_dispositions'][0]['amount_scope']='total_leave';self.assertFalse(self.audit().pending)
        self.record['facts'][1]['amount']=1;self.assertTrue(self.audit().errors)
    def component_case(self):
        self.decision('本人事假清算本月已计入核定组成明细。','adopted')
        rows=[dict(self.person,month='2099-01',component_id='SETTLEMENT',amount=-80,note='本人事假清算本月已计入核定组成明细。',total_leave=-100),dict(self.person,month='2099-01',component_id='SICK',amount=-20,note='本人本月病假既有已核明细。',total_leave=-100)]
        self.table('existing',rows)
        decision=self.record['comment_dispositions'][0]
        decision['component_refs']=[{'table':'existing','row':n,'column':'amount'} for n in (1,2)]
        decision['total_leave_ref']={'table':'existing','row':1,'column':'total_leave'}
        book=load_workbook(self.output);book['工资表']['S2']=-100;book['工资表']['W2']=-99;book.save(self.output);self.record['sources']['main']['sha256']=CROSS.sha256(self.output)
        self.record['facts'][1]['amount']=-100;self.record['facts'][2]['amount']=-99
        return decision
    def test_settlement_component_and_sick_component_reconcile_to_full_s(self):
        self.component_case();audit=self.audit();self.assertFalse(audit.errors);self.assertFalse(audit.pending)
        item=audit.evidence['comment_business_review']['items'][0];self.assertEqual(item['adopted_amount'],-80);self.assertEqual(item['adopted_total_leave'],-100)
    def test_component_alone_without_complete_reconciliation_is_pending_not_wrong_s_arithmetic(self):
        self.component_case();self.record['comment_dispositions'][0].pop('component_refs')
        audit=self.audit();self.assertTrue(audit.pending);self.assertFalse(any('完整请假扣款总额与实际S事实不符' in e for e in audit.errors))
    def test_omitted_settlement_or_other_actual_component_cannot_pass(self):
        for omitted in (0,1):
            decision=self.component_case();decision['component_refs'].pop(omitted);self.assertTrue(self.audit().pending)
    def test_duplicate_settlement_reference_cannot_pass(self):
        decision=self.component_case();decision['component_refs'].append(copy.deepcopy(decision['component_refs'][0]));self.assertTrue(self.audit().pending)
    def test_actual_s_missing_or_repeating_settlement_is_rejected(self):
        self.component_case()
        for value in (-20,-180):
            book=load_workbook(self.output);book['工资表']['S2']=value;book['工资表']['W2']=1+value;book.save(self.output);self.record['sources']['main']['sha256']=CROSS.sha256(self.output)
            self.record['facts'][1]['amount']=value;self.record['facts'][2]['amount']=1+value
            self.assertTrue(any('完整请假扣款总额与实际S事实不符' in e for e in self.audit().errors))
    def test_total_scope_requires_actual_explicit_total_wording(self):
        self.decision('本人事假清算本月已计入。','adopted');self.record['comment_dispositions'][0]['amount_scope']='total_leave';self.assertTrue(self.audit().pending)
    def test_wrong_person_evidence_stays_pending(self):
        self.decision(person=dict(self.person,employee_key='OTHER',name='OTHER'));self.assertTrue(self.audit().pending)
    def test_wrong_month_evidence_stays_pending(self):
        self.decision(month='2098-12');self.assertTrue(self.audit().pending)
    def test_self_declared_source_month_without_actual_row_period_is_insufficient(self):
        self.decision();self.table('existing',[dict(self.person,note='本人事假清算已结清，本月不重复扣款。',amount=0)])
        self.assertTrue(self.audit().pending)
    def test_generic_all_staff_unconfirmed_is_not_person_resolution(self):
        self.decision('全员未确认，统一以后处理。');self.assertTrue(self.audit().pending)
    def test_self_reported_action_without_actual_evidence_stays_pending(self):
        self.decision();self.record['comment_dispositions'][0]['evidence_refs']=[];self.assertTrue(self.audit().pending)
    def test_ordinary_note_needs_no_new_approval(self):
        book=load_workbook(self.source);book['原工资']['U2'].comment=Comment('人员备注：电子资料已归档。','SYNTHETIC');book.save(self.source)
        book=load_workbook(self.output);book['工资表']['S2'].comment=Comment('人员备注：电子资料已归档。','SYNTHETIC');book.save(self.output);self.record['sources']['main']['sha256']=CROSS.sha256(self.output)
        self.inventory['sources'][0]['source']['sha256']=CROSS.sha256(self.source);self.inventory['sources'][0]['items'][0]['text']='人员备注：电子资料已归档。'
        audit=self.audit();self.assertFalse(audit.pending);self.assertFalse(audit.errors)
    def test_current_month_explanation_can_use_existing_actual_evidence(self):
        self.decision('本人事假清算仅说明历史背景，不影响本月。','explanation');self.assertFalse(self.audit().pending)
    def test_review_csv_contains_real_original_comment_and_person(self):
        audit=self.audit();path=self.root/'review.csv';CROSS.write_comment_review_csv({'evidence':audit.evidence},path)
        text=path.read_text(encoding='utf-8-sig');self.assertIn('SYNTHETIC-ONE',text);self.assertIn('含事假清算扣款',text);self.assertIn('pending',text)

if __name__=='__main__':unittest.main(verbosity=2)
