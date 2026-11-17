"""明确合成原假别与最终人资交接绑定，不含真实人员或工资。"""
import copy,importlib.util,json,tempfile,unittest
from pathlib import Path
from openpyxl import Workbook,load_workbook

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts','payroll-operations/scripts','skills/payroll-operations/scripts','candidate/skills/payroll-operations/scripts'):
            path=parent/relative/'attendance_fact_handoff.py'
            if path.is_file():return path
    raise RuntimeError('未找到相邻技能')
spec=importlib.util.spec_from_file_location('_handoff_public',locate());PROGRAM=importlib.util.module_from_spec(spec);spec.loader.exec_module(PROGRAM)

class Handoff(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(prefix='synthetic-handoff-');self.root=Path(self.temp.name)
        people=[{'employee_key':key,'name':key,'company':'SYNTHETIC'} for key in ('SYNTHETIC-ONE','SYNTHETIC-TWO')]
        scope=self.root/'scope.json';scope.write_text(json.dumps(people),encoding='utf8')
        source=self.root/'daily.xlsx';book=Workbook();ws=book.active;ws.title='明确合成原件';ws.append(['姓名','日期','产检假(天)','调休假(分钟)']);ws.append(['SYNTHETIC-ONE','2099-07-08',1,0]);ws.append(['SYNTHETIC-TWO','2099-07-27',0,450]);book.save(source)
        self.record={'month':'2099-07','through_stage':'payroll','scope_table':'scope','sources':{'scope':{'path':str(scope),'sha256':PROGRAM.sha(scope),'month':'2099-07','role':'raw'},'daily':{'path':str(source),'sha256':PROGRAM.sha(source),'month':'2099-07','role':'raw'}},'tables':{'scope':{'source_id':'scope','format':'json','start_row':1,'end_row':2,'columns':{k:k for k in people[0]}},'daily':{'source_id':'daily','format':'xlsx','sheet':'明确合成原件','start_row':2,'end_row':3,'columns':{'name':'A','date':'B','prenatal':'C','compensatory':'D'}}},'raw_attendance':{'table':'daily','header_row':1,'leave_columns':{'prenatal':{'kind':'other','unit':'days'},'compensatory':{'kind':'other','unit':'minutes'}}},'facts':[{'employee_key':p['employee_key'],'stage':'payroll','item':'other_leave','status':'pending','amount':None} for p in people]}
    def tearDown(self):self.temp.cleanup()
    def render(self):
        output=self.root/'rendered';PROGRAM.render_handoff(self.record,self.root,output);return json.loads((output/'run-with-handoff.json').read_text(encoding='utf8'))
    def modify_text(self,record,old,new):
        sid=record['attendance_handoff']['targets'][0]['source_id'];source=record['sources'][sid];path=Path(source['path']);path.write_text(path.read_text(encoding='utf8').replace(old,new),encoding='utf8');source['sha256']=PROGRAM.sha(path)
    def test_same_source_auto_render_preserves_original_facts_with_unknown_treatment(self):
        record=self.render();result=PROGRAM.verify_handoff(record,self.root)
        self.assertTrue(result['passed']);self.assertEqual(len(result['data_pending']),2)
        facts=result['facts'];self.assertEqual([(f['date'],f['original_label'],f['raw_value'],f['raw_unit']) for f in facts],[('2099-07-08','产检假(天)','1','天'),('2099-07-27','调休假(分钟)','450','分钟')])
        self.assertEqual(facts[1]['source']['value_cell'],'D3')
    def test_date_tamper_with_fresh_final_sha_is_rejected(self):
        record=self.render();self.modify_text(record,'2099-07-08','2099-07-28');self.assertFalse(PROGRAM.verify_handoff(record,self.root)['passed'])
    def test_leave_label_tamper_with_fresh_final_sha_is_rejected(self):
        record=self.render();self.modify_text(record,'产检假','哺乳假');self.assertFalse(PROGRAM.verify_handoff(record,self.root)['passed'])
    def test_unit_tamper_with_fresh_final_sha_is_rejected(self):
        record=self.render();self.modify_text(record,'450分钟','450小时');self.assertFalse(PROGRAM.verify_handoff(record,self.root)['passed'])
    def test_quantity_tamper_with_fresh_final_sha_is_rejected(self):
        record=self.render();self.modify_text(record,'450分钟','480分钟');self.assertFalse(PROGRAM.verify_handoff(record,self.root)['passed'])
    def test_correct_side_log_without_final_binding_does_not_pass(self):
        PROGRAM.collect_facts(self.record,self.root);self.assertFalse(PROGRAM.verify_handoff(self.record,self.root)['passed'])
    def test_free_non_data_explanation_does_not_require_extra_approval(self):
        record=self.render();sid=record['attendance_handoff']['targets'][0]['source_id'];source=record['sources'][sid];path=Path(source['path']);path.write_text('其他处理说明可以自由书写。\n'+path.read_text(encoding='utf8'),encoding='utf8');source['sha256']=PROGRAM.sha(path)
        self.assertTrue(PROGRAM.verify_handoff(record,self.root)['passed'])
    def test_excluded_raw_rows_are_still_read_as_original_facts(self):
        self.record['tables']['daily']['excluded_rows']={'2':'明确合成排除声明不能藏原事实'};self.assertEqual(len(PROGRAM.collect_facts(self.record,self.root)['facts']),2)
    def test_actual_final_xlsx_cell_is_replayed_not_a_correct_adjacent_log(self):
        facts=PROGRAM.collect_facts(self.record,self.root)['facts'];ids=[f['fact_id'] for f in facts]
        book=Workbook();book.active.title='合成归并清单';book.active['F10']='错误的旧摘要，仅在隔离修正副本替换';summary=self.root/'summary.xlsx';book.save(summary)
        notes=self.root/'notes.md';start,end=PROGRAM.block_markers('leave');notes.write_text(start+'\n{{attendance_facts:leave}}\n'+end,encoding='utf8')
        for sid,path,role in [('summary',summary,'hr_handoff'),('notes',notes,'run_explanation')]:self.record['sources'][sid]={'path':str(path),'sha256':PROGRAM.sha(path),'month':'2099-07','role':role}
        self.record['attendance_handoff']={'final_source_ids':['summary','notes'],'targets':[{'source_id':'summary','format':'xlsx','sheet':'合成归并清单','cell':'F10','fact_ids':ids},{'source_id':'notes','format':'text','block_id':'leave','fact_ids':ids}]}
        before=PROGRAM.sha(summary);record=self.render();self.assertEqual(PROGRAM.sha(summary),before);self.assertTrue(PROGRAM.verify_handoff(record,self.root)['passed'])
        target=Path(record['sources']['summary']['path']);book=load_workbook(target);book['合成归并清单']['F10']=book['合成归并清单']['F10'].value.replace('450分钟','480分钟');book.save(target);record['sources']['summary']['sha256']=PROGRAM.sha(target)
        self.assertFalse(PROGRAM.verify_handoff(record,self.root)['passed'])
        ctx=PROGRAM.context_for(record,self.root);ctx.attendance_handoff_review();self.assertTrue(ctx.technical_errors)
    def test_existing_final_file_cannot_be_silently_replaced_by_side_log(self):
        path=self.root/'existing.md';path.write_text('原最终说明',encoding='utf8');self.record['sources']['notes']={'path':str(path),'sha256':PROGRAM.sha(path),'month':'2099-07','role':'run_explanation'}
        with self.assertRaisesRegex(ValueError,'已有最终交接文件'):self.render()

if __name__=='__main__':unittest.main(verbosity=2)
