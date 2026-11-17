"""明确合成T业务金额依赖与技术交付语义，不含真实人员或真实工资。"""
import copy
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from openpyxl import Workbook,load_workbook

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts','payroll-operations/scripts','skills/payroll-operations/scripts','candidate/skills/payroll-operations/scripts'):
            path=parent/relative/'attendance_link.py'
            if path.is_file():return path.parent
    raise RuntimeError('未找到相邻技能')
def load(name):
    spec=importlib.util.spec_from_file_location('_test_'+name,locate()/(name+'.py'));value=importlib.util.module_from_spec(spec);spec.loader.exec_module(value);return value
LINK=load('attendance_link');CROSS=load('crossstage_audit')

def caches(path,values):
    with ZipFile(path) as z:parts={n:z.read(n) for n in z.namelist()}
    ns='{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    root=ET.fromstring(parts['xl/worksheets/sheet1.xml'])
    for cell in root.findall('.//'+ns+'c'):
        if cell.get('r') not in values:continue
        v=cell.find(ns+'v')
        if v is None:v=ET.SubElement(cell,ns+'v')
        value=values[cell.get('r')]
        if value is None:cell.set('t','str');v.text=''
        else:cell.attrib.pop('t',None);v.text=str(value)
    parts['xl/worksheets/sheet1.xml']=ET.tostring(root,encoding='utf-8')
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        for key,value in parts.items():z.writestr(key,value)

class AttendanceLinks(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(prefix='synthetic-attendance-link-');self.root=Path(self.tmp.name).resolve()
        self.source=self.root/'source.xlsx';b=Workbook();s=b.active;s.title='SOURCE';s.append(['name','deduction','date']);s.append(['SYNTHETIC-ONE',-3,'2099-01-01']);s.append(['SYNTHETIC-TWO',None,'2099-01-01']);b.save(self.source)
        self.output=self.root/'output.xlsx';b=Workbook();s=b.active;s.title='工资表';s['B1']='姓名';s['T1']='考勤扣款';v=b.create_sheet('考勤 依据');v.append(['name','deduction']);v.append(['SYNTHETIC-ONE',-3]);v.append(['SYNTHETIC-TWO',None])
        rows=[];links=[]
        for r,identity,amount in ((2,'SYNTHETIC-ONE',-3),(3,'SYNTHETIC-TWO',None)):
            s[f'B{r}']=identity;s[f'T{r}']=LINK.formula_for_value(v.title,f'B{r}')
            rows.append({'source_row':r,'view_row':r,'identity':identity,'values':{'name':identity,'deduction':amount},'source_cells':{'name':f'A{r}','deduction':f'B{r}'}})
            links.append({'view_id':'attendance','target_sheet':s.title,'target_cell':f'T{r}','identity':identity,'view_sheet':v.title,'view_row':r,'value_cell':f'B{r}','source_row':r,'formula':s[f'T{r}'].value})
        b.save(self.output);caches(self.output,{'T2':-3,'T3':None})
        self.inventory={'sources':[{'id':'attendance','source':{'path':str(self.source),'sha256':LINK.sha256(self.source),'locator':'明确合成A2:B3'},'source_sheet':'SOURCE','source_columns':{'name':'A','deduction':'B'},'sheet':v.title,'columns':{'name':'A','deduction':'B'},'rows':rows}],'links':links}
        self.targets=[{'sheet':'工资表','cell':f'T{r}','identity_cell':f'B{r}'} for r in (2,3)]
    def tearDown(self):
        if self.root.parent!=Path(tempfile.gettempdir()).resolve():raise RuntimeError('隔离目录越界')
        self.tmp.cleanup()
    def audit(self):return LINK.audit_attendance_links(self.output,self.inventory,self.targets)
    def test_real_value_reference_and_blank_source_are_distinguished(self):
        a=self.audit();self.assertFalse(a['errors']);self.assertTrue(a['technical_passed']);self.assertTrue(a['data_pending'])
    def test_identity_lookup_constant_array_is_not_business_link(self):
        b=load_workbook(self.output);b['工资表']['T2']='=IF(B2="SYNTHETIC-ONE",-3,0)';b.save(self.output);caches(self.output,{'T2':-3,'T3':None})
        self.assertTrue(any('身份/常量数组' in x for x in self.audit()['errors']))
    def test_external_source_change_invalidates_sha_without_refreshing_workbook(self):
        before=LINK.sha256(self.output);b=load_workbook(self.source);b.active['B2']=-6;b.save(self.source)
        self.assertTrue(any('SHA失效' in x for x in self.audit()['errors']));self.assertEqual(before,LINK.sha256(self.output))
    def test_other_person_amount_cannot_substitute(self):
        b=load_workbook(self.output);b['考勤 依据']['A2']='SYNTHETIC-TWO';b.save(self.output);caches(self.output,{'T2':-3,'T3':None})
        self.assertTrue(any('不是同人' in x for x in self.audit()['errors']))
    def test_source_blank_cannot_turn_into_numeric_zero(self):
        caches(self.output,{'T2':-3,'T3':0});self.assertTrue(any('空考勤金额' in x for x in self.audit()['errors']))
    def test_business_input_must_be_editable_value_not_formula_constant(self):
        b=load_workbook(self.output);b['考勤 依据']['B2']='=-3';b.save(self.output);caches(self.output,{'T2':-3,'T3':None})
        self.assertTrue(any('可编辑实际数字' in x for x in self.audit()['errors']))
    def test_optional_sheet_quotes_do_not_change_reference_but_sheet_spaces_do(self):
        self.assertTrue(LINK.same_formula('=IF(考勤依据!B2="","",考勤依据!B2)', '=IF(\'考勤依据\'!B2="","",\'考勤依据\'!B2)'))
        self.assertFalse(LINK.same_formula('=IF(\'考勤 依据\'!B2="","",\'考勤 依据\'!B2)', '=IF(\'考勤依据\'!B2="","",\'考勤依据\'!B2)'))
    def record(self):
        person={'employee_key':'SYNTHETIC-KEY','name':'SYNTHETIC-ONE','company':'SYNTHETIC'}
        raw=dict(person,date='2099-01-01',attendance_type='需要打卡',**{k:0 for k in CROSS.LEAVE_ITEMS})
        sources={'original':{'path':str(self.source),'sha256':LINK.sha256(self.source),'month':'2099-01','role':'raw'}};tables={'original':{'source_id':'original','format':'xlsx','sheet':'SOURCE','start_row':2,'end_row':3,'columns':{'name':'A','date':'C'}}}
        for key,rows in (('scope',[person]),('raw',[raw])):
            path=self.root/(key+'.json');path.write_text(json.dumps(rows),encoding='utf-8');sources[key]={'path':str(path),'sha256':LINK.sha256(path),'month':'2099-01','role':'raw'};tables[key]={'source_id':key,'format':'json','start_row':1,'end_row':1,'columns':{k:k for k in rows[0]}}
        return {'schema_version':'crossstage-1','month':'2099-01','policy_version':'SYNTHETIC','mode':'draft','through_stage':'attendance','scope_table':'scope','sources':sources,'tables':tables,'stages':{'attendance':{'dispositions':[{'employee_key':person['employee_key'],'route':'pending','reason':'只测技术语义'}]}},'facts':[],'confirmations':[],'raw_attendance':{'table':'raw','original_table':'original','leave_columns':{k:{'kind':k,'unit':'hours'} for k in CROSS.LEAVE_ITEMS}}}
    def test_default_template_draft_cannot_pass_with_only_calculation_json(self):
        a=CROSS.audit_crossstage(self.record());self.assertTrue(a['calculation_passed']);self.assertFalse(a['passed']);self.assertFalse(a['technical_delivery']['ready']);self.assertEqual(a['technical_delivery']['required_template_stages'],['attendance'])
    def test_self_declared_empty_delivery_subset_does_not_bypass_through_stage(self):
        record=self.record();record['delivery_stages']=[];a=CROSS.audit_crossstage(record)
        self.assertFalse(a['passed']);self.assertEqual(a['technical_delivery']['required_template_stages'],['attendance'])
    def test_explicit_details_mode_preserves_independent_observations(self):
        record=self.record();record['purpose']='calculation_details';a=CROSS.audit_crossstage(record)
        self.assertTrue(a['passed']);self.assertEqual(a['delivery_status'],'calculation_details_only');self.assertFalse(a['formal_ready']);self.assertFalse(a['technical_delivery']['all_five_stages_verified']);self.assertTrue(a['evidence']['raw_attendance']['SYNTHETIC-KEY']['raw_rows'])

if __name__=='__main__':unittest.main(verbosity=2)
