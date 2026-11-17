"""明确合成人员原型测试，不含实际人员、月份、工资或机器绝对路径。"""
from pathlib import Path
from copy import deepcopy
import sys,tempfile,unittest
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Border,Side

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts','payroll-operations/scripts','skills/payroll-operations/scripts','candidate/skills/payroll-operations/scripts'):
            p=parent/relative/'template_guard.py'
            if p.is_file():return p.parent
    raise RuntimeError('未找到相邻技能template_guard.py')

sys.path.insert(0,str(locate()))
import template_guard as GUARD

class PersonnelPrototype(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(prefix='synthetic-personnel-prototype-');self.root=Path(self.tmp.name).resolve()
        source=self.root/'合成依据.txt';source.write_text('明确合成工程依据，非业务人工确认',encoding='utf8')
        self.ev={'path':str(source),'sha256':GUARD.sha256(source),'locator':'全文件合成依据'}
        self.template=self.root/'合成母版.xlsx';b=Workbook();s=b.active;s.title='人员表';s['A1']='明确合成母版';s['D1']='姓名'
        s['D2']='从《花名册》提取';s.row_dimensions[2].height=96
        s['D3'].border=Border(bottom=Side(style='thin'));s.row_dimensions[3].height=18
        s['D4']='示例员工甲';s['D4'].border=Border(bottom=Side(style='double'));s.row_dimensions[4].height=24
        s['A5']='合成尾行';b.save(self.template)
        self.policy={'sheets':{'人员表':{'personnel':{'start_row':2,'end_row':4,'identity_columns':['D']},'writable_ranges':['D2:D4']}}};self.count=0
    def tearDown(self):
        if self.root.parent!=Path(tempfile.gettempdir()).resolve():raise RuntimeError('临时目录越界')
        self.tmp.cleanup()
    def build(self,plan,policy=None,template=None):
        self.count+=1;out=self.root/str(self.count);out.mkdir();template=template or self.template
        registry=out/'registry.json';pp=out/'plan.json';GUARD.write_json(pp,plan)
        r=GUARD.register(template,GUARD.sha256(template),policy or self.policy,registry)
        output=out/'output.xlsx';receipt=out/'receipt.json';GUARD.build(registry,r['registry_sha256'],pp,output,receipt)
        audit=GUARD.verify(registry,r['registry_sha256'],pp,output,receipt);self.assertTrue(audit['passed'],audit['errors'])
        return load_workbook(output),GUARD.read_json(registry)
    def write(self,row,value,old=None):
        return {'sheet':'人员表','cell':f'D{row}','expected_old':old,'value':value,'source':self.ev,'authorization':self.ev,'reason':'明确合成原型测试'}
    def plan(self,rows,writes):
        return {'row_maps':{'人员表':rows},'row_change_sources':{'人员表':self.ev},'writes':writes}
    def test_instruction_row_cannot_be_new_employee_prototype(self):
        with self.assertRaisesRegex(ValueError,'说明行'):
            self.build(self.plan([1,{'source_row':2,'mode':'new'},5],[self.write(2,'合成新员工甲')]))
    def test_normal_row_three_can_generate_first_person_in_row_two(self):
        b,reg=self.build(self.plan([1,{'source_row':3,'mode':'new'},5],[self.write(2,'合成新员工甲')]))
        self.assertEqual(b['人员表']['D2'].value,'合成新员工甲');self.assertEqual(b['人员表'].row_dimensions[2].height,18)
        self.assertTrue(any(x['source_row']==2 for x in reg['personnel_instruction_rows']['人员表']))
    def test_identity_write_on_instruction_outside_personnel_region_is_rejected(self):
        p=deepcopy(self.policy);p['sheets']['人员表']['personnel']['start_row']=3
        with self.assertRaisesRegex(ValueError,'说明行'):self.build({'writes':[self.write(2,'合成新员工甲','从《花名册》提取')]},policy=p)
    def test_preserving_instruction_unchanged_is_allowed(self):
        b,_=self.build({});self.assertEqual(b['人员表']['D2'].value,'从《花名册》提取');self.assertEqual(b['人员表'].row_dimensions[2].height,96)
    def test_two_legal_prototypes_and_example_name_remain_allowed(self):
        b,_=self.build(self.plan([1,{'source_row':3,'mode':'new'},{'source_row':4,'mode':'new'},5],
            [self.write(2,'合成新员工甲'),self.write(3,'合成新员工乙')]))
        s=b['人员表'];self.assertEqual(s['D2'].border.bottom.style,'thin');self.assertEqual(s['D3'].border.bottom.style,'double')
        self.assertEqual(s.row_dimensions[2].height,18);self.assertEqual(s.row_dimensions[3].height,24)
    def test_registering_only_instruction_region_does_not_authorize_employee(self):
        p=deepcopy(self.policy);p['sheets']['人员表']['personnel']['end_row']=2
        with self.assertRaisesRegex(ValueError,'说明行'):
            self.build(self.plan([1,{'source_row':2,'mode':'new'},3,4,5],[self.write(2,'合成新员工甲')]),policy=p)
    def test_another_row_and_identity_column_are_detected(self):
        path=self.root/'不同布局.xlsx';b=Workbook();s=b.active;s.title='不同布局';s['A1']='合成头'
        s['H4']='请填写员工姓名';s['H5'].border=Border(bottom=Side(style='thin'));s['A6']='合成尾';b.save(path)
        p={'sheets':{'不同布局':{'personnel':{'start_row':4,'end_row':5,'identity_columns':['H']},'writable_ranges':['H4:H5']}}}
        plan={'row_maps':{'不同布局':[1,2,3,{'source_row':4,'mode':'new'},6]},'row_change_sources':{'不同布局':self.ev},
            'writes':[{'sheet':'不同布局','cell':'H4','expected_old':None,'value':'合成新员工甲','source':self.ev,'reason':'明确合成不同布局'}]}
        with self.assertRaisesRegex(ValueError,'原行4'):self.build(plan,policy=p,template=path)

if __name__=='__main__':unittest.main(verbosity=2)
