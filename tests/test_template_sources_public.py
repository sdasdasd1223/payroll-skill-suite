"""可移植的明确合成来源页/批注对象测试；不含真实人员、工资或机器绝对路径。

普通运行不需要原生应用；设置PAYROLL_TEST_NATIVE=1可额外跑隔离原生重算与严格缓存搬运。
"""
from pathlib import Path
from copy import deepcopy
from contextlib import contextmanager
import json,os,shutil,sys,tempfile,unittest
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Border,Side
from openpyxl.comments import Comment


def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts','payroll-operations/scripts','skills/payroll-operations/scripts','candidate/skills/payroll-operations/scripts'):
            path=parent/relative/'template_guard.py'
            if path.is_file():return path.parent
    raise RuntimeError('未找到相邻技能的template_guard.py')


sys.path.insert(0,str(locate()))
import template_guard as GUARD
import template_cache as CACHE
from template_source_views import source_view_formula


@contextmanager
def native_application(test):
    """共用生产互斥锁；已有工作簿即跳过，仅恢复属性，从不退出应用。"""
    import ctypes
    from ctypes import wintypes
    import win32com.client,win32process
    import native_recalc_verify as native
    kernel=ctypes.WinDLL('kernel32',use_last_error=True)
    kernel.CreateMutexW.argtypes=[wintypes.LPVOID,wintypes.BOOL,wintypes.LPCWSTR]
    kernel.CreateMutexW.restype=wintypes.HANDLE
    kernel.WaitForSingleObject.argtypes=[wintypes.HANDLE,wintypes.DWORD]
    kernel.ReleaseMutex.argtypes=[wintypes.HANDLE];kernel.CloseHandle.argtypes=[wintypes.HANDLE]
    mutex=kernel.CreateMutexW(None,False,'Local\\CodexPayrollNativeRecalcV1')
    acquired=False;app=None;settings={}
    try:
        acquired=bool(mutex) and kernel.WaitForSingleObject(mutex,0) in (0,0x80)
        if not acquired:test.skipTest('原生应用被其他任务使用，未打开工作簿')
        before=set(win32process.EnumProcesses());app=win32com.client.DispatchEx('Excel.Application')
        pid,executable=native.process_engine(app.Hwnd)
        if not native.engine_identity(executable):test.skipTest('实际应用身份无法证实，未打开工作簿')
        if int(app.Workbooks.Count):test.skipTest('办公实例已有其他工作簿，未打开隔离副本')
        for name in ('Visible','DisplayAlerts','EnableEvents','AskToUpdateLinks','AutomationSecurity','Calculation'):
            try:settings[name]=getattr(app,name)
            except Exception:pass
        if pid not in before:app.Visible=False
        app.DisplayAlerts=False;app.EnableEvents=False;app.AskToUpdateLinks=False;app.AutomationSecurity=3
        yield app
    finally:
        if app is not None:
            for name,value in settings.items():
                try:setattr(app,name,value)
                except Exception:pass
            # DispatchEx可能复用用户或其他测试实例；不对任何应用调用Quit。
            app=None
        if mutex:
            if acquired:kernel.ReleaseMutex(mutex)
            kernel.CloseHandle(mutex)


class TemplateSources(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory(prefix='synthetic-template-sources-');self.root=Path(self.tmp.name).resolve()
        self.source=self.root/'合成来源.xlsx';book=Workbook();sheet=book.active;sheet.title='本月结果'
        sheet.append(['姓名','已算扣款','备注1']);sheet.append(['合成员工甲',-3,None]);sheet.append(['合成员工乙',None,None])
        sheet['C2'].comment=Comment('明确合成原注：保留原文及作者。','合成作者甲');book.save(self.source)
        self.zero=self.root/'整本零批注原件.xlsx';book=Workbook();book.active['A1']='明确合成零批注原件';book.save(self.zero)
        self.template=self.root/'合成指定母版.xlsx';book=Workbook();sheet=book.active;sheet.title='工资表'
        sheet['A1']='明确合成指定母版';sheet['B1']='姓名';sheet['T1']='考勤扣款';sheet['AG1']='备注1'
        sheet['B2'].border=Border(bottom=Side(style='thin'));sheet['AG2'].border=Border(bottom=Side(style='thin'))
        sheet['T2'].number_format='0.00';book.save(self.template)
        self.ev={'path':str(self.source),'sha256':GUARD.sha256(self.source),'locator':'明确合成A2:C3，非真实工资'}
        self.policy={'sheet_policy':{'source_views':[{'id':'attendance','sheet':'考勤依据','source':self.ev,
            'source_sheet':'本月结果','first_row':2,'last_row':3,'fields':{'name':'A','deduction':'B'},
            'link':{'target_sheet':'工资表','identity_column':'B','target_column':'T'}}]},
            'require_source_comments':True,'source_comments':[{'id':'有批注原件','source':self.ev,
              'regions':[{'sheet':'本月结果','first_row':2,'last_row':3,'identity_columns':['A'],'target_sheet':'工资表',
                'target_identity_columns':['B'],'header_rows':[1],'target_header_rows':[1],'fallback_column':'AG'}]},
              {'id':'零批注原件','source':{'path':str(self.zero),'sha256':GUARD.sha256(self.zero),'locator':'全工作簿实际盘点'}}],
            'sheets':{'工资表':{'personnel':{'start_row':2,'end_row':2,'identity_columns':['B']},'writable_ranges':['B2']}}}
        self.plan={'row_maps':{'工资表':[1,2,{'source_row':2,'mode':'new'}]},'row_change_sources':{'工资表':self.ev},
            'writes':[{'sheet':'工资表','cell':f'B{r}','expected_old':None,'value':name,'source':self.ev,
                'reason':'明确合成来源/批注工程测试'} for r,name in [(2,'合成员工甲'),(3,'合成员工乙')]]}
        self.counter=0

    def tearDown(self):
        if self.root.parent!=Path(tempfile.gettempdir()).resolve():raise RuntimeError('临时隔离目录越界')
        self.tmp.cleanup()

    def build(self,title='考勤依据',plan=None):
        self.counter+=1;folder=self.root/str(self.counter);folder.mkdir()
        policy=deepcopy(self.policy);policy['sheet_policy']['source_views'][0]['sheet']=title
        registry=folder/'registry.json';p=folder/'plan.json';GUARD.write_json(p,plan or self.plan)
        r=GUARD.register(self.template,GUARD.sha256(self.template),policy,registry)
        output=folder/'output.xlsx';receipt=folder/'receipt.json'
        result=GUARD.build(registry,r['registry_sha256'],p,output,receipt)
        return {'registry':registry,'registry_sha256':r['registry_sha256'],'plan':p,'output':output,'receipt':receipt},result

    def verify(self,args):return GUARD.verify(**args)

    def forge_receipt_after_change(self,args):
        receipt=GUARD.read_json(args['receipt']);receipt['passed']=True;receipt['status']='passed'
        receipt['output']['sha256']=GUARD.sha256(args['output'])
        args['receipt'].write_text(json.dumps(receipt,ensure_ascii=False),encoding='utf8')

    def test_production_snapshot_links_comments_and_whole_zero_inventory(self):
        args,result=self.build();self.assertTrue(self.verify(args)['passed'])
        book=load_workbook(args['output']);self.assertEqual(book['工资表']['T2'].value,source_view_formula('考勤依据','B2'))
        self.assertIsNone(book['考勤依据']['B3'].value)
        comment=book['工资表']['AG2'].comment
        self.assertEqual((comment.text,comment.author),('明确合成原注：保留原文及作者。','合成作者甲'))
        counts=[x['comment_count'] for x in result['source_comment_inventory']['sources']]
        self.assertEqual(counts,[1,0]);self.assertEqual(len(result['source_view_inventory']['links']),2)

    def test_plain_and_space_names_use_the_same_production_function(self):
        for title in ['考勤依据','考勤 依据']:
            args,_=self.build(title);book=load_workbook(args['output'])
            self.assertEqual(book['工资表']['T2'].value,source_view_formula(title,'B2'))
            self.assertTrue(self.verify(args)['passed'])
        self.assertEqual(source_view_formula('考勤依据','B2'),'=IF(考勤依据!B2="","",考勤依据!B2)')
        self.assertEqual(source_view_formula("考勤'依据",'B2'),"=IF('考勤''依据'!B2=\"\",\"\",'考勤''依据'!B2)")

    def test_snapshot_tamper_and_forged_passed_sha_are_rejected(self):
        args,_=self.build();book=load_workbook(args['output']);book['考勤依据']['B2']=-99;book.save(args['output'])
        self.forge_receipt_after_change(args);self.assertFalse(self.verify(args)['passed'])

    def test_wrong_employee_reference_and_name_constants_are_rejected(self):
        for formula in [source_view_formula('考勤依据','B3'),'=IF(B2="合成员工甲",-3,"")']:
            args,_=self.build();book=load_workbook(args['output']);book['工资表']['T2']=formula;book.save(args['output'])
            self.forge_receipt_after_change(args);self.assertFalse(self.verify(args)['passed'])

    def test_missing_actual_employee_is_rejected(self):
        plan=deepcopy(self.plan);plan['writes'][0]['value']='合成员工丙'
        with self.assertRaises(ValueError):self.build(plan=plan)

    def test_comment_cannot_be_copied_to_other_employee(self):
        plan=deepcopy(self.plan);plan['comment_copies']=[{'source_id':'有批注原件','source_sheet':'本月结果',
            'source_cell':'C2','target_sheet':'工资表','target_cell':'AG3'}]
        with self.assertRaises(ValueError):self.build(plan=plan)

    def test_text_without_comment_object_is_rejected(self):
        args,_=self.build();book=load_workbook(args['output']);cell=book['工资表']['AG2']
        cell.value=cell.comment.text;cell.comment=None;book.save(args['output'])
        self.forge_receipt_after_change(args);self.assertFalse(self.verify(args)['passed'])

    def test_source_sha_change_is_rejected(self):
        args,_=self.build();book=load_workbook(self.source);book.active['B2']=-66;book.save(self.source)
        self.assertFalse(self.verify(args)['passed'])

    def test_required_original_inventory_cannot_be_omitted_even_when_zero(self):
        policy=deepcopy(self.policy);policy['source_comments']=[]
        with self.assertRaises(ValueError):GUARD.register(self.template,GUARD.sha256(self.template),policy,self.root/'missing-registry.json')

    @unittest.skipUnless(os.environ.get('PAYROLL_TEST_NATIVE')=='1','显式选择原生应用测试后才打开隔离副本')
    def test_native_plain_and_space_names_strict_cache_carry(self):
        with native_application(self) as app:
            for title in ['考勤依据','考勤 依据']:
                args,_=self.build(title);base=args['output'];calculated=base.parent/'仅原生隔离副本.xlsx';shutil.copy2(base,calculated)
                book=None
                try:
                    book=app.Workbooks.Open(str(calculated),UpdateLinks=0,ReadOnly=False);app.CalculateFullRebuild()
                    self.assertEqual(book.Worksheets('工资表').Range('T2').Value,-3)
                    self.assertIn(book.Worksheets('工资表').Range('T3').Value,('',None))
                    book.Save();book.Close(False);book=None
                    self.assertEqual(load_workbook(calculated)['工资表']['T2'].value,source_view_formula(title,'B2'))
                    CACHE.carry(base,calculated,base.parent/'保持母版与真实缓存.xlsx')
                    print(json.dumps({'明确合成页名':title,'实际应用路径':str(app.Path),'生产公式重算后逐字符一致':True,
                        '严格缓存搬运通过':True,'Microsoft_Excel已核实':False,'仅关闭本测试工作簿且不退出应用':True},ensure_ascii=False))
                finally:
                    if book is not None:book.Close(False)


if __name__=='__main__':unittest.main(verbosity=2)
