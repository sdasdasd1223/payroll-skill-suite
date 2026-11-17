"""六项明确合成范围反查；不决定HR归属，不自动加入工资/税表。"""
import copy
import csv
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from openpyxl import Workbook

def locate():
    for parent in Path(__file__).resolve().parents:
        for relative in ('scripts','payroll-operations/scripts','skills/payroll-operations/scripts','candidate/skills/payroll-operations/scripts'):
            path = parent/relative/'crossstage_audit.py'
            if path.is_file(): return path
    raise RuntimeError('未找到相邻技能程序')

def load(path):
    spec = importlib.util.spec_from_file_location('_synthetic_scope_' + str(id(path)), path)
    value = importlib.util.module_from_spec(spec); spec.loader.exec_module(value); return value

PROGRAM = load(locate())

def build(root):
    root = Path(root).resolve(); root.mkdir(parents=True, exist_ok=True)
    record = {'schema_version':'crossstage-1','month':'2099-01','policy_version':'SYNTHETIC','mode':'draft','through_stage':'attendance','scope_table':'scope','sources':{},'tables':{},'facts':[],'confirmations':[],'stages':{'attendance':{'dispositions':[{'employee_key':'SYNTHETIC-KEY','route':'pending','reason':'仅范围反查'}]}},'raw_attendance':{'table':'normalized','original_table':'original','leave_columns':{kind:{'kind':kind,'unit':'hours'} for kind in PROGRAM.LEAVE_ITEMS}}}
    person = {'employee_key':'SYNTHETIC-KEY','name':'SYNTHETIC-IN','company':'SYNTHETIC'}
    def table(key, rows):
        path=root/(key+'.json'); path.write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
        record['sources'][key]={'path':str(path),'sha256':PROGRAM.sha256(path),'month':record['month'],'role':'synthetic-source'}
        record['tables'][key]={'source_id':key,'format':'json','start_row':1,'end_row':len(rows),'columns':{k:k for r in rows for k in r}}
    table('scope',[person]); table('normalized',[dict(person,date='2099-01-02',attendance_type='需要打卡',**{k:0 for k in PROGRAM.LEAVE_ITEMS})])
    table('ownership',[{'name':'SYNTHETIC-OUT','reason':'明确合成：人工归属依据'},{'name':'SYNTHETIC-IN','reason':'明确合成：其他人'}])
    original=root/'original.xlsx'; book=Workbook(); sheet=book.active; sheet.title='SYNTHETIC'; sheet.append(['name','date']); sheet.append(['SYNTHETIC-IN','2099-01-02']); sheet.append(['SYNTHETIC-OUT','2099-01-03']); book.save(original)
    record['sources']['original']={'path':str(original),'sha256':PROGRAM.sha256(original),'month':record['month'],'role':'raw_attendance'}
    record['tables']['original']={'source_id':'original','format':'xlsx','sheet':'SYNTHETIC','start_row':2,'end_row':3,'columns':{'name':'A','date':'B'}}
    return record

def check(record, program=PROGRAM):
    audit=program.Audit(record,Path('.')); audit.load_sources(); audit.load_tables(); audit.scope(); audit.raw_scope_dispositions(); return audit

class ScopeReverse(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory(prefix='synthetic-scope-reverse-'); self.root=Path(self.temp.name).resolve(); self.record=build(self.root)
    def tearDown(self):
        if self.root.parent!=Path(tempfile.gettempdir()).resolve(): raise RuntimeError('隔离目录越界')
        self.temp.cleanup()
    def test_normalized_json_without_original_is_pending(self):
        self.record['raw_attendance'].pop('original_table'); a=check(self.record)
        self.assertEqual(a.evidence['scope_reverse_status'],'pending_original'); self.assertTrue(a.pending)
    def test_original_excluded_rows_still_get_named_pending_and_csv(self):
        self.record['tables']['original']['excluded_rows']={'3':'不能借表过滤丢弃候选'}
        a=check(self.record); row=a.evidence['scope_review'][0]
        self.assertEqual((row['name'],row['route'],row['original_rows']),('SYNTHETIC-OUT','pending',[3]))
        path=self.root/'review.csv'; PROGRAM.write_scope_review_csv({'evidence':a.evidence},path)
        with path.open(encoding='utf-8-sig',newline='') as stream: exported=list(csv.DictReader(stream))
        self.assertEqual(exported[0]['name'],'SYNTHETIC-OUT'); self.assertEqual(set(a.people),{'SYNTHETIC-KEY'})
    def test_explicit_pending_is_not_added_to_fixed_people(self):
        self.record['scope_dispositions']=[{'name':'SYNTHETIC-OUT','route':'pending','reason':'负责人范围待核'}]
        original=copy.deepcopy(self.record['stages']); a=check(self.record)
        self.assertEqual(a.evidence['scope_review'][0]['reason'],'负责人范围待核'); self.assertEqual(self.record['stages'],original); self.assertEqual(set(a.people),{'SYNTHETIC-KEY'})
    def test_excluded_needs_actual_same_person_reference(self):
        self.record['scope_dispositions']=[{'name':'SYNTHETIC-OUT','route':'excluded','reason':'按实际同人归属依据人工排除','source_refs':[{'table':'ownership','row':1,'column':'reason'}]}]
        a=check(self.record); self.assertFalse(a.pending); self.assertEqual(a.evidence['scope_reverse_status'],'complete'); self.assertEqual(a.evidence['scope_review'][0]['route'],'excluded')
        # 已有合法排除时，原始适配数据中保留该行也不能被误算为本人工资。
        path=Path(self.record['sources']['normalized']['path']); rows=json.loads(path.read_text(encoding='utf-8')); rows.append(dict(rows[0],employee_key='SYNTHETIC-OTHER',name='SYNTHETIC-OUT'))
        path.write_text(json.dumps(rows),encoding='utf-8'); self.record['sources']['normalized']['sha256']=PROGRAM.sha256(path); self.record['tables']['normalized']['end_row']=2
        a=check(self.record); a.raw_attendance(); self.assertFalse(any('范围或身份待核' in x for x in a.pending))
    def test_circular_reason_or_other_person_reference_stays_pending(self):
        for reason,row in (('不在已选名单',1),('按归属依据排除',2)):
            with self.subTest(reason=reason):
                self.record['scope_dispositions']=[{'name':'SYNTHETIC-OUT','route':'excluded','reason':reason,'source_refs':[{'table':'ownership','row':row,'column':'reason'}]}]
                self.assertEqual(check(self.record).evidence['scope_review'][0]['route'],'pending')
    def test_direct_original_workbook_does_not_need_duplicate_binding(self):
        self.record['raw_attendance']={'table':'original'}; a=check(self.record)
        self.assertEqual(a.evidence['scope_review'][0]['original_rows'],[3]); self.assertEqual(a.evidence['scope_reverse_status'],'pending_candidates')

if __name__=='__main__': unittest.main(verbosity=2)
