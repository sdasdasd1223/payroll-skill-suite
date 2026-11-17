"""明確合成案件；不含真實薪資、不連線 GitHub，不讀取其他聊天。"""
from pathlib import Path
import base64, copy, json, os, shutil, sys, tempfile, threading, unittest, urllib.error, urllib.request, uuid
from unittest.mock import patch
P=Path(__file__).resolve().parents[1];sys.path.insert(0,str(P/'diagnostics'))
import admin, case_admin, case_archive, case_store, collector, common, conversation_reader, install_client
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from openpyxl import Workbook,load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

class FakeCloud:
    def __init__(self,folder):
        self.files={};self.reads=0
        for path in Path(folder).rglob('*'):
            if path.is_file() and 'cases' in path.relative_to(folder).parts:
                raw=path.read_bytes();name=path.relative_to(folder).as_posix()
                self.files[name]=(dict(path=name,size=len(raw),sha=case_archive.git_sha(raw)),raw)
    def all_entries(self):return 'a'*40,{name:meta for name,(meta,raw) in self.files.items()}
    def blob_bytes(self,entry):self.reads+=1;return self.files[entry['path']][1]

class Cases(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.shared=tempfile.TemporaryDirectory();cls.shared_path=Path(cls.shared.name)
        cls.org=admin.init(cls.shared_path/'admin',case_archive.REPOSITORY,'[ssh.github.com]:443 test\n')
    @classmethod
    def tearDownClass(cls):cls.shared.cleanup()
    def setUp(self):
        self.t=tempfile.TemporaryDirectory();self.root=Path(self.t.name).resolve();self.ah=self.root/'admin';common.private_dir(self.ah)
        for name in ['admin-key.secure','organization.json']:shutil.copyfile(self.shared_path/'admin'/name,self.ah/name)
        (self.ah/'devices.json').write_text('{}');self.profile=admin.create_profile(self.ah,'合成人資甲')
        self.profile_path=self.root/'profile.json';self.profile_path.write_bytes(common.canonical(self.profile));self.home=self.root/'client'
        install_client.install(self.home,self.profile_path,self.org,True)
        self.case=self.root/'case';self.people=[{'employee_key':'E01','name':'合成甲','company':'測試公司'}]
        self.thread=str(uuid.uuid4());self.turn='turn-synthetic-1'
        self.context=case_store.begin(self.case,'2099-09',self.people,self.home,'合成人資甲',thread_id=self.thread,turn_ids=[self.turn])
        self.path=self.case/case_store.CONTEXT
        self.source=self.case/'原始考勤.xlsx';book=Workbook();sheet=book.active;sheet['A1']='合成甲';sheet['B1']=3000;sheet['C1']='=B1+200';sheet['B1'].fill=PatternFill('solid',fgColor='FFE699');sheet['A1'].comment=Comment('合成批註','測試作者');book.save(self.source)
        self.result=self.case/'計算結果.json';self.result.write_text('{"合成獎勵":200}',encoding='utf-8')
        self.conversation={'schema':'payroll-case-conversation-1','thread_id':self.thread,'turn_ids':[self.turn],
                           'messages':[{'role':'user','text':'合成測試，請核算九月。'},{'role':'assistant','text':'已核對合成獎勵 200 元。'}],
                           'completed_turns':True,'omitted_nontext_items':0,'coverage':'bound_turns_visible_text_only'}
    def tearDown(self):self.t.cleanup()
    def capture(self,**kwargs):
        return case_store.capture(self.path,[(self.source,'source'),(self.result,'result')],self.home,self.org,conversation=kwargs.pop('conversation',self.conversation),add_current_turn=False,**kwargs)
    def envelope(self,snapshot_id):
        with case_store.db(self.home) as connection:return json.loads(connection.execute('SELECT envelope FROM snapshots WHERE id=?',(snapshot_id,)).fetchone()[0])
    def cloud(self):
        folder=self.root/'cloud'
        with case_store.db(self.home) as connection:rows=connection.execute('SELECT id,envelope FROM snapshots').fetchall()
        for identifier,payload in rows:
            target=folder/'cases/manifests'/self.profile['device_id']/(identifier+'.json');target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(payload.encode())
        for source in (self.home/'case-objects').glob('*.bin'):
            target=folder/'cases/objects'/source.stem[:2]/source.name;target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,target)
        return folder
    def test_roundtrip_original_result_conversation_preserves_formulas_styles_comments(self):
        original=self.source.read_bytes();result=self.capture();folder=self.cloud();cases,rejected=case_admin.load_cases(folder,self.ah)
        self.assertFalse(rejected);self.assertEqual(len(cases),1)
        for item in cases[0]['files']:
            meta,raw=case_admin.file_bytes(cases[0],item['file_id'],folder,self.ah)
            if meta['role']=='source':self.assertEqual(raw,original)
            if meta['role']=='result':self.assertEqual(raw,self.result.read_bytes())
            if meta['role']=='conversation':self.assertEqual(json.loads(raw)['messages'],self.conversation['messages'])
        self.assertEqual(self.source.read_bytes(),original)
        book=load_workbook(self.source);self.assertEqual(book.active['C1'].value,'=B1+200');self.assertEqual(book.active['A1'].comment.text,'合成批註');self.assertEqual(book.active['B1'].fill.fgColor.rgb,'00FFE699');book.close()
    def test_repeat_checkpoint_creates_no_duplicate_snapshot(self):
        self.capture();self.assertEqual(self.capture()['state'],'unchanged')
        with case_store.db(self.home) as connection:self.assertEqual(connection.execute('SELECT COUNT(*) FROM snapshots').fetchone()[0],1)
    def test_changed_result_preserves_before_and_after(self):
        self.capture();old=self.result.read_bytes();self.result.write_text('{"合成獎勵":400}',encoding='utf-8');self.capture()
        folder=self.cloud();cases,rejected=case_admin.load_cases(folder,self.ah);self.assertEqual(len(cases),2)
        originals=[]
        for case in cases:
            meta=next(x for x in case['files'] if x['role']=='result');originals.append(case_admin.file_bytes(case,meta['file_id'],folder,self.ah)[1])
        self.assertIn(old,originals);self.assertIn(self.result.read_bytes(),originals)
    def test_plaintext_absent_from_cloud_names_and_bytes(self):
        self.capture();folder=self.cloud();raw=b''.join(p.read_bytes() for p in folder.rglob('*') if p.is_file())
        for word in ['合成甲','原始考勤','合成獎勵','已核對合成獎勵']:self.assertNotIn(word.encode(),raw)
        self.assertNotIn('原始考勤',' '.join(str(p) for p in folder.rglob('*')))
    def test_unapproved_source_is_rejected(self):
        other=self.root/'其他案件.txt';other.write_text('不可讀取',encoding='utf-8')
        with self.assertRaises(ValueError):case_store.capture(self.path,[(other,'source')],self.home,self.org,conversation=self.conversation,add_current_turn=False)
    def test_credentials_filename_rejected(self):
        secret=self.case/'api-token.json';secret.write_text('不能上傳',encoding='utf-8')
        with self.assertRaises(ValueError):case_store.capture(self.path,[(secret,'source')],self.home,self.org,conversation=self.conversation,add_current_turn=False)
    def test_whole_drive_scope_rejected(self):
        with self.assertRaises(ValueError):case_store.begin(self.root/'other','2099-09',self.people,self.home,allowed_roots=[str(self.root.anchor)],thread_id=self.thread,turn_ids=[self.turn])
    def test_wrong_conversation_binding_rejected(self):
        other=dict(self.conversation,thread_id=str(uuid.uuid4()))
        with self.assertRaises(ValueError):self.capture(conversation=other)
    def test_missing_conversation_remains_unavailable_not_complete(self):
        with patch.object(conversation_reader,'extract',side_effect=RuntimeError):result=self.capture(conversation=None)
        self.assertEqual(result['conversation_state'],'unavailable');self.assertTrue(result['issues'])
    def test_unfinished_turn_is_pending(self):
        result=self.capture(conversation=dict(self.conversation,completed_turns=False));self.assertEqual(result['conversation_state'],'in_progress')
    def test_nontext_attachment_is_reported(self):
        result=self.capture(conversation=dict(self.conversation,omitted_nontext_items=1));self.assertTrue(any('非文字' in x for x in result['issues']))
    def test_oversize_is_rejected_without_truncation(self):
        with patch.object(case_store,'MAX_FILE',1):
            with self.assertRaises(ValueError):self.capture()
    def test_missing_source_does_not_claim_saved(self):
        self.source.unlink()
        with self.assertRaises(FileNotFoundError):self.capture()
    def test_duplicate_employee_key_rejected(self):
        with self.assertRaises(ValueError):case_store.begin(self.root/'bad','2099-09',self.people*2,self.home,thread_id=self.thread,turn_ids=[self.turn])
    def test_same_name_different_company_remains_two_people(self):
        case=case_store.begin(self.root/'other','2099-09',self.people+[dict(self.people[0],company='第二公司')],self.home,thread_id=self.thread,turn_ids=[self.turn]);self.assertEqual(len(case['employees']),2)
    def test_month_employee_operator_filters(self):
        self.capture();cases,_=case_admin.load_cases(self.cloud(),self.ah)
        self.assertEqual(len(case_admin.filter_cases(cases,'2099-09','E01','合成人資甲')),1)
        self.assertEqual(case_admin.filter_cases(cases,'2099-08'),[]);self.assertEqual(case_admin.filter_cases(cases,employee='不存在'),[])
    def test_cipher_tamper_is_rejected(self):
        self.capture();folder=self.cloud();target=next(folder.glob('cases/objects/*/*.bin'));raw=bytearray(target.read_bytes());raw[0]^=1;target.write_bytes(raw)
        cases,rejected=case_admin.load_cases(folder,self.ah);self.assertFalse(cases);self.assertEqual(len(rejected),1)
    def test_manifest_signature_tamper_is_rejected(self):
        self.capture();folder=self.cloud();path=next(folder.glob('cases/manifests/*/*.json'));env=common.read_json(path);raw=bytearray(base64.b64decode(env['signature']));raw[0]^=1;env['signature']=base64.b64encode(raw).decode();path.write_bytes(common.canonical(env));self.assertEqual(len(case_admin.load_cases(folder,self.ah)[1]),1)
    def test_other_admin_cannot_restore(self):
        self.capture();cases,_=case_admin.load_cases(self.cloud(),self.ah);meta=cases[0]['files'][0];cipher=(self.home/'case-objects'/(meta['file_id']+'.bin')).read_bytes();key=rsa.generate_private_key(public_exponent=65537,key_size=3072);pem=key.private_bytes(serialization.Encoding.PEM,serialization.PrivateFormat.PKCS8,serialization.NoEncryption())
        with self.assertRaises(ValueError):case_store.decrypt_attachment(cipher,meta,cases[0]['case_id'],pem)
    def test_archive_twice_no_duplicate_cloud_reads_for_existing(self):
        self.capture();reader=FakeCloud(self.cloud());target=self.root/'archive';first=case_archive.archive_cases(target,self.ah,reader,0);reads=reader.reads;second=case_archive.archive_cases(target,self.ah,reader,0)
        self.assertEqual(first['verified_snapshots'],1);self.assertEqual(second['new_files'],0);self.assertEqual(reader.reads,reads)
    def test_archive_missing_cloud_object_keeps_prior_success(self):
        self.capture();reader=FakeCloud(self.cloud());target=self.root/'archive';case_archive.archive_cases(target,self.ah,reader,0);before=(target/'last-success.json').read_bytes();del reader.files[next(name for name in reader.files if '/objects/' in name)]
        with self.assertRaises(ValueError):case_archive.archive_cases(target,self.ah,reader,0)
        self.assertEqual((target/'last-success.json').read_bytes(),before);self.assertEqual(common.read_json(target/'status.json')['state'],'failed')
    def test_archive_offline_keeps_prior_success(self):
        self.capture();reader=FakeCloud(self.cloud());target=self.root/'archive';case_archive.archive_cases(target,self.ah,reader,0);before=(target/'last-success.json').read_bytes()
        with patch.object(reader,'all_entries',side_effect=ConnectionError):
            with self.assertRaises(ConnectionError):case_archive.archive_cases(target,self.ah,reader,0)
        self.assertEqual((target/'last-success.json').read_bytes(),before)
    def test_archive_full_disk_does_not_finish(self):
        self.capture();reader=FakeCloud(self.cloud());target=self.root/'archive'
        with patch.object(case_archive.shutil,'disk_usage') as usage:
            usage.return_value.free=0
            with self.assertRaises(OSError):case_archive.archive_cases(target,self.ah,reader,0)
        self.assertFalse((target/'last-success.json').exists())
    def test_queue_stays_pending_when_transport_fails(self):
        self.capture()
        with patch.object(case_store,'cache',side_effect=RuntimeError):
            with self.assertRaises(RuntimeError):case_store.sync(self.home,self.org)
        with case_store.db(self.home) as connection:self.assertEqual(connection.execute('SELECT state FROM snapshots').fetchone()[0],'queued')
    def test_conversation_refresh_survives_original_cleanup(self):
        self.capture(conversation=dict(self.conversation,completed_turns=False));self.source.unlink();self.result.unlink()
        with patch.object(conversation_reader,'extract',return_value=self.conversation):
            results=case_store.refresh(self.home,self.org)
        self.assertEqual(results[0].get('conversation_state'),'completed_visible_text')
    def test_viewer_filters_and_restores_download(self):
        self.capture();folder=self.cloud();server,url=case_admin.create_server(folder,self.ah);worker=threading.Thread(target=server.serve_forever,daemon=True);worker.start()
        try:
            data=json.load(urllib.request.urlopen(url+'api/cases?month=2099-09'));self.assertEqual(len(data['cases']),1);case=data['cases'][0];meta=next(x for x in case['files'] if x['role']=='source')
            response=urllib.request.urlopen(url+'api/file/'+case['event_id']+'/'+meta['file_id']);self.assertEqual(response.read(),self.source.read_bytes());self.assertTrue(response.headers['Content-Disposition'].startswith('attachment;'))
            with self.assertRaises(urllib.error.HTTPError):urllib.request.urlopen(url.replace('/session/','/invalid/')+'api/cases')
        finally:server.shutdown();server.server_close();worker.join()
    def test_conversation_reader_excludes_unbound_turns_and_reasoning(self):
        class FakeRPC:
            def request(self,*args,**kwargs):return {'data':[{'id':'other','status':'completed','items':[{'type':'userMessage','content':[{'type':'text','text':'不屬本案'}]}]},
              {'id':'wanted','status':'completed','items':[{'type':'reasoning','summary':['不保存思考']},{'type':'userMessage','content':[{'type':'text','text':'本案問題'}]},{'type':'agentMessage','text':'本案回答'}]}],'nextCursor':None}
        result=conversation_reader.extract(self.thread,['wanted'],FakeRPC());self.assertEqual([x['text'] for x in result['messages']],['本案問題','本案回答'])
    def test_conversation_missing_turn_is_not_called_full(self):
        class FakeRPC:
            def request(self,*args,**kwargs):return {'data':[],'nextCursor':None}
        with self.assertRaises(ValueError):conversation_reader.extract(self.thread,['missing'],FakeRPC())
    def test_more_than_old_24_files_and_large_manifest_roundtrip(self):
        files=[]
        for index in range(70):
            path=self.case/('合成文件'+str(index)+'.txt');path.write_text(str(index));files.append((path,'source'))
        result=case_store.capture(self.path,files,self.home,self.org,conversation=self.conversation,add_current_turn=False)
        self.assertEqual(result['files'],71);cases,rejected=case_admin.load_cases(self.cloud(),self.ah);self.assertFalse(rejected);self.assertEqual(len(cases[0]['files']),71)
    def test_cross_month_is_not_silently_indexed_under_wrong_month(self):
        with patch.object(case_store,'find_context',return_value=self.path):
            with self.assertRaises(ValueError):case_store.record_entry('process_attendance',['--month','2099-08','--attendance',str(self.source)],self.home,self.org)
    def test_legitimate_private_run_record_can_be_archived(self):
        path=self.case/'run-record.private.json';path.write_text('{"month":"2099-09"}')
        result=case_store.capture(self.path,[(path,'run_record')],self.home,self.org,conversation=self.conversation,add_current_turn=False)
        self.assertEqual(result['files'],2)
    def test_automatic_entry_preserves_pending_exit_and_saves_available_inputs(self):
        with patch.object(conversation_reader,'extract',return_value=self.conversation):
            value=collector.observed(lambda:2,'process_attendance',['--month','2099-09','--attendance',str(self.source),'--output',str(self.case/'missing-result.json')],self.home,self.org,False)
        self.assertEqual(value,2)
        with case_store.db(self.home) as connection:self.assertEqual(connection.execute('SELECT COUNT(*) FROM snapshots').fetchone()[0],1)
        health=common.read_json(self.home/'case-health.json');self.assertTrue(health['issues']);self.assertEqual(health['state'],'queued')
    def test_three_months_are_separately_searchable(self):
        self.capture();folder=self.cloud()
        for month,worker in [('2099-07','合成人資乙'),('2099-08','合成人資丙')]:
            other=self.root/month;context=case_store.begin(other,month,self.people,self.home,worker,thread_id=self.thread,turn_ids=[self.turn]);raw=other/'結果.txt';raw.write_text(month)
            case_store.capture(other/case_store.CONTEXT,[(raw,'result')],self.home,self.org,conversation=self.conversation,add_current_turn=False)
        folder=self.cloud();cases,rejected=case_admin.load_cases(folder,self.ah);self.assertFalse(rejected);self.assertEqual(len(cases),3)
        self.assertEqual(len(case_admin.filter_cases(cases,'2099-07','合成甲','合成人資乙')),1);self.assertEqual(case_admin.filter_cases(cases,'2099-08',operator='合成人資乙'),[])
    def test_conversation_credentials_are_masked(self):
        text='Bearer '+('a'*30)+' sk-'+('b'*30)
        result=conversation_reader.redact(text);self.assertNotIn('a'*30,result);self.assertNotIn('b'*30,result)
    def test_background_refresh_preserves_missing_result_warning(self):
        self.capture(capture_issues=['本次計算沒有產生指定結果文件，未當成已保存'])
        with patch.object(conversation_reader,'extract',return_value=self.conversation):result=case_store.refresh(self.home,self.org)[0]
        if result['state']=='unchanged':
            context=common.read_json(self.path);self.assertTrue(context['pending_capture_issues'])
        else:self.assertTrue(result['issues'])
    def test_missing_case_is_recorded_in_admin_diagnostic(self):
        with patch.object(case_store,'find_context',return_value=None):
            collector.observed(lambda:2,'workflow_audit',['--month','2099-09'],self.home,self.org,False)
        with common.database(self.home) as connection:payload=connection.execute('SELECT envelope FROM events').fetchone()[0]
        event=common.decrypt(json.loads(payload),common.load_secret(self.ah/'admin-key.secure'))
        self.assertEqual(event['case_archive_state'],'context_missing')

if __name__=='__main__':unittest.main(verbosity=2)
