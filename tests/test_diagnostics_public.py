"""公开合成测试：不连接云端，不含任何真实员工资料或管理员密钥。"""
from pathlib import Path
import base64, concurrent.futures, copy, json, os, shutil, sqlite3, subprocess, sys, tempfile, time, unittest, uuid
from unittest.mock import patch

PACKAGE=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(PACKAGE/'diagnostics'))
import admin, collector, common, install_client
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric import rsa

def append_in_process(spec):
    home,event,org=spec
    collector.emit(event,home,org)
    return event['event_id']

class Diagnostics(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root=tempfile.TemporaryDirectory();cls.base=Path(cls.root.name)
        cls.shared=cls.base/'shared-admin';cls.org=admin.init(cls.shared,'synthetic-owner/private-diagnostics','[ssh.github.com]:443 ssh-ed25519 AAAAC3NzaC1lZDI1NTE5AAAAITest\n')
    @classmethod
    def tearDownClass(cls):cls.root.cleanup()
    def setUp(self):
        self.folder=self.base/uuid.uuid4().hex;self.folder.mkdir();self.ah=self.folder/'admin';common.private_dir(self.ah)
        (self.ah/'admin-key.secure').write_bytes((self.shared/'admin-key.secure').read_bytes());(self.ah/'organization.json').write_bytes(common.canonical(self.org));(self.ah/'devices.json').write_text('{}')
        self.profile=admin.create_profile(self.ah,'明确合成设备');self.profile_file=self.folder/'profile.json';self.profile_file.write_bytes(common.canonical(self.profile))
        self.home=self.folder/'client';install_client.install(self.home,self.profile_file,self.org,True)
        self.event=collector.event_for('workflow_audit',['--month','2099-09'],time.time(),0,client=self.profile)
    def signed(self,event=None):return common.sign_envelope(common.encrypt(event or self.event,self.org['public_key_pem']),self.profile['device_id'],self.profile['ssh_private_pem'].encode())
    def stored(self):
        with common.database(self.home) as db:return db.execute('SELECT event_id,envelope,state FROM events').fetchall()
    def remote_file(self,envelope):
        p=self.folder/'cloud/events'/envelope['event_id'][:2]/(envelope['event_id']+'.json');p.parent.mkdir(parents=True,exist_ok=True);p.write_bytes(common.canonical(envelope));return p
    def test_roundtrip_and_sender_identity(self):
        env=self.signed();self.assertEqual(common.decrypt(env,common.load_secret(self.ah/'admin-key.secure')),self.event);self.assertEqual(common.verify_sender(env,self.profile['ssh_public_key']),self.profile['device_id'])
    def test_wrong_admin_key_cannot_decrypt(self):
        key=rsa.generate_private_key(public_exponent=65537,key_size=3072);pem=key.private_bytes(serialization.Encoding.PEM,serialization.PrivateFormat.PKCS8,serialization.NoEncryption())
        with self.assertRaises(ValueError):common.decrypt(self.signed(),pem)
    def test_ciphertext_tampering_is_rejected(self):
        env=self.signed();changed=bytearray(base64.b64decode(env['ciphertext']));changed[0]^=1;env['ciphertext']=base64.b64encode(changed).decode()
        with self.assertRaises(Exception):common.decrypt(env,common.load_secret(self.ah/'admin-key.secure'))
    def test_sender_signature_tampering_is_rejected(self):
        env=self.signed();env['sender_device']=str(uuid.uuid4())
        with self.assertRaises(Exception):common.verify_sender(env,self.profile['ssh_public_key'])
    def test_event_id_is_authenticated(self):
        env=self.signed();env['event_id']=str(uuid.uuid4())
        with self.assertRaises(Exception):common.decrypt(env,common.load_secret(self.ah/'admin-key.secure'))
    def test_oversized_event_is_not_silently_truncated(self):
        event=dict(self.event,unexpected='x'*40000)
        with self.assertRaises(ValueError):common.encrypt(event,self.org['public_key_pem'])
    def test_unsigned_device_profile_is_rejected(self):
        profile=dict(self.profile);profile.pop('admin_signature')
        with self.assertRaises(Exception):admin.verify_profile(profile,self.org)
    def test_altered_device_profile_is_rejected(self):
        profile=dict(self.profile,repository='different/target')
        with self.assertRaises(ValueError):admin.verify_profile(profile,self.org)
    def test_generic_private_repo_flag_must_be_true(self):
        p=self.folder/'organization.json';p.write_bytes(common.canonical(dict(self.org,private_repository=False)))
        with self.assertRaises(ValueError):common.organization(p)
    def test_record_redacts_identity_paths_messages_and_amounts(self):
        secret='虚构测试人甲_工资88888_账号999000111_模拟密令SECRET'
        raw=self.folder/(secret+'.csv');raw.write_text(secret,encoding='utf-8')
        record=self.folder/'record.json';record.write_text(json.dumps({'month':'2099-09','policy_version':'2099.1','employees':[{'name':secret,'amount':88888}],'sources':{'secret':{'path':str(raw),'role':secret}}}),encoding='utf-8')
        output=self.folder/'result.json';output.write_text(json.dumps({'passed':False,'errors':['奖励申请不匹配:'+secret],'pending':[secret]}),encoding='utf-8')
        event=collector.event_for('workflow_audit',['--record',str(record),'--output',str(output),'--token',secret],time.time(),2,client=self.profile)
        payload=common.canonical(event).decode();self.assertNotIn(secret,payload);self.assertNotIn(str(self.folder),payload);self.assertNotIn('88888',payload);self.assertEqual(event['employee_count'],1);self.assertEqual(event['source_versions'][0]['sha256'],common.digest(raw.read_bytes()))
    def test_no_audit_is_unknown_not_zero_errors(self):
        event=collector.event_for('insurance_merge',[],time.time(),0,client=self.profile)
        self.assertIsNone(event['error_count']);self.assertIsNone(event['pending_count']);self.assertEqual(event['state'],'completed_unverified')
    def test_attendance_sidecar_is_read_without_scanning_other_files(self):
        output=self.folder/'out.csv';output.write_text('x');Path(str(output)+'.audit.json').write_text(json.dumps({'rows':3,'pending':[{'name':'明确合成','detail':'奖励待核'}]}),encoding='utf-8')
        event=collector.event_for('process_attendance',['--output',str(output)],time.time(),2,client=self.profile)
        self.assertEqual(event['pending_count'],1);self.assertEqual(event['employee_count'],3)
    def test_errors_are_failed_even_when_exit_two(self):
        output=self.folder/'result.json';output.write_text(json.dumps({'errors':['公式无效'],'pending':[]}),encoding='utf-8')
        event=collector.event_for('workflow_audit',['--output',str(output)],time.time(),2,client=self.profile);self.assertEqual(event['state'],'failed')
    def test_observed_keeps_success_return_and_queues(self):
        result=collector.observed(lambda:17,'workflow_audit',[],self.home,self.org,False)
        self.assertEqual(result,17);self.assertEqual(len(self.stored()),1)
    def test_observed_preserves_pending_exit(self):
        def operation():raise SystemExit(2)
        with self.assertRaises(SystemExit) as cm:collector.observed(operation,'workflow_audit',[],self.home,self.org,False)
        self.assertEqual(cm.exception.code,2);event=common.decrypt(json.loads(self.stored()[0][1]),common.load_secret(self.ah/'admin-key.secure'));self.assertEqual(event['exit_code'],2)
    def test_observed_preserves_original_exception_without_message(self):
        def operation():raise ValueError('虚构敏感原文不能上传')
        with self.assertRaises(ValueError):collector.observed(operation,'workflow_audit',[],self.home,self.org,False)
        event=common.decrypt(json.loads(self.stored()[0][1]),common.load_secret(self.ah/'admin-key.secure'));self.assertEqual(event['exception_kind'],'ValueError');self.assertNotIn('敏感原文',common.canonical(event).decode())
    def test_logging_failure_does_not_replace_calculation_return(self):
        with patch.object(collector,'emit',side_effect=OSError('synthetic')):self.assertEqual(collector.observed(lambda:7,'workflow_audit',[],self.home,self.org,False),7)
    def test_unconfigured_machine_is_not_falsely_recorded(self):
        home=self.folder/'not-installed';collector.observed(lambda:None,'workflow_audit',[],home,self.org,False);self.assertFalse(home.exists())
    def test_queue_is_encrypted_at_rest(self):
        event=dict(self.event,synthetic_marker='明确合成不能出现在数据库明文')
        collector.emit(event,self.home,self.org)
        for p in self.home.glob('*.sqlite3*'):self.assertNotIn(event['synthetic_marker'].encode(),p.read_bytes())
    def test_duplicate_event_cannot_replace_previous_record(self):
        collector.emit(self.event,self.home,self.org)
        with self.assertRaises(sqlite3.IntegrityError):collector.emit(self.event,self.home,self.org)
        self.assertEqual(len(self.stored()),1)
    def test_transport_failure_preserves_unsent_queue(self):
        collector.emit(self.event,self.home,self.org)
        with patch.object(common,'cache',side_effect=RuntimeError('synthetic offline')):
            with self.assertRaises(RuntimeError):common.sync_client(self.home,self.org)
        self.assertEqual(self.stored()[0][2],'queued')
    def test_empty_queue_does_not_contact_github(self):
        with patch.object(common,'git_run') as transport:self.assertEqual(common.sync_client(self.home,self.org)['sent'],0);transport.assert_not_called()
    def test_foreign_cache_is_not_reset(self):
        foreign=self.home/'cloud-cache';foreign.mkdir();keep=foreign/'keep.txt';keep.write_text('keep')
        with patch.object(common,'git_run') as transport:
            with self.assertRaises(ValueError):common.cache(self.home,'ssh://example.invalid/repo')
            transport.assert_not_called()
        self.assertEqual(keep.read_text(),'keep')
    def test_admin_ingest_and_repeat_are_idempotent(self):
        self.remote_file(self.signed());first=admin.ingest(self.ah,self.folder/'cloud');second=admin.ingest(self.ah,self.folder/'cloud');self.assertEqual(first['added'],1);self.assertEqual(second['added'],0);self.assertFalse(first['rejected'])
    def test_admin_rejects_unregistered_sender(self):
        env=self.signed();self.remote_file(env);(self.ah/'devices.json').write_text('{}');result=admin.ingest(self.ah,self.folder/'cloud');self.assertEqual(result['added'],0);self.assertEqual(len(result['rejected']),1)
    def test_admin_detects_replacement_of_same_event(self):
        self.remote_file(self.signed());admin.ingest(self.ah,self.folder/'cloud');changed=dict(self.event,state='failed');self.remote_file(self.signed(changed));result=admin.ingest(self.ah,self.folder/'cloud');self.assertEqual(result['added'],0);self.assertEqual(len(result['rejected']),1)
    def test_dashboard_output_is_local_and_escapes_embedded_script(self):
        registry=common.read_json(self.ah/'devices.json');registry[self.profile['device_id']]['label']='</script><script>BAD</script>';(self.ah/'devices.json').write_bytes(common.canonical(registry))
        self.remote_file(self.signed());admin.ingest(self.ah,self.folder/'cloud');out=self.ah/'查看.html';self.assertEqual(admin.dashboard(self.ah,out),1);self.assertNotIn('</script><script>BAD</script>',out.read_text('utf-8'))
    def test_dashboard_cannot_be_written_to_public_package(self):
        self.remote_file(self.signed());admin.ingest(self.ah,self.folder/'cloud')
        with self.assertRaises(ValueError):admin.dashboard(self.ah,self.folder/'public.html')
    def test_install_upgrade_preserves_identity_and_queue(self):
        collector.emit(self.event,self.home,self.org);before=(self.home/'client.secure').read_bytes();install_client.install(self.home,self.profile_file,self.org,True);self.assertEqual((self.home/'client.secure').read_bytes(),before);self.assertEqual(len(self.stored()),1)
    def test_install_cannot_silently_change_existing_device(self):
        another=admin.create_profile(self.ah,'另一合成设备');p=self.folder/'another.json';p.write_bytes(common.canonical(another))
        with self.assertRaises(ValueError):install_client.install(self.home,p,self.org,True)
    def test_multiple_threads_do_not_lose_queue_records(self):
        def operation(_):collector.emit(dict(self.event,event_id=str(uuid.uuid4())),self.home,self.org)
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:list(pool.map(operation,range(12)))
        self.assertEqual(len(self.stored()),12)
    def test_runtime_ssh_key_is_readable_by_actual_openssh(self):
        binary=shutil.which('ssh-keygen')
        if not binary:self.skipTest('此环境无OpenSSH，不能宣称原生密钥检查通过')
        with common.client_git_environment(self.home,self.profile,self.org):
            path=next(self.home.glob('ssh-*/identity'))
            result=subprocess.run([binary,'-y','-f',str(path)],capture_output=True,text=True,timeout=10)
            self.assertEqual(result.returncode,0,result.stderr)
            self.assertEqual(result.stdout.strip(),self.profile['ssh_public_key'])
    def test_multiple_processes_do_not_lose_queue_records(self):
        specs=[(self.home,dict(self.event,event_id=str(uuid.uuid4())),self.org) for _ in range(16)]
        with concurrent.futures.ProcessPoolExecutor(max_workers=4) as pool:ids=list(pool.map(append_in_process,specs))
        self.assertEqual(len(set(ids)),16);self.assertEqual(len(self.stored()),16)
    def test_offline_admin_view_is_explicitly_cached_not_fresh(self):
        self.remote_file(self.signed());admin.ingest(self.ah,self.folder/'cloud');out=self.ah/'offline.html'
        with patch.object(admin,'cache',side_effect=RuntimeError('synthetic offline')):result=admin.sync(self.ah,out)
        self.assertFalse(result['cloud_synced']);self.assertIn('不是最新云端状态',out.read_text('utf-8'))
    def test_offline_without_cache_cannot_claim_view_success(self):
        with patch.object(admin,'cache',side_effect=RuntimeError('synthetic offline')):
            with self.assertRaises(RuntimeError):admin.sync(self.ah,self.ah/'not-created.html')
        self.assertFalse((self.ah/'not-created.html').exists())

if __name__=='__main__':unittest.main(verbosity=2)
