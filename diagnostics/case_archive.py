"""從指定私有 GitHub 倉庫下載案件密文；完整校驗後本機歸檔，永不刪除雲端。"""
from __future__ import annotations
import datetime, json, os, re, shutil, tempfile, urllib.request, uuid
from pathlib import Path
from common import canonical, decrypt, digest, load_secret, lock, organization, private_dir, read_json, verify_sender
from diagnostic_archive import GitHubReader, REPOSITORY, OWNER, SHA, git_sha, safe_path, atomic_json
from case_store import MAX_FILE, CASE_EVENT
from case_admin import verify_snapshot

UUID=r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}'
MANIFEST=re.compile(r'cases/manifests/('+UUID+r')/('+UUID+r')\.json\Z')
OBJECT=re.compile(r'cases/objects/([0-9a-f]{2})/('+UUID+r')\.bin\Z')

class CaseReader(GitHubReader):
    def all_entries(self):
        if self.get('/user').get('login')!=OWNER:raise ValueError('不是公司管理員帳戶')
        base='/repos/'+REPOSITORY;repo=self.get(base)
        if repo.get('private') is not True or repo.get('full_name')!=REPOSITORY:raise ValueError('不是指定私有倉庫')
        commit=self.get(base+'/git/ref/heads/main')['object']['sha']
        if not SHA.fullmatch(commit):raise ValueError('雲端版本格式錯誤')
        tree_sha=self.get(base+'/git/commits/'+commit)['tree']['sha']
        if not SHA.fullmatch(tree_sha):raise ValueError('雲端目錄版本格式錯誤')
        tree=self.get(base+'/git/trees/'+tree_sha+'?recursive=1')
        if tree.get('truncated'):raise ValueError('雲端索引不完整')
        entries={}
        for entry in tree['tree']:
            name=entry['path']
            if not name.startswith('cases/') or entry['type']=='tree':continue
            manifest=MANIFEST.fullmatch(name);obj=OBJECT.fullmatch(name)
            if not (manifest or obj) or entry.get('mode')!='100644' or not SHA.fullmatch(entry['sha']):raise ValueError('案件雲端路徑或類型異常')
            limit=1024*1024 if manifest else MAX_FILE+16
            if not 0<=entry.get('size',limit+1)<=limit:raise ValueError('案件雲端檔案超過上限')
            if obj and obj[2][:2]!=obj[1]:raise ValueError('附件目錄不符')
            entries[name]={'path':name,'sha':entry['sha'],'size':entry['size']}
        return commit,entries
    def blob_bytes(self,entry):
        if not SHA.fullmatch(entry['sha']) or not 0<=entry['size']<=MAX_FILE+16:raise ValueError('附件參數錯誤')
        request=urllib.request.Request('https://api.github.com/repos/'+REPOSITORY+'/git/blobs/'+entry['sha'],headers={
            'Authorization':'Bearer '+self.token,'Accept':'application/vnd.github.raw+json',
            'X-GitHub-Api-Version':'2026-03-10','User-Agent':'company-payroll-case-archive'})
        with self.opener.open(request,timeout=40) as response:raw=response.read(entry['size']+1)
        if len(raw)!=entry['size'] or git_sha(raw)!=entry['sha']:raise ValueError('案件下載摘要不符')
        return raw

def archive_cases(root,admin_home,reader,reserve_bytes=10*2**30):
    original=Path(root).absolute();root=original.resolve()
    if root!=original or root==Path(root.anchor):raise ValueError('請指定獨立案件歸檔資料夾')
    safe_path(root,'status.json');private_dir(root,harden=True)
    with lock(root) as acquired:
        if not acquired:raise RuntimeError('案件歸檔正在執行')
        try:
            home=Path(admin_home);org=organization(home/'organization.json');registry=read_json(home/'devices.json');pem=load_secret(home/'admin-key.secure')
            if org['repository']!=REPOSITORY:raise ValueError('歸檔組織不符')
            commit,entries=reader.all_entries();needed={};snapshots=[];receipts=[];pending=[]
            with tempfile.TemporaryDirectory(prefix='.pending-',dir=root) as temp:
                staging=Path(temp)
                def obtain(entry):
                    existing=safe_path(root,entry['path'])
                    if existing.exists():raw=existing.read_bytes()
                    else:
                        raw=reader.blob_bytes(entry);target=safe_path(staging,entry['path']);target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(raw);pending.append((target,existing))
                    if len(raw)!=entry['size'] or git_sha(raw)!=entry['sha']:raise ValueError('已歸檔或下載的案件內容不符')
                    needed[entry['path']]=entry
                    return raw
                for path,entry in entries.items():
                    match=MANIFEST.fullmatch(path)
                    if not match:continue
                    raw=obtain(entry);envelope=json.loads(raw);sender=envelope['sender_device']
                    if sender not in registry:raise ValueError('未知設備案件')
                    verify_sender(envelope,registry[sender]['public_key']);event=decrypt(envelope,pem,CASE_EVENT)
                    if match[1]!=event['device_id'] or match[2]!=event['event_id']:raise ValueError('案件路徑身份不符')
                    snapshots.append(envelope)
                    for meta in event['files']:
                        object_path=meta['object_path']
                        if not OBJECT.fullmatch(object_path) or object_path not in entries:raise ValueError('案件附件缺失')
                        needed[object_path]=entries[object_path]
                new_size=sum(entry['size'] for entry in needed.values() if not safe_path(root,entry['path']).exists())
                if shutil.disk_usage(root).free<reserve_bytes+new_size+1024*1024:raise OSError('案件歸檔磁碟空間不足')
                for path,entry in list(needed.items()):
                    if OBJECT.fullmatch(path):obtain(entry)
                for envelope in snapshots:verify_snapshot(envelope,staging,org,registry,pem,True,fallback=root)
                for source,target in pending:
                    target.parent.mkdir(parents=True,exist_ok=True)
                    with target.open('xb') as stream:stream.write(source.read_bytes());stream.flush();os.fsync(stream.fileno())
                    if digest(target.read_bytes())!=digest(source.read_bytes()):raise OSError('案件落盤校驗失敗')
                for path in sorted(needed):
                    raw=safe_path(root,path).read_bytes();receipts.append({'path':path,'sha256':digest(raw),'bytes':len(raw)})
            metadata={'organization':org,'registered_devices':registry};meta_path=safe_path(root,'metadata/'+digest(canonical(metadata))+'.json')
            if not meta_path.exists():atomic_json(meta_path,metadata)
            result={'schema':'company-payroll-case-archive-1','state':'success','at':datetime.datetime.now(datetime.timezone.utc).isoformat(),
                    'cloud_commit':commit,'verified_snapshots':len(snapshots),'verified_objects':sum(1 for p in needed if OBJECT.fullmatch(p)),
                    'new_files':len(pending),'files':receipts,'cloud_deleted':False,'metadata':meta_path.relative_to(root).as_posix()}
            atomic_json(safe_path(root,'receipts/'+datetime.datetime.now().strftime('%Y%m%d-%H%M%S-')+uuid.uuid4().hex[:8]+'.json'),result)
            atomic_json(safe_path(root,'last-success.json'),result);atomic_json(safe_path(root,'status.json'),result)
            return result
        except Exception as exc:
            atomic_json(safe_path(root,'status.json'),{'state':'failed','at':datetime.datetime.now(datetime.timezone.utc).isoformat(),'category':type(exc).__name__,'cloud_deleted':False})
            raise
