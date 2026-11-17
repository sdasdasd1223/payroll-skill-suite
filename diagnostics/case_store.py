"""公司核算案件：限定來源、密文附件、版本與明確綁定的對話。"""
from __future__ import annotations
import argparse, base64, contextlib, datetime, getpass, json, os, re, sqlite3, subprocess, sys, time, uuid
from pathlib import Path
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from common import canonical, client_git_environment, cache, decrypt, digest, encrypt, git_run, home_path, load_secret, lock, organization, private_dir, private_key, protect, public_key, read_json, read_json_secret, sign_envelope, stage_blobs, commit_index, unprotect

MAX_FILE=32*1024*1024
MAX_TOTAL=256*1024*1024
MAX_FILES=128
CASE_EVENT=512*1024
EXTENSIONS={'.xlsx','.xls','.xlsm','.csv','.tsv','.json','.txt','.md','.pdf','.docx','.zip','.png','.jpg','.jpeg'}
SECRET=re.compile(r'(^|[._\- ])(token|credentials?|password|secret|auth)([._\- ]|$)',re.I)
DENIED={'.git','.codex','.ssh','node_modules','.venv'}
ROLES={'source','result','conversation','run_record','template'}
CONTEXT='.payroll-case.json'

def now():return datetime.datetime.now(datetime.timezone.utc).isoformat()

@contextlib.contextmanager
def db(home):
    home=private_dir(Path(home));connection=sqlite3.connect(home/'case-queue.sqlite3',timeout=10)
    try:
        connection.execute('PRAGMA busy_timeout=10000')
        connection.execute('CREATE TABLE IF NOT EXISTS snapshots(id TEXT PRIMARY KEY,case_id TEXT,envelope TEXT,state TEXT,created REAL)')
        connection.execute('CREATE TABLE IF NOT EXISTS file_cache(cache_key TEXT PRIMARY KEY,metadata BLOB)')
        connection.execute('CREATE TABLE IF NOT EXISTS watches(case_id TEXT PRIMARY KEY,path TEXT,done INTEGER DEFAULT 0)')
        with connection:yield connection
    finally:connection.close()

def identifier(value):return str(uuid.UUID(value))

def safe_under(path, roots, must_exist=True):
    original=Path(path).absolute();resolved=original.resolve()
    if original!=resolved or any(part in DENIED for part in resolved.parts):raise ValueError('案件路徑不可經過連結或私有設定目錄')
    if not any(resolved==Path(root).resolve() or resolved.is_relative_to(Path(root).resolve()) for root in roots):raise ValueError('文件不在已指定的公司案件範圍')
    if SECRET.search(resolved.name) or resolved.name=='device-profile.private.json' or resolved.suffix.lower() not in EXTENSIONS:raise ValueError('不接受憑據或未支援文件類型')
    if must_exist and not resolved.is_file():raise FileNotFoundError('案件文件不存在')
    return resolved

def save_context(path, context):
    path=Path(path);temporary=path.with_name('.case-writing-'+uuid.uuid4().hex)
    try:
        temporary.write_bytes(canonical(context));os.replace(temporary,path)
    finally:temporary.unlink(missing_ok=True)

def begin(root, month, employees, home=None, operator_name=None, allowed_roots=None, thread_id=None, turn_ids=None):
    if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',month):raise ValueError('案件月份格式錯誤')
    root=Path(root).absolute()
    if root!=root.resolve() or root==Path(root.anchor):raise ValueError('請指定獨立案件資料夾')
    root.mkdir(parents=True,exist_ok=True);path=root/CONTEXT
    if path.exists():raise ValueError('案件已存在，請使用原案件補存，不自動替換人員或月份')
    if not isinstance(employees,list) or not 1<=len(employees)<=300:raise ValueError('須綁定 1 至 300 名已核定候選人員')
    people=[];keys=set()
    for person in employees:
        item={k:str(person.get(k,'')).strip() for k in ('employee_key','name','company')}
        if not all(item.values()) or any(len(v)>160 for v in item.values()):raise ValueError('人員須有穩定編號、姓名及公司')
        key=(item['company'],item['employee_key'])
        if key in keys:raise ValueError('同公司人員編號重複')
        keys.add(key);people.append(item)
    home=home_path(home);client=read_json_secret(home/'client.secure')
    roots=[str(root)]+[str(Path(p).resolve()) for p in allowed_roots or []]
    for value in roots:
        candidate=Path(value)
        if candidate==Path(candidate.anchor) or any(part in DENIED for part in candidate.parts):raise ValueError('不能指定整顆磁碟或私有設定目錄')
    thread_id=thread_id or os.environ.get('CODEX_THREAD_ID');conversation={'thread_id':None,'turn_ids':[],'bind_state':'not_bound'}
    if thread_id:
        identifier(thread_id)
        if turn_ids is None:
            try:
                from conversation_reader import latest_turn
                turn_ids=[latest_turn(thread_id)]
            except Exception:turn_ids=[]
        conversation={'thread_id':thread_id,'turn_ids':list(dict.fromkeys(turn_ids)),
                      'bind_state':'bound' if turn_ids else 'read_unavailable'}
    context={'schema':'payroll-case-1','case_id':str(uuid.uuid4()),'month':month,'root':str(root),
             'allowed_roots':roots,'employees':people,'operator':{'device_id':client['device_id'],'name':operator_name or getpass.getuser(),'identity_source':'claimed_name_or_os_login'},
             'conversation':conversation,'tracked_files':[],'created_at':now(),'closed':False}
    save_context(path,context)
    with db(home) as connection:connection.execute('INSERT INTO watches VALUES(?,?,0)',(context['case_id'],str(path)))
    return context

def arg(args, key):
    try:return args[args.index(key)+1]
    except (ValueError,IndexError):return None

def find_context(args):
    specified=os.environ.get('PAYROLL_CASE_CONTEXT')
    if specified:return Path(specified)
    for flag in ('--record','--output','--workbook','--attendance'):
        value=arg(args,flag)
        if not value:continue
        parent=Path(value).resolve().parent
        for _ in range(5):
            candidate=parent/CONTEXT
            if candidate.is_file():return candidate
            if parent==parent.parent:break
            parent=parent.parent
    return None

def explicit_files(args):
    files=[];record_path=arg(args,'--record')
    if record_path:
        path=Path(record_path).resolve();files.append((path,'run_record'))
        record=read_json(path)
        for source in record.get('sources',{}).values():
            if isinstance(source,dict) and source.get('path'):
                value=Path(source['path'])
                files.append((value if value.is_absolute() else path.parent/value,'result' if source.get('role')=='calculation' else 'source'))
    for flag in ('--workbook','--attendance','--roster','--history','--overtime','--plan','--input','--pension','--unemployment','--medical','--housing-fund','--source-periods'):
        value=arg(args,flag)
        if value:files.append((Path(value),'source'))
    output=arg(args,'--output')
    if output:
        files.append((Path(output),'result'))
        sidecar=Path(output+'.audit.json')
        if sidecar.is_file():files.append((sidecar,'result'))
    return files

def encrypt_attachment(raw, filename, role, case_id, org, file_id=None):
    if len(raw)>MAX_FILE:raise ValueError('單一文件超過 32 MiB，本版不截斷歸檔')
    if role not in ROLES or Path(filename).name!=filename or '\\' in filename or '/' in filename:raise ValueError('附件名稱或類型錯誤')
    file_id=identifier(file_id or str(uuid.uuid4()));key=os.urandom(32);nonce=os.urandom(12)
    aad=('payroll-case-file-1:'+case_id+':'+file_id).encode()
    cipher=AESGCM(key).encrypt(nonce,raw,aad)
    wrapped=public_key(org['public_key_pem']).encrypt(key,padding.OAEP(mgf=padding.MGF1(hashes.SHA256()),algorithm=hashes.SHA256(),label=None))
    meta={'file_id':file_id,'name':filename,'role':role,'bytes':len(raw),'sha256':digest(raw),'cipher_sha256':digest(cipher),
          'nonce':base64.b64encode(nonce).decode(),'wrapped_key':base64.b64encode(wrapped).decode(),
          'object_path':'cases/objects/'+file_id[:2]+'/'+file_id+'.bin'}
    return meta,cipher

def decrypt_attachment(cipher, meta, case_id, pem):
    identifier(case_id);file_id=identifier(meta['file_id'])
    if meta['object_path']!='cases/objects/'+file_id[:2]+'/'+file_id+'.bin' or len(cipher)>MAX_FILE+16 or digest(cipher)!=meta['cipher_sha256']:raise ValueError('密文附件損壞')
    if Path(meta['name']).name!=meta['name'] or '\\' in meta['name'] or '/' in meta['name']:raise ValueError('附件名稱越界')
    key=private_key(pem).decrypt(base64.b64decode(meta['wrapped_key'],validate=True),padding.OAEP(mgf=padding.MGF1(hashes.SHA256()),algorithm=hashes.SHA256(),label=None))
    raw=AESGCM(key).decrypt(base64.b64decode(meta['nonce'],validate=True),cipher,('payroll-case-file-1:'+case_id+':'+file_id).encode())
    if len(raw)!=meta['bytes'] or digest(raw)!=meta['sha256']:raise ValueError('原件還原摘要不符')
    return raw

def read_bound_conversations(bind):
    from conversation_reader import extract
    primary=extract(bind['thread_id'],bind['turn_ids']);additional=bind.get('additional_threads',[])
    if not additional:return primary
    bindings=[{'thread_id':bind['thread_id'],'turn_ids':bind['turn_ids']}]+additional
    if len(bindings)>10 or sum(len(item['turn_ids']) for item in bindings)>100:raise ValueError('案件綁定超過 10 個對話或 100 輪')
    groups=[primary]+[extract(item['thread_id'],item['turn_ids']) for item in additional]
    result=dict(primary,bindings=bindings,messages=[],tool_summary=[],completed_turns=all(item['completed_turns'] for item in groups),omitted_nontext_items=sum(item.get('omitted_nontext_items',0) for item in groups))
    for group in groups:
        result['messages'] += [dict(item,thread_id=group['thread_id']) for item in group['messages']]
        result['tool_summary'] += [dict(item,thread_id=group['thread_id']) for item in group.get('tool_summary',[])]
    if len(canonical(result))>4*1024*1024:raise ValueError('合併案件對話超過 4 MiB，未截斷冒充完整')
    return result

def _capture(context_path, files=None, home=None, org=None, entry='manual_checkpoint', conversation=None, close=False, add_current_turn=True, refresh_only=False, capture_issues=None):
    home=private_dir(home_path(home));org=org or organization();context_path=Path(context_path).resolve();context=read_json(context_path)
    if context.get('schema')!='payroll-case-1' or context_path.parent!=Path(context['root']).resolve():raise ValueError('案件設定位置不符')
    case_id=identifier(context['case_id']);client=read_json_secret(home/'client.secure')
    if context['operator']['device_id']!=client['device_id']:raise ValueError('案件不屬於本設備')
    if org['key_id']!=client['key_id'] or org['repository']!=client['repository']:raise ValueError('案件加密組織不符')
    preserved=context.get('pending_capture_issues',[]) if refresh_only else list(capture_issues or [])
    context['pending_capture_issues']=list(preserved)
    issues=list(preserved);tracked={(item['path'],item['role']):item for item in context.get('tracked_files',[])}
    for path,role in files or []:
        if role not in ROLES or role=='conversation':raise ValueError('文件用途錯誤')
        resolved=safe_under(path,context['allowed_roots']);tracked[(str(resolved),role)]={'path':str(resolved),'role':role}
    if len(tracked)>MAX_FILES:raise ValueError('案件文件超過 128 份，本版需分案，沒有略過多餘文件')
    context['tracked_files']=list(tracked.values());bind=context['conversation']
    if add_current_turn and os.environ.get('CODEX_THREAD_ID'):
        try:
            from conversation_reader import latest_turn
            current_thread=identifier(os.environ['CODEX_THREAD_ID']);current=latest_turn(current_thread)
            if not bind.get('thread_id'):bind.update(thread_id=current_thread,turn_ids=[],bind_state='bound')
            if current_thread==bind['thread_id']:target=bind
            else:
                additional=bind.setdefault('additional_threads',[])
                target=next((item for item in additional if item['thread_id']==current_thread),None)
                if target is None:
                    if len(additional)>=9:raise ValueError('案件超過 10 個綁定對話')
                    target={'thread_id':current_thread,'turn_ids':[]};additional.append(target)
            if current not in target['turn_ids']:target['turn_ids'].append(current)
        except Exception:issues.append('本輪對話讀取未完成')
    if conversation is None and bind.get('thread_id') and bind.get('turn_ids'):
        try:
            conversation=read_bound_conversations(bind)
        except Exception:issues.append('綁定對話匯出失敗，原件可獨立保存並等待重試')
    if conversation is not None:
        if conversation.get('thread_id')!=bind.get('thread_id') or conversation.get('turn_ids')!=bind.get('turn_ids'):raise ValueError('匯出的對話不屬於綁定輪次')
        if bind.get('additional_threads') and conversation.get('bindings')!=[{'thread_id':bind['thread_id'],'turn_ids':bind['turn_ids']}]+bind['additional_threads']:raise ValueError('跨對話匯出沒有完整綁定範圍')
        convo_state='completed_visible_text' if conversation.get('completed_turns') else 'in_progress'
        if conversation.get('omitted_nontext_items'):issues.append('對話包含尚未保存的非文字內容')
    else:convo_state='not_bound' if not bind.get('thread_id') else 'unavailable'
    descriptors=[];total=0;cached_files=[];file_keys=[]
    if refresh_only and context.get('last_file_cache_keys'):
        with db(home) as connection:
            for cache_key in context['last_file_cache_keys']:
                row=connection.execute('SELECT metadata FROM file_cache WHERE cache_key=?',(cache_key,)).fetchone()
                if not row:raise ValueError('原件保存索引缺失')
                meta=json.loads(unprotect(row[0]));stored=home/'case-objects'/(meta['file_id']+'.bin')
                if not stored.is_file() or digest(stored.read_bytes())!=meta['cipher_sha256']:raise ValueError('已保存原件的密文損壞')
                cached_files.append(meta);file_keys.append(cache_key)
    else:
        for item in context['tracked_files']:
            path=safe_under(item['path'],context['allowed_roots']);before=path.stat()
            if before.st_size>MAX_FILE:raise ValueError('單一文件超過 32 MiB')
            raw=path.read_bytes();after=path.stat()
            if before.st_mtime_ns!=after.st_mtime_ns or before.st_size!=after.st_size or len(raw)!=before.st_size:raise ValueError('文件正在變更，等待穩定後重試')
            if path.is_relative_to(Path(context['root'])):source_label=path.relative_to(Path(context['root'])).as_posix()
            else:
                root_index=next(i for i,root in enumerate(context['allowed_roots']) if path==Path(root) or path.is_relative_to(Path(root)))
                relative=path.relative_to(Path(context['allowed_roots'][root_index])).as_posix()
                source_label='指定來源'+str(root_index)+'/'+(relative if relative!='.' else path.name)
            total+=len(raw);descriptors.append((path.name,item['role'],raw,source_label))
    if conversation is not None:descriptors.append(('案件對話.json','conversation',canonical(conversation),'對話/案件對話.json'))
    if total>MAX_TOTAL:raise ValueError('案件超過 256 MiB，未截斷歸檔')
    context['closed']=context['closed'] or close
    fingerprint=digest(canonical({'files':[(meta['name'],meta['role'],meta['sha256'],meta.get('source_label',meta['name'])) for meta in cached_files]+[(name,role,digest(raw),label) for name,role,raw,label in descriptors],
                                 'people':context['employees'],'conversation':convo_state,'issues':issues,'closed':context['closed']}))
    if context.get('last_snapshot_fingerprint')==fingerprint:
        save_context(context_path,context)
        with db(home) as connection:connection.execute('UPDATE watches SET done=? WHERE case_id=?',(int(context['closed'] and convo_state=='completed_visible_text'),case_id))
        return {'state':'unchanged','case_id':case_id,'conversation_state':convo_state,'files':len(descriptors)}
    attachments=list(cached_files);object_dir=private_dir(home/'case-objects')
    with db(home) as connection:
        for name,role,raw,source_label in descriptors:
            cache_key=digest(canonical([case_id,source_label,name,role,digest(raw)]))
            old=connection.execute('SELECT metadata FROM file_cache WHERE cache_key=?',(cache_key,)).fetchone()
            if old:
                meta=json.loads(unprotect(old[0]));stored=object_dir/(meta['file_id']+'.bin')
                if not stored.is_file() or digest(stored.read_bytes())!=meta['cipher_sha256']:raise ValueError('本機附件隊列損壞')
            else:
                meta,cipher=encrypt_attachment(raw,name,role,case_id,org)
                meta['source_label']=source_label
                target=object_dir/(meta['file_id']+'.bin')
                with target.open('xb') as stream:stream.write(cipher);stream.flush();os.fsync(stream.fileno())
                connection.execute('INSERT INTO file_cache VALUES(?,?)',(cache_key,protect(canonical(meta))))
            attachments.append(meta)
            if role!='conversation':file_keys.append(cache_key)
        event={'schema':'payroll-case-snapshot-1','event_id':str(uuid.uuid4()),'case_id':case_id,'device_id':client['device_id'],
               'month':context['month'],'operator':context['operator'],'employees':context['employees'],'created_at':now(),
               'entry':entry,'files':attachments,'conversation_state':convo_state,'conversation_binding':bind,
               'issues':issues,'closed':context['closed'],'synthetic_only':context.get('synthetic_only',False)}
        envelope=sign_envelope(encrypt(event,org['public_key_pem'],CASE_EVENT),client['device_id'],client['ssh_private_pem'].encode())
        connection.execute('INSERT INTO snapshots VALUES(?,?,?,?,?)',(event['event_id'],case_id,canonical(envelope).decode(),'queued',time.time()))
    context['last_snapshot_fingerprint']=fingerprint;context['last_snapshot_id']=event['event_id'];context['last_file_cache_keys']=file_keys
    save_context(context_path,context)
    with db(home) as connection:connection.execute('UPDATE watches SET done=? WHERE case_id=?',(int(context['closed'] and convo_state=='completed_visible_text'),case_id))
    return {'state':'queued','case_id':case_id,'snapshot_id':event['event_id'],'files':len(attachments),'conversation_state':convo_state,'issues':issues}

def capture(context_path, files=None, home=None, org=None, entry='manual_checkpoint', conversation=None, close=False, add_current_turn=True, refresh_only=False, capture_issues=None):
    context=read_json(context_path);case_id=identifier(context['case_id']);target=home_path(home)
    with lock(target/'case-locks'/case_id) as acquired:
        if not acquired:raise RuntimeError('同一案件正在保存，保留待重試')
        return _capture(context_path,files,target,org,entry,conversation,close,add_current_turn,refresh_only,capture_issues)

def record_entry(entry, args, home=None, org=None):
    home=home_path(home);path=find_context(args)
    if not path:
        result={'at':now(),'state':'context_missing','message':'尚未建立案件，未保存原件及對話'}
        (home/'case-health.json').write_bytes(canonical(result));return result
    context=read_json(path);months=[]
    if arg(args,'--month'):months.append(arg(args,'--month'))
    if arg(args,'--record'):
        month=read_json(arg(args,'--record')).get('month')
        if month:months.append(month)
    if any(month!=context['month'] for month in months):raise ValueError('程式所屬月份與保存案件不符')
    selected=[];issues=[]
    for file,role in explicit_files(args):
        if role=='result' and not file.is_file():issues.append('本次計算沒有產生指定結果文件，未當成已保存');continue
        selected.append((file,role))
    result=capture(path,selected,home,org,entry,capture_issues=issues)
    (home/'case-health.json').write_bytes(canonical(dict(result,at=now())))
    return result

def refresh(home=None,org=None):
    home=home_path(home);results=[]
    with db(home) as connection:items=connection.execute('SELECT path FROM watches WHERE done=0 ORDER BY rowid DESC LIMIT 5').fetchall()
    for (path,) in items:
        try:results.append(capture(path,[],home,org,'conversation_refresh',add_current_turn=False,refresh_only=True))
        except Exception as exc:results.append({'state':'pending','category':type(exc).__name__})
    return results

def sync(home=None,org=None):
    home=home_path(home);org=org or organization();client=read_json_secret(home/'client.secure')
    if client['repository']!=org['repository'] or client['key_id']!=org['key_id']:raise ValueError('案件上傳組織與密鑰不符')
    with lock(home) as acquired:
        if not acquired:return {'state':'busy','sent':0}
        with db(home) as connection:rows=connection.execute("SELECT id,envelope FROM snapshots WHERE state='queued' ORDER BY created LIMIT 10").fetchall()
        if not rows:return {'state':'up_to_date','sent':0}
        with client_git_environment(home,client,org,use_http_proxy=True) as env:
            with db(home) as connection:cached=connection.execute('SELECT metadata FROM file_cache').fetchall()
            def payloads():
                for (protected,) in cached:
                    meta=json.loads(unprotect(protected));source=home/'case-objects'/(meta['file_id']+'.bin');raw=source.read_bytes()
                    if digest(raw)!=meta['cipher_sha256']:raise ValueError('附件隊列摘要不符')
                    yield meta['object_path'],raw
                for event_id,payload in rows:yield 'cases/manifests/'+client['device_id']+'/'+identifier(event_id)+'.json',payload.encode('utf-8')
            for attempt in range(3):
                repo=cache(home,'ssh://git@ssh.github.com:443/'+org['repository']+'.git',env,transport_timeout=180,thin=True)
                if not stage_blobs(repo,payloads(),env):break
                commit_index(repo,'追加加密核算案件及附件',env)
                try:git_run(['push','origin','HEAD:main'],cwd=repo,env=env,timeout=900);break
                except (RuntimeError,subprocess.TimeoutExpired):
                    if attempt==2:raise
        with db(home) as connection:connection.executemany("UPDATE snapshots SET state='synced' WHERE id=?",[(row[0],) for row in rows])
        return {'state':'synced','sent':len(rows)}

def main():
    for stream in (sys.stdout,sys.stderr):
        if hasattr(stream,'reconfigure'):stream.reconfigure(encoding='utf-8')
    parser=argparse.ArgumentParser(description='公司薪資案件保存')
    parser.add_argument('command',choices=['begin','checkpoint','flush','status']);parser.add_argument('--root',type=Path);parser.add_argument('--context',type=Path);parser.add_argument('--month');parser.add_argument('--people',type=Path);parser.add_argument('--operator');parser.add_argument('--allow-root',action='append',default=[]);parser.add_argument('--file',action='append',default=[]);parser.add_argument('--result',action='append',default=[]);parser.add_argument('--thread-id');parser.add_argument('--close',action='store_true');parser.add_argument('--home',type=Path)
    args=parser.parse_args();home=home_path(args.home)
    if args.command=='begin':result=begin(args.root,args.month,read_json(args.people),home,args.operator,args.allow_root,args.thread_id);result={'結果':'案件已建立','案件編號':result['case_id'],'對話綁定':result['conversation']['bind_state']}
    elif args.command=='checkpoint':result=capture(args.context or args.root/CONTEXT,[(Path(x),'source') for x in args.file]+[(Path(x),'result') for x in args.result],home,close=args.close)
    elif args.command=='flush':result={'refresh':refresh(home),'sync':sync(home)}
    else:
        with db(home) as connection:result={'snapshots':dict(connection.execute('SELECT state,COUNT(*) FROM snapshots GROUP BY state'))}
    print(json.dumps(result,ensure_ascii=False))

if __name__=='__main__':
    try:main()
    except Exception as exc:print(json.dumps({'結果':'案件保存未完成','錯誤類別':type(exc).__name__,'說明':'原件未改動；不能宣稱已完整保存。'},ensure_ascii=False));raise SystemExit(2)
