"""管理員限定：驗證案件、按人員月份查看、按需還原原件。"""
from __future__ import annotations
import argparse, hmac, http.server, json, mimetypes, os, re, secrets, sqlite3, sys, threading, urllib.parse, uuid
from pathlib import Path
from common import canonical, decrypt, digest, load_secret, private_dir, read_json, verify_sender
from case_store import MAX_FILE, MAX_FILES, CASE_EVENT, ROLES, decrypt_attachment, identifier

def safe_file(root, relative):
    root=Path(root).resolve();path=root/relative
    if not path.resolve().is_relative_to(root) or path.is_symlink() or path.resolve()==root:raise ValueError('案件檔案路徑越界')
    for part in path.parents:
        if part==root:break
        if part.is_symlink() or getattr(part,'is_junction',lambda:False)():raise ValueError('案件檔案不可經過連結')
    return path

def verify_snapshot(envelope, folder, org, registry, pem, restore=True, fallback=None):
    sender=identifier(envelope['sender_device']);device=registry.get(sender)
    if not device or envelope.get('key_id')!=org['key_id']:raise ValueError('來源設備或管理金鑰不符')
    verify_sender(envelope,device['public_key']);event=decrypt(envelope,pem,CASE_EVENT)
    if event.get('schema')!='payroll-case-snapshot-1' or event['device_id']!=sender or event['operator']['device_id']!=sender:raise ValueError('案件身份不符')
    identifier(event['case_id']);identifier(event['event_id'])
    if not re.fullmatch(r'\d{4}-(0[1-9]|1[0-2])',event.get('month','')) or not isinstance(event.get('employees'),list) or not 1<=len(event['employees'])<=300:raise ValueError('案件月份或人員名單錯誤')
    people=set()
    for person in event['employees']:
        if any(not isinstance(person.get(k),str) or not person[k].strip() or len(person[k])>160 for k in ('employee_key','name','company')):raise ValueError('案件人員資料不完整')
        key=(person['company'],person['employee_key'])
        if key in people:raise ValueError('案件人員編號重複')
        people.add(key)
    if not isinstance(event.get('files'),list) or len(event['files'])>MAX_FILES+1:raise ValueError('案件附件數錯誤')
    seen=set()
    for meta in event['files']:
        file_id=identifier(meta['file_id'])
        if file_id in seen or meta['role'] not in ROLES or not 0<=meta['bytes']<=MAX_FILE:raise ValueError('附件編號或範圍錯誤')
        seen.add(file_id)
        if meta['object_path']!='cases/objects/'+file_id[:2]+'/'+file_id+'.bin':raise ValueError('密文路徑錯誤')
        path=safe_file(folder,meta['object_path'])
        if not path.is_file() and fallback is not None:path=safe_file(fallback,meta['object_path'])
        if path.stat().st_size>MAX_FILE+16:raise ValueError('密文超出上限')
        raw=path.read_bytes()
        if digest(raw)!=meta['cipher_sha256']:raise ValueError('密文摘要不符')
        if restore:decrypt_attachment(raw,meta,event['case_id'],pem)
    return event

def load_cases(folder, home, restore=True):
    folder=Path(folder).resolve();home=Path(home).resolve();org=read_json(home/'organization.json');registry=read_json(home/'devices.json');pem=load_secret(home/'admin-key.secure')
    cases=[];rejected=[]
    for path in sorted(folder.glob('cases/manifests/*/*.json')):
        try:
            relative=path.relative_to(folder).as_posix();safe_file(folder,relative)
            if path.stat().st_size>1024*1024:raise ValueError('案件清單過大')
            event=verify_snapshot(read_json(path),folder,org,registry,pem,restore)
            if path.stem!=event['event_id'] or path.parent.name!=event['device_id']:raise ValueError('案件路徑與身份不符')
            event['executor_label']=registry[event['device_id']].get('label',event['device_id'])
            cases.append(event)
        except Exception as exc:rejected.append({'id':path.stem,'category':type(exc).__name__})
    cases.sort(key=lambda x:x['created_at'],reverse=True)
    return cases,rejected

def filter_cases(cases, month='', employee='', operator=''):
    result=[]
    for case in cases:
        if month and case['month']!=month:continue
        if employee and not any(employee.casefold() in (' '.join(str(p.get(k,'')) for k in ('name','employee_key','company'))).casefold() for p in case['employees']):continue
        if operator and operator.casefold() not in (case.get('executor_label','')+' '+case['operator'].get('name','')+' '+case['device_id']).casefold():continue
        result.append(case)
    return result

def file_bytes(case, file_id, folder, home):
    meta=next((item for item in case['files'] if item['file_id']==file_id),None)
    if not meta:raise FileNotFoundError('案件未包含此附件')
    cipher=safe_file(folder,meta['object_path']).read_bytes()
    if len(cipher)>MAX_FILE+16:raise ValueError('附件過大')
    return meta,decrypt_attachment(cipher,meta,case['case_id'],load_secret(Path(home)/'admin-key.secure'))

def create_server(folder, home, port=0):
    folder=Path(folder).resolve();home=private_dir(Path(home).resolve());cases,rejected=load_cases(folder,home)
    if rejected:raise ValueError('存在驗證失敗案件，請先查看歸檔狀態，未建立誤導性的完整頁面')
    token=secrets.token_urlsafe(32);prefix='/session/'+token;lookup={case['event_id']:case for case in cases}
    class Handler(http.server.BaseHTTPRequestHandler):
        def log_message(self,*args):pass
        def do_GET(self):
            path=urllib.parse.urlsplit(self.path);parts=path.path.split('/')
            if len(parts)<3 or parts[1]!='session' or not hmac.compare_digest(parts[2],token):self.send_error(404);return
            host=self.headers.get('Host','')
            if host not in {'127.0.0.1:'+str(self.server.server_port),'localhost:'+str(self.server.server_port)}:self.send_error(403);return
            route='/'+'/'.join(parts[3:])
            try:
                if route in {'/',''}:
                    body=Path(__file__).with_name('case_dashboard.html').read_bytes();mime='text/html; charset=utf-8';disposition=None
                elif route=='/api/cases':
                    params=urllib.parse.parse_qs(path.query)
                    def value(key):return params.get(key,[''])[0]
                    selected=filter_cases(cases,value('month'),value('employee'),value('operator'))
                    body=canonical({'cases':selected,'rejected':len(rejected),'snapshot_only':True});mime='application/json; charset=utf-8';disposition=None
                elif route.startswith('/api/file/'):
                    unused,api,kind,snapshot_id,file_id=route.split('/')
                    case=lookup.get(identifier(snapshot_id))
                    if not case:raise FileNotFoundError()
                    meta,body=file_bytes(case,identifier(file_id),folder,home)
                    # 原件僅供下載；不把上傳的 HTML/腳本當成本機網站執行。
                    mime='application/octet-stream';disposition="attachment; filename*=UTF-8''"+urllib.parse.quote(meta['name'],safe='')
                elif route.startswith('/api/conversation/'):
                    snapshot_id=identifier(route.rsplit('/',1)[1]);case=lookup.get(snapshot_id)
                    if not case:raise FileNotFoundError()
                    meta=next((item for item in case['files'] if item['role']=='conversation'),None)
                    if not meta:raise FileNotFoundError()
                    _,body=file_bytes(case,meta['file_id'],folder,home);mime='application/json; charset=utf-8';disposition=None
                else:self.send_error(404);return
                self.send_response(200);self.send_header('Content-Type',mime);self.send_header('Content-Length',str(len(body)))
                self.send_header('Cache-Control','no-store');self.send_header('X-Content-Type-Options','nosniff');self.send_header('Referrer-Policy','no-referrer')
                self.send_header('Content-Security-Policy',"default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; connect-src 'self'; base-uri 'none'; frame-ancestors 'none'; form-action 'none'")
                if disposition:self.send_header('Content-Disposition',disposition)
                self.end_headers();self.wfile.write(body)
            except (ValueError,KeyError,FileNotFoundError,StopIteration):self.send_error(404)
            except Exception:self.send_error(500,'案件驗證或還原失敗')
    server=http.server.ThreadingHTTPServer(('127.0.0.1',port),Handler)
    return server,'http://127.0.0.1:'+str(server.server_port)+prefix+'/'

def main():
    parser=argparse.ArgumentParser(description='管理員公司案件查看')
    parser.add_argument('command',choices=['check','view']);parser.add_argument('--folder',required=True,type=Path);parser.add_argument('--home',type=Path);parser.add_argument('--port',type=int,default=0)
    args=parser.parse_args();home=args.home or Path(os.environ.get('CODEX_HOME',str(Path.home()/'.codex')))/'payroll-diagnostics-admin'
    if args.command=='check':
        cases,rejected=load_cases(args.folder,home);print(json.dumps({'驗證案件版本數':len(cases),'拒收數':len(rejected),'拒收':rejected},ensure_ascii=False));return 2 if rejected else 0
    server,url=create_server(args.folder,home,args.port)
    output=private_dir(home)/'案件查看入口.txt';output.write_text(url,encoding='utf-8')
    print('管理員本機案件查看入口：'+url,flush=True)
    print('僅本機臨時入口；關閉此程式後失效。原件按需解密下載，未寫入公開網站。',flush=True)
    timer=threading.Timer(3600,server.shutdown);timer.daemon=True;timer.start()
    try:server.serve_forever()
    except KeyboardInterrupt:pass
    finally:timer.cancel();server.server_close()
    return 0
if __name__=='__main__':
    try:raise SystemExit(main())
    except Exception as exc:print('案件查看未完成：'+type(exc).__name__+'；不能當成驗證成功。');raise SystemExit(2)
