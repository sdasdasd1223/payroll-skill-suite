"""只讀取明確綁定的 Codex 對話與指定輪次；不列出其他聊天。"""
from __future__ import annotations
import json, os, queue, re, shutil, subprocess, threading, time, uuid

def redact(text):
    text=re.sub(r'-----BEGIN [^-]*PRIVATE KEY-----.*?-----END [^-]*PRIVATE KEY-----','[已遮蔽私鑰]',str(text),flags=re.S)
    text=re.sub(r'(?i)\bBearer\s+[A-Za-z0-9._~+/=-]{12,}','Bearer [已遮蔽憑據]',text)
    text=re.sub(r'\bsk-[A-Za-z0-9_-]{20,}','[已遮蔽 API 金鑰]',text)
    return text

class RPC:
    ALLOWED={'initialize','thread/turns/list','thread/read'}
    def __init__(self, executable=None):
        executable=executable or shutil.which('codex')
        if not executable:raise RuntimeError('未找到 Codex 應用服務')
        self.p=subprocess.Popen([executable,'app-server','--stdio'],stdin=subprocess.PIPE,stdout=subprocess.PIPE,stderr=subprocess.DEVNULL,creationflags=0x08000000 if os.name=='nt' else 0)
        self.q=queue.Queue();self.number=0
        def read():
            try:
                while True:
                    line=self.p.stdout.readline(32*1024*1024+1)
                    if not line:break
                    if len(line)>32*1024*1024:self.q.put(RuntimeError('對話回應超出上限'));break
                    try:self.q.put(json.loads(line))
                    except ValueError:continue
            finally:self.q.put(EOFError())
        threading.Thread(target=read,daemon=True).start()
        try:
            self.request('initialize',{'clientInfo':{'name':'company_payroll_case_archive','version':'2026.11.17.2'},'capabilities':{'experimentalApi':True}})
            self.p.stdin.write(b'{"method":"initialized"}\n');self.p.stdin.flush()
        except BaseException:self.close();raise
    def request(self,method,params,timeout=20):
        if method not in self.ALLOWED:raise ValueError('不允許變更對話')
        self.number+=1;identifier=self.number
        self.p.stdin.write((json.dumps({'id':identifier,'method':method,'params':params})+'\n').encode());self.p.stdin.flush()
        deadline=time.monotonic()+timeout
        while True:
            item=self.q.get(timeout=max(.01,deadline-time.monotonic()))
            if isinstance(item,Exception):raise RuntimeError('Codex 對話服務中斷')
            if item.get('id')!=identifier:continue
            if 'error' in item:raise RuntimeError('對話讀取不受支援：'+str(item['error'].get('code')))
            return item['result']
    def close(self):
        if getattr(self,'p',None):
            self.p.terminate()
            try:self.p.wait(timeout=3)
            except subprocess.TimeoutExpired:self.p.kill();self.p.wait(timeout=3)
    def __enter__(self):return self
    def __exit__(self,*args):self.close()

def latest_turn(thread_id, rpc=None):
    str(uuid.UUID(thread_id))
    if rpc:return rpc.request('thread/turns/list',{'threadId':thread_id,'limit':1,'sortDirection':'desc','itemsView':'full'})['data'][0]['id']
    with RPC() as client:return latest_turn(thread_id,client)

def extract(thread_id, turn_ids, rpc=None):
    str(uuid.UUID(thread_id));wanted=set(turn_ids)
    if not wanted:raise ValueError('沒有綁定對話輪次')
    if len(wanted)>100:raise ValueError('本案件超過 100 輪，須分段歸檔')
    if rpc is None:
        with RPC() as client:return extract(thread_id,turn_ids,client)
    found={};cursor=None;seen=set()
    for _ in range(40):
        params={'threadId':thread_id,'limit':50,'sortDirection':'desc','itemsView':'full'}
        if cursor:params['cursor']=cursor
        page=rpc.request('thread/turns/list',params)
        for turn in page['data']:
            if turn['id'] in wanted:found[turn['id']]=turn
        if wanted<=found.keys():break
        cursor=page.get('nextCursor')
        if not cursor or cursor in seen:break
        seen.add(cursor)
    if wanted!=found.keys():raise ValueError('綁定輪次缺失，不能當成完整對話')
    messages=[];omitted=0;tool_events=[];complete=True
    for identifier in turn_ids:
        turn=found[identifier];complete=complete and turn.get('status')=='completed'
        for item in turn.get('items',[]):
            kind=item.get('type')
            if kind=='userMessage':
                fragments=[]
                for part in item.get('content',[]):
                    if part.get('type')=='text':fragments.append(redact(part.get('text','')))
                    else:omitted+=1
                messages.append({'turn_id':identifier,'role':'user','text':'\n'.join(fragments)})
            elif kind=='agentMessage':messages.append({'turn_id':identifier,'role':'assistant','text':redact(item.get('text','')),'phase':item.get('phase')})
            elif kind in {'commandExecution','fileChange','mcpToolCall','dynamicToolCall'}:
                tool_events.append({'turn_id':identifier,'kind':kind,'status':item.get('status'),'exit_code':item.get('exitCode')})
            # 不保存 reasoning、隱藏思考、任意工具輸出或原始命令。
    result={'schema':'payroll-case-conversation-1','thread_id':thread_id,'turn_ids':turn_ids,'messages':messages,
            'tool_summary':tool_events,'completed_turns':complete,'omitted_nontext_items':omitted,
            'coverage':'bound_turns_visible_text_only'}
    if len(json.dumps(result,ensure_ascii=False).encode())>4*1024*1024:raise ValueError('案件對話超過 4 MiB，未截斷冒充完整')
    return result

if __name__=='__main__':
    current=os.environ.get('CODEX_THREAD_ID')
    with RPC() as client:
        identifier=latest_turn(current,client);data=extract(current,[identifier],client)
    print(json.dumps({'讀取入口':'本機 Codex 應用服務，只讀','綁定輪次':1,'可見訊息數':len(data['messages']),
                      '目前輪次已完成':data['completed_turns'],'非文字項目未保存數':data['omitted_nontext_items']},ensure_ascii=False))
