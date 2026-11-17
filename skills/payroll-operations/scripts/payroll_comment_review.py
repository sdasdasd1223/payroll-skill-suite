"""原工资批注的业务处置复核；对象复制不等于清算事项已处理。"""
import hashlib
from pathlib import Path
import re
from decimal import Decimal,InvalidOperation
from openpyxl import load_workbook

ACTIONS={'adopted','explanation','historical_not_applicable','pending'}

def settlement_comment(text):
    return bool(re.search(r'事假[\s\S]*(?:清算|结算|結算|结清|結清)|(?:清算|结算|結算|结清|結清)[\s\S]*事假',text))

def review_comments(inventory, people, month, dispositions, resolve_ref, resolve_components=None):
    """inventory必须来自实际template_guard.verify；resolve_ref返回实际值/行/源。"""
    rows=[]
    for group in (inventory or {}).get('sources',[]):
        source=group['source'];path=Path(source['path'])
        if hashlib.sha256(path.read_bytes()).hexdigest()!=source['sha256']:raise ValueError('业务批注源SHA已变化')
        book=load_workbook(path,data_only=False)
        try:
            for item in group.get('items',[]):
                original=book[item['sheet']][item['cell']].comment
                if original is None or original.text!=item['text']:raise ValueError('实际原批注与盘点不符')
                text=original.text;identities=item.get('identity') or []
                owners=[key for key,p in people.items() if p['name'] in identities]
                risky=settlement_comment(text) and item.get('status')!='ignored_historical_sheet'
                row={'source_path':str(path),'source_sha256':source['sha256'],'source_sheet':item['sheet'],'source_cell':item['cell'],'original_text':text,'object_status':item.get('status'),'employee_key':owners[0] if len(owners)==1 else None,'name':people[owners[0]]['name'] if len(owners)==1 else ';'.join(identities),'requires_business_resolution':risky,'action':'explanation','effective_status':'resolved','reason':'普通说明；未识别到本门槛限定的事假清算事项','evidence_refs':[]}
                rows.append(row)
                if not risky:continue
                row.update(action='pending',effective_status='pending',reason='原批注涉及事假清算，尚无实际本人本月处置依据')
                if len(owners)!=1:row['reason']='清算批注不能唯一定位本次人员，需核人员去向';continue
                key=owners[0]
                actual_names=[str(c.value or '').strip() for c in book[item['sheet']][item['row']]]
                if people[key]['name'] not in actual_names:raise ValueError('原批注实际行身份与盘点不符')
                choices=[d for d in dispositions if d.get('employee_key')==key and d.get('source_sha256')==source['sha256'] and d.get('source_sheet')==item['sheet'] and d.get('source_cell')==item['cell']]
                if len(choices)!=1:continue
                decision=choices[0];action=decision.get('action')
                if action not in ACTIONS or not decision.get('reason'):row['reason']='逐条处置缺明确采用/仅说明/历史不适用/待确认及理由';continue
                row.update(action=action,reason=decision['reason'])
                if action=='pending':continue
                refs=decision.get('evidence_refs',[])
                try:
                    if not refs:raise ValueError('缺既有实际处置来源')
                    texts=[]
                    for ref in refs:
                        value,actual,origin=resolve_ref(ref,key)
                        if origin.get('month')!=month or actual.get('month')!=month:raise ValueError('处置证据缺实际本月归属或属于其他月份')
                        texts.append(str(value))
                    evidence='\n'.join(texts)
                    if not settlement_comment(evidence):raise ValueError('本人来源未定位事假清算事项；全员泛称未确认不能代替')
                    if any(t in evidence for t in ('待确认','待確認','未确认','未確認','待核','尚未')):raise ValueError('实际来源仍为未决，不能自报已处置')
                    tokens={'adopted':('已计入','已計入','本月采用','本月採用','本月结算','本月結算'),'explanation':('仅说明','僅說明','仅备注','僅備註','不影响本月','不影響本月'),'historical_not_applicable':('已结算','已結算','已结清','已結清','历史不适用','歷史不適用','不重复','不重複')}
                    if not any(t in evidence for t in tokens[action]):raise ValueError('实际处置文字不支持所选处理')
                    if action=='adopted':
                        ref=decision.get('amount_ref')
                        if not ref:raise ValueError('采用清算须指向本人已核扣款金额，程序不代算')
                        if not any((r.get('table'),r.get('row'))==(ref.get('table'),ref.get('row')) for r in refs):raise ValueError('采用金额须来自同一本人处置记录行，不能引用无关数字')
                        value,actual,origin=resolve_ref(ref,key)
                        if origin.get('month')!=month or actual.get('month')!=month or isinstance(value,bool) or not isinstance(value,(int,float)):raise ValueError('采用金额不是本人本月已核数字')
                        row['adopted_amount']=value;row['amount_ref']=ref
                        scope=decision.get('amount_scope','component')
                        if scope=='total_leave':
                            if not any(t in evidence for t in ('完整请假扣款','完整請假扣款','请假扣款总额','請假扣款總額','S总额','S總額')):raise ValueError('原处置未明确核定完整S总额，不能把清算子项当完整请假扣款')
                            row['adopted_total_leave']=value
                        elif scope=='component':
                            components=decision.get('component_refs',[]);total_ref=decision.get('total_leave_ref')
                            if not components or not total_ref or resolve_components is None:raise ValueError('清算仅为子项，缺完整请假扣款组成及总额依据，继续待核')
                            locations=[(r.get('table'),r.get('row'),r.get('column')) for r in components]
                            if len(locations)!=len(set(locations)):raise ValueError('清算组成引用重复，不能重复计入')
                            if (ref.get('table'),ref.get('row'),ref.get('column')) not in locations:raise ValueError('完整组成未包含本人已核清算子项')
                            tables={r.get('table') for r in components}
                            if len(tables)!=1:raise ValueError('完整组成须引用同一实际明细表，以核全项覆盖')
                            expected=resolve_components(next(iter(tables)),key)
                            if set(locations)!={(r['table'],r['row'],r['column']) for r in expected}:raise ValueError('完整请假扣款组成遗漏或越出实际同人本月明细')
                            subtotal=Decimal(0);ids=set()
                            for part in components:
                                amount,actual,origin=resolve_ref(part,key)
                                component_id=actual.get('component_id')
                                if not component_id or component_id in ids:raise ValueError('实际明细项目缺唯一标识或重复，不能重复计入清算')
                                ids.add(component_id)
                                if origin.get('month')!=month or actual.get('month')!=month or isinstance(amount,bool) or not isinstance(amount,(int,float)):raise ValueError('完整组成仍有非本人本月已核数字')
                                subtotal+=Decimal(str(amount))
                            total,actual,origin=resolve_ref(total_ref,key)
                            if total_ref.get('column')!='total_leave' or origin.get('month')!=month or actual.get('month')!=month or isinstance(total,bool) or not isinstance(total,(int,float)):raise ValueError('缺实际本人本月完整请假扣款总额定位')
                            if not subtotal.is_finite() or not Decimal(str(total)).is_finite() or subtotal!=Decimal(str(total)):raise ValueError('实际完整组成合计与已核完整请假扣款总额不符')
                            row.update(adopted_total_leave=total,component_refs=components,total_leave_ref=total_ref,component_sum=str(subtotal))
                        else:raise ValueError('清算金额范围须为component或明确total_leave')
                    row.update(effective_status='resolved',evidence_refs=refs)
                except (ValueError,KeyError,TypeError,InvalidOperation) as exc:row.update(effective_status='pending',reason=str(exc))
        finally:book.close()
        if hashlib.sha256(path.read_bytes()).hexdigest()!=source['sha256']:raise ValueError('读取业务批注期间源SHA变化')
    return rows
