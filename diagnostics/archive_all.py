"""管理員每日入口：診斷與案件均成功才回報整體成功。"""
from pathlib import Path
import argparse, datetime, json, os
from diagnostic_archive import archive as archive_diagnostics, atomic_json
from case_archive import CaseReader,archive_cases

def main():
    parser=argparse.ArgumentParser(description='公司薪資文件、結果、對話及診斷歸檔')
    parser.add_argument('--config',required=True,type=Path);args=parser.parse_args();config=json.loads(args.config.read_text('utf-8'))
    output=Path(config['status_file']);reader=None
    try:
        reader=CaseReader()
        diagnostic=archive_diagnostics(config['diagnostic_root'],config['admin_home'],reader)
        cases=archive_cases(config['case_root'],config['admin_home'],reader)
        result={'state':'success','at':datetime.datetime.now(datetime.timezone.utc).isoformat(),'diagnostic_events':diagnostic['verified_events'],
                'case_snapshots':cases['verified_snapshots'],'case_objects':cases['verified_objects'],'cloud_deleted':False}
        atomic_json(output,result);print(json.dumps({'結果':'整體歸檔成功','診斷紀錄':result['diagnostic_events'],'案件版本':result['case_snapshots'],'附件':result['case_objects'],'雲端刪除':False},ensure_ascii=False));return 0
    except Exception as exc:
        atomic_json(output,{'state':'failed','at':datetime.datetime.now(datetime.timezone.utc).isoformat(),'category':type(exc).__name__,'cloud_deleted':False})
        print('整體歸檔未完成：'+type(exc).__name__+'；已成功的個別歸檔保留，不能當成本次全部成功。');return 2
if __name__=='__main__':raise SystemExit(main())
