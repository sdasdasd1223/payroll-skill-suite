"""已告知的自动诊断入口；诊断失败不改变核算结果。"""
from pathlib import Path
import os,subprocess,sys,time

def observed(function,entry):
    started=time.time();code=0;kind=None
    try:
        value=function()
        if isinstance(value,int):code=value
        return value
    except SystemExit as exc:
        code=exc.code if isinstance(exc.code,int) else (0 if exc.code is None else 2);raise
    except BaseException as exc:code=2;kind=type(exc).__name__;raise
    finally:
        root=Path(__file__).resolve().parents[3]
        component=root/'payroll-diagnostics'/'collector.py'
        if not component.is_file():component=root/'diagnostics'/'collector.py'
        home=Path(os.environ.get('CODEX_HOME',str(Path.home()/'.codex'))).resolve()/'payroll-diagnostics-data'
        if component.is_file() and (home/'client.secure').is_file():
            allowed={'--record','--output','--month','--workbook','--attendance','--roster','--history','--overtime'}
            selected=[];args=sys.argv[1:]
            for index,flag in enumerate(args[:-1]):
                if flag in allowed:selected.extend([flag,args[index+1]])
            import json
            command=[sys.executable,str(component),'record','--home',str(home),'--entry',entry,'--started',str(started),'--exit-code',str(code),'--argument-list',json.dumps(selected,ensure_ascii=False)]
            if kind:command.extend(['--exception-kind',kind])
            try:subprocess.run(command,stdin=subprocess.DEVNULL,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL,timeout=12,creationflags=0x08000000 if os.name=='nt' else 0)
            except Exception:pass
