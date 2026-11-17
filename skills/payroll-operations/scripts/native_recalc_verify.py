"""只在隔离副本执行真实Excel复核；不接收外部passed作为证据。
先核ProgID注册及实际进程的厂商/文件身份，记录真实引擎，无法证实则待核。
"""
import hashlib
import json
import math
import re
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from zipfile import ZipFile
from xml.etree import ElementTree as ET

def sha256(path): return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def same(a, b):
    if isinstance(a, (int, float)) and not isinstance(a, bool) and isinstance(b, (int, float)) and not isinstance(b, bool):
        return math.isfinite(a) and math.isfinite(b) and abs(a-b) <= max(0.000001, abs(a)*0.000000001)
    return a == b

def observe_cell(cell):
    """保存真实显示及ISERROR类型；####仅表示显示不足，不能当计算错误。"""
    value=cell.Value2
    try:text=str(cell.Text or '')
    except Exception:text=None
    formula=bool(cell.HasFormula);is_error=None;method='unavailable'
    try:
        address=cell.Address;address=address(False,False) if callable(address) else str(address)
        detected=cell.Parent.Evaluate('ISERROR('+address+')')
        if isinstance(detected,bool):is_error=detected;method='worksheet_ISERROR'
    except Exception:pass
    if is_error is None:
        # pywin32可把VT_ERROR暴露为有符号COM错误码；须同时有错误显示证据。
        error_text=bool(re.fullmatch(r'#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|UNKNOWN!|CONNECT!|BUSY!)',text or '',re.I))
        if isinstance(value,int) and not isinstance(value,bool) and (value & 0xFFFF0000)==0x800A0000 and error_text:
            is_error=True;method='COM_error_code_and_display'
    return {'value':value,'text':text,'has_formula':formula,'is_error':is_error,'error_type_source':method}

def assess_linkage_output(before,after,expected):
    """先判真实技术错误，再判未知预期；未知金额不能掩盖公式错误。"""
    if before['is_error'] is True or after['is_error'] is True:return 'failed','实际联动输出含Excel公式错误'
    if not after['has_formula']:return 'failed','联动输出不是实际公式'
    if before['value'] in (None,'') and after['value'] in (None,''):return 'pending','联动输出仍为空白，其他依赖或口径未决，未证实完整联动'
    if expected is None:return 'pending','缺核定数值预期，不能用本机计算结果反推未知规则'
    if same(before['value'],after['value']) or not (after['value'] in (None,'') if expected=='' else same(after['value'],expected)):return 'failed','实际公式未按核定预期联动'
    return 'verified',None

def formula_cache_metadata(worksheet_xml):
    """空字符串缓存须同时有公式、t=str及v节点，区别于尚未计算的空v/无v。"""
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    result = {}
    for cell in ET.fromstring(worksheet_xml).findall('.//' + ns + 'c'):
        if cell.find(ns + 'f') is None: continue
        value = cell.find(ns + 'v')
        empty_string = value is not None and cell.get('t') == 'str' and value.text in (None, '')
        result[cell.get('r')] = {'xml_type': cell.get('t'), 'has_v': value is not None, 'empty_string_cache': empty_string, 'cache_present': value is not None and (empty_string or value.text not in (None, ''))}
    return result

def engine_identity(path):
    import win32api
    path = Path(path)
    if not path.is_file(): return None
    translations = win32api.GetFileVersionInfo(str(path), '\\VarFileInfo\\Translation')
    for language, encoding in translations:
        try: company = win32api.GetFileVersionInfo(str(path), f'\\StringFileInfo\\{language:04x}{encoding:04x}\\CompanyName')
        except Exception: continue
        try: original = win32api.GetFileVersionInfo(str(path), f'\\StringFileInfo\\{language:04x}{encoding:04x}\\OriginalFilename')
        except Exception: original = path.name
        if 'Microsoft' in company and original.upper() == 'EXCEL.EXE' and path.name.upper() == 'EXCEL.EXE': return 'Microsoft Excel'
        if any(text in company.lower() for text in ('kingsoft', '金山')) and path.name.upper() in ('ET.EXE', 'WPS.EXE', 'KSOLAUNCH.EXE', 'WPSOFFICE.EXE'): return 'WPS Office'
    return None

def registered_engine():
    """优先让COM解析ProgID，再核32/64注册视图；缺键不代表引擎不存在。"""
    import pythoncom
    import winreg
    result = {'requested_progid': 'Excel.Application', 'registered_executable': None, 'registration_gaps': []}
    try:
        clsid = str(pythoncom.CLSIDFromProgID(result['requested_progid'])); result['clsid_resolver'] = 'pythoncom'
    except Exception as exc:
        result['registration_gaps'].append('pythoncom未能预先解析ProgID:' + str(exc))
        try:
            import ctypes
            import uuid
            ole32 = ctypes.OleDLL('ole32')
            identifier = (ctypes.c_ubyte * 16)()
            ole32.CLSIDFromProgID.argtypes = [ctypes.c_wchar_p, ctypes.c_void_p]
            ole32.CLSIDFromProgID(result['requested_progid'], ctypes.byref(identifier))
            clsid = '{' + str(uuid.UUID(bytes_le=bytes(identifier))).upper() + '}'
            result['clsid_resolver'] = 'ole32.CLSIDFromProgID'
        except Exception as fallback:
            result['registration_gaps'].append('Windows COM解析也未完成:' + str(fallback)); return result
    result['resolved_clsid'] = clsid
    for view in (winreg.KEY_WOW64_32KEY, winreg.KEY_WOW64_64KEY):
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 'CLSID\\' + clsid + r'\LocalServer32', 0, winreg.KEY_READ | view) as key: command = winreg.QueryValue(key, None)
            text = str(command).strip()
            executable = text.split('"')[1] if text.startswith('"') else text[:text.lower().index('.exe')+4]
            if Path(executable).is_file():
                result.update(registered_executable=str(Path(executable).resolve()), registration_view=view); return result
            result['registration_gaps'].append('注册命令没有可核实执行档:' + str(view))
        except (OSError, ValueError, IndexError) as exc: result['registration_gaps'].append('注册视图定位缺口:' + str(view) + ':' + str(exc))
    return result

def process_engine(window):
    import ctypes
    from ctypes import wintypes
    import win32process
    _, pid = win32process.GetWindowThreadProcessId(int(window))
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = kernel.OpenProcess(0x1000, False, pid)
    if not handle: raise OSError('无法读取实际Excel进程身份')
    try:
        size = wintypes.DWORD(32768); buffer = ctypes.create_unicode_buffer(size.value)
        if not kernel.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)): raise OSError('无法解析实际Excel进程路径')
        return pid, buffer.value
    finally: kernel.CloseHandle(handle)

def native_child(config):
    result = {'verified': False, 'errors': [], 'pending': [], 'formula_cache_count': 0, 'linkage_checks': []}
    app = book = None; independent = False; original_settings = {}; mutex = None; kernel = None
    try:
        import ctypes
        from ctypes import wintypes
        kernel = ctypes.WinDLL('kernel32', use_last_error=True)
        kernel.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
        kernel.CreateMutexW.restype = wintypes.HANDLE
        kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        kernel.ReleaseMutex.argtypes = [wintypes.HANDLE]; kernel.CloseHandle.argtypes = [wintypes.HANDLE]
        mutex = kernel.CreateMutexW(None, False, 'Local\\CodexPayrollNativeRecalcV1')
        if not mutex or kernel.WaitForSingleObject(mutex, 0) not in (0, 0x80):
            result['pending'].append('原生复核正在被另一任务使用，本次未打开工作簿')
            if mutex: kernel.CloseHandle(mutex)
            mutex = None; return result
        registration = registered_engine(); result.update(registration)
        executable = registration['registered_executable']
        if executable:
            try: result['registered_engine'] = engine_identity(executable)
            except Exception as exc: result['registration_gaps'].append('注册执行档版本资料无法读取:' + str(exc))
        # 注册定位不足时仍可在未开工作簿之前核实际进程，不凭缺键推断没有引擎。
        import win32com.client
        import win32process
        import win32api
        before_processes = set(win32process.EnumProcesses()); result['process_count_before_dispatch'] = len(before_processes)
        app = win32com.client.DispatchEx('Excel.Application')
        pid, actual = process_engine(app.Hwnd); result.update(process_id=pid, actual_executable=actual)
        independent = pid not in before_processes
        initial_books = int(app.Workbooks.Count)
        result.update(independent_new_process=independent, workbooks_before_open=initial_books, executable_sha256=sha256(actual), executable_version=win32api.GetFileVersionInfo(actual, '\\'))
        engine_name = engine_identity(actual); result['actual_engine'] = engine_name
        if not engine_name:
            result['pending'].append('实际进程身份无法证实；未打开工作簿')
            return result
        result['microsoft_excel'] = engine_name == 'Microsoft Excel'
        if initial_books:
            result['pending'].append('该办公实例已有其他工作簿，本次未打开隔离文件'); return result
        for setting in ('Visible', 'DisplayAlerts', 'EnableEvents', 'AskToUpdateLinks', 'AutomationSecurity', 'Calculation'):
            try: original_settings[setting] = getattr(app, setting)
            except Exception: pass
        if independent: app.Visible = False
        app.DisplayAlerts = False; app.EnableEvents = False; app.AskToUpdateLinks = False
        app.AutomationSecurity = 3
        try: app.Calculation = -4135
        except Exception: pass
        # 配置只含事先复制的隔离文件，子进程没有原件路径。
        book = app.Workbooks.Open(str(Path(config['isolated_path']).resolve()), UpdateLinks=0, ReadOnly=False)
        try: app.Calculation = -4135
        except Exception: pass
        formulas = config['formulas']
        if config['format'] == 'xls':
            for sheet in book.Worksheets:
                for cell in sheet.UsedRange:
                    if cell.HasFormula:
                        address = cell.Address(False, False)
                        formulas.append({'sheet': sheet.Name, 'cell': address, 'formula': str(cell.Formula), 'cached': config['xls_values'].get(sheet.Name, {}).get(address),'cached_error':config.get('xls_errors',{}).get(sheet.Name,{}).get(address)})
        app.CalculateFullRebuild()
        if int(app.CalculationState) != 0: result['pending'].append('原生计算尚未结束')
        for item in formulas:
            cell = book.Worksheets(item['sheet']).Range(item['cell'])
            cache_present = item.get('cache_present', item['cached'] is not None)
            observation=observe_cell(cell);actual=observation['value']
            if item.get('xml_type')=='e' or item.get('cached_error'):
                result['errors'].append('实际原文件公式缓存为Excel错误:'+item['sheet']+'!'+item['cell']+':'+str(item.get('cached_error') or item.get('cached')))
            if observation['is_error'] is True:
                result['errors'].append('实际原生重算为Excel公式错误:'+item['sheet']+'!'+item['cell']+':'+str(observation['text']))
            if item.get('xml_type')=='e' or item.get('cached_error') or observation['is_error'] is True:
                result.setdefault('formula_errors',[]).append({'sheet':item['sheet'],'cell':item['cell'],'cached':item.get('cached'),'cache_xml_type':item.get('xml_type'),'native_observation':observation})
            empty_string = item.get('empty_string_cache', item['cached'] == '')
            if not cache_present or not (actual in (None, '') if empty_string else same(item['cached'], actual)):
                result['errors'].append('实际原文件公式缓存缺失或与原生重算不符:' + item['sheet'] + '!' + item['cell'])
            if empty_string and cache_present: result['typed_empty_cache_count'] = result.get('typed_empty_cache_count', 0) + 1
            result['formula_cache_count'] += 1
        for check in config['linkage_checks']:
            sheet = book.Worksheets(check['sheet']); target = sheet.Range(check['input_cell'])
            if target.HasFormula or not same(target.Value2, check['input_expected']) or same(check['test_value'], check['input_expected']):
                result['errors'].append('联动输入不是核定可编辑原值或测试值未变化'); continue
            output_key=lambda item:item.get('sheet',check['sheet'])+'!'+item['cell']
            output_cell=lambda item:book.Worksheets(item.get('sheet',check['sheet'])).Range(item['cell'])
            before = {output_key(item):observe_cell(output_cell(item)) for item in check['outputs']}
            target.Value2 = check['test_value']; app.CalculateFullRebuild()
            output_results = []
            for item in check['outputs']:
                cell = output_cell(item);observation=observe_cell(cell);previous=before[output_key(item)]
                status,message=assess_linkage_output(previous,observation,item.get('expected'))
                if message:result['errors' if status=='failed' else 'pending'].append(message+':'+item.get('sheet',check['sheet'])+'!'+item['cell'])
                output_results.append({'sheet':item.get('sheet',check['sheet']),'cell':item['cell'],'status':status,'before':previous['value'],'after':observation['value'],'before_text':previous['text'],'after_text':observation['text'],'before_observation':previous,'after_observation':observation,'expected':item.get('expected'),'expected_provided':'expected' in item})
            result['linkage_checks'].append({'sheet': check['sheet'], 'input_cell': check['input_cell'], 'outputs': output_results, 'status': 'failed' if any(x['status'] == 'failed' for x in output_results) else ('pending' if any(x['status'] == 'pending' for x in output_results) else 'verified')})
            target.Value2 = check['input_expected']; app.CalculateFullRebuild()
        result['verified'] = not result['errors'] and not result['pending']
    except Exception as exc:
        result['pending'].append('原生引擎或缓存无法证实:' + str(exc))
    finally:
        if book is not None:
            try: book.Close(False)
            except Exception: pass
        if app is not None:
            for setting, value in original_settings.items():
                try: setattr(app, setting, value)
                except Exception: pass
            # WPS的DispatchEx可能复用实例，只关闭本工具打开的副本。
            # 仅实际新进程且已无其他工作簿时退出应用，不影响既有会话。
            try:
                if independent and int(app.Workbooks.Count) == 0: app.Quit()
            except Exception: pass
        if mutex is not None:
            kernel.ReleaseMutex(mutex); kernel.CloseHandle(mutex)
    return result

def verify_native_copy(path, expected_sha256, linkage_checks=None):
    path = Path(path).resolve()
    result = {'verified': False, 'errors': [], 'pending': [], 'source_sha256': expected_sha256}
    if not path.is_file() or sha256(path) != expected_sha256:
        result['errors'].append('原生复核前实际文件SHA不符'); return result
    config = {'format': path.suffix.lower().lstrip('.'), 'formulas': [], 'xls_values': {}, 'linkage_checks': linkage_checks or []}
    if config['format'] == 'xlsx':
        from openpyxl import load_workbook
        formulas = load_workbook(path, read_only=True, data_only=False); cached = load_workbook(path, read_only=True, data_only=True)
        try:
            with ZipFile(path) as package:
                for ws in formulas:
                    ws.reset_dimensions(); values = cached[ws.title]; values.reset_dimensions()
                    value_rows = {i: row for i, row in enumerate(values.values, 1)}
                    metadata = formula_cache_metadata(package.read(ws._worksheet_path.lstrip('/')))
                    for row in ws:
                        for cell in row:
                            if cell.data_type == 'f':
                                values_row = value_rows.get(cell.row, ())
                                value = values_row[cell.column-1] if cell.column <= len(values_row) else None
                                meta = metadata.get(cell.coordinate, {'cache_present': False})
                                if meta.get('empty_string_cache'): value = ''
                                config['formulas'].append({'sheet': ws.title, 'cell': cell.coordinate, 'formula': cell.value, 'cached': value, **meta})
        finally: formulas.close(); cached.close()
    elif config['format'] == 'xls':
        import xlrd
        from openpyxl.utils.cell import get_column_letter
        book = xlrd.open_workbook(path)
        config['xls_values'] = {s.name: {f'{get_column_letter(c+1)}{r+1}': s.cell_value(r, c) for r in range(s.nrows) for c in range(s.ncols)} for s in book.sheets()}
        config['xls_errors'] = {s.name:{f'{get_column_letter(c+1)}{r+1}':xlrd.error_text_from_code.get(int(s.cell_value(r,c)),'Excel错误') for r in range(s.nrows) for c in range(s.ncols) if s.cell_type(r,c)==xlrd.XL_CELL_ERROR} for s in book.sheets()}
    else:
        result['pending'].append('实际文件不是支持原生复核的xlsx/xls'); return result
    # 临时目录限定在系统临时目录中，仅删除本次创建的隔离目录。
    parent = Path(tempfile.gettempdir()).resolve()
    with tempfile.TemporaryDirectory(prefix='payroll-native-audit-', dir=parent) as temporary:
        root = Path(temporary).resolve()
        if root.parent != parent or path.is_relative_to(root): raise ValueError('隔离目录越界')
        isolated = root / ('isolated' + path.suffix); shutil.copy2(path, isolated)
        config['isolated_path'] = str(isolated)
        input_file, output_file = root / 'config.json', root / 'result.json'
        input_file.write_text(json.dumps(config, ensure_ascii=False, default=str), encoding='utf-8')
        try:
            process = subprocess.run([sys.executable, '-B', str(Path(__file__).resolve()), '--private-child', str(input_file), str(output_file)], capture_output=True, text=True, timeout=45)
            if process.returncode or not output_file.is_file(): result['pending'].append('隔离原生复核未完成')
            else: result.update(json.loads(output_file.read_text(encoding='utf-8')))
        except subprocess.TimeoutExpired: result['pending'].append('隔离原生复核超过45秒，不能视为通过')
    result['source_sha256_after'] = sha256(path)
    if result['source_sha256_after'] != expected_sha256:
        result['errors'].append('原生复核后原文件SHA变化'); result['verified'] = False
    return result

if __name__ == '__main__':
    if len(sys.argv) != 4 or sys.argv[1] != '--private-child': raise SystemExit('此程序由crossstage_audit在隔离副本上调用')
    configuration = json.loads(Path(sys.argv[2]).read_text(encoding='utf-8'))
    Path(sys.argv[3]).write_text(json.dumps(native_child(configuration), ensure_ascii=False, default=str, indent=2), encoding='utf-8')
