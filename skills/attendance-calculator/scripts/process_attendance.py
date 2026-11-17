#!/usr/bin/env python3
"""Generate the HR attendance summary from normalized attendance inputs."""

from __future__ import annotations

import argparse
import csv
import math
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple


OUTPUT_COLUMNS_TEMPLATE = [
    "UserID",
    "所属公司",
    "一级部门",
    "三级部门",
    "职级",
    "姓名",
    "迟到分钟数（分钟）",
    "漏打卡次数",
    "迟到扣款情况",
    "漏打卡扣款情况",
    "考勤扣款情况（合计）",
    "本月事假时数（小时）",
    "累计事假时长（小时）",
    "累计加班结余（小时）",
    "病假（小时）",
    "育儿假",
    "奖励金",
    "本月应上班总小时数（h）",
    "入职时间",
    "离职时间",
    "上月病假时长",
    "病假天数合计(年/小时)",
    "{yy}年年假（天） 在职时间（{yy}.1.1-{yy}.12.31）",
    "上月年假剩余天数",
    "本月使用年假（天）",
    "{yy}年年假（天） 在职结余时间（{yy}.1.1-{yy}.12.31）",
    "考勤情况",
    "缺卡情况",
    "早退次数",
    "早退时长（分钟）",
    "异常提示",
]


def normalize_attendance_type(value: Any) -> str:
    """Normalize observed source labels; missing/unknown is never a no-punch exemption."""
    text = str(value or "").strip()
    return {"无需打卡": "不需要打卡", "無需打卡": "不需要打卡"}.get(text, text)


def read_table(path: Optional[str]) -> List[Dict[str, Any]]:
    if not path:
        return []
    table_path = Path(path)
    if not table_path.exists():
        raise FileNotFoundError(f"Input file not found: {table_path}")

    suffix = table_path.suffix.lower()
    if suffix in {".csv", ".txt"}:
        return read_csv(table_path)
    if suffix in {".xlsx", ".xlsm"}:
        return read_xlsx(table_path)
    raise ValueError(f"Unsupported file type: {table_path.suffix}. Use CSV or XLSX.")


def read_csv(path: Path) -> List[Dict[str, Any]]:
    last_error: Optional[Exception] = None
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            with path.open("r", encoding=encoding, newline="") as handle:
                return [normalize_row(row) for row in csv.DictReader(handle)]
        except UnicodeDecodeError as exc:
            last_error = exc
    raise UnicodeDecodeError("csv", b"", 0, 1, f"Could not decode {path}: {last_error}")


def read_xlsx(path: Path) -> List[Dict[str, Any]]:
    try:
        from openpyxl import load_workbook
    except ImportError as exc:
        raise RuntimeError("Reading XLSX requires openpyxl. Convert the file to CSV or install openpyxl.") from exc

    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook.active
    # Some Feishu exports declare A1 while retaining thousands of actual XML rows.
    # Read-only mode otherwise trusts that declaration and silently drops the data.
    sheet.reset_dimensions()
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    if not rows:
        return []
    headers = [str(value).strip() if value is not None else "" for value in rows[0]]
    records: List[Dict[str, Any]] = []
    for values in rows[1:]:
        row = {headers[index]: values[index] if index < len(values) else "" for index in range(len(headers)) if headers[index]}
        if any(str(value).strip() for value in row.values() if value is not None):
            records.append(normalize_row(row))
    return records


def present(row, key):
    value = row.get(key)
    if value is None or str(value).strip() == "": return False
    try: return math.isfinite(float(value))
    except (TypeError, ValueError): return False


def normalize_row(row: Dict[str, Any]) -> Dict[str, Any]:
    return {str(key).strip(): value for key, value in row.items()}


def write_table(path: str, rows: List[Dict[str, Any]], columns: List[str]) -> None:
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    suffix = output_path.suffix.lower()
    if suffix == ".csv":
        with output_path.open("w", encoding="utf-8-sig", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=columns)
            writer.writeheader()
            writer.writerows({column: row.get(column, "") for column in columns} for row in rows)
        return
    if suffix in {".xlsx", ".xlsm"}:
        write_xlsx(output_path, rows, columns)
        return
    raise ValueError("Output must be .csv or .xlsx.")


def write_xlsx(path: Path, rows: List[Dict[str, Any]], columns: List[str]) -> None:
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font
    except ImportError as exc:
        raise RuntimeError("Writing XLSX requires openpyxl. Use a .csv output path or install openpyxl.") from exc

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "综合汇总"
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column, "") for column in columns])

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    header_index = {name: index + 1 for index, name in enumerate(columns)}
    for row_index in range(2, sheet.max_row + 1):
        rank = sheet.cell(row_index, header_index["职级"]).coordinate
        late = sheet.cell(row_index, header_index["迟到分钟数（分钟）"]).coordinate
        missed = sheet.cell(row_index, header_index["漏打卡次数"]).coordinate
        late_deduction = sheet.cell(row_index, header_index["迟到扣款情况"])
        missed_deduction = sheet.cell(row_index, header_index["漏打卡扣款情况"])
        total_deduction = sheet.cell(row_index, header_index["考勤扣款情况（合计）"])
        deduction_rank = rows[row_index - 2].get('_deduction_rank')
        rate = {'组员': -3, '管理层': -5}.get(deduction_rank)
        if rows[row_index - 2].get('迟到扣款情况') == '' or rate is None:
            late_deduction.value = ''
        else:
            late_deduction.value = f'=IF({late}="","",{late}*{rate})'
        missed_deduction.value = '' if rows[row_index - 2].get('漏打卡扣款情况') == '' else f'=IF({missed}="","",MAX(0,{missed}-2)*-50)'
        total_deduction.value = '' if rows[row_index - 2].get('考勤扣款情况（合计）') == '' else f'=IF(OR({late_deduction.coordinate}="",{missed_deduction.coordinate}=""),"",SUM({late_deduction.coordinate}:{missed_deduction.coordinate}))'
        sheet.cell(row_index, header_index["UserID"]).number_format = "@"
    for index, column in enumerate(columns, start=1):
        width = min(max(len(column) + 4, 12), 32)
        sheet.column_dimensions[sheet.cell(row=1, column=index).column_letter].width = width
    workbook.save(path)


def parse_number(value: Any, default: float = 0.0) -> float:
    if value is None:
        return default
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return default
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return default
    try:
        parsed = float(text)
        if not math.isfinite(parsed): raise ValueError("非有限数字")
        return parsed
    except (ValueError, TypeError) as exc:
        raise ValueError(f"标准化数字无法解析:{text}") from exc


def parse_int(value: Any, default: int = 0) -> int:
    return int(round(parse_number(value, float(default))))


def parse_date(value: Any) -> Optional[date]:
    if value is None or value == "":
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, (int, float)):
        try:
            base = datetime(1899, 12, 30)
            return (base + timedelta(days=float(value))).date()
        except Exception:
            return None
    text = str(value).strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def format_date(value: Optional[date]) -> str:
    return value.isoformat() if value else ""


def key_for(row: Dict[str, Any]) -> str:
    employee_id = str(row.get("UserID", "") or row.get("user_id", "") or row.get("employee_id", "") or "").strip()
    if employee_id:
        return f"id:{employee_id}"
    return f"name:{str(row.get('name', '') or '').strip()}"


def make_index(rows: Iterable[Dict[str, Any]]) -> Tuple[Dict[str, Dict[str, Any]], Dict[str, int]]:
    index: Dict[str, Dict[str, Any]] = {}
    counts: Dict[str, int] = {}
    for row in rows:
        key = key_for(row)
        if key in {"id:", "name:"}:
            continue
        counts[key] = counts.get(key, 0) + 1
        if key not in index:
            index[key] = row
    return index, counts


def add_years(day: date, years: int) -> date:
    try:
        return day.replace(year=day.year + years)
    except ValueError:
        return day.replace(month=2, day=28, year=day.year + years)


def ceil_to_half(value: float) -> float:
    if value <= 0:
        return 0.0
    return max(0.5, math.ceil(value * 2) / 2)


def annual_leave_entitlement(hire_date: Optional[date], year: int, month_start: Optional[date] = None) -> float:
    if not hire_date:
        return 0.0
    jan1 = date(year, 1, 1)
    one_year = add_years(hire_date, 1)
    three_years = add_years(hire_date, 3)
    five_years = add_years(hire_date, 5)

    if jan1 >= five_years:
        return 7.0
    if jan1 >= three_years:
        return 5.0
    if jan1 >= one_year:
        return 3.0
    if one_year.year == year and (month_start is None or month_start.month >= one_year.month):
        prorated = (12 - one_year.month) / 12 * 3
        return ceil_to_half(prorated)
    return 0.0


def split_birthdates(value: Any) -> List[date]:
    if value is None:
        return []
    parts = re.split(r"[,，;；\n]+", str(value))
    dates = [parse_date(part.strip()) for part in parts if part.strip()]
    return [item for item in dates if item]


def month_index(day: date) -> int:
    return day.year * 12 + day.month


def month_from_index(index: int) -> Tuple[int, int]:
    return (index - 1) // 12, (index - 1) % 12 + 1


def childcare_note(
    childcare_days: float,
    birthdates: List[date],
    month_start: date,
    prior_child_year_days: float,
    prior_note: str = "",
) -> Tuple[str, float]:
    current_index = month_index(month_start)
    active = [day for day in birthdates if 0 <= current_index - month_index(day) <= 36]
    prior_match = re.search(r"已使用\s*([0-9.]+)\s*天", prior_note)
    prior_used = parse_number(prior_match.group(1)) if prior_match else prior_child_year_days

    if not active:
        total = prior_used + childcare_days
        if prior_note and childcare_days:
            return re.sub(r"已使用\s*[0-9.]+\s*天", f"已使用{clean_number(total)}天", prior_note, count=1), total
        return prior_note, total

    newest = max(active)
    birth_index = month_index(newest)
    months_since_birth = current_index - birth_index
    if months_since_birth <= 12:
        start_index, end_index = birth_index, birth_index + 12
    elif months_since_birth <= 24:
        start_index, end_index = birth_index + 13, birth_index + 24
    else:
        start_index, end_index = birth_index + 25, birth_index + 36

    prior_end_match = re.search(r"截至\s*(\d{4})年\s*(\d{1,2})月", prior_note)
    prior_end_index = int(prior_end_match.group(1)) * 12 + int(prior_end_match.group(2)) if prior_end_match else None
    entered_new_cycle = current_index == start_index or (prior_end_index is not None and prior_end_index != end_index)
    total = childcare_days if entered_new_cycle else prior_used + childcare_days
    end_year, end_month = month_from_index(end_index)
    note = f"截至{end_year}年{end_month}月（含），有10天育儿假，已使用{clean_number(total)}天"
    if not prior_note and childcare_days <= 0 and months_since_birth > 1:
        return "", total
    return note, total


def clean_number(value: float) -> Any:
    if value == "" or value is None: return ""
    if abs(value - round(value)) < 0.00001:
        return int(round(value))
    return float(f"{value:.2f}")


def detail_text(prefix: str, value: Any) -> str:
    text = str(value or "").strip()
    return f"{prefix}：{text}" if text else ""


def join_attendance_detail(late_detail: Any, personal_leave_detail: Any, sick_detail: Any = "", annual_detail: Any = "") -> str:
    parts = [detail_text("迟到", late_detail), detail_text("事假", personal_leave_detail), detail_text("病假", sick_detail), detail_text("年假", annual_detail)]
    return "\n".join(part for part in parts if part)


def output_columns(year: int) -> List[str]:
    yy = f"{year % 100:02d}"
    return [column.format(yy=yy) for column in OUTPUT_COLUMNS_TEMPLATE]


def build_summary(args: argparse.Namespace) -> Tuple[List[Dict[str, Any]], List[str]]:
    month = datetime.strptime(args.month, "%Y-%m").date()
    year = month.year
    columns = output_columns(year)
    annual_entitlement_column = next(column for column in columns if "年年假（天） 在职时间" in column)
    annual_remaining_column = next(column for column in columns if "年年假（天） 在职结余时间" in column)

    attendance_rows = read_table(args.attendance)
    if not attendance_rows:
        raise ValueError("本月标准化考勤无员工行，不能将空表认定核算完成")
    if args.paid_days <= 0 or args.hours_per_day <= 0:
        raise ValueError("计薪日和每日工时须大于0")
    normalized_keys = [key_for(r) for r in attendance_rows]
    if len(normalized_keys) != len(set(normalized_keys)):
        raise ValueError("标准化考勤一人多行，须先核对去重或按日汇总")
    for record in attendance_rows:
        if not record.get("name"):
            raise ValueError("标准化考勤缺姓名")
        for field in ("late_minutes", "missed_punch_count", "personal_leave_hours", "sick_leave_hours", "annual_leave_days", "childcare_leave_days"):
            if present(record, field) and parse_number(record[field]) < 0:
                raise ValueError(f"{field}不能为负值，负结余应放结余字段")
        if present(record, "required_work_hours") and parse_number(record["required_work_hours"]) <= 0:
            raise ValueError("本月统一计薪小时数须大于0")
    roster_index, roster_counts = make_index(read_table(args.roster))
    history_index, history_counts = make_index(read_table(args.history))
    overtime_index, overtime_counts = make_index(read_table(args.overtime))

    result_rows: List[Dict[str, Any]] = []
    for attendance in attendance_rows:
        key = key_for(attendance)
        roster = roster_index.get(key, {})
        history = history_index.get(key, {})
        overtime = overtime_index.get(key, {})
        exceptions: List[str] = []

        if roster_counts.get(key, 0) > 1 or history_counts.get(key, 0) > 1 or overtime_counts.get(key, 0) > 1:
            exceptions.append("员工标识重复，需复核")
        if not roster:
            exceptions.append("缺少花名册信息")
        if roster_counts.get(key, 0) > 1: roster = {}
        if history_counts.get(key, 0) > 1: history = {}
        if overtime_counts.get(key, 0) > 1: overtime = {}

        attendance_type = normalize_attendance_type(attendance.get("attendance_type"))
        if attendance_type not in ("需要打卡", "不需要打卡"):
            exceptions.append("考勤类型不明，结果待核")
        for field in ("late_minutes", "missed_punch_count", "personal_leave_hours", "sick_leave_hours", "annual_leave_days", "childcare_leave_days"):
            if not present(attendance, field): exceptions.append(f"缺少标准化输入字段{field}，结果待核")
        for source, fields, label in [(history, ("prior_sick_leave_hours_year", "prior_annual_leave_remaining_days", "prior_annual_leave_entitlement_days"), "上月确认资料"), (overtime, ("overtime_balance_hours", "prior_leave_balance_hours"), "加班结余资料")]:
            for field in fields:
                if field == "prior_leave_balance_hours" and attendance_type == "不需要打卡":
                    continue  # 明确不适用；原始资料缺失但类型未知者仍不可豁免。
                if not present(source, field): exceptions.append(f"缺少{label}:{field}，不得默认零或完整额度")
        reward_status = attendance.get("reward_status", "")
        reward_amount = attendance.get("reward_amount", "")
        if reward_status == "confirmed_zero": reward_amount = 0
        elif reward_status == "not_applicable" and attendance.get("reward_reason"): reward_amount = ""
        elif reward_status != "calculated" or not present(attendance, "reward_amount"):
            reward_amount = ""
            exceptions.append("奖励金未完成核验或缺核准金额，待核")
        if len(split_birthdates(roster.get("child_birthdates"))) > 1:
            exceptions.append("多子女育儿假须按子女逐笔核对；本程序不合并推定额度")
        expected_prior = (month.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
        if history.get("history_month") and history["history_month"] != expected_prior:
            exceptions.append("上月资料月份不连续，累计余额待核")
        needs_punch = attendance_type == "需要打卡"

        job_level = str(roster.get("job_level", "") or "").strip()
        rank = str(roster.get("rank", "") or job_level).strip()
        late_minutes = parse_int(attendance.get("late_minutes"))
        missed_punch_count = parse_int(attendance.get("missed_punch_count"))
        late_deduction: Any = ""
        if needs_punch:
            if job_level == "组员":
                late_deduction = late_minutes * -3
            elif job_level == "管理层":
                late_deduction = late_minutes * -5
            else:
                exceptions.append("缺少或无法识别职级，迟到扣款未计算")
        if needs_punch and not present(attendance, "late_minutes"): late_deduction = ""
        missed_deduction = max(0, missed_punch_count - 2) * -50 if needs_punch else 0
        if needs_punch and not present(attendance, "missed_punch_count"): missed_deduction = ""
        total_deduction = late_deduction + missed_deduction if isinstance(late_deduction, (int, float)) and isinstance(missed_deduction, (int, float)) else ""

        overtime_balance = parse_number(overtime.get("overtime_balance_hours")) if present(overtime, "overtime_balance_hours") else ""
        monthly_personal_leave = parse_number(attendance.get("personal_leave_hours")) if needs_punch else 0.0
        prior_leave_balance = parse_number(overtime.get("prior_leave_balance_hours")) if present(overtime, "prior_leave_balance_hours") else ""
        cumulative_personal_leave = prior_leave_balance + monthly_personal_leave if prior_leave_balance != "" else ""

        monthly_sick_leave = parse_number(attendance.get("sick_leave_hours")) if needs_punch else 0.0
        prior_sick_leave = parse_number(history.get("prior_sick_leave_hours_year")) if present(history, "prior_sick_leave_hours_year") else ""
        ytd_sick_leave = prior_sick_leave + monthly_sick_leave if prior_sick_leave != "" else ""

        hire_date = parse_date(roster.get("hire_date"))
        annual_entitlement = annual_leave_entitlement(hire_date, year, month)
        annual_used = parse_number(attendance.get("annual_leave_days"))
        prior_annual_remaining = parse_number(history.get("prior_annual_leave_remaining_days")) if present(history, "prior_annual_leave_remaining_days") else ""
        prior_annual_entitlement = parse_number(history.get("prior_annual_leave_entitlement_days")) if present(history, "prior_annual_leave_entitlement_days") else ""
        annual_remaining = max(0.0, prior_annual_remaining + max(0.0, annual_entitlement - prior_annual_entitlement) - annual_used) if prior_annual_remaining != "" and prior_annual_entitlement != "" else ""

        childcare_days = parse_number(attendance.get("childcare_leave_days"))
        birthdates = split_birthdates(roster.get("child_birthdates"))
        prior_childcare = parse_number(history.get("childcare_leave_days_in_child_year"))
        prior_childcare_note = str(history.get("childcare_leave_note", "") or history.get("育儿假", "") or "")
        childcare, childcare_total = childcare_note(childcare_days, birthdates, month, prior_childcare, prior_childcare_note)
        if childcare_days > 1:
            exceptions.append("育儿假本月超过1天，需复核")
        if childcare_total > 10:
            exceptions.append(f"育儿假额度超限：本周期10天，实际累计{clean_number(childcare_total)}天；保留原记录，不自动转换假别")

        department = str(attendance.get("department", "") or roster.get("department", "") or "").strip()
        missed_detail = detail_text("缺卡", attendance.get("missed_punch_detail"))
        attendance_detail = join_attendance_detail(attendance.get("late_detail"), attendance.get("personal_leave_detail"), attendance.get("sick_leave_detail"), attendance.get("annual_leave_detail"))
        user_id = str(attendance.get("UserID", "") or attendance.get("user_id", "") or roster.get("UserID", "") or roster.get("user_id", "") or history.get("UserID", "") or "").strip()
        if not re.fullmatch(r"[0-9]{19}", user_id):
            exceptions.append("UserID缺失或不是完整19位文本标识，禁止猜填")
        if present(attendance, "required_work_hours") and parse_number(attendance["required_work_hours"]) != args.paid_days * args.hours_per_day:
            exceptions.append("个人输入工时与统一计薪日历不一致，待核")
        row = {
            "_deduction_rank": job_level,
            "UserID": user_id,
            "所属公司": str(roster.get("company", "") or roster.get("所属公司", "") or "").strip(),
            "一级部门": str(roster.get("level1_department", "") or roster.get("1级部门", "") or "").strip(),
            "三级部门": str(roster.get("level3_department", "") or roster.get("3级部门", "") or "").strip(),
            "部门": department,
            "职级": rank,
            "姓名": str(attendance.get("name", "") or roster.get("name", "") or "").strip(),
            "迟到分钟数（分钟）": late_minutes if needs_punch else "",
            "漏打卡次数": missed_punch_count if needs_punch else "",
            "迟到扣款情况": late_deduction,
            "漏打卡扣款情况": missed_deduction if needs_punch else "",
            "考勤扣款情况（合计）": total_deduction if needs_punch else "",
            "本月事假时数（小时）": clean_number(monthly_personal_leave) if needs_punch else "",
            "累计事假时长（小时）": clean_number(cumulative_personal_leave) if needs_punch else "",
            "累计加班结余（小时）": clean_number(overtime_balance),
            "病假（小时）": clean_number(monthly_sick_leave) if needs_punch else "",
            "育儿假": childcare,
            "奖励金": reward_amount,
            "本月应上班总小时数（h）": clean_number(parse_number(attendance.get("required_work_hours"), args.paid_days * args.hours_per_day)),
            "入职时间": format_date(hire_date),
            "离职时间": format_date(parse_date(roster.get("resignation_date"))),
            "上月病假时长": clean_number(prior_sick_leave),
            "病假天数合计(年/小时)": clean_number(ytd_sick_leave),
            annual_entitlement_column: clean_number(annual_entitlement),
            "上月年假剩余天数": clean_number(prior_annual_remaining),
            "本月使用年假（天）": clean_number(annual_used),
            annual_remaining_column: clean_number(annual_remaining),
            "考勤情况": attendance_detail,
            "缺卡情况": missed_detail,
            "早退次数": parse_int(attendance.get("early_leave_count")),
            "早退时长（分钟）": clean_number(parse_number(attendance.get("early_leave_minutes"))),
            "异常提示": "；".join(exceptions),
        }
        result_rows.append(row)
    return result_rows, columns


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate HR attendance summary from normalized inputs.")
    parser.add_argument("--month", required=True, help="Calculation month in YYYY-MM format.")
    parser.add_argument("--paid-days", required=True, type=float, help="Actual paid workdays in the month.")
    parser.add_argument("--hours-per-day", default=8.0, type=float, help="Work hours per day. Default: 8.")
    parser.add_argument("--attendance", required=True, help="Normalized monthly attendance CSV/XLSX.")
    parser.add_argument("--roster", required=True, help="Normalized roster CSV/XLSX.")
    parser.add_argument("--history", help="Normalized history CSV/XLSX.")
    parser.add_argument("--overtime", help="Normalized overtime CSV/XLSX.")
    parser.add_argument("--output", required=True, help="Output CSV/XLSX path.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    sources = [getattr(args, k, None) for k in ("attendance", "roster", "history", "overtime")]
    output_path = Path(args.output).resolve()
    if any(p and Path(p).resolve() == output_path for p in sources):
        raise SystemExit("输出不得覆盖任何输入源文件")
    if output_path.exists():
        raise SystemExit("输出文件已存在，请另存新版本，不覆盖确认结果")
    rows, columns = build_summary(args)
    write_table(args.output, rows, columns)
    import json
    audit = {"month": args.month, "rows": len(rows), "kind": "intermediate_draft", "pending": [{"name": r["姓名"], "detail": r["异常提示"]} for r in rows if r["异常提示"]]}
    Path(str(args.output) + ".audit.json").write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {len(rows)} intermediate draft rows to {args.output}; pending: {len(audit['pending'])}")
    if audit["pending"]: raise SystemExit(2)


if __name__ == "__main__":
    from diagnostics_bridge import observed
    try: observed(main,'process_attendance')
    except (ValueError, FileNotFoundError) as exc:
        import sys
        print(str(exc), file=sys.stderr)
        raise SystemExit(2)
