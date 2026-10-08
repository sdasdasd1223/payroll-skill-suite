#!/usr/bin/env python3
"""Build a payroll-ready five-insurance-one-fund workbook from monthly exports."""

from __future__ import annotations

import argparse
import csv
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import re
import sys
import zipfile
import xml.etree.ElementTree as ET


INSURANCE_DIRS = {
    "pension": "养老",
    "unemployment": "失业",
    "medical": "医保",
    "housing_fund": "公积金",
}

PENSION_EMPLOYER_RATE = Decimal("0.16")
UNEMPLOYMENT_EMPLOYER_RATE = Decimal("0.005")
MEDICAL_EMPLOYER_RATE = Decimal("0.08")
MATERNITY_RATE = Decimal("0.007")


@dataclass
class EmployeeRecord:
    company: str
    name: str
    certificate_number: str = ""
    pension_employee: Decimal = Decimal("0")
    pension_employer: Decimal = Decimal("0")
    unemployment_employee: Decimal = Decimal("0")
    unemployment_employer: Decimal = Decimal("0")
    medical_employee: Decimal = Decimal("0")
    medical_employer_basic: Decimal = Decimal("0")
    maternity: Decimal = Decimal("0")
    housing_employee: Decimal = Decimal("0")
    housing_employer: Decimal = Decimal("0")
    present: set[str] = field(default_factory=set)
    notes: list[str] = field(default_factory=list)


def money(value: Decimal) -> Decimal:
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def parse_decimal(value: object) -> Decimal:
    text = str(value or "").strip().replace(",", "").replace("，", "")
    if text in {"", "-", "--"}:
        return Decimal("0")
    if text.endswith("%"):
        text = text[:-1]
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal("0")


def display_decimal(value: Decimal) -> object:
    value = money(value)
    if value == value.to_integral():
        return int(value)
    return float(value)


def col_to_num(col: str) -> int:
    number = 0
    for char in col:
        number = number * 26 + ord(char) - 64
    return number


def read_xlsx(path: Path) -> list[list[str]]:
    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    }
    with zipfile.ZipFile(path) as zf:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("main:si", ns):
                shared.append(
                    "".join(
                        node.text or ""
                        for node in si.iter(
                            "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
                        )
                    )
                )

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        first_sheet = workbook.find("main:sheets/main:sheet", ns)
        if first_sheet is None:
            return []
        rid = first_sheet.attrib.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        target = relmap[rid]
        if not target.startswith("xl/"):
            target = "xl/" + target

        sheet = ET.fromstring(zf.read(target))
        rows: list[list[str]] = []
        for row in sheet.findall(".//main:sheetData/main:row", ns):
            values: list[str] = []
            for cell in row.findall("main:c", ns):
                ref = cell.attrib.get("r", "")
                match = re.match(r"([A-Z]+)", ref)
                if not match:
                    continue
                index = col_to_num(match.group(1)) - 1
                while len(values) <= index:
                    values.append("")
                values[index] = read_xlsx_cell(cell, shared, ns)
            rows.append(values)
        return rows


def read_xlsx_cell(cell: ET.Element, shared: list[str], ns: dict[str, str]) -> str:
    cell_type = cell.attrib.get("t")
    value_node = cell.find("main:v", ns)
    inline_node = cell.find("main:is", ns)
    if cell_type == "s" and value_node is not None and value_node.text is not None:
        return shared[int(value_node.text)]
    if cell_type == "inlineStr" and inline_node is not None:
        return "".join(
            node.text or ""
            for node in inline_node.iter(
                "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"
            )
        )
    if value_node is not None and value_node.text is not None:
        return value_node.text
    return ""


def read_xls_with_excel(path: Path) -> list[list[str]]:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError("读取 .xls 需要 Windows Excel 或请先另存为 .xlsx") from exc

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Open(str(path.resolve()), ReadOnly=True)
        try:
            sheet = workbook.Worksheets(1)
            used = sheet.UsedRange
            rows: list[list[str]] = []
            for r in range(1, used.Rows.Count + 1):
                row: list[str] = []
                for c in range(1, used.Columns.Count + 1):
                    row.append(str(sheet.Cells(r, c).Text or "").strip())
                rows.append(row)
            return rows
        finally:
            workbook.Close(False)
    finally:
        excel.Quit()


def read_table(path: Path) -> list[list[str]]:
    suffix = path.suffix.lower()
    if suffix == ".xlsx":
        return read_xlsx(path)
    if suffix == ".xls":
        return read_xls_with_excel(path)
    raise ValueError(f"Unsupported file type: {path}")


def find_header(rows: list[list[str]], required: list[str]) -> tuple[int, dict[str, int]]:
    for row_index, row in enumerate(rows):
        mapping = {str(value).strip(): idx for idx, value in enumerate(row) if str(value).strip()}
        if all(name in mapping for name in required):
            return row_index, mapping
    raise ValueError(f"Cannot find header with required columns: {', '.join(required)}")


def cell(row: list[str], header: dict[str, int], name: str) -> str:
    index = header.get(name)
    if index is None or index >= len(row):
        return ""
    return str(row[index]).strip()


def company_from_path(path: Path, insurance_type: str) -> str:
    stem = path.stem
    suffix = INSURANCE_DIRS[insurance_type]
    if stem.endswith("-" + suffix):
        return stem[: -(len(suffix) + 1)]
    return stem.split("-")[0]


def discover_files(input_root: Path) -> dict[str, dict[str, Path]]:
    result: dict[str, dict[str, Path]] = {}
    for insurance_type, folder_name in INSURANCE_DIRS.items():
        folder = input_root / folder_name
        if not folder.exists():
            continue
        for path in sorted(folder.glob("*")):
            if path.suffix.lower() not in {".xlsx", ".xls"}:
                continue
            company = company_from_path(path, insurance_type)
            result.setdefault(company, {})[insurance_type] = path
    return result


def add_record(records: dict[tuple[str, str], EmployeeRecord], company: str, name: str) -> EmployeeRecord:
    key = (company, name)
    if key not in records:
        records[key] = EmployeeRecord(company=company, name=name)
    return records[key]


def parse_social_file(
    records: dict[tuple[str, str], EmployeeRecord],
    detail_rows: dict[str, list[list[object]]],
    exceptions: list[list[object]],
    insurance_type: str,
    company: str,
    path: Path,
) -> None:
    rows = read_table(path)
    header_index, header = find_header(rows, ["姓名", "证件号码", "缴费基数", "应缴费额"])
    seen: set[str] = set()

    for source_row in rows[header_index + 1 :]:
        name = cell(source_row, header, "姓名")
        if not name or name == "合计":
            continue
        base = parse_decimal(cell(source_row, header, "缴费基数"))
        employee_amount = money(parse_decimal(cell(source_row, header, "应缴费额")))
        certificate = cell(source_row, header, "证件号码")
        fee_rate = cell(source_row, header, "费率")
        wage = cell(source_row, header, "缴费工资")
        cert_type = cell(source_row, header, "证件类型")
        if name in seen:
            exceptions.append([company, name, insurance_type, "duplicate_name_in_source", str(path), "同一险种源表中姓名重复"])
        seen.add(name)
        if base == 0:
            exceptions.append([company, name, insurance_type, "missing_base", str(path), "缴费基数为空或无法解析"])
        if employee_amount == 0:
            exceptions.append([company, name, insurance_type, "missing_employee_amount", str(path), "个人应缴费额为空或为 0"])

        record = add_record(records, company, name)
        if certificate and not record.certificate_number:
            record.certificate_number = certificate
        record.present.add(insurance_type)

        if insurance_type == "pension":
            employer_amount = money(base * PENSION_EMPLOYER_RATE)
            record.pension_employee = employee_amount
            record.pension_employer = employer_amount
            detail_rows["养老"].append([company, name, cert_type, certificate, wage, base, fee_rate, employee_amount, employer_amount])
        elif insurance_type == "unemployment":
            employer_amount = money(base * UNEMPLOYMENT_EMPLOYER_RATE)
            record.unemployment_employee = employee_amount
            record.unemployment_employer = employer_amount
            detail_rows["失业"].append([company, name, cert_type, certificate, wage, base, fee_rate, employee_amount, employer_amount])
        elif insurance_type == "medical":
            medical_basic = money(base * MEDICAL_EMPLOYER_RATE)
            maternity = money(base * MATERNITY_RATE)
            unit_total = money(medical_basic + maternity)
            record.medical_employee = employee_amount
            record.medical_employer_basic = medical_basic
            record.maternity = maternity
            detail_rows["医保"].append([company, name, cert_type, certificate, wage, base, fee_rate, employee_amount, medical_basic, maternity, unit_total])


def parse_housing_file(
    records: dict[tuple[str, str], EmployeeRecord],
    detail_rows: dict[str, list[list[object]]],
    exceptions: list[list[object]],
    company: str,
    path: Path,
) -> None:
    rows = read_table(path)
    header_index, header = find_header(rows, ["姓名", "单位缴存额", "个人月缴存额"])
    seen: set[str] = set()
    for source_row in rows[header_index + 1 :]:
        name = cell(source_row, header, "姓名")
        if not name or name == "合计":
            continue
        employer = money(parse_decimal(cell(source_row, header, "单位缴存额")))
        employee = money(parse_decimal(cell(source_row, header, "个人月缴存额")))
        certificate = cell(source_row, header, "证件号码")
        if name in seen:
            exceptions.append([company, name, "housing_fund", "duplicate_name_in_source", str(path), "同一公积金源表中姓名重复"])
        seen.add(name)
        if "***" in certificate:
            exceptions.append([company, name, "housing_fund", "masked_certificate", str(path), "公积金证件号码被打码，只能按姓名匹配"])
        if employer == 0 or employee == 0:
            exceptions.append([company, name, "housing_fund", "missing_housing_amount", str(path), "公积金单位或个人金额为空/为 0"])

        record = add_record(records, company, name)
        record.present.add("housing_fund")
        record.housing_employer = employer
        record.housing_employee = employee
        detail_rows["公积金"].append([
            company,
            cell(source_row, header, "个人账号"),
            name,
            cell(source_row, header, "证件类型"),
            certificate,
            cell(source_row, header, "当前工资"),
            employer,
            employee,
            cell(source_row, header, "补贴工资"),
            cell(source_row, header, "补贴月缴"),
            cell(source_row, header, "合计月缴额"),
            cell(source_row, header, "个人状态"),
        ])


def build_records(company_files: dict[str, dict[str, Path]]) -> tuple[dict[tuple[str, str], EmployeeRecord], dict[str, list[list[object]]], list[list[object]]]:
    records: dict[tuple[str, str], EmployeeRecord] = {}
    detail_rows: dict[str, list[list[object]]] = {"养老": [], "失业": [], "医保": [], "公积金": []}
    exceptions: list[list[object]] = []
    required = ["pension", "unemployment", "medical", "housing_fund"]

    for company, files in sorted(company_files.items()):
        for insurance_type in required:
            path = files.get(insurance_type)
            if not path:
                exceptions.append([company, "", insurance_type, "missing_source_file", "", "缺少该险种源文件"])
                continue
            if insurance_type == "housing_fund":
                parse_housing_file(records, detail_rows, exceptions, company, path)
            else:
                parse_social_file(records, detail_rows, exceptions, insurance_type, company, path)

    for record in records.values():
        for insurance_type in required:
            if insurance_type not in record.present:
                exceptions.append([record.company, record.name, insurance_type, "employee_missing_insurance_type", "", "员工未出现在该险种源表中"])
                record.notes.append(f"缺少{INSURANCE_DIRS[insurance_type]}")

    return records, detail_rows, exceptions


def summary_rows(records: dict[tuple[str, str], EmployeeRecord]) -> list[list[object]]:
    rows: list[list[object]] = []
    for index, record in enumerate(sorted(records.values(), key=lambda r: (r.company, r.name)), start=1):
        social_employer = money(record.pension_employer + record.unemployment_employer)
        social_employee = money(record.pension_employee + record.unemployment_employee)
        medical_employer = money(record.medical_employer_basic + record.maternity)
        rows.append([
            index,
            record.company,
            record.name,
            display_decimal(record.pension_employer),
            display_decimal(record.pension_employee),
            display_decimal(record.unemployment_employer),
            display_decimal(record.unemployment_employee),
            display_decimal(social_employer),
            display_decimal(social_employee),
            display_decimal(medical_employer),
            display_decimal(record.medical_employee),
            display_decimal(record.housing_employer),
            display_decimal(record.housing_employee),
            "；".join(record.notes),
        ])
    return rows


def write_csv_fallback(output: Path, sheets: dict[str, list[list[object]]]) -> None:
    output.mkdir(parents=True, exist_ok=True)
    for sheet_name, rows in sheets.items():
        with (output / f"{sheet_name}.csv").open("w", newline="", encoding="utf-8-sig") as handle:
            writer = csv.writer(handle)
            writer.writerows(rows)


def write_xlsx_with_excel(output: Path, sheets: dict[str, list[list[object]]]) -> None:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError("写入 .xlsx 需要 Windows Excel；可改用 --csv-output-dir 输出 CSV") from exc

    output.parent.mkdir(parents=True, exist_ok=True)
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Add()
        while workbook.Worksheets.Count < len(sheets):
            workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        for index, (sheet_name, rows) in enumerate(sheets.items(), start=1):
            sheet = workbook.Worksheets(index)
            sheet.Name = sheet_name[:31]
            for row_index, row in enumerate(rows, start=1):
                for col_index, value in enumerate(row, start=1):
                    sheet.Cells(row_index, col_index).Value = value
            sheet.Columns.AutoFit()
        workbook.SaveAs(str(output.resolve()), FileFormat=51)
        workbook.Close(False)
    finally:
        excel.Quit()


def make_sheets(
    records: dict[tuple[str, str], EmployeeRecord],
    detail_rows: dict[str, list[list[object]]],
    exceptions: list[list[object]],
) -> dict[str, list[list[object]]]:
    sheets: dict[str, list[list[object]]] = {}
    sheets["五险一金"] = [
        ["序号", "所属公司（用于区分未本、金碳成本核算）", "姓名", "养老单位缴纳", "养老个人缴纳", "就业险单位缴纳", "就业险个人缴纳", "社保单位缴纳", "社保个人缴纳", "医保单位缴纳", "医保个人缴纳", "公积金单位缴纳", "公积金个人缴纳", "备注"],
        *summary_rows(records),
    ]
    sheets["养老"] = [["主体", "姓名", "证件类型", "证件号码", "缴费工资", "缴费基数", "费率", "个人缴费金额", "单位缴费金额"], *detail_rows["养老"]]
    sheets["失业"] = [["主体", "姓名", "证件类型", "证件号码", "缴费工资", "缴费基数", "费率", "个人缴费金额", "单位缴费金额"], *detail_rows["失业"]]
    sheets["医保"] = [["主体", "姓名", "证件类型", "证件号码", "缴费工资", "缴费基数", "费率", "基本医疗保险（个人）", "基本医疗保险（单位）", "生育保险", "单位合计"], *detail_rows["医保"]]
    sheets["公积金"] = [["主体", "个人账号", "姓名", "证件类型", "证件号码", "当前工资", "单位缴存额", "个人月缴存额", "补贴工资", "补贴月缴", "合计月缴额", "个人状态"], *detail_rows["公积金"]]
    sheets["异常"] = [["主体", "姓名", "险种", "异常类型", "来源文件", "说明"], *exceptions]
    return sheets


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--month", required=True, help="Payroll month, for example 2026-06.")
    parser.add_argument("--input-root", type=Path, help="Folder containing 养老/失业/医保/公积金 subfolders.")
    parser.add_argument("--company", help="Company name for single-company mode.")
    parser.add_argument("--pension", type=Path)
    parser.add_argument("--unemployment", type=Path)
    parser.add_argument("--medical", type=Path)
    parser.add_argument("--housing-fund", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--csv-output-dir", type=Path, help="Fallback CSV output folder when Excel automation is unavailable.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.input_root:
        company_files = discover_files(args.input_root)
    else:
        if not args.company:
            raise SystemExit("--company is required in single-company mode")
        company_files = {
            args.company: {
                "pension": args.pension,
                "unemployment": args.unemployment,
                "medical": args.medical,
                "housing_fund": args.housing_fund,
            }
        }
        company_files[args.company] = {k: v for k, v in company_files[args.company].items() if v}

    if not company_files:
        raise SystemExit("No source files discovered.")

    records, detail_rows, exceptions = build_records(company_files)
    sheets = make_sheets(records, detail_rows, exceptions)
    try:
        write_xlsx_with_excel(args.output, sheets)
    except Exception:
        if not args.csv_output_dir:
            raise
        write_csv_fallback(args.csv_output_dir, sheets)
    print(f"Generated {args.output} for {args.month}. Employees: {len(records)}. Exceptions: {len(exceptions)}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
