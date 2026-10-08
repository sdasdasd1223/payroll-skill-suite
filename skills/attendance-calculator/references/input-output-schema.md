# 输入和输出字段说明

运行脚本前，先把源文件整理成以下标准字段。字段名区分大小写。

## Required Inputs

### Monthly Attendance

One row per employee for the calculation month.

Required columns:

- `employee_id`: stable employee identifier. May be blank if `name` is unique.
- `department`
- `name`
- `attendance_type`: `需要打卡` or `不需要打卡`.
- `late_minutes`: monthly late minutes.
- `missed_punch_count`: monthly missed-punch count.
- `personal_leave_hours`: current-month personal leave hours from attendance.
- `sick_leave_hours`: current-month sick leave hours from attendance.
- `annual_leave_days`: current-month annual leave used.
- `childcare_leave_days`: current-month childcare leave used.
- `required_work_hours`: current-month required work hours for the employee. Prefer summing daily `应出勤天数` and multiplying by `8`; only when that field is missing, sum `应出勤时长(分钟)` and divide by `60`.

Optional columns:

- `late_detail`: preformatted late-arrival detail text without the `迟到：` prefix.
- `personal_leave_detail`: preformatted personal-leave detail text without the `事假：` prefix.
- `missed_punch_detail`: preformatted missed-punch detail text without the `缺卡：` prefix.
- `early_leave_count`: current-month early-departure count.
- `early_leave_minutes`: current-month early-departure minutes.
- `early_leave_detail`: preformatted early-departure detail text without the `早退：` prefix.

### Roster

One row per employee.

Required columns:

- `employee_id`
- `name`
- `job_level`: `组员` or `管理层`.
- `hire_date`: date parseable by Excel or ISO format.

Optional columns:

- `department`
- `rank`: output value for `职级`. If blank, `job_level` is used.
- `resignation_date`: used to filter records after the resignation date. The resignation date itself remains included by default.
- `child_birthdates`: one or more child birth dates separated by comma, Chinese comma, semicolon, Chinese semicolon, or newline. Do not split a single date such as `2024/09/11` on its internal slash.

### History

One row per employee. Use the previous month summary or a separately maintained history table.

Required columns:

- `employee_id`
- `name`

Optional columns:

- `prior_personal_leave_hours`: historical cumulative personal leave hours; retained for comparison but not the preferred source for the current cumulative personal-leave field.
- `prior_sick_leave_hours_year`: sick leave hours before the current month in the same year.
- `prior_annual_leave_remaining_days`: annual leave balance before the current month.
- `childcare_leave_days_in_child_year`: childcare leave already used in the active child-birthday year.

### Overtime

One row per employee.

Required columns:

- `employee_id`
- `name`
- `overtime_balance_hours`: cumulative overtime balance.
- `prior_leave_balance_hours`: prior-month leave balance used for the current cumulative personal-leave field.

Do not automatically offset current-month overtime against cumulative personal leave. If a negative overtime conversion is present in a source workflow, keep it visible in the exception or calculation notes.

## Final Output Columns

Generate columns in this order:

- `部门`
- `职级`
- `姓名`
- `迟到分钟数（分钟）`
- `漏打卡次数`
- `迟到扣款情况`
- `漏打卡扣款情况`
- `考勤扣款情况（合计）`
- `本月事假时数（小时）`
- `累计事假时长（小时）`
- `累计加班结余（小时）`
- `病假（小时）`
- `育儿假`
- `奖励金`
- `本月应上班总小时数（h）`
- `入职时间`
- `离职时间`
- `上月病假时长`
- `病假天数合计(年/小时)`
- `26年年假（天） 在职时间（26.1.1-26.12.31）`
- `上月年假剩余天数`
- `本月使用年假（天）`
- `26年年假（天） 在职结余时间（26.1.1-26.12.31）`
- `考勤情况`
- `缺卡情况`
- `早退次数`
- `早退时长（分钟）`
- `异常提示`

## Exception Report Columns

Generate a separate exception report whenever any employee has data that cannot be matched or fully validated. Use one row per issue, not one row per employee, because one employee may have multiple issues.

Recommended columns:

- `姓名`
- `飞书部门`
- `花名册部门`
- `加班表部门`
- `职级`
- `异常类型`
- `涉及数据源`
- `飞书考勤是否存在`
- `花名册是否匹配`
- `上月汇总是否匹配`
- `加班表是否匹配`
- `当前处理方式`
- `建议人工核对`
- `详细说明`

Common exception types:

- `花名册匹配失败`: Feishu attendance has the employee, but the roster did not match.
- `上月汇总匹配失败`: Feishu attendance has the employee, but previous summary data did not match.
- `职级缺失或无法识别`: late deduction cannot be calculated reliably.
- `育儿假资格未校验`: childcare leave exists, but child birth date or approval data is missing.
- `姓名重复`: name-only matching may merge different employees.

For years other than 2026, keep the same business meaning and update the year text if HR asks for exact header names.

## Mapping Guidance

When using raw Feishu exports:

- Inspect the header row first.
- Preserve the raw file.
- Create a normalized copy with the canonical column names.
- Do not infer employee identity from name alone when duplicates exist.
- Keep dates as actual dates where possible.

If a source file contains one row per attendance event instead of one row per employee, filter and aggregate it before running the script:

- Exclude events after the employee's resignation date.
- On the hire date, exclude late arrival and upper-shift missed punch, but keep lower-shift missed punch.
- Preserve every excluded event in a rule-exclusion sheet or exception report.
- Sum late minutes by employee.
- Count missed punches by employee.
- Sum early-departure count and minutes by employee.
- Sum leave hours or days by employee and leave type, rounding each natural day upward and capping each day at `8` hours.
- Build detail strings by sorting events by date ascending or by Feishu export order if HR prefers.
- Sum daily `应出勤天数` by employee and multiply by `8`; use `应出勤时长(分钟)` only as the fallback.
