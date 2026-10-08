---
name: social-insurance-fund
description: 按月汇总养老、失业、医保、生育险和公积金数据，支持从本地社保费管理客户端或用户上传文件中提取五险一金明细，合并多公司数据，计算单位和个人金额，生成工资可用的五险一金汇总表，并输出异常清单供复核。
---

# Social Insurance Fund

## Purpose

使用本 skill 将月度社保和公积金导出数据整理为可用于工资核算的 `五险一金` 汇总表。

公司范围中包含深圳分公司 `广思`。广思与其他公司不是同一缴纳主体，通常需要单独提供五险数据；不得因为其他公司的源文件已齐全，就默认广思已经包含在其中。

The current supported workflow is based on company-specific exports for:

- `养老`
- `失业`
- `医保`
- `公积金`

Do not include `工伤` in the payroll deduction summary because it is paid by the company and has no employee contribution in the current company workflow.

## Local Client Export

If the user has already installed the local social-insurance client, prefer downloading the monthly `养老`、`失业`、`医保` raw data from that client before asking for manual uploads.

Known local environment:

- Client launcher: `E:\Program Files (x86)\EPPortal_Soft\EPPortal_SB\EPEvenue_SB.exe`
- Company accounts are expected to have been pre-added by the user.
- Login password must be requested from the user in the current session and must not be hardcoded into the skill.

Preferred operational sequence:

1. Open the client launcher.
2. Log in to the target company account from the `操作` column.
3. If the `操作` column is hidden, move the horizontal scrollbar all the way to the right first.
4. Enter the temporary login password when prompted.
5. Close popups and message-center notices as they appear.
6. Open `社保费申报` -> `申报记录`.
7. Click `查询` to load the current company's records.
8. Open the target month row through `查看`.
9. Click the `缴费人数` cell and confirm downloading employee detail data.
10. Export the detail view as `PDF`.
11. Save the file locally on the desktop with a clear name including date, company, and content.

If the client export fails or the file cannot be read, ask the user to provide the exported file directly and continue with workbook parsing from that file.

## Standard Workflow

1. Confirm the payroll month and company scope.
2. 每月开始并表前，先询问使用者：本月是否需要把深圳分公司广思的数据一起加入；如果使用者确认需要，要求提供本月广思的独立五险数据，以及广思公积金数据（如广思另有单独公积金文件）。
3. Collect the source files. Prefer the local client export first when available; otherwise use user-provided files. Prefer a folder containing these subfolders:
   - `养老`
   - `失业`
   - `医保`
   - `公积金`
4. 如果纳入广思，优先读取使用者单独提供的广思申报明细表。广思文件可能是一张横向合并表，同时包含养老、失业、医保、生育、工伤等项目，不得强行按普通的“养老/失业/医保”三张表处理。
5. Read source files without changing them.
6. Derive company name from filenames like `感观-养老.xlsx` or from the client export context. 广思单独文件统一归属为 `广思`，不能并入 `未本` 或其他公司的五险记录；如业务要求称为“未本广思”，在输出中保留清晰的公司归属说明。
7. Normalize each insurance export into a standard detail table.
8. Merge employees by company plus name. Use certificate number as supporting evidence when available.
9. Calculate employer-side values from the confirmed company rules.
10. Generate:
    - `五险一金` summary sheet.
    - Detail sheets for `养老`, `失业`, `医保`, `公积金`.
    - `异常` sheet for missing files, duplicate names, missing amounts, or incomplete matches.
11. 在每个阶段完成后，主动向用户说明已完成什么，并引导下一步操作。不要只停留在“结果已生成”，要明确提示用户下一步该提供什么、确认什么、或继续做哪一段流程，例如“请确认本月是否纳入广思”“请先人工校准考勤再进入工资测算”“请提供五险一金模板或公积金文件”。

## Current Calculation Rules

Use these contribution rules unless the user provides a different month-specific policy:

- Pension employee contribution: use source `应缴费额`.
- Pension employer contribution: `缴费基数 * 16%`.
- Unemployment employee contribution: use source `应缴费额`.
- Unemployment employer contribution: `缴费基数 * 0.5%`.
- Medical employee contribution: use source `应缴费额`.
- Medical employer basic contribution: `缴费基数 * 8%`.
- Maternity insurance contribution: `缴费基数 * 0.7%`.
- Medical employer total: medical employer basic contribution plus maternity insurance contribution.
- Housing fund employer contribution: use source `单位缴存额`.
- Housing fund employee contribution: use source `个人月缴存额`.

Round calculated contribution amounts to 2 decimal places for payroll output.

### 广思专项规则

- 广思源表中的养老、失业、医保、生育、工伤费率和应缴金额可能与其他公司不同，必须以广思当月源表中的费率和应缴费额为准，不得套用其他公司的默认费率。
- 广思的单位承担金额按源表拆分记录；工伤仍只计入公司承担，不计入员工个人扣款。
- 广思源表中常见的地方补充医疗、职业年金、公务员医疗补助、家属统筹医疗、地方补充养老等额外项目，只有在工资模板或用户明确要求纳入时才计入对应汇总；否则保留在原始明细或异常说明中，不得擅自并入基本五险一金合计。
- 广思五险文件与公积金文件可能不是同一份文件。若广思五险文件中出现员工，但公积金源文件中没有该员工，不得把公积金静默填为 `0`；应保留五险数据，并在 `异常` 表注明“公积金源文件未匹配，待人工确认”。
- 如果使用者本月确认不纳入广思，则不读取、不追加广思数据，并在计算说明中记录本月未纳入广思。

## Source Formats

For `养老`, `失业`, and `医保`, expect either a `.xlsx` sheet or a client-exported `PDF` table that can be normalized into the same columns. The header row should contain:

```text
序号, 姓名, 证件类型, 证件号码, 缴费工资, 缴费基数, 费率, 应缴费额, 减免费额, 应补(退)费额, 人员编号
```

For `公积金`, expect an `.xls` or `.xlsx` export where the header row contains:

```text
个人账号, 姓名, 证件类型, 证件号码, 当前工资, 单位缴存额, 个人月缴存额, 补贴工资, 补贴月缴, 合计月缴额, 个人状态
```

Some housing-fund exports use old `.xls` format and require Microsoft Excel automation on Windows. If the script cannot read `.xls`, ask the user to resave the file as `.xlsx`.

If the local client only provides `PDF`, do not invent missing numbers. Read the table if possible; otherwise ask the user to re-export or resave as `xlsx`.

广思专项申报明细表可能使用横向多级表头：员工姓名、证件号码、缴费所属期位于左侧，养老、医保、生育、失业、工伤等项目按“费率”和“应缴费额”成对出现。读取时必须按表头定位列，不能仅依赖固定列号；应将每个险种的单位/个人金额、费率、缴费工资和证件号码保留到 `五险原始明细`。

## Script

Use `scripts/build_social_insurance_fund.py` for deterministic workbook generation.

Batch-folder mode:

```bash
python scripts/build_social_insurance_fund.py \
  --month 2026-06 \
  --input-root Downloads \
  --output outputs/social_insurance_fund_2026-06.xlsx
```

Single-company mode:

```bash
python scripts/build_social_insurance_fund.py \
  --month 2026-06 \
  --company 感观 \
  --pension 感观-养老.xlsx \
  --unemployment 感观-失业.xlsx \
  --medical 感观-医保.xlsx \
  --housing-fund 感观-公积金.xls \
  --output outputs/感观_五险一金_2026-06.xlsx
```

## Validation Rules

Always create an exception row when:

- A required source file is missing for a company.
- 使用者确认纳入广思，但未提供广思独立五险文件或广思公积金文件。
- The same company and employee name appears multiple times in the same insurance type.
- A contribution base is missing for a source row that needs calculated employer contribution.
- Housing-fund certificate number is masked and name-only matching is used.
- An employee appears in one insurance type but not in another.
- 广思员工出现在五险文件中，但没有在公积金文件中匹配到。
- 广思源表包含无法归类的额外项目或无法识别的多级表头。
- A source amount is blank, zero, or not parseable.

Do not silently fill missing employees with `0` unless the source file confirms they have no contribution for that month.

## Output Meaning

In the `五险一金` summary:

- `社保单位缴纳` equals pension employer plus unemployment employer.
- `社保个人缴纳` equals pension employee plus unemployment employee.
- `医保单位缴纳` equals medical employer basic contribution plus maternity insurance.
- `医保个人缴纳` equals medical employee contribution.
- `公积金单位缴纳` and `公积金个人缴纳` come directly from the housing-fund export.

Keep raw detail sheets in the output so HR can audit how each summary value was derived.

## Conversation Guidance

- 每次完成一个可交付阶段后，主动接下一步，而不是等待用户重新发起。
- 用一句话总结当前进度，再给出一个明确的下一步问题或材料清单。
- 若当前阶段已经结束但后续流程还未开始，优先询问用户要继续哪一步：继续补资料、做复核、还是进入下一阶段计算。
