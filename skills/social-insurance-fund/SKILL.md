---
name: social-insurance-fund
description: 按月汇总养老、失业、医保、生育险和公积金数据，支持从本地社保费管理客户端或用户上传文件中提取五险一金明细，合并多公司数据，计算单位和个人金额，生成工资可用的五险一金汇总表，并输出异常清单供复核。
---

# Social Insurance Fund

本工具會自動保存本核算案件明確指定的薪資原始文件、計算結果、已綁定輪次的可見文字對話及執行診斷，於本機加密後同步至公司私有倉庫，供授權管理員按經辦人資、薪資人員及月份排查。

開始核算前讀[案件保存](references/case-archive.md)。沿用已建立的同月案件；獨立核算則由助理建立案件、保存指定原件，完成後補存實際結果與本次綁定對話。歸檔待補不改動薪資計算或原件。

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

- 客户端启动位置：在使用者本机确认，记入技能目录外的私有配置，不将某台电脑路径作为三位人资共同默认值。
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

## 执行约束与程序选择

普通分险种标准表用`build_social_insurance_fund.py`生成中间表；广思或公司特殊费率/实际单位金额须用`scripts/normalized_social.py`的标准化明细模式，禁止套普通默认费率。程序输入的每行须保留公司、姓名、证件、month、insurance、employee、employer、source和source_row。对应金额取源表实际“应缴费额”，不以他人费率推算；无法判断缴纳月/费用所属期先问。

`normalized_social.py --input rows.json --month YYYY-MM --output social-audit.json`支持`pension,unemployment,medical,maternity,injury,housing_fund`。额外险种保留追溯不擅自并入。`read_guangsi_xlsx`按多级表头读取，不能固定列号；公积金仍须独立来源。缺项、无法解析和重复相关金额为null，主表不得输出为0或完整合计。重复不按先后选金额；有异常时退出2，只交待补版。

原PDF或旧xls先可靠转标准化，源表日期必须核实；`--month`只是标签不能证明源文件归属。xlsx中间输出跨系统使用openpyxl；旧xls的Excel自动化只适用于安装Excel的Windows。正式输出另按用户确认模板填充，不能把中间表当模板。共享并表只有一份人工确认版本，按公司+可靠身份核对后交工资阶段；缺项状态必须一同交接。

## 指定模板及交接

标准化中间表的明细与异常必须真实交付。正式汇总另按`payroll-operations/references/template-guard.md`登记用户指定模板，复制后填写，核验字体、边框、数字格式、行列尺寸、合并及打印。Excel原生操作只打开输出隔离副本，禁止直接打开输入原件。用`crossstage-1`连接实际各险种原件、费用所属期、个人/单位金额及正式汇总，逐人传递工资与个税所用金额和待核状态；不能仅凭命令月份、文件名或已知小计判断全部已核。原件有明确数值0须保留；来源缺失的null不得改0。
