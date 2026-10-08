# 工资核算 Skill 完整内容

生成日期：2026-10-08

## 文档用途

本文档完整收录公司工资核算 Skill 套件的当前正式内容，用于业务审阅、版本归档、内部交接和云端发布前复核。

本文档本身不是可直接安装的 Skill。Codex 安装时需要保留每个 Skill 的目录结构、`SKILL.md`、`agents/openai.yaml`、引用文件和脚本。正式安装请使用同目录下的完整在线安装包和 `install.ps1`。

## 套件组成

- `payroll-operations`：工资发放总流程和阶段引导。
- `attendance-calculator`：月度考勤计算和审计。
- `attendance-settlement`：考勤确认结果与工资结算衔接。
- `social-insurance-fund`：多公司五险一金并表。

## 安装包目录

```text
工资核算Skill在线安装包/
├─ install.ps1
├─ 发布说明.md
├─ 在线安装命令模板.txt
└─ skills/
   ├─ payroll-operations/
   ├─ attendance-calculator/
   ├─ attendance-settlement/
   └─ social-insurance-fund/
```

## 总流程 Skill

# payroll-operations

### payroll-operations/agents/openai.yaml

```yaml
interface:
  display_name: "工资发放总流程"
  short_description: "按月统筹考勤、五险一金、薪资、个税和飞书工资单。"
  default_prompt: "请按目标月份执行月度工资发放流程，在每个节点索要缺失资料，并生成可复核的工资、个税和飞书工资单文件。"
```

### payroll-operations/references/三位人资部署与月度操作教程.md

````markdown
# 三位人资部署与月度操作教程

## 1. 教程用途

本教程供艾美兰、范艳艳、俞霞共同执行公司月度工资核算。目标是让每位人资只需按阶段准备自己负责范围的资料，Codex 负责读取共享资料、核对人员范围、计算、生成异常清单并提示下一步。

不要在流程开始时一次性收集全部材料。完成一个阶段、人工确认一个阶段，再进入下一阶段。

## 2. 首次部署

### 一键在线安装

每台电脑首次使用时，打开 PowerShell，复制下面整条命令并按回车：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$z=Join-Path $env:TEMP 'payroll-skill-suite.zip'; $d=Join-Path $env:TEMP 'payroll-skill-suite-online'; Invoke-WebRequest 'https://github.com/sdasdasd1223/payroll-skill-suite/archive/refs/heads/main.zip' -OutFile $z; if(Test-Path -LiteralPath $d){Remove-Item -LiteralPath $d -Recurse -Force}; Expand-Archive -LiteralPath $z -DestinationPath $d -Force; $i=Get-ChildItem -LiteralPath $d -Filter install.ps1 -Recurse | Select-Object -First 1; if(-not $i){throw '下载包中没有找到 install.ps1'}; & $i.FullName"
```

看到“安装成功”后，完全退出并重新启动 Codex。以后更新也运行同一条命令，安装程序会先备份旧版本再安装最新版。

公开仓库地址：`https://github.com/sdasdasd1223/payroll-skill-suite`

### 安装结果检查

安装完成后确认：

1. 确认 `payroll-operations`、`attendance-calculator`、`attendance-settlement`、`social-insurance-fund` 已出现在 `C:\Users\<当前用户名>\.codex\skills\`。
2. 确认每个 skill 根目录都有 `SKILL.md`；`payroll-operations\agents\openai.yaml` 存在且界面文字为中文。
3. 确认飞书 CLI 已安装，并使用本公司有权读取花名册和通讯录的账号登录。
4. 重启 Codex，使新安装或更新的 skill 生效。
5. 第一次运行时说：`请使用 payroll-operations，开始 <工资月份> 的工资核算，并按阶段引导我准备资料。`

飞书登录、授权失效或实时读取失败时必须暂停。不要用上月缓存花名册代替实时花名册继续正式计算。

## 3. 每月目录

建议在桌面建立一个月度总文件夹：

```text
桌面\<YYYY-MM>-工资核算\
├─ 00-共享模板\
│  ├─ 工资表模板.xlsx
│  ├─ 个税模板.xls
│  ├─ 飞书工资单模板.xlsx
│  ├─ 考勤表模板.xlsx
│  └─ 五险一金汇总表模板.xlsx
├─ 01-考勤原始数据\
│  ├─ 本月每日统计.xlsx
│  ├─ 上月考勤确认汇总.xlsx
│  └─ 本月加班数据.xlsx（仅云端表不适合目标月份时提供）
├─ 02-考勤确认版\
│  ├─ 艾美兰\
│  ├─ 范艳艳\
│  └─ 俞霞\
├─ 03-五险一金\
│  ├─ 各公司五险源文件\
│  ├─ 各公司公积金源文件\
│  ├─ 广思独立资料（本月纳入时）\
│  └─ 本月五险一金总表.xlsx
├─ 04-工资资料\
│  ├─ 艾美兰\
│  │  └─ <公司或管理范围>\
│  ├─ 范艳艳\
│  │  └─ <公司或管理范围>\
│  └─ 俞霞\
│     └─ <公司或管理范围>\
├─ 05-绩效与奖励\
├─ 06-人工调整\
├─ 07-输出\
└─ 08-异常与确认\
```

现有文件夹不符合这个名称时不必重新搬运。告诉 Codex 实际路径即可；Codex 必须先复述实际读取路径和识别到的资料范围，再开始计算。

## 4. 阶段一：考勤初算

### 人资要准备

- 工资月份。
- `01-考勤原始数据` 中的本月原始考勤。
- 上月最终确认的考勤汇总。
- 目标月份对应的加班数据。云端加班表混有后续月份、无法还原月末状态时，必须提供目标月份截止日的本地版本。

云端花名册、标准加班表和审批数据由 Codex 每月重新读取，不要求人资重复导出。需要登录或重新授权时，Codex 应停在授权步骤等待人工完成。

### Codex 要做

1. 说明读取的本地文件和云端来源。
2. 调用 `attendance-calculator`，核对入职、离职、迟到、缺卡、早退、请假、年假、育儿假、加班和累计结余。
3. 先解释差异或异常来源，再决定是否改数据；不得为了对齐历史结果直接修改。
4. 输出考勤初算表和异常清单到 `07-输出\考勤初算`。
5. 明确提示这不是工资结算版。

### 人工确认

人资将初算表发给员工确认，处理特殊情况后，把修正版分别放入 `02-考勤确认版\<负责人>`。人工修正后的版本是后续工资计算唯一考勤依据，不得被重新计算覆盖。

## 5. 阶段二：五险一金并表

### 人资要准备

- 各公司养老、失业、医保和公积金源文件，放入 `03-五险一金`。
- 本月是否纳入深圳广思。纳入时同时提供广思独立五险资料和公积金资料。

### Codex 要做

1. 调用 `social-insurance-fund` 合并多公司数据。
2. 保留原始明细和异常说明，不把缺失匹配静默填成 0。
3. 输出唯一一份共享的本月五险一金总表，不要求三位人资分别维护副本。
4. 提示用户先确认未匹配、同名、缺险种和广思资料是否完整。

确认后的总表放在 `03-五险一金\本月五险一金总表.xlsx`。

## 6. 阶段三：三位人资工资资料

每位人资只维护 `04-工资资料\<本人姓名>` 下自己负责的公司或管理范围。

每个公司或范围建议包含：

- 上月或当月人资维护的工资底稿，用于读取人员历史薪资和人工字段。
- `02-考勤确认版` 中对应范围的确认版考勤。
- 本月新入职员工的 offer 或录用通知。
- 本月调薪、提前转正、岗位变动、特殊补贴等已确认资料。
- 本范围绩效、奖励金和其他人工调整资料；也可以统一放在 `05-绩效与奖励`、`06-人工调整`。

只有本月新入职且工资底稿没有薪资标准的员工需要提供 offer。历史月份已入职员工不要求每月重复提供 offer，应从已确认工资记录、本人历史递增方案和实时花名册继续计算。

共享的正式工资模板只放在 `00-共享模板`。模板负责列、顺序、样式和公式，不负责提供人员名单、部门、组别或薪资金额。

## 7. 阶段四：工资计算和纠错

计算开始前，Codex 必须重新读取实时花名册，并先汇报：

- 三位人资各自识别到哪些公司或管理范围。
- 使用哪份确认版考勤、工资底稿、共享模板和五险一金总表。
- 哪些当月新人资料、绩效和人工调整已找到。
- 哪些资料缺失，是否需要暂停。

执行时遵守以下长期规则：

- 人员是否在发薪范围，以及所属公司、部门、小组，以实时花名册为准。
- 现有员工在最新小组内保持原相对顺序；新员工放在所属小组末尾。
- 工资金额必须回查本人底稿、本人 offer 或本人已确认历史记录，不能根据另一名员工案例猜测。
- 普通在职非实习员工的事假时间默认不显示、不自动扣款，但所有人的请假扣款公式都保留，便于人资人工补录后联动。
- 实习生当月事假当月结算，分母为本月应出勤小时数；即使当月离职仍使用实习生口径。
- 非实习离职员工结算累计事假，分母为 174；累计值已经包含本月时不得重复相加。
- 请假扣款、考勤扣款、入离职差额等计算列保留公式，不写死。
- 产假按正常薪资与产假薪资分段；月中产假差额放在 `其他个人扣款`，不得放进入离职差额。复工后恢复正常薪资。
- 递增薪资先更新本月薪资档位，再计算入离职差额；月中生效的当期补差放入其他补贴。
- 数字写为真正的数字格式；0 按模板习惯显示为空白；负数不得出现两个减号。
- 不擅自增加模板没有的列，不改变字体、列宽、行高、边框、颜色或隐藏状态。
- 备注只保留必要的入离职时间和确需人工复核的事项，避免重复解释正常规则。

工资结果按负责人和公司范围输出到 `07-输出\工资表`，原始模板和底稿保持不变。

## 8. 阶段五：工资表人工确认

三位人资分别复核自己负责的工资表，重点检查：

- 人员、公司、部门和小组。
- 新人薪资、递增和补差。
- 实习生与离职员工事假。
- 病假、产假、入离职差额和个人扣款。
- 五险一金、绩效、奖励金和人工调整。
- 修改输入单元格后，公式是否自动更新。
- 字体、顺序、模板列和数字格式。

人资明确回复“工资表已确认”后，该版本成为本月正式工资基准。后续不得自行重算或回退。

## 9. 阶段六：个税表

### 需要的材料

- 已确认工资表。
- `00-共享模板` 中的政府个税原始模板。
- 实时花名册中的证件和所属公司信息。

### 输出规则

- 直接复制政府模板，不新建表格，不删除任何列。
- 只填写实际黄色列；`工号`是唯一经确认的非黄色例外，表头仍为`工号`，内容填写花名册所属公司。
- 每位人资输出一份个税表。
- 缺证件、公司为空或匹配不唯一时写入异常，不猜测。

人工确认通过后，将个税表标记为本月正式版本。

## 10. 阶段七：飞书工资单

### 需要的材料

- 已确认工资表。
- `00-共享模板` 中的飞书工资单模板。
- 公司总账号的实时飞书通讯录授权。

### 输出规则

- 最终工资表中的当月有薪人员仍保留。
- 只有在公司总账号通讯录中匹配到唯一且已激活账号的人进入飞书工资单。
- 查不到、未激活、跨租户或同名歧义人员排除，并生成原因清单。
- 不在工资单阶段重新计算工资金额。

## 11. 阶段八：最终交付

`07-输出` 最终应包含：

- 各负责人或各公司的最终工资表。
- 五险一金汇总表。
- 三位人资个税模板。
- 飞书工资单导入表。
- 飞书工资单排除与复核清单。
- 全流程异常清单。

Codex 最终必须汇报使用了哪些文件、各表人数、排除人数及原因、仍待确认的事项，并主动说明流程是否已经结束。

## 12. 每月启动话术

人资只需先说：

> 请使用 payroll-operations 开始核算 YYYY 年 MM 月工资。月度资料在桌面的“YYYY-MM-工资核算”文件夹，请先检查第一阶段资料，并一步一步提示我。

Codex 应先检查考勤阶段的最小材料，不应立刻要求提供后续全部工资、个税和工资单资料。

## 13. 临时情况如何记录

以下内容放入 `08-异常与确认\当月运行说明.md` 或异常清单，不写进永久 skill：

- 某名员工当月特殊金额或人工处理结论。
- 某个月固定的应出勤小时数、人数或日期。
- 临时调组、跨公司、补发补扣和特殊发薪方式。
- 当月某份表缺列、错列或人工修订。
- 一次性排除人员及其原因。

只有人资明确确认“以后每个月都这样执行”的口径，才升级为通用规则并更新 skill。
````

### payroll-operations/SKILL.md

```markdown
---
name: payroll-operations
description: "统筹公司月度工资发放全流程，从资料收集、考勤初算、人资和员工确认、五险一金并表、工资计算、绩效和奖励金核对、政府个税模板生成、飞书工资单筛选到最终交付。当需要执行或解释月度薪资流程、向用户索要缺失工资资料、核对人资确认后的考勤、计算工资表、生成个税导入模板、准备飞书工资单导入文件时使用。"
---

# 工资发放总流程

## 目标

使用本 skill 作为公司月度工资发放的总调度流程。

本 skill 不替代已有专项 skill，而是负责编排它们：

- 原始飞书考勤计算使用 `attendance-calculator`。
- 考勤结果和工资结算字段衔接使用 `attendance-settlement`。
- 五险一金并表使用 `social-insurance-fund`。
- 读取飞书云端花名册、通讯录、审批、表格时，使用已安装的飞书 CLI 及相关 skill，例如 `lark-sheets`、`lark-contact`、`lark-approval`。

本 skill 的核心职责是：按阶段推进、索要资料、设置人工确认节点、执行工资口径、输出异常清单，并确保不要跳过人工复核环节。

## 总原则

工资发放不是一次性计算，而是一个分阶段的月度项目。

每完成一个阶段，都要：

1. 说明本阶段完成了什么。
2. 说明生成了哪些文件。
3. 说明有哪些异常、缺失资料或需要人工复核的点。
4. 如果下一阶段缺资料，要明确向用户索要具体文件或约定文件夹。
5. 遇到人工确认节点时要暂停，除非用户明确说只是测试，否则不要用未确认的草稿数据继续算工资。

所有源文件默认只读，不要覆盖。所有处理结果应输出为新文件。

用户明确表示工资表、个税表或其他阶段结果已经人工确认通过后，应将该版本视为本月正式基准。除非用户明确提出变更，不得重新计算、替换或回退已确认数据；后续个税、工资单和最终交付都应从该确认版本继续。

## 通用规则与当月例外

更新或执行本 skill 时，必须先判断规则属于哪一类：

- `通用规则`：每月都适用的计算口径、数据权威顺序、模板保真要求、公式要求、人工确认节点和异常处理纪律。通用规则写入本 skill 或对应专项 skill。
- `模板规则`：由当前正式模板本身决定的字段、列顺序、颜色、样式和公式。每次以用户提供的最新正式模板为准，不把旧模板的字段增删写成永久规则。
- `当月例外`：只与某个月、某名员工、某笔金额、某次调组、某份临时表或某个人工确认有关。只写入当月运行说明、异常清单或确认记录，不写成其他月份自动套用的永久规则。

具体姓名、固定金额、固定人数、固定工时和某个月的文件名通常都是当月例外。已确认的个人历史薪资或递增方案可以作为该员工后续月份的可追溯数据来源，但不得被抽象成适用于其他员工的规则。

执行完整月度流程或帮助新的人资部署时，读取 [references/三位人资部署与月度操作教程.md](references/三位人资部署与月度操作教程.md)，按其中的目录约定、分阶段材料清单、确认节点和提示话术推进。

## 温馨提示与阶段引导

用户进入本 skill 时，要用温和、清晰、一步一步的方式引导，不要一次性索要全流程所有材料。除非用户明确要直接跑完整工资流程，否则先从当前阶段需要的最小材料开始。

每次提示下一步时必须同时说清楚：当前处于哪个阶段、用户需要提供哪些文件、建议放到哪个文件夹、将读取哪些云端数据、完成后会生成什么文件、是否需要人工确认。用户采用其他目录时，以实际目录为准并复述已识别的路径，不能因为目录名不同而要求用户重复整理。

初次启动或开始新月份时，先提示：

“我们先从本月考勤初算开始。请告诉我计算几月份，并提供上月考勤汇总、本月原始考勤数据、本月加班或调休数据。我会每个月重新读取云端花名册；如果本月加班表不用云端版本，也请把对应的本地加班表发给我。”

考勤初算完成后，必须暂停并提示：

“这份是考勤初算版，还不能直接进入工资计算。请先发给人资和员工确认；如果有人有特殊情况，请人工校准后把确认版 Excel 再发给我。我后续会以确认版为准，不覆盖人工修改。”

如果用户想跳过人工确认，要提醒：

“如果只是测试，我可以继续用初算版往下跑；如果是正式发薪，不建议跳过员工和人资确认节点。”

收到人工确认版考勤后，再提示进入下一步：

“接下来可以进入五险一金并表和工资资料收集。请提供本月五险一金源文件、工资表模板、绩效或奖励金资料、人工补贴/扣款/第三方代发/现金等调整资料；云端花名册我会重新读取。”

五险一金并表完成后，要提示用户先确认异常清单，再进入工资表计算。

工资表计算完成后，要提示用户确认最终工资表；确认后再生成政府个税模板和飞书工资单导入文件。

## 标准月度流程

### 阶段 0：工资项目初始化

先确认：

- 工资月份，例如 `2026-07`。
- 发薪日期，例如 `2026-08-10`。
- 本月公司范围。
- 本月人资负责人范围。
- 资料是用户手动上传，还是统一放到约定文件夹。
- 当前云端花名册链接。当前标准云端花名册为：
  `https://l5rd0uzm0z.feishu.cn/sheets/CUA7sW7WChtjsgt5c1xcWIvCnsd`
- 当前加班表链接。当前标准加班表为：
  `https://l5rd0uzm0z.feishu.cn/wiki/Yzs4wnk2CilJYAkRRNdcmSidnhf?table=tblG5IjKRK8l5jgo&view=vewCi8JgYW`
- 飞书 CLI 是否已经安装、配置并登录。

推荐资料文件夹约定：

- `Desktop/<月份>-payroll/attendance-raw/`：飞书原始考勤。
- `Desktop/<月份>-payroll/attendance-confirmed/`：人资或员工确认后的考勤修正版。
- `Desktop/<月份>-payroll/social-insurance/`：五险一金源文件。
- `Desktop/<月份>-payroll/payroll-templates/`：工资表模板或各人资工资表。
- `Desktop/<月份>-payroll/performance/`：绩效数据。
- `Desktop/<月份>-payroll/offers/`：offer、录用函、新人薪资资料。
- `Desktop/<月份>-payroll/manual-adjustments/`：人工补贴、扣款、第三方代发、现金、补发补扣等资料。
- `Desktop/<月份>-payroll/outputs/`：最终输出。

三位人资共同使用时，优先采用教程中的统一目录：共享模板、五险一金总表和云端数据放在共享区；艾美兰、范艳艳、俞霞各自的确认版考勤、工资底稿、当月新人 offer、绩效和人工调整放在各自负责人目录。不要把共享五险一金总表复制成三份后分别维护。

如果用户使用了其他文件夹，以用户实际文件夹为准，并在回复里说明实际读取路径。

统一模板与核算数据来源分离：用户提供独立的“薪资核算模板表”文件夹时，该文件夹中的工资模板只决定工资表字段、字段顺序、样式和模板公式，不作为人员、薪资或人工金额的数据来源。员工名单、部门/组别、工资基数、人工字段及考勤数据仍从各人资负责人原资料文件夹、实时花名册、确认版考勤和其他已确认数据源读取。模板中的示例姓名、示例金额、历史演示工作表不得混入正式结果。输出以统一模板为底稿并填入核算数据，不覆盖模板和原始数据文件；交付前核对字段与模板完全对应，确保没有多余旧列、隐藏列或错位字段。

### 阶段 1：考勤初算

需要收集：

- 飞书月度原始考勤导出。
- 最新花名册，包含入职时间、离职时间、职级、部门、员工类型、子女出生日期等字段。
- 如果可以拿到，花名册或考勤输出中应包含飞书 user ID / open_id。
- 上月考勤汇总。
- 加班结余表或飞书多维表格导出。
- 加班、请假、调休、补班、育儿假等相关表格或审批数据。

调用 `attendance-calculator` 进行考勤初算。

考勤汇总输出要求：

- 如果能拿到飞书 user ID / open_id，应把它放在第一列，排在部门和姓名前面。
- 这样后续可以避免同名、改名、离职状态、飞书工资单发送对象等匹配问题。

考勤细则：

- 事假源数据按分钟统计，汇总按小时展示。
- 病假按小时统计。
- 一天事假最多不超过 8 小时。
- 年假按天统计，最小单位 0.5 天。
- 育儿假按天统计，最小单位 1 天。
- 缺卡按次数统计。
- 本月应上班小时数按当月应出勤工作日乘以 8 小时，除非用户提供了特殊日历口径。
- 育儿假列要写成累计说明，例如：`截至2026年7月（含），有10天育儿假，已使用9天`。
- 育儿假不要写进“考勤情况”列。
- “考勤情况”列只写迟到、事假、病假、年假。
- 缺卡单独放在“缺卡情况”列。
- 员工满一年当月可能新增年假额度，不能因为上月余额为 0 就把本月使用年假后的余额错误算成 0。
- 育儿假和年假都要继承上月累计记录；本月没使用，也要保留可用和已用情况，除非花名册显示资格已经失效或人员已离职。

阶段结束后要暂停，并提示用户：请把考勤初算结果发给人资和员工确认。该输出只是初算版，不是工资计算确认版。

### 阶段 2：考勤人工确认回收

工资计算必须使用“人工确认后的考勤修正版”，不能直接使用考勤初算草稿。

原因：

- 员工可能解释迟到、缺卡、请假等特殊情况。
- 人资可能手工调整考勤扣款、请假时长、缺卡记录。
- 工资表应以最终确认版考勤为准。

如果缺少确认版考勤，要向用户索要：

- “请把确认后的考勤修正版放到约定文件夹。”
- 或者“请上传本月人资确认后的考勤表。”

不要用重新计算的原始考勤覆盖确认版考勤中的人工修改。

如果确认版和初算版不一致，要先说明差异来源、规则原因或人工修改痕迹，再由用户判断是否采纳；不要为了对齐结果直接改数据。

### 阶段 3：五险一金并表

需要收集各公司源文件：

- 养老。
- 失业。
- 医保。
- 公积金。

调用 `social-insurance-fund`。

优先顺序：

1. 如果本机已安装社保费管理客户端，优先从客户端下载 `养老`、`失业`、`医保` 明细。
2. 如果客户端不可用或导出失败，再使用用户上传的 Excel 或 PDF 文件。
3. 公积金仍按单独平台导出，不和社保费管理客户端混在一起。

当前规则：

- 医保表中包含医保和生育险。
- 工伤由公司承担，员工不缴，不进入个人扣款汇总。
- 养老单位金额 = 缴费基数 × 16%。
- 失业单位金额 = 缴费基数 × 0.5%。
- 生育险 = 缴费基数 × 0.7%。

校验要求：

- 不能随意把缺失人员填 0。
- 要检查同一个人是否只出现在部分险种里。
- 要检查同名重复。
- 有证件号码时，用证件号码辅助判断。
- 证件号码脱敏或缺失时，要写异常。
- 输出中要保留明细 sheet，方便人资复核。

### 阶段 4：工资资料收集

需要收集：

- 各人资负责人或各公司的工资表模板。
- 阶段 2 的考勤确认版。
- 阶段 3 的五险一金汇总。
- 云端花名册。
- offer、录用函、新人薪资资料。
- 绩效表。
- 人工补贴、人工扣款、第三方代发、现金或其他银行、补发补扣等资料。
- 飞书审批数据，例如调薪、提前转正、岗位调整、特殊补贴等。

如果缺资料，要明确问用户，不要猜。

如果用户按人资负责人提供资料文件夹，例如 `桌面/范艳艳/`，则默认该文件夹就是本阶段的资料包。不要要求用户重复解释公司范围；应主动扫描文件夹结构并识别：

- 人资负责人名称，通常来自文件夹名。
- 该负责人管理的公司或工资范围，通常来自子文件夹名，例如 `小竹-工资资料收集`、`未本咨询-工资资料收集`。
- 每个子文件夹里的人工确认后考勤表。
- 每个子文件夹里的工资底稿或工资模板。
- 父文件夹里的五险一金总表；五险一金总表可能包含全公司人员，不只包含该负责人范围。

识别后应向用户简短说明本次读取到的资料范围，并直接按子文件夹分别生成该负责人管理范围内的工资表。只有在文件夹里缺少考勤确认版、工资底稿或五险一金总表时，才向用户索要对应缺失资料。

关键边界：

- 考勤表里有但工资表里没有的人，要写入异常清单。
- 不要自动新增工资表人员，除非已经有薪资标准、公司主体、岗位信息、人工字段等必要资料。
- offer 可以帮助发现新人，但最终进入工资表仍需要确认薪资和公司主体。

### 阶段 5：工资计算

以工资表模板为目标格式，尽量保留原公式、原格式和人工字段。

开始工资计算前，必须重新实时读取云端花名册，不能使用旧本地缓存或上次快照作为正式人员范围依据。如果飞书 CLI、授权或网络导致实时读取失败，要停止计算并明确告知用户；不要回退旧花名册继续生成工资表。

人员所属公司、部门和小组也必须以本月实时花名册为权威来源。组织架构调整应先维护在花名册，再同步到工资表；工资底稿和考勤表中的旧组别不得覆盖花名册。组别取值优先使用花名册 `三级部门`，为空时依次使用 `二级部门`、`一级部门`，仍为空才列入异常并请人资确认。

常见工资字段：

- 基本工资。
- 岗位工资。
- 管理津贴。
- 岗位津贴。
- 绩效工资。
- 抽成。
- 薪资。
- 公司承担社保。
- 公司承担医保。
- 公司承担公积金。
- 其他补贴。
- 岗位晋升考核奖金。
- 绩效/激励。
- 奖励金。
- 事假时间。
- 病假时间。
- 请假扣款。
- 考勤扣款。
- 入离职差额。
- 应出勤小时数。
- 应发工资。
- 其他。
- 财务核算薪资。
- 其他个人扣款。
- 个人社保。
- 个人医保。
- 个人公积金。
- 薪资个税。
- 实发金额。
- 合并计算个税。
- 原应税金额。
- 第三方代发。
- 现金/其他银行。
- 招行。

核心公式：

- 应发工资 = 固定薪资类项目 + 其他补贴 + 绩效/激励 + 奖励金 + 请假扣款 + 考勤扣款 + 入离职差额。
- 财务核算薪资 = 应发工资 + 其他。
- 实发金额 = 财务核算薪资 - 个人扣款 - 个人五险一金 - 薪资个税。
- 原应税金额 = 财务核算薪资 - 个人扣款 + 合并计算个税。
- 招行 = 实发金额 - 第三方代发 - 现金/其他银行。

要保留的人工字段：

- 其他。
- 其他个人扣款。
- 第三方代发。
- 现金/其他银行。
- 备注。
- 用户未要求重算的既有人工值。

按人资负责人文件夹计算工资时：

- 每个公司或管理范围输出一份新的已填工资表，不覆盖原始工资底稿。
- 工资结果的字段完全以当月正式工资模板为准。模板不存在的列不得自行新增；旧底稿中多出的列不得带入正式结果；模板要求删除的列应实际删除而不是仅隐藏。某位人资、某家公司或旧月份的字段差异只作为当月模板事实处理，不固化成跨月份规则。
- 必须保留工资模板原有字体、字号、粗体、对齐、边框、填充、行高、列宽及合并方式。重排、增加或删除人员后，同类字段必须保持统一格式；不得因为员工顺序变化而改变字体。
- 部门或小组标题必须复制原模板中部门/小组标题单元格的格式，不能复制该组第一位员工的空白单元格格式。例如原组首人员被剔除或迁移后，新组首人员只提供数据行格式，组标题仍应沿用原组标题字体和对齐方式。
- 输出前要扫描同类标题的字体名称、字号、粗体和对齐方式；发现同级部门或小组标题格式不一致时，先统一到模板原有标题格式，再交付。
- 每月实时花名册是公司、部门和小组归属的唯一权威来源；工资模板只用于保留样式以及员工在其最新花名册小组内的原有相对顺序。花名册组别发生变化时，应将该员工移动到新组；不得因工资模板仍保留旧组别而拒绝移动，也不得自行猜测组别。
- 同一最新花名册小组内，工资模板已有员工保持原相对顺序；新增人员插入该小组末尾。例如模板中营销策划2组依次为吴小小、严文静、胡诗涵，新人彭欢在花名册中属于营销策划2组，则顺序应为吴小小、严文静、胡诗涵、彭欢。除花名册组别变化涉及的人员外，不得重排其他员工。
- 使用工资表中的公司字段映射五险一金总表，例如 `福州小竹 -> 小竹`、`未本共生 -> 共生`、`未本万思 -> 万思`、`未本上海 -> 未本`、`感观共识 -> 感观`、`广思 -> 广思`。
- 五险一金优先按公司主体加姓名匹配；匹配不到但姓名唯一时可弱匹配，并写入异常；同名多条不得自动匹配。
- 考勤确认版按姓名匹配到工资表。工资表有、考勤表没有的人，保留工资底稿原值并写异常，不要自动删除或补造考勤。
- 考勤确认版和最新花名册都存在的人，即使工资底稿没有，也不能直接当作遗失人员剔除；应先纳入工资计算范围，再按 `offer/录用函 -> 已确认调薪或薪资资料 -> 同公司同岗位唯一历史薪资标准` 的顺序寻找薪资来源。使用同岗历史薪资只能作为暂算，必须在异常清单中标记 `薪资标准暂算` 并要求人资确认。
- 工资底稿里有但最新花名册不存在，或最新花名册显示不在本月发薪范围的人员，应剔除并写入异常清单；不能继续给已从花名册消失的历史离职人员生成工资。
- `事假时间`是供人资校准的数字输入：仅对按规则需要结算事假的人员填写相应小时数；其他员工保持空白。该单元格不使用外部链接公式。
- `病假时间`填本月病假小时数；`请假扣款`按病假规则计算，超过 9 天或缺少最低工资口径时写异常。
- `考勤扣款`填确认版考勤扣款合计；若确认版中该列是公式且没有缓存结果，应按确认版表内公式口径重新计算。
- `应出勤小时数`以确认版考勤表为准；如果工资底稿中已有过期月份小时数，也应更新为本月确认值。
- `公司承担`中的社保、医保、公积金，以及 `个人扣款`中的社保、医保、公积金，从五险一金总表填入。
- 不要把 `其他个人扣款` 空白单元格自动改成公式；该列是人工字段，已有值或已有公式才保留。
- 应发工资、财务核算薪资、实发金额、原应税金额、招行等列可按工资模板既有公式补齐，并让 Excel 重新计算后保存。
- `请假扣款`、`迟到扣款情况`、`漏打卡扣款情况`、`考勤扣款情况`、`入离职差额` 等计算列必须保留公式，方便人工校准后自动联动；不得把计算结果写死为数值。
- 所有员工的 `请假扣款` 公式都必须引用本行 `事假时间` 和 `病假时间`，不能因为本月默认不结算、初始输入为空或数值为零，就把引用替换成固定的 `0`。这样人资后续手工补录事假或病假时间时，扣款、应发工资和实发金额会自动联动。
- `入离职差额` 只处理实际入职前或离职后的未出勤天数，不能混入产假、请假或其他工资调整；该列不能直接写死数值，必须保留公式或可追溯的计算路径。
- 产假人员按产假薪资计算时，复工后要从复工日起恢复正常薪资；月中开始或结束的产假，要按工作日拆分正常薪资和产假薪资，不能整月一刀切。
- 月中产假差额必须以正数扣款公式放在 `其他个人扣款`，由实发金额公式扣除；不得放入 `入离职差额`、`请假扣款`或 `考勤扣款`。例如正常月薪 10004 元、产假月薪 2000 元、当月 21 个工作日、产假 11 个工作日，则 `其他个人扣款=(10004-2000)/21*11=4192.57`，应发工资仍按模板正常形成，实发工资再扣除 4192.57 元。
- 实习生事假扣款分母使用本月应出勤小时数，该值按当月工作日和节假日重新计算，不是固定常数。

### 阶段 6：实习、试用、新人、入离职规则

先区分员工类型：

- 实习生：花名册类型为实习或实习员工。
- 试用期员工：全职，但尚未转正。
- 正式员工：全职且已转正。

实习或试用薪资递增：

- 本月新入职、工资底稿尚无薪资标准时，读取本月 offer 确认初始薪资和递增约定。
- 上月或更早已入职的员工，不要每月重新索要当初的 offer；以本月工资底稿和上月已确认薪资记录为基准，结合每月重新读取的花名册入职日期、转正日期，延续已确认的递增计划。不要仅因本月资料包没有历史 offer 就漏算递增，或标记为缺少 offer。
- 历史递增额或原薪资标准确实无可追溯依据时，列出具体缺失项请人资核对；不要凭“试用六个月”擅自推断转正薪资或每月增额，也不要把底稿已经包含的递增再次加一遍。
- 对当月有 `入离职差额` 的新人也必须单独检查薪资递增，二者不能互相替代：入离职差额处理未在岗天数，薪资递增处理跨入新薪资周期。先根据 offer 或已确认的历史递增记录确定本月薪资，再用该薪资计算入离职差额。
- 已经由人资确认过的历史新人递增方案要沿用确认结果，不得因本月缺少历史 offer 而恢复旧薪资。整月已生效的递增额进入 `薪资`，本月月中生效的当期补差进入 `其他补贴`。
- 每名员工的工资基数必须回查该员工本月工资底稿、本人 offer 或本人已确认的历史调薪记录；不得从其他员工案例反推，也不得把一名员工的人工确认套到同名规则下的其他员工。人工确认必须与姓名绑定。
- 存在按月递增约定时，月初前已经完整生效的递增档位并入 `薪资`；本月中途生效的档位按本月工作日比例计入 `其他补贴`。例如员工每周期递增 200 元，本月月初已有一档完整生效，则先把 200 元并入本月薪资；本月中途又进入下一档时，只将下一档在本月生效工作日对应的比例放入其他补贴。每个人的初始薪资、入职日、递增额和周期必须回查本人资料，不能套用其他员工的金额或案例。
- 如果 offer 写试用或实习薪资 4500，转正薪资 6000，则月增额 = `(转正薪资 - 初始薪资) / 5`。
- 第一个周期从入职日开始，到下个月同日前一天结束。
- 例如 2026-06-15 入职，则 2026-06-15 至 2026-07-14 用初始薪资，2026-07-15 起进入第二周期。
- 如果递增从月中开始，只按本月递增生效期间的应出勤天数折算。
- 例如 7 月应出勤 23 天，月增额 200，7 月 15 日至 7 月 31 日有 13 个应出勤日，则本月补贴增额 = `200 / 23 * 13 = 113.04`。
- 不要因为员工跨入第二个月，就直接把整月增额全部加上。

入离职差额：

- 入职或离职导致没有上满整月时，要计算入离职差额。
- 使用确认版考勤里的应出勤天数或应出勤小时数。
- 例如 7 月 6 日入职，差额 = 固定薪资 / 当月应出勤天数 × 入职前未出勤天数，通常作为负数扣减。
- 符号要符合工资表原有口径。
- 入离职差额要保留公式，不要直接写死结果，方便人资后续校准。

事假扣款：

- 一般情况下，事假只在考勤汇总中累计；在职非实习员工的事假不显示在工资表，也不在当月工资中扣。
- 只有实习生和当月离职员工的事假进入工资表并计算扣款；试用期全职员工不等于实习生，不能仅因尚未转正就显示或扣除事假。其他特殊清算必须有当月人资明确确认。
- 如果不扣款，不要把考勤里的事假小时数同步到工资表扣款字段。
- 是否自动填写事假时间与是否保留计算公式是两件事：普通在职非实习员工的 `事假时间` 默认留空，但 `请假扣款` 仍必须保留本行事假计算公式，供人资遇到请假过多、需要月度结算等特殊情况时直接填写并自动计算。
- 所有员工的 `事假时间` 都是可编辑数字输入；修改后，`请假扣款`、应发工资和实发金额必须自动联动更新。
- 实习生事假扣款分母使用本月应出勤小时数；这个小时数每月都要按当月工作日重新计算，不要写成固定 168。实习生当月离职时仍按实习生口径，不切换为 174。
- 所有非实习员工的事假公式分母使用年平均小时数 `174`；当月离职员工自动填写累计事假，普通在职员工默认不填，但人工补填后也按 `174` 自动结算。累计事假已包含本月事假时，不要把本月事假再次相加。

病假扣款：

- 全年病假不超过 9 天：病假扣款 = 固定薪资 / 应出勤小时数 × 病假小时数 / 2。
- 全年病假超过 9 天：病假扣款 = `(固定薪资 - 福州最低工资 × 80%) / 应出勤小时数 × 病假小时数`。
- 如果缺最低工资或全年病假累计数据，要写异常并请人工确认。

### 阶段 7：绩效

绩效来源可能变化，不要把单一表格格式写死。

绩效可能来自：

- 问卷。
- 组长或负责人评价。
- 部门负责人确认。
- 人资维护的绩效表。
- 当月临时绩效方案。

处理原则：

- 读取当月用户提供的绩效来源。
- 按姓名或员工 ID 匹配到绩效金额。
- 对缺失、重复、无法匹配的绩效写异常。
- 如果用户确认某个人资或部门没有绩效，则保持空白或按工资表习惯填 0。

### 阶段 8：工作日加班次日准时到奖励金

奖励金不能只靠原始考勤自动挖。

必须以奖励金问卷作为申报入口，再校验：

- 员工是否申报了该日期。
- 实际下班时间是否达到申报档位。
- 次日是否工作日。
- 次日是否准时到岗。
- 12 点后档位是否有加班申请。

奖励档位：

- 晚上 10 点后下班，次日 9 点前到：奖励金 50 元，可报销上班打车，不需要加班申请。
- 晚上 10 点后下班，次日 9:30 前到：奖励金 0 元，可报销上班打车，不需要加班申请。
- 晚上 12 点后下班，次日 9 点前到：奖励金 200 元，可报销上班打车，需要加班申请。
- 晚上 12 点后下班，次日 9:30 前到：奖励金 0 元，可报销上班打车，需要加班申请。

注意事项：

- 有人晚上回公司拿东西，也会刷脸导致下班时间被重置，所以不能只看考勤自动发奖励。
- 周五晚上或次日为休息日、节假日时，不发次日准时到奖励。
- 12 点后奖励必须匹配加班申请。
- 如果加班申请结束时间写成同一天 `00:00`，但开始时间是晚上，要理解为次日 `00:00`。
- 匹配加班申请时，优先按员工和加班开始日期匹配。
- 如果申请证明到 0 点，不要因为实际打卡 `00:02` 比申请 `00:00` 晚几分钟就判定失败。
- 同人同日重复提交问卷，要去重，只保留一条可计奖记录，并在明细里说明。

### 阶段 9：个税模板

必须严格使用政府个税模板。

规则：

- 必须以用户提供的政府原始模板为母版直接复制并填充，不得自行新建或重新设计表格。
- 模板的工作表、字段名、列顺序、列数量、列宽、行高、字体、边框、底色、数值格式和填表说明都不能改；政府系统依赖完整模板结构识别文件。
- 所有列都必须保留。默认只填写模板中标为黄色的列，非黄色列即使有数据来源也保持空白，不得删除。
- 当前个税模板黄色列为：`姓名`、`证照类型`、`证照号码`、`本期收入`、`基本养老保险费`、`基本医疗保险费`、`失业保险费`、`住房公积金`。生成前应以模板实际底色再次识别和校验，不能只靠固定列号。
- `工号`列是经人工确认的唯一非黄色例外：列名仍必须保持`工号`，但单元格填写最新云端花名册中的`所属公司`，供三位人资筛选分类；不得把表头改成`公司`。
- `工号`列的公司必须每月重新读取云端花名册，不得使用工资表公司字段替代。花名册所属公司为空时，该单元格留空并写入异常清单，不得猜测或静默补值。
- 不要额外增加公司列、人资列等模板没有的字段。
- 没有的数据可以空着。
- 需要填的核心字段包括：姓名、证件类型、证照号码、本期收入、基本养老保险、基本医疗保险费、失业保险费、住房公积金。
- 人员范围来自最终工资表。
- 身份证信息优先从云端花名册、本地花名册等来源补充。
- 如果缺身份证，写异常，不要编造。
- 按人资生成时，一个人资一份 Excel。

如果个税模板缺人，先检查最终工资表里有没有这个人。

- 如果最终工资表没有，根因是工资表人员范围问题，不是个税导出问题。
- 如果最终工资表有但个税模板没有，再排查个税生成逻辑。

### 阶段 10：飞书工资单导入

从最终工资表生成飞书工资单导入表。

模板规则：

- 保留飞书工资单模板原始表头和字段。
- 不要新增人资负责人、公司、审计字段到导入 sheet。
- 姓名来自最终工资表。
- 手机号来自云端花名册，必要时用本地资料辅助补充。
- 其他字段直接从工资表映射。
- 不重新计算工资字段，除非本流程明确规定。

`薪酬总计` 规则：

- `薪酬总计` = 飞书工资单模板中从 `薪资` 到 `奖励金` 之间所有列的合计。
- 不是简单的 `薪资 + 奖励金`。

飞书可发送人员筛选：

- 最终工资表要保留当月有工资的人，即使这个人月底已经离职。
- 但飞书工资单导入表只保留能收到飞书工资单的人。
- 工资单在公司总账号的飞书中发放，因此公司总账号通讯录是发送范围的权威来源。员工即使属于公司，只要使用其他飞书组织或不在该总账号通讯录中，就不得进入工资单导入表。
- 纳入工资单的人员必须在公司总账号通讯录中匹配到唯一且已激活的账号。查询不到、账号未激活、跨租户或存在多个同名活跃账号时都不得自动纳入；同名歧义必须进入复核清单，等待人工确认具体账号。
- 云端花名册离职日期小于或等于发薪日期的人，要从飞书工资单导入表排除。
- 即使云端花名册没有离职日期，如果飞书通讯录里查不到可用账号，也要从飞书工资单导入表排除。
- 这类人员可能是异地员工、不在福州线下打卡、不需要飞书账号，或团队未开通飞书付费账号。
- 使用 `lark-cli contact +search-user` 通过用户身份查询飞书账号。
- 每次生成工资单前都要重新检查飞书 CLI 当前登录身份和公司总账号通讯录，不能沿用上月或旧会话的账号匹配结果。
- 要输出飞书工资单排除清单，包含姓名、排除原因、离职日期、飞书 open_id、是否激活等信息。

### 阶段 11：最终交付

最终通常要交付：

- 最终工资总表。
- 各人资或各公司拆分工资表。
- 五险一金汇总表。
- 政府个税模板。
- 飞书工资单导入表。
- 飞书工资单排除和复核清单。
- 工资异常清单。

最终汇报必须说明：

- 使用了哪些源文件。
- 最终工资表人数。
- 飞书工资单可发送人数。
- 飞书工资单排除人数及原因。
- 哪些数据仍需人工确认。

## 飞书 CLI 使用要求

需要云端数据时，使用飞书 CLI。

常用命令：

- 检查身份：`lark-cli whoami`
- 检查授权：`lark-cli auth status --json --verify`
- 读取云端花名册元信息：`lark-cli sheets +workbook-info --url <url>`
- 读取云端花名册数据：`lark-cli sheets +csv-get --url <url> --sheet-name 员工数据 --range A1:T200`
- 搜索飞书用户：`lark-cli contact +search-user --queries "张三,李四" --as user --json`
- 读取审批数据：使用 `lark-approval` 或 `lark-cli approval`

如果飞书 CLI 未登录，要暂停并要求用户完成 `lark-cli auth login`。当用户明确要求使用云端花名册或飞书账号状态时，不要退回到过期的本地花名册。

## 异常处理纪律

不能静默猜测以下内容：

- 薪资标准。
- 公司主体。
- offer 薪资拆分。
- 身份证号码。
- 飞书工资单所需手机号。
- 调薪审批是否通过。
- 人工补贴或扣款。
- 没有飞书账号的人是否还需要通过其他渠道发工资单。

遇到缺失资料时，要写异常，并向用户索要具体来源。

推荐异常字段：

- 阶段。
- 人资负责人。
- 公司。
- 部门。
- 员工 ID / open_id。
- 姓名。
- 异常类型。
- 当前处理方式。
- 数据来源。
- 详细说明。
- 是否需要人工复核。

## 后续更新规则

用户确认稳定工资规则后，应更新本 skill 或相关专项 skill，不要只依赖对话记忆。

如果规则只是某个月临时适用，应写在当月运行说明或异常清单里，不要永久写入 skill。
```

# attendance-calculator

### attendance-calculator/agents/openai.yaml

```yaml
interface:
  display_name: "考勤核算助手"
  short_description: "汇总飞书考勤、假期、加班并生成综合表"
  default_prompt: "使用 attendance-calculator 处理月度飞书考勤、花名册、加班和育儿假数据，生成可供 HR 复核的考勤汇总表与异常说明。"
```

### attendance-calculator/references/business-rules.md

```markdown
# 考勤业务规则

以下规则用于当前版本的月度考勤初算与复核。

## 已确认的生产规则

- 最终汇总行按“部门拼音，再按姓名拼音”排序。
- 花名册是输出资格的主名单。飞书考勤里存在、但花名册无法匹配的员工，不进入汇总表，并在异常清单中记录“花名册未匹配，已从汇总表排除”。
- 月度应上班总小时数按当月实际应出勤工作日动态计算，即“当月应出勤天数 × 8 小时”，优先使用飞书日明细里的 `应出勤天数` 汇总；如果没有该字段，再用 `应出勤时长(分钟)` 除以 60。该值每月可能不同，不能把某个月的 `168` 小时写成固定常量。
- 个人请假、病假等按自然日先向上取整，再汇总月度；单日最多按 8 小时。事假、病假、缺勤等小时制项目不保留 0.5 小时单位，任何零头都按 1 小时计。
- 累计加班结余按滚动月计算：上月考勤汇总里的累计加班结余，加上本月加班净变动，得到本月累计加班结余。
- 累计事假时长按滚动月计算：本月事假时数，加上加班表中带来的上月请假结余，得到本月累计事假时长。不要再用“上月累计事假 + 本月事假”去覆盖这个口径，也不要自动用本月加班去抵扣事假。
- 育儿假要保留上月已有记录；本月有使用时要补充说明；花名册里出现新的孩子信息时，也要同步更新说明。育儿假按自然月口径滚动，到月末截止，不按具体出生日期卡点。每个周期 10 天，每个自然月最多 1 天。若该员工既没有上月育儿假记录、当前月也没有使用记录，只是花名册里有较早的孩子信息，则不要凭空新建“已使用 0 天”的说明。
- `考勤情况` 只写迟到、事假、病假、年假等日明细，不写育儿假。`缺卡情况` 只写缺卡明细。早退先单独统计展示，不自动扣款，除非 HR 后续明确同一扣款口径。
- 新员工入职当天的上午到岗异常放宽处理：若考勤日期与入职日期相同，则当天迟到和上班缺卡不计入主表汇总、不计扣款、不写入考勤或缺卡明细；当天离开公司时的下班缺卡仍照常统计。
- 员工离职后，离职日期之后的迟到、缺卡、请假、育儿假等考勤记录都不进主表汇总，也不进扣款和明细；离职当天按正常规则处理，除非 HR 另有说明。
- 花名册是工资和考勤人员范围的主名单。工资表或考勤源数据中存在、但最新花名册中不存在的人员，按历史离职处理：从本月工资表或考勤汇总中剔除，不保留空工资行；同时在异常清单中保留“花名册不存在，视为历史离职”的追溯记录。
- 输出表里的时长、次数、扣款字段必须保持数字格式，不能写成文本。`迟到扣款情况`、`漏打卡扣款情况`、`考勤扣款情况` 优先保留公式，方便人工校准后自动联动。

## 全局假设

- 一次只算一个自然月。
- 默认每人 8 小时 = 1 个工作日；当月应上班小时数必须根据当月工作日和国家法定节假日重新确定，不得沿用上月固定值。
- 版本 1 不做人为工龄折算，也不做人为离职比例分摊。
- 能用员工 ID 就优先用员工 ID；如果没有可靠 ID，再退回姓名，并标记可能重名。
- 不在本技能里计算整薪、奖励金、病假工资、事假工资扣款、离职日期折算。

## 考勤类型

员工在飞书考勤里的类型通常分为：

- `需要打卡`：计算迟到、缺卡、事假、病假、年假等。
- `不需要打卡`：迟到、缺卡不计算为扣款，相关数值列输出为 0；版本 1 不记录事假和病假，但仍可保留年假和育儿假。

## 职级

从花名册读取职级：

- `组员`：迟到扣款按 `迟到分钟数 × -3`。
- `管理层`：迟到扣款按 `迟到分钟数 × -5`。

扣款统一输出为负数，和现有 HR 汇总表保持一致。

如果职级缺失或无法识别，迟到扣款留空，并写入异常说明。

## 迟到

- 使用整月迟到分钟数。
- 没有迟到宽限。
- `需要打卡` 的员工按迟到分钟全量计入。
- 迟到发生在入职当天的，忽略，不进主表汇总，也不进扣款和考勤明细，但要保留在可追溯的排除说明里。
- 入职当天上午没有实际工作，上午迟到和上班缺卡放宽；入职当天晚上离开公司的下班打卡仍有严格要求，下班缺卡不能按入职当天放宽。

## 缺卡

- 每月前 2 次缺卡免费。
- 免费次数每个月重置。
- 第 3 次起，每次缺卡扣 `50` 元。
- 公式：`max(0, 缺卡次数 - 2) × -50`。

缺卡扣款统一输出为负数。

- 上班缺卡发生在入职当天的，忽略，不进主表汇总，也不进扣款和缺卡明细。
- 入职当天的下班缺卡不放宽，仍按正常缺卡规则统计和扣款。

## 早退

- 原始数据里如果有早退次数或早退分钟数，先统计出来。
- 当前版本只做展示和留档，不自动扣款。
- 若 HR 后续确认早退要沿用迟到口径，再单独补扣款规则。

## 事假

- `本月事假时数`：按本月日明细汇总，先按自然日向上取整，再按 8 小时封顶。
- `累计事假时长`：`本月事假时数 + 加班表中的上月请假结余`。
- 不把本月加班自动拿来抵扣事假，人工修正优先。

工资衔接时，实习生事假扣款分母使用本月动态应上班小时数；例如 2026 年 8 月为 168 小时，但其他月份应按当月工作日和节假日重新计算，不能固定写 168。

如果加班表中的结余本身是负值，按原值带入，并在异常说明里保留来源。

## 病假

对 `需要打卡` 的员工：

- 记录本月病假时数。
- 记录年度累计病假时数。
- 不计算病假工资。
- 病假按自然日先向上取整，再按 8 小时封顶后汇总。

对 `不需要打卡` 的员工：

- 版本 1 不记录病假。

## 年假

记录 3 个值：

- 当年年假应得天数。
- 本月使用年假天数。
- 年假剩余天数。

年假单位：

- 最小单位为 0.5 天。
- 应得天数向上取整到 0.5 天。
- 若计算结果大于 0 但小于 0.5，则按 0.5 天处理。

按工龄划分：

- 不满 1 年：0 天。
- 满 1 年不满 3 年：3 天。
- 满 3 年不满 5 年：5 天。
- 满 5 年：7 天。

年假重置规则：

- 每年 1 月 1 日按当年工龄判断该年的年假档位。
- 同一年内通常不在周年日再次升级，除非是满 1 年后的首个可享受周期。

满 1 年后的首年规则：

- 员工在当年满 1 年时，周年月之前仍为 0 天。
- 从周年月开始，按周年月而不是精确日期计算首年年假。
- 公式：`(12 - 周年月份) / 12 × 3`。
- 例如入职日是 `2025-08-xx`，2026 年周年月是 8 月，则首年可得 `1` 天。

## 育儿假

- 每个孩子按 10 天/周期计算。
- 周期按自然月滚动，不按精确日期卡点。
- 出生月整月计入；到期按月末截止。
- 例如孩子 2024 年 7 月出生，第一周期可按 2024 年 7 月到 2025 年 7 月月末计算；下一周期从 2025 年 8 月开始。
- 每个自然月最多 1 天。
- 如果上一周期没休完，到新周期时清零，不结转。
- 花名册新增孩子信息时，要同步更新育儿假说明。
- 花名册新增孩子信息只是资格和周期的更新来源，仍需结合本月及上月使用记录判断是否可用，不能仅凭花名册直接写成已使用或已用完。

## 考勤情况文本

- `考勤情况` 只放迟到、事假、病假、年假等日明细。
- 不放育儿假。
- 早退默认不放进这里，除非 HR 后续明确要并入某个扣款口径。
- `缺卡情况` 只放缺卡明细。
- 没有明细时留空。

示例：

- `迟到：27日迟到1分钟，21日迟到1分钟`
- `事假：27日事假3小时`
- `缺卡：21日缺卡1次，20日缺卡1次，6日缺卡1次，15日缺卡1次`

## 异常处理

- 任何字段缺失、主键冲突、部门不一致、飞书读取失败、离职后记录仍然出现、育儿假资格不清楚，都不要乱填。
- 先写异常说明，再继续后续计算。
- 发现无法可靠判断的项，优先让 HR 复核，不要自动猜。
```

### attendance-calculator/references/input-output-schema.md

```markdown
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
```

### attendance-calculator/scripts/process_attendance.py

```python
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
    "部门",
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
    "异常提示",
]


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
    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        return []
    headers = [str(value).strip() if value is not None else "" for value in rows[0]]
    records: List[Dict[str, Any]] = []
    for values in rows[1:]:
        row = {headers[index]: values[index] if index < len(values) else "" for index in range(len(headers)) if headers[index]}
        if any(str(value).strip() for value in row.values() if value is not None):
            records.append(normalize_row(row))
    return records


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
            writer.writerows(rows)
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
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    return float(match.group(0)) if match else default


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
    employee_id = str(row.get("employee_id", "") or "").strip()
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
    if one_year.year == year and (month_start is None or month_start >= one_year):
        prorated = (12 - one_year.month) / 12 * 3
        return ceil_to_half(prorated)
    return 0.0


def split_birthdates(value: Any) -> List[date]:
    if value is None:
        return []
    parts = re.split(r"[,，;；/\n]+", str(value))
    dates = [parse_date(part.strip()) for part in parts if part.strip()]
    return [item for item in dates if item]


def child_year_window_for_month(birthdate: date, month_start: date) -> Optional[Tuple[date, date]]:
    third_birthday = add_years(birthdate, 3)
    if month_start > third_birthday:
        return None
    anchor_year = month_start.year
    anchor = birthdate.replace(year=anchor_year)
    if month_start < anchor:
        anchor = birthdate.replace(year=anchor_year - 1)
    end = add_years(anchor, 1)
    if anchor <= third_birthday and month_start <= third_birthday:
        return anchor, min(end, third_birthday)
    return None


def childcare_note(childcare_days: float, birthdates: List[date], month_start: date, prior_child_year_days: float) -> str:
    if childcare_days <= 0:
        return ""
    active_windows = [window for birthdate in birthdates if (window := child_year_window_for_month(birthdate, month_start))]
    notes: List[str] = [f"{clean_number(childcare_days)}天"]
    if not active_windows:
        notes.append("需复核：未找到3岁内子女")
    if childcare_days > 1:
        notes.append("需复核：本月超过1天")
    if prior_child_year_days + childcare_days > 10:
        notes.append("需复核：出生周年内超过10天")
    return "；".join(notes)


def clean_number(value: float) -> str:
    if abs(value - round(value)) < 0.00001:
        return str(int(round(value)))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def detail_text(prefix: str, value: Any) -> str:
    text = str(value or "").strip()
    return f"{prefix}：{text}" if text else ""


def join_attendance_detail(late_detail: Any, personal_leave_detail: Any) -> str:
    parts = [detail_text("迟到", late_detail), detail_text("事假", personal_leave_detail)]
    return "\n".join(part for part in parts if part)


def output_columns(year: int) -> List[str]:
    yy = f"{year % 100:02d}"
    return [column.format(yy=yy) for column in OUTPUT_COLUMNS_TEMPLATE]


def build_summary(args: argparse.Namespace) -> Tuple[List[Dict[str, Any]], List[str]]:
    month = datetime.strptime(args.month, "%Y-%m").date()
    year = month.year
    columns = output_columns(year)

    attendance_rows = read_table(args.attendance)
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

        attendance_type = str(attendance.get("attendance_type", "") or "").strip()
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
        missed_deduction = max(0, missed_punch_count - 2) * -50 if needs_punch else 0
        total_deduction = (late_deduction if isinstance(late_deduction, int) else 0) + missed_deduction

        overtime_balance = parse_number(overtime.get("overtime_balance_hours"))
        overtime_to_personal_leave = abs(overtime_balance) if overtime_balance < 0 else 0.0

        monthly_personal_leave = parse_number(attendance.get("personal_leave_hours")) if needs_punch else 0.0
        monthly_personal_leave += overtime_to_personal_leave
        prior_personal_leave = parse_number(history.get("prior_personal_leave_hours"))
        cumulative_personal_leave = prior_personal_leave + monthly_personal_leave

        monthly_sick_leave = parse_number(attendance.get("sick_leave_hours")) if needs_punch else 0.0
        prior_sick_leave = parse_number(history.get("prior_sick_leave_hours_year"))
        ytd_sick_leave = prior_sick_leave + monthly_sick_leave

        hire_date = parse_date(roster.get("hire_date"))
        annual_entitlement = annual_leave_entitlement(hire_date, year, month)
        annual_used = parse_number(attendance.get("annual_leave_days"))
        prior_annual_remaining = parse_number(history.get("prior_annual_leave_remaining_days"), annual_entitlement)
        annual_remaining = max(0.0, prior_annual_remaining - annual_used)

        childcare_days = parse_number(attendance.get("childcare_leave_days"))
        birthdates = split_birthdates(roster.get("child_birthdates"))
        prior_childcare = parse_number(history.get("childcare_leave_days_in_child_year"))
        childcare = childcare_note(childcare_days, birthdates, month, prior_childcare)

        department = str(attendance.get("department", "") or roster.get("department", "") or "").strip()
        missed_detail = detail_text("缺卡", attendance.get("missed_punch_detail"))
        attendance_detail = join_attendance_detail(attendance.get("late_detail"), attendance.get("personal_leave_detail"))
        if overtime_to_personal_leave:
            addition = f"加班结余为负转事假{clean_number(overtime_to_personal_leave)}小时"
            attendance_detail = "\n".join(part for part in [attendance_detail, f"事假：{addition}"] if part)

        row = {
            "部门": department,
            "职级": rank,
            "姓名": str(attendance.get("name", "") or roster.get("name", "") or "").strip(),
            "迟到分钟数（分钟）": late_minutes if needs_punch else "",
            "漏打卡次数": missed_punch_count if needs_punch else "",
            "迟到扣款情况": late_deduction,
            "漏打卡扣款情况": missed_deduction if needs_punch else "",
            "考勤扣款情况（合计）": total_deduction if needs_punch else "",
            "本月事假时数（小时）": clean_number(monthly_personal_leave) if needs_punch or overtime_to_personal_leave else "",
            "累计事假时长（小时）": clean_number(cumulative_personal_leave) if needs_punch or overtime_to_personal_leave else "",
            "累计加班结余（小时）": clean_number(overtime_balance),
            "病假（小时）": clean_number(monthly_sick_leave) if needs_punch else "",
            "育儿假": childcare,
            "奖励金": "",
            "本月应上班总小时数（h）": clean_number(parse_number(attendance.get("required_work_hours"), args.paid_days * args.hours_per_day)),
            "入职时间": format_date(hire_date),
            "离职时间": "",
            "上月病假时长": clean_number(prior_sick_leave),
            "病假天数合计(年/小时)": clean_number(ytd_sick_leave),
            columns[19]: clean_number(annual_entitlement),
            "上月年假剩余天数": clean_number(prior_annual_remaining),
            "本月使用年假（天）": clean_number(annual_used),
            columns[22]: clean_number(annual_remaining),
            "考勤情况": attendance_detail,
            "缺卡情况": missed_detail,
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
    rows, columns = build_summary(args)
    write_table(args.output, rows, columns)
    print(f"Wrote {len(rows)} rows to {args.output}")


if __name__ == "__main__":
    main()
```

### attendance-calculator/SKILL.md

```markdown
---
name: attendance-calculator
description: 处理和审计公司月度考勤汇总，涵盖花名册、考勤原始数据、上月汇总、加班结余、育儿假、年假、事假、病假、迟到、缺卡、早退、离职过滤与异常说明。
---

# 考勤核算

## 作用

用于把月度飞书考勤导出、花名册、上月汇总、加班表和育儿假数据整理成一份可交 HR 复核的考勤初算表。

本技能只负责考勤统计、异常归因和扣款口径测算，不负责整薪、奖励金、病假工资、事假工资或最终发薪。

## 温馨提示流程

1. 先确认本次计算月份，例如 `2026-08`。
2. 先提醒用户准备这些材料：
   - 本月原始考勤数据。
   - 上月考勤汇总表。
   - 最新花名册。
   - 加班结余表。
   - 育儿假或相关更新说明。
   - 如有人工校准结果，也一并提供。
   - 本月应上班小时数会随节假日变化，不是固定常数；如果用户提到 `168`，要理解为某个月的示例值。
3. 先读 `references/input-output-schema.md`，再读 `references/business-rules.md`。
4. 先把源数据整理成脚本需要的标准字段，再开始计算。
5. 先输出考勤初算版，再输出异常清单。
6. 如果发现字段缺失、主键冲突、离职后数据、早退未定义等情况，先提示用户，不要乱写。

## 计算范围

自动计算这些内容：

- 迟到分钟数和迟到扣款。
- 漏打卡次数和漏打卡扣款。
- 考勤扣款合计。
- 本月事假时数。
- 累计事假时长。
- 累计加班结余。
- 病假时数。
- 病假年度累计。
- 育儿假使用情况和限制提示。
- 年假应得、当月使用和剩余。
- 考勤情况文本。
- 缺卡情况文本。
- 早退次数和早退时长，仅统计展示，不自动扣款，除非用户后续明确规则。
- 独立异常清单，用于记录匹配失败、字段缺失、离职过滤和人工复核事项。

以下内容先保留人工值或留空，除非用户明确新增规则：

- 奖励金。
- 整体工资测算。
- 病假工资。
- 事假工资扣款。
- 离职日期工资折算。

## 处理顺序

1. 确认月份。
2. 读取花名册、上月汇总、本月考勤、加班表和育儿假来源。
3. 先做身份匹配，再做规则过滤，再做汇总计算。
4. 先过滤离职后数据、入职当天放宽项和不应计入的异常。
5. 再生成考勤初算版和异常清单。
6. 最后保留公式、数字格式和可追溯说明。

## 输出要求

- 迟到扣款情况、漏打卡扣款情况、考勤扣款情况尽量保留 Excel 公式。
- 计时字段必须写成数字，不要写成文本。
- 任何被过滤的记录要单独留下说明，不能直接抹掉。
- 同一员工如果有重名、主键异常或飞书读取异常，要先提示，不要硬填。
```

# attendance-settlement

### attendance-settlement/agents/openai.yaml

```yaml
interface:
  display_name: "Attendance Settlement"
  short_description: "Prepare attendance summaries for HR settlement and payroll handoff."
  default_prompt: "Use $attendance-settlement to prepare this monthly attendance summary for HR settlement and flag fields that need manual confirmation."

policy:
  allow_implicit_invocation: true
```

### attendance-settlement/SKILL.md

````markdown
---
name: attendance-settlement
description: Prepare, audit, and explain monthly HR attendance settlement outputs after attendance statistics are calculated. Use when Codex needs to turn attendance summary data into a settlement-ready workbook or checklist, reconcile attendance deductions with payroll handoff fields, review reward/payment/sick-pay/personal-leave deduction fields, preserve manual salary-sensitive values, and produce exception notes for HR confirmation.
---

# Attendance Settlement

## Purpose

Use this skill after monthly attendance statistics have already been calculated. It prepares the attendance result for HR settlement or payroll handoff, checks whether salary-sensitive fields are complete, and separates confirmed automated values from values that require HR approval.

Use `attendance-calculator` first when the user provides raw Feishu/Lark attendance exports, rosters, leave records, childcare data, overtime balances, or asks to compute attendance totals from source files.

## Settlement Workflow

1. Confirm the settlement month in `YYYY-MM`.
2. Collect the current attendance summary workbook and any payroll or settlement template the user wants to fill.
3. Identify whether the workbook was generated by `attendance-calculator` or is an HR-maintained summary.
4. Preserve all existing manual salary-sensitive values unless the user gives explicit rules and asks to recalculate them.
5. Reconcile automated attendance values against settlement fields:
   - Late-arrival deduction.
   - Missed-punch deduction.
   - Total attendance deduction.
   - Personal leave hours.
   - Sick leave hours.
   - Annual leave used and remaining balance.
   - Overtime balance after negative-balance conversion.
   - Required monthly work hours.
6. Flag missing, inconsistent, or policy-dependent fields in an exception report instead of guessing.
7. Produce a settlement-ready workbook or a clear checklist of fields that HR must confirm before payroll use.

## Scope Boundary

Calculate or verify these values when the input summary already contains enough data:

- Attendance deduction total as late-arrival deduction plus missed-punch deduction.
- Whether negative overtime balance has already been converted to personal leave.
- Whether personal leave, sick leave, annual leave, and childcare leave are internally consistent with detail text.
- Whether employees in the settlement file match the attendance summary by employee ID or unique name.
- Whether required work hours are present for every active employee.

Do not invent or calculate these values unless the user provides exact policy rules:

- Full salary.
- Reward payment.
- Sick-pay amount.
- Personal-leave salary deduction.
- Tax, social insurance, housing fund, or net pay.
- Resignation proration.
- Any disciplinary deduction not present in the attendance summary.

When a value is missing and no rule is provided, keep it blank or preserve the existing value, then add an exception note.

## Required Inputs

Minimum useful inputs:

- Settlement month.
- Attendance summary workbook or CSV.
- Payroll/settlement template, if the user needs a specific output format.

Recommended inputs:

- Employee roster with employee ID, name, department, hire date, and resignation date if applicable.
- Previous month summary when cumulative leave or annual leave balances must be checked.
- Written HR policy for salary-sensitive calculations if the user wants monetary settlement beyond attendance deductions.

## Reconciliation Rules

Use employee ID as the primary key when available. If employee ID is missing, match by name only when names are unique and add an exception for name-only matching.

Treat the attendance summary as the source of truth for attendance-derived values. Treat the payroll or settlement template as the target format unless the user says the template contains authoritative manual overrides.

Preserve manual fields by default. If the same employee has different values in the attendance summary and settlement template:

- Use the attendance summary for automated attendance-derived fields.
- Use the settlement template for manual salary-sensitive fields.
- Add an exception explaining the conflict and the value chosen.

Attendance deduction total should equal:

```text
late_arrival_deduction + missed_punch_deduction
```

Values are normally negative numbers when they represent deductions. If a target workbook expects positive deduction amounts, convert signs only after confirming the target convention from existing rows or user instruction.

## Exception Report

Create or update an exception report whenever any value cannot be safely settled. Use one row per issue.

Recommended columns:

- `employee_id`
- `name`
- `department`
- `exception_type`
- `source_field`
- `summary_value`
- `settlement_value`
- `chosen_value`
- `current_action`
- `manual_review_required`
- `detail`

Common exception types:

- `missing_employee_id`
- `name_only_match`
- `duplicate_name`
- `employee_missing_from_summary`
- `employee_missing_from_settlement_template`
- `deduction_total_mismatch`
- `manual_salary_field_preserved`
- `salary_rule_missing`
- `required_work_hours_missing`
- `leave_balance_mismatch`
- `sign_convention_unclear`

## Output Guidance

When creating a settlement workbook:

- Keep the user's original workbook intact and write a new output file.
- Preserve sheet names, column order, formulas, and formatting when practical.
- Add a timestamp or month to the output filename.
- Include a separate exception sheet or separate exception file.
- Keep calculation notes concise and auditable.

When no target template is provided:

- Produce a normalized settlement table with one row per employee.
- Include only confirmed automated attendance fields and preserved manual fields.
- Add clear placeholders for fields that require HR policy confirmation.

## Iteration Guidance

If the user later provides stable salary calculation rules or a recurring settlement template, add a script under `scripts/` to perform the deterministic workbook transformation. Until then, prefer explicit reconciliation steps and exception reporting over hidden assumptions.
````

# social-insurance-fund

### social-insurance-fund/agents/openai.yaml

```yaml
interface:
  display_name: "五险一金汇总"
  short_description: "按月并表五险一金并输出工资可用汇总表。"
  default_prompt: "使用 $social-insurance-fund 将本月养老、失业、医保、公积金及广思专项五险文件并入工资可用的五险一金汇总表。"

policy:
  allow_implicit_invocation: true
```

### social-insurance-fund/scripts/build_social_insurance_fund.py

```python
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
```

### social-insurance-fund/SKILL.md

````markdown
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
````

# 安装与发布文件

### README.md

````markdown
# 工资核算 Skill 套件

这是一套用于公司月度工资核算的 Codex Skills，覆盖考勤初算、人工确认、五险一金并表、工资计算、个税模板和飞书工资单生成。

## 一键安装

适用于 Windows 10 或 Windows 11。

1. 在电脑左下角搜索并打开 **PowerShell**。
2. 复制下面整条命令，粘贴后按回车。
3. 看到“安装成功”后，完全退出并重新启动 Codex。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -Command "$z=Join-Path $env:TEMP 'payroll-skill-suite.zip'; $d=Join-Path $env:TEMP 'payroll-skill-suite-online'; Invoke-WebRequest 'https://github.com/sdasdasd1223/payroll-skill-suite/archive/refs/heads/main.zip' -OutFile $z; if(Test-Path -LiteralPath $d){Remove-Item -LiteralPath $d -Recurse -Force}; Expand-Archive -LiteralPath $z -DestinationPath $d -Force; $i=Get-ChildItem -LiteralPath $d -Filter install.ps1 -Recurse | Select-Object -First 1; if(-not $i){throw '下载包中没有找到 install.ps1'}; & $i.FullName"
```

以后更新到最新版时，重新运行同一条命令即可。安装程序会先把电脑上的旧 Skill 备份到 `.codex\skill-backups`，再安装新版本。

## 安装内容

- `payroll-operations`：工资发放总流程、材料引导和人工确认节点。
- `attendance-calculator`：月度考勤计算与异常审计。
- `attendance-settlement`：考勤确认结果与工资结算衔接。
- `social-insurance-fund`：多公司五险一金汇总。

## 开始使用

重启 Codex 后输入：

> 请使用 payroll-operations，开始核算 YYYY 年 MM 月工资，并按阶段引导我准备资料。

Codex 会先引导考勤阶段，不会一次性要求提供全部工资资料。

## 文档

- [三位人资部署与月度操作教程](docs/三位人资工资核算部署与月度操作教程.docx)
- [工资核算 Skill 完整内容](工资核算Skill完整内容.md)
- [在线安装命令](在线安装命令模板.txt)
- [发布说明](发布说明.md)

## 更新纪律

通用规则进入正式 Skill；员工姓名、固定金额、某个月工时和一次性例外只进入当月运行说明，避免影响以后月份。已人工确认的工资表、个税表和工资单是当月正式基准，未经明确要求不得重新计算覆盖。
````

### install.ps1

```powershell
[CmdletBinding()]
param(
    [string]$CodexHome = $(if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $HOME ".codex" })
)

$ErrorActionPreference = "Stop"
$skillNames = @(
    "payroll-operations",
    "attendance-calculator",
    "attendance-settlement",
    "social-insurance-fund"
)

function Write-Step([string]$Message) {
    Write-Host "`n==> $Message" -ForegroundColor Cyan
}

function Assert-SkillPackage([string]$SkillPath, [string]$ExpectedName) {
    $skillFile = Join-Path $SkillPath "SKILL.md"
    $agentFile = Join-Path $SkillPath "agents\openai.yaml"
    if (-not (Test-Path -LiteralPath $skillFile -PathType Leaf)) {
        throw "安装包缺少 $ExpectedName\SKILL.md"
    }
    if (-not (Test-Path -LiteralPath $agentFile -PathType Leaf)) {
        throw "安装包缺少 $ExpectedName\agents\openai.yaml"
    }

    $skillText = Get-Content -LiteralPath $skillFile -Raw -Encoding UTF8
    if ($skillText -notmatch "(?m)^name:\s*$([regex]::Escape($ExpectedName))\s*$") {
        throw "$ExpectedName 的 SKILL.md 缺少正确的 name。"
    }
    if ($skillText -notmatch "(?m)^description:\s*.+$") {
        throw "$ExpectedName 的 SKILL.md 缺少 description。"
    }

    $agentText = Get-Content -LiteralPath $agentFile -Raw -Encoding UTF8
    foreach ($field in @("display_name", "short_description", "default_prompt")) {
        if ($agentText -notmatch "(?m)^\s*${field}:\s*.+$") {
            throw "$ExpectedName 的 agents\openai.yaml 缺少 $field。"
        }
    }
}

$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceRoot = Join-Path $packageRoot "skills"
$skillsRoot = Join-Path $CodexHome "skills"
$backupRoot = Join-Path $CodexHome ("skill-backups\payroll-skill-suite_" + (Get-Date -Format "yyyyMMdd_HHmmss"))

Write-Step "检查安装包"
foreach ($name in $skillNames) {
    Assert-SkillPackage -SkillPath (Join-Path $sourceRoot $name) -ExpectedName $name
}

New-Item -ItemType Directory -Path $skillsRoot -Force | Out-Null
$existing = @($skillNames | Where-Object { Test-Path -LiteralPath (Join-Path $skillsRoot $_) })

if ($existing.Count -gt 0) {
    Write-Step "备份电脑上的旧 Skill"
    New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
    foreach ($name in $existing) {
        Copy-Item -LiteralPath (Join-Path $skillsRoot $name) -Destination (Join-Path $backupRoot $name) -Recurse -Force
        Write-Host "已备份：$name"
    }
}

Write-Step "安装工资核算 Skill 套件"
foreach ($name in $skillNames) {
    $destination = Join-Path $skillsRoot $name
    if (Test-Path -LiteralPath $destination) {
        Remove-Item -LiteralPath $destination -Recurse -Force
    }
    Copy-Item -LiteralPath (Join-Path $sourceRoot $name) -Destination $destination -Recurse -Force
    Write-Host "已安装：$name"
}

Write-Step "核对安装结果"
foreach ($name in $skillNames) {
    Assert-SkillPackage -SkillPath (Join-Path $skillsRoot $name) -ExpectedName $name
}

Write-Host "`n安装成功。" -ForegroundColor Green
if ($existing.Count -gt 0) {
    Write-Host "旧版本备份位置：$backupRoot"
}
Write-Host "请完全退出并重新启动 Codex，重启后 Skill 才会生效。" -ForegroundColor Yellow
```

### 发布说明.md

````markdown
# 工资核算 Skill 套件发布说明

## 发布内容

本安装包包含以下四个 Skill：

- `payroll-operations`
- `attendance-calculator`
- `attendance-settlement`
- `social-insurance-fund`

`install.ps1` 会在覆盖前备份电脑上的旧版本，然后安装并检查四个 Skill。

## 在线仓库

本套件已按使用者确认发布到公开仓库：

`https://github.com/sdasdasd1223/payroll-skill-suite`

发布到私有 GitHub 时，仓库根目录应直接包含：

```text
payroll-skill-suite/
├─ install.ps1
├─ skills/
│  ├─ payroll-operations/
│  ├─ attendance-calculator/
│  ├─ attendance-settlement/
│  └─ social-insurance-fund/
└─ 在线安装命令模板.txt
```

最终一键安装命令见 `在线安装命令模板.txt`。发布后必须在独立测试目录执行一次从 GitHub 下载的在线安装测试。

## 版本更新

每次发布前：

1. 更新四个 Skill 的正式文件。
2. 运行各 Skill 的校验。
3. 更新发布版本或 Git 标签。
4. 用一台测试电脑执行在线安装命令。
5. 确认旧版本已备份、四个 Skill 均安装成功，并重启 Codex 验证。
````

### 在线安装命令模板.txt

```text
工资核算 Skill 一键在线安装

适用系统：Windows 10 或 Windows 11

安装步骤：
1. 在电脑左下角搜索 PowerShell 并打开。
2. 复制下面整条命令，粘贴后按回车。
3. 看到“安装成功”后，完全退出并重新启动 Codex。

powershell -NoProfile -ExecutionPolicy Bypass -Command "$z=Join-Path $env:TEMP 'payroll-skill-suite.zip'; $d=Join-Path $env:TEMP 'payroll-skill-suite-online'; Invoke-WebRequest 'https://github.com/sdasdasd1223/payroll-skill-suite/archive/refs/heads/main.zip' -OutFile $z; if(Test-Path -LiteralPath $d){Remove-Item -LiteralPath $d -Recurse -Force}; Expand-Archive -LiteralPath $z -DestinationPath $d -Force; $i=Get-ChildItem -LiteralPath $d -Filter install.ps1 -Recurse | Select-Object -First 1; if(-not $i){throw '下载包中没有找到 install.ps1'}; & $i.FullName"

仓库地址：
https://github.com/sdasdasd1223/payroll-skill-suite

安装脚本会先检查安装包；电脑已有旧版 Skill 时会自动备份，然后安装 payroll-operations、attendance-calculator、attendance-settlement 和 social-insurance-fund。
```

## 维护纪律

- 每月通用规则更新到对应 Skill；个人姓名、固定金额、当月工时和临时例外只放入当月运行说明。
- 更新正式 Skill 后，应重新生成本文档和发布包，再执行校验及测试安装。
- 在线仓库应使用公司私有仓库或受控内部下载地址，不应未经确认公开发布工资规则。

