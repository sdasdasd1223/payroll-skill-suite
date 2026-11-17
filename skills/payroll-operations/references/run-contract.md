# 月度实际文件与跨阶段交接契约

本月私有记录用`schema_version:"crossstage-1"`，放实际资料目录，不进入技能发布包。运行`python scripts/crossstage_audit.py --record 本月记录.json --output 新审计结果.json`；`workflow_audit.py`接收同一契约。程序只读取明确列出的来源，不扫描任意目录、不修改来源、不生成真实确认。

`mode`为`draft/formal`；`purpose`默认`template_delivery`，仅计算明细须显式写`calculation_details`。默认草稿也必须实际核验本次交付阶段的模板、公式和重算，才可`passed:true`；数据待核与技术未完成分别报告。`calculation_passed`只说明计算检查，不能冒充合规模板交付；纯明细始终`formal_ready:false`并显示“仅计算明细”。缺必要资料只使相应项目待核，不因缺加班结余清空独立可算的奖励、薪资或保险。正式模式须无错误、无缺项、无未确认前置，否则拒绝正式税表或可发工资单。以下新契约优先，文末旧契约只用于历史诊断，不能正式通过。

## 最小可运行入口

不建议手填整份事实台账。先运行下面的完整合成演示：

```powershell
python scripts/crossstage_scaffold.py --demo --output-dir ./明确合成演示
python scripts/crossstage_audit.py --record ./明确合成演示/run.json --output ./明确合成演示/另存复核.json
```

演示生成实际可读的名单、日明细、奖励、保险、薪资来源、各阶段计算文件及完整台账；全部数据明确合成，没有指定真实模板核验或真实确认，所以脚手架显式生成`purpose:calculation_details`，只交付计算明细，不声称模板合规或正式通过。目录必须尚不存在。它用于学习契约，不得当真实核算资料。

真实冷启动提供实际核定候选名单JSON数组（每人`employee_key,name,company`，可附入离职日期和员工类型），以及按实际原件明确填写的日明细映射；可参照演示生成的`attendance-config.json`，保留完整列及实际范围。执行：

```powershell
python scripts/crossstage_scaffold.py --scope ./本月候选名单.json --month 2026-09 --attendance-config ./实际日明细映射.json --policy-version 已核定规则版本 --output-dir ./本月私有核算草稿
```

映射JSON结构为`{path:实际原件路径,table:本契约表映射,raw_attendance:本契约假别映射}`。脚手架读取实际全范围原件，独立产生能证实的本月假别计算文件和逐行观察；不能证实的项目才保留待核。没有日明细时仍可去掉`--attendance-config`生成固定人员及待核交接，但执行者须继续完成已具备来源的其他独立核算，不能只交一份JSON计划。随后将奖励、保险及工资核算程序的实际产出接入台账和正式模板。脚手架不会猜历史余额、人工确认、工资口径或账号资格。

## 固定名单与实际文件

顶层为`schema_version,month,policy_version,mode,through_stage,scope_table,sources,tables,stages,facts,confirmations,raw_attendance,raw_rewards,raw_insurance,template_bindings`，另按实际来源提供`scope_dispositions`和`payroll_original_source_ids`。

- `through_stage`依次为`attendance/reward/insurance/payroll/tax/payslip`，前面各阶段均须交接。`scope_table`指向实际核定候选名单表，映射`employee_key,name,company`。可靠员工标识优先；无ID来源只按公司+唯一姓名弱匹配，同名冲突不得合并。`expected_employee_keys`只作额外断言，不能代替实际名单。
- `sources[id]={path,sha256,month,role}`；路径相对记录或绝对路径。逐文件重算SHA，来源变化使旧确认失效。角色包括`scope_approved/attendance_confirmed/payroll_confirmed/raw/calendar/approval/manual/calculation`。上月结余等来源按用途核期，原始日期及保险归属期仍按内容核验。
- `tables[id]={source_id,format,sheet,start_row,end_row,columns}`。`xlsx/xls`列为`A/B/...`；`csv`为从1起的列号；`json`为实际字段名，`list_path`可逐层定位数组。Excel/CSV用物理行号，JSON数组位置从1起。xlsx先`reset_dimensions()`再读真实行，不能信任导出文件宣告的A1尺寸；原xls用xlrd只读全范围及日期类型，不转换格式。
- 从`start_row`扫描全部非空映射行，`end_row`仅是断言，不能截断新增人员。`excluded_rows`须逐行说明真实摘要/非人员/范围外依据，不能隐藏候选人。`constants`只填有真实依据的来源上下文，不能覆盖实际字段。缺公司原件使用常量公司时须实际核定范围依据；测试“归属未核”不能转成已核定范围。
- `stages[阶段]={table,dispositions}`；每名候选人每阶段恰好一个`{employee_key,route,row,reason,source_refs}`。`route`只可`output/pending/excluded`。输出须有实际行；待核可留草稿行或无行；排除须原因、来源且无输出行。反向扫描实际产出每行也必须唯一对应，拒绝漏人、多余人、重复行与一行映射多人。工资未核人员不能进入正式税表/可发工资单。

实际引用固定为`{table,row,column}`。默认核对本人身份；引用通用政策或日历时才显式加`person_bound:false`。映射范围也是确认时必须复核的输入；程序不能独立证明操作者选择的是所有业务源。

原始姓名必须反向核对去向：`raw_attendance.table`若是标准化JSON，另以`raw_attendance.original_table`绑定实际xlsx/xls原件，至少映射`name/date`。程序按原始本月全行姓名反查，不让`excluded_rows`或规范化过程隐藏人员。原件出现、但不在固定scope的姓名，须在`scope_dispositions:[{name,route:"pending/excluded",reason,source_refs}]`明确去向；排除须实际同人来源引用，不能仅以“不在已选名单”循环说明。未声明或证据不足自动列范围待核，绝不自动加入工资、税表或固定名单，也不替人资判断归属。缺原始绑定明确列待核，不宣称范围已核全。CLI在审计JSON旁另存`.scope.csv`，脚手架输出`scope-review.csv`，含姓名、去向、原因及原始物理行号；私有范围复核清单不进入技能发布包。

## 逐人事实与金额

### 原假别到最终人资交接

原件有本范围本月非零假别时，日期、原表头假别、原值与单位必须用`scripts/attendance_fact_handoff.py`从现有record的`raw_attendance.table/original_table`及实际姓名/日期映射读取；表头取已有header_row或数据起始行上一行。程序保留实际行列、源SHA与稳定fact_id，分类与工资待遇另列，不把原标签改写为另一假别，不将原分钟进位，不推定工资待遇。

运行`python scripts/attendance_fact_handoff.py render --record 本月记录.json --output 尚不存在的新目录`生成简体事实CSV、短句、最终交接/说明及`run-with-handoff.json`；已有最终清单时先在record填实际目标定位，工具只另存修正副本，不改原件。`attendance_handoff:{final_source_ids:[实际最终来源id],targets:[{source_id,format:"xlsx",sheet,cell,fact_ids},{source_id,format:"text",block_id,fact_ids}]}`；两类来源角色分别为`hr_handoff/run_explanation`，每类须覆盖全部原事实，所有最终交接用途来源都须列入。fact_id由工具生成，不要求用户重填参数；目标定位由执行者从本次实际交付清单取得。默认无既有清单时工具直接生成这两件最终文件，并自动写新record。

UTF8说明中使用唯一的`<!-- attendance-facts:leave:BEGIN -->`和`<!-- attendance-facts:leave:END -->`控制段（block_id可替换leave）；原日期/假别/数值短句只在程序控制段或绑定格渲染，非数据影响说明可以自由书写，不手工重复编造原事实。工具不在旧摘要旁静默附一份正确日志：已有最终文件必须绑定真实格/事实段，不能以旁边的source-facts.json冒充最终清单。最终gate重新读实际原件并重生成短句，再读最终文件绑定逐字核验；即使重新填写最终文件SHA，篡改日期/假别/单位/数量仍拒绝。自由区不做泛自然语言扫描，程序只证明全部已声明最终交接的绑定事实一致，不证明任意未登记文件或自由文字都无误。

`verify --record 新record.json --output 新检查.json`可独立复核；`workflow_audit`自动执行同一检查。原值正确而待遇未知时，事实交接可以技术通过，工资待遇继续业务pending，不新增人工审批或补零。

`facts`每项为`{employee_key,stage,item,status,amount,source_refs,value_refs,reason}`，人员+阶段+项目唯一。状态为`calculated/confirmed/confirmed_zero/not_applicable/pending`。缺事实等同待核；待核金额必须null，实际输出格必须空白，不能补0。不适用须原因和真实来源，确认零须数值0。`value_refs`定位实际产出，阶段表映射了项目列就不得漏绑。各阶段必需项目见`crossstage_audit.py REQUIRED`。

只对绩效有已确认制度的显示例外：本人本月已确认无绩效时，辅助事实保持`status:"confirmed_zero",amount:0`参与核心计算，主表可空白，不能把空白误判缺数。显式记录`display_exception:{kind:"confirmed_zero_performance",render:"blank",confirmation_source_id,amount_ref,month_ref}`及实际`policy_refs`。`amount_ref/month_ref`必须同一本人行，原确认来源、确认记录与实际核算月一致，确认绑定实际原件SHA和本人scope依赖；缺确认、其他人、月份错、依赖变化均使该绩效待核，不能进入完整应发/实发。此例外不扩大到其他工资桶。

复用已经取得的本人本月“无绩效”确认，不要求用户另做数值0表或再确认一次。原确认文件已有本人0/月可直接定位；原确认是聊天或文字时，可在私有规范化辅助文件登记amount0/month，`confirmation_source_id`及`confirmations`仍指原确认留痕，另加`original_evidence:{source_id,line,quote}`（或`ref`定位实际原文表格）。实际原文必须能定位本人或已核定整份人员范围，并明确确认无绩效；规范化文件不能冒充用户确认的新文件。“未提供/待核”原文不能转成0。原始引用、规范化引用和政策都纳入后续工资确认的来源快照。

只有模板或人资明确要求显示0时用`render:"zero"`，另以`zero_required_ref`定位实际要求（如明确的`zero_required/填0`记录）。检查程序只读，不会把原主表空白自动写成0；实际填0不代表缺资料已获确认。绩效尚未提供或尚未核算仍为`pending,amount:null`，主表留白，其余独立可算项目继续。

薪资、调薪、递增、绩效和人工补扣已完成时，另填`derivation:{operation:"copy/sum/negate",refs:[实际本人核定金额引用]}`。重读金额计算，拒绝工资输出自证原始金额。复杂递增、绩效政策仍由专用核算程序保留条款、计算过程及待核口径，本检查不凭状态推定政策正确。可选`depends_on:[{stage,item}]`及`operation`表示其他明确计算，固定交接关系不能靠删依赖绕过。

固定核对奖励逐笔→个人→考勤→工资；假别与加班待核状态→工资；个人养老+失业→工资个人社保，医疗/公积金分别→工资及个税；单位养老+失业→工资`employer_social`、单位医疗含生育→`employer_medical`、单位公积金→`employer_fund`；工资应发/实发及个税→工资单。单位三栏是工资模板原有的公司社保/医保/公积金栏（按实际模板列映射），不能只填个人三栏替代，也不能加入gross、个人扣款或个税二次计算。假期时数与人民币扣款单位不同，只比较适用/待核状态，不硬比金额。

金额复核直接调用既有`payroll_core.payroll_totals`，不另建工资引擎：应发`gross`只取薪资、其他补贴、绩效、奖励、请假扣款、考勤扣款、入离职差额七个实际工资桶；`finance=gross+other_income`；`net=finance-other_deduction-个人社保-医疗-公积金-个税`；`taxable=finance-other_deduction+combined_tax`；`bank=net-third_party-cash_other`。原始核心中的`fixed_salary/other/personal_deduction`分别映射台账`salary/other_income/other_deduction`。

调薪、递增、人工补扣是组成及适用性说明，不作为额外加项。非零事实必须用`included_in`说明已包含在哪个实际工资桶：整月已生效递增可已含`salary`，月中补差可已含`other_allowance`，人工补扣依本人核定来源归入适当桶。不得把这些金额重复加到已包含它们的薪资/补贴。假别待遇分别按既有专门制度映射，未明口径待核，不统一归入请假扣款。月中产假差额必须为正数、`included_in:"other_deduction"`，由实发公式扣一次，不得放入请假、考勤或入离职差额。用`manual_original_ref`定位原其他个人扣款格；原来非空时必须提供`merge_confirmation_source_id`（角色`manual_approval`的真实确认来源）及`merge_amount_ref`，确认要绑定原人工值和本次差额来源SHA，实际合并金额须与本人确认一致。不能假定原人工值已包含产假，也不能重复扣除。原来空白时也要在实际桶`depends_on`中绑定该专项差额。

未决项目的已知应发小计只存于`evidence.known_subtotals`并列出未知项，不写成正式`gross/finance/net`。加班余额缺项不影响无依赖的独立金额。

## 全范围原始事实

`raw_attendance:{table,header_row,leave_columns}`。映射`name,date,attendance_type`及公司或可靠ID；`leave_columns`逐列为`别名:{kind,unit}`。`kind`仅`personal/sick/annual/childcare/maternity/other`；`unit`仅`minutes/hours/days`。Excel实际表头全部含“假”的列均须映射。产检、丧假、婚假、陪产、调休、哺乳等归`other`仍保留原列、行号、原值，不得丢弃或默认为工资不适用。不是此明确宽表结构的来源先做保留全部来源行的适配器，无法证实则待核。

程序从实际全部日明细独立筛出本月记录。`raw_hours`是仅换算单位的原始观察量；`business_hours`是按已确认制度的业务量，二者分开。不需要打卡者制度不记录的事病假，业务量为null而非计入0小时，原始观察不丢失。事假、病假按自然日向上取整到1小时、每天8小时封顶后汇总，30分钟为1小时、61分钟为2小时、跨日各30分钟为2小时；不保留0.5小时。其他假别保留原始单位换算，不在此另创舍入。考勤事实带`unit:"hours/days"`和`observed_rows:[{row,column}]`；原始行须完整，但实际主表`amount`依制度适用性核验，不能强迫等于原始观察量。

工资同项目也保留完整`observed_rows`，金额是工资处理而非小时。已完成且有原始记录的项目须实际`policy_refs`。`treatment:"not_recorded_no_punch"`表示已证实不需要打卡，制度不记录事病假月度项；`treatment:"no_auto_deduction"`只适用于实际花名册证实普通在职非实习员工，事假不自动填时数和扣款；实习生、当月离职人员按原制度计算，特殊确认用`policy_confirmed`并绑定本人实际依据。没有适用性证据就待核，不能只改成不适用。

原始导出含入职前/离职后占位或假别记录时，按实际核定scope中的`hire_date/departure_date`过滤业务计入；`raw_rows/raw_hours/leave`仍完整保留，`excluded_employment_rows`逐行记录岗外日期和原因，不能只缩小预期日历而仍把岗外时数加进业务量。身份与日期必须来自绑定确认链的实际名单，不自行填写推测日期。

打卡类型先用考勤模块`normalize_attendance_type`；`无需打卡/無需打卡`统一为`不需要打卡`。只对本人在职日期范围内完整逐日覆盖且全部不需要打卡者，依已确认制度不记录事假和病假月度项，可留空；原始观察仍完整保留并通过政策引用说明。日明细缺失、未知或混合类型不得套用此豁免，年假、育儿假等不能一并豁免。原始空白默认未知；只有实际确认该导出空白表示零才填`blank_means_zero:true`和`blank_zero_refs`，不能掩盖未提取列。

`raw_rewards:{table,enrichment,collection_complete_refs}`直接读原始问卷全行。映射`name,date,tier,next_workday,checkout,next_arrival`及必要审批字段，缺字段保留空与待核。`enrichment`按原始行号及字段绑定实际打卡/日历/审批引用，禁止直接填结论。重新调用`reward_policy.audit_claims`逐笔判断、重复处理、个人汇总；个人奖励事实带完整`claim_rows`，金额和待核状态贯穿。无申报者须完整收取依据才可能确认零，空表本身不证明无奖励。

`raw_insurance:{table,components,medical_employer_mode_column}`是明确映射的实际分项宽表，映射身份、`month`。原标准必需四类`pension/unemployment/medical/housing_fund`，每类`components[险种]={employee:员工金额字段,employer:单位金额字段}`；真正缺列明确null，不删必需类、不补零。`maternity/injury`只在实际有明细时作为可选项目保留；无独立工伤源不阻断工资所需保险检查，不能伪造工伤单位0。可选工伤仍不进个人扣款或上述单位三栏。

事实`amount/status`为个人金额/状态，`employer_amount/employer_status`为单位金额/状态；单位状态省略时仅对有数值的单位金额兼容视为calculated，无金额即待核。个人与单位分别读取实际列验证，缺单位只使对应单位栏待核，不抹空已核个人项；缺个人不妨碍已证单位金额。声明单位pending时单位金额也须null；可选单位险种没有个人字段则保留原始null，不编造原件0。单位三栏必须另有工资事实及实际主表列和值引用；工资列没有映射时不因辅助表金额正确而通过。费用所属期须本月，错公司/人/月、多行补退费未确认仍拒绝。

`medical_employer_mode_column`指实际来源宽表里表示医疗单位组成的字段（默认`medical_employer_mode`，必须为实际映射列，不能用常量代替）：`combined`表示medical单位值已经包含生育，直接取该总额；`split`表示medical单位值仅基本医保，另加实际maternity单位值。普通原制度8%+0.7%可分别得800+70或合并870，二者都只计870，不能把合并870再加70；广思用本人源表实际金额及拆项，不能套普通费率。适配文件须保留源文件SHA、行号、表头和组成依据，不能只凭猜测选模式。模式未核只使单位医保待核，其他独立金额继续。额外单位项目保留原明细和核验状态，不把不属于工资必需项的未知金额伪装成个人扣款或正式三栏金额。

这些口径沿用`social-insurance-fund/SKILL.md`的标准四类、Current Calculation Rules与Output Meaning：单位社保为养老+失业，单位医保为基本医疗+生育，单位基金来自原单位缴存额；本检查器不新增险种费率或工资纳入规则。

## 确认及模板连接

`confirmations`每条为`{source_id,sha256,confirmed_by,confirmed_at,dependencies:{前置来源id:当时SHA}}`，仅记录实际明确人为确认。正式名单要`scope_approved`；正式工资要`attendance_confirmed`；正式个税/工资单要`payroll_confirmed`。确认须绑定实际名单/阶段产出及全部前置来源快照；重新填写当前源SHA不能使旧确认快照重新有效，来源变更后必须重核、重确认。

`crossstage-1`按`through_stage`对应前缀中考勤、保险、工资、个税、工资单阶段核验指定模板，不要求尚未请求的未来阶段；不存在省略`template_required`或自行列子集即可豁免的路径。报告明确列`required_template_stages/verified_template_stages/all_five_stages_verified`；仅考勤请求只核考勤，工资核考勤/保险/工资三类，完整工资单回放须覆盖五类。缺资料只使对应金额待核，其他独立计算继续；缺模板绑定或仅JSON表则标技术未完成，默认`passed:false`。已确认输入不得为过技术检查重算或覆盖。旧`template_receipts`自报回执不能替代下列实物绑定。

`template_bindings[阶段]`的xlsx绑定为`{kind:"xlsx",registry_source_id,registry_sha256,plan_source_id,plan_sha256,receipt_source_id,receipt_sha256}`；每个来源还须在`sources`提供实际路径和SHA。程序直接只读调用`template_guard.verify(...,report=None)`，重新读取指定登记、封存原模板、计划和成品，核真实公式/样式/打印/封装及人员映射，不信任回执的`passed`。重算后同一文件SHA变化时另提供`recalculation_log_source_id/recalculation_log_sha256`用于原生成SHA→最终SHA连接；日志字符串本身不证明原生核验。登记/计划/回执等实际来源也纳入前置确认快照，换计划或登记须重核重确认。

工资登记还须`registry.policy.require_source_comments:true`，并在`policy.source_comments[*].source`按实际path/SHA覆盖`payroll_original_source_ids`列出的本次全部工资原件（及sources中明确角色`payroll_original/wage_original`的来源）。没有批注也要登记0条盘点，不能把新成品当原件代替盘点。跨环节gate检查实际登记盘点及实际verify能力字段；缺原件清单、覆盖或能力只列对应待核，不声称批注已保真，其他独立草稿继续。批注逐条保留/迁移的实物验证仍由模板模块实际执行。

原批注对象保留不等于业务事项已处理。程序从实际`source_comment_inventory`复读原件SHA、原Comment和同人行，逐条生成`comment_business_review`及`.comments.csv`供人资复核。本门槛只自动识别明确“事假清算/结算/结清”类事项，不把原注小时数当制度或自动填扣；普通说明列“仅说明”，不新增人工审批。清算类若无本人本月实际处置，S请假扣款、W应发及其财务/实发/税表/工资单依赖均为业务待核；事实仍称已算或不适用会报矛盾，实际金额格仍显示数值（含0）也拒绝技术交付，不能只改审计状态。执行者须让实际受影响输出待核，独立已知组成仍留辅助明细。

既有逐人处置用`comment_dispositions:[{employee_key,source_sha256,source_sheet,source_cell,action,reason,evidence_refs}]`绑定精确原注；`action`为`adopted`（采用）、`explanation`（仅说明）、`historical_not_applicable`（历史不适用）、`pending`（待确认）。`evidence_refs`复用已有确认文件/聊天留痕的实际同人本月表行，引用文字须实际说明该事假清算如何处理；来源月及该行实际`month`均须本月，不能只给旧证据新贴月份，错人、错月或“全员未确认”不能代替。采用还提供`amount_ref`指向同一处置行的本人当月已核金额；该金额默认仅为清算子项，不能强制等于含病假等其他项目的完整S。

清算子项用`component_refs`引用同一实际完整明细表的各项`amount`，其中`amount_ref`必须恰好出现一次；明细表实际映射`component_id,month,amount`及本人身份，每人本月项目标识唯一，程序扫描全部实际本人当月行核覆盖（不能删引用或排除行藏项目）。`total_leave_ref`定位本人当月实际`total_leave`完整扣款总额，先核全部组成求和，再与实际S核对。例如明确合成清算-80、病假-20、完整S-100可以通过，漏清算、重复计入或遗漏其他明细不能通过。只有原处置文字明确说明已核的是“完整请假扣款/S总额”，且声明`amount_scope:"total_leave"`，才可把`amount_ref`直接当整S核对。缺组成或总额依据继续待核，不把子项不同于S说成算术错。执行者可从既有有日期的原确认与明细规范化身份、month和原文定位，不要求用户重做表格或重复批准，不能自造确认。处置及组成来源纳入后续工资确认快照。

政府原xls用`{kind:"legacy_xls",template_source_id,template_sha256,policy_source_id,policy_sha256}`；程序直接调用`legacy_template_guard.inspect(...,report=None)`读取真实22列、说明页、样式、打印，`policy`按实际确认的清底色要求填写`clear_all_fills/clear_original_yellow`，不转换格式或重建原表。

工资绑定另须`formula_audit_plan_source_id/formula_audit_plan_sha256`，程序直接调用实际`audit_payroll_workbook.audit`。计划须覆盖实际工资sheet及全部人员行；`pending_formula_exemptions`兼容逐格待核政策及实际来源，空白待核可以保留，但不能写死常数或当正式通过。

原生核验用绑定内`native_recalculation:{run:true,linkage_checks:[...]}`实际执行`native_recalc_verify`。它先复制到隔离目录，子进程没有原件路径，只在隔离副本用实际兼容办公引擎重算；核实际进程路径/厂商/版本/执行档SHA，报告实际WPS或Microsoft Excel，不冒称产品身份。它读取原文件全部公式缓存并与原生计算值比较，工资还需在实际同一人员行编辑事假输入，核`leave_deduction/gross/net`真实公式按计划联动；每个测试为`{sheet,input_cell,input_expected,test_value,outputs:[{item,sheet,cell,expected}]}`，输出sheet默认输入sheet。工资另须`kind:"attendance_snapshot"`测试：编辑实际受控来源页本人金额格，核工资T及同人应发/实发跨页联动；实际guard映射决定允许输入和目标，不接受只改姓名的常量数组伪联动。依赖未决导致合理空白时逐格待核，不称完整联动通过。字符串`passed/engine/events`不能代替实际运行。具名锁串行本工具调用；已有其他工作簿则待核，只关闭本次副本，只有确认独立新进程且无其他工作簿时才退出应用。45秒未完成保留待核，不杀共享办公进程。工资单账号资格仍另用`workflow_audit.payslip_eligibility`核对。

随包最小合成反例位于`crossstage-tests/public_tests/test_crossstage_public.py`，动态寻找相邻技能程序，不含真实人员、工资或用户绝对路径。真实复制件与其他私有测试证据不能放进发布包。

原生缓存按实际xlsx XML核验：公式格同时有`t="str"`和空`v`节点时是合法空字符串缓存，不能因openpyxl读为None就判缺缓存；没有`v`或未声明字符串的空`v`仍是未证实缓存，不补0。S输入改变后，扣款可已联动而应发/实发因其他资料未决仍空白，此时输出逐格状态和待核，不能冒称完整联动通过，也不把合理草稿当公式错误。`expected:null`或缺核定预期时不猜数字答案。ProgID优先由COM解析，核实际32/64注册视图；注册定位缺口可在未开工作簿前用真实受支持进程身份补证，保留定位缺口及实际应用信息。

实际公式错误独立于金额预期：原生复核保存单格`Value2/Text`及同格`ISERROR`结果，缓存`t=e`/xls错误类型或原生真实错误均列技术errors，即使`expected:null`也不能吞成一般待核。正确空白或正常数字但预期未知仍pending，不猜答案；列宽不足显示`####`本身不算计算错误。原锁、隔离副本和共享办公会话保护不变，不自动修改工资或输入。

程序保障给定契约内的实际行、金额与状态一致，不证明已收齐现实中所有来源，也不验证现实确认人身份、不补造历史政策。九月原始复制验证只验证读取与事实提取；缺确认及后续前置保持待核，不能称新月份真实核算完成。

## 旧版诊断契约（不用于正式通过）

记录放本月实际资料目录，不能放Skill安装/发布目录。JSON顶层：`month,policy_version,stage,expected_employee_keys,employees,sources,confirmations`。

- `stage`为`attendance/payroll/tax/payslip`；预期名单来自已核定负责人范围，`employee_key`优先可靠员工标识；无ID且姓名唯一才用公司+姓名并记录弱匹配，不把OpenID当UserID。
- `employees`每项有`employee_key`和`items`。每项目含`status,source_ids,reason`，可加金额、来源行、规则、计算过程。状态：`calculated,confirmed_zero,not_applicable,confirmed,pending`；缺项等同pending，空值不能代表零。项目列表由`workflow_audit.py REQUIRED`维护。
- 工资阶段每人另带 `observations.maternity`：`record_count` 为从本人本月原始假别/请假明细实际筛出的产假记录数，`coverage_complete` 表示原始覆盖已核实，`source_ids` 绑定原件，并保留日期及行号。只查事假、病假汇总不能证明没有产假；原表其他假别也保留并交接，不因考勤模板没有专栏而丢弃。原件有产假时，待遇或分段未核定就标产假待核，不能在工资交接再写“不适用”。记录数量/覆盖未知就保留未知，不能为通过检查补0。检查程序会拒绝缺覆盖证据，以及记录数大于0却标不适用的矛盾；计算明细、异常清单、工资及交接记录须指向同一结论。
- `sources`是来源ID至记录的映射，记录`path,sha256,month,role`；原始月份可以不同（如上月余额），必须按用途核对，不要求全部等于本月。角色如`attendance_confirmed/payroll_confirmed/roster/history/calendar/claim/approval/manual`。路径可相对该JSON；文件发生变化即使仍叫确认版也须重新检查。
- `confirmations`逐条含`source_id,sha256,confirmed_by,confirmed_at`，只能记录用户实际明确确认，不得由程序或AI自动确认。进入工资须已确认考勤；进入个税/工资单须已确认工资。任何受影响源变化后重做对应检查与确认。

程序不会自动盘点任意目录或生成真实确认。执行者必须按源文件整理记录，逐项对明细并运行检查。核算程序负责金额，检查程序负责完整性；状态字段不反向驱动金额。正式模板不增审计列，计算/核验记录另存。

标准化输入要核对：员工行不漏不重、负责人范围与顺序、日期过滤、日假时长向上取整及8小时上限、跨年病假/年假初始化、结余归属期、姓名与ID冲突、奖励申报覆盖。没有完整原始文件就标资料不足，不能只凭文件名或空表证明零。

工资T连接由登记的`sheet_policy.source_views[*]`及其`link:{target_sheet,identity_column,target_column}`生成：先核考勤原件path/SHA/实际同人行，将真实数字或空白快照入工资内来源页，再引用同人金额。实际verify返回`source_view_inventory`，跨阶段检查直接注入公式审计，不采信自造回执。`attendance_link.link_for_employee(mapping,identity)`可取得同一生成式/源定位；无来源、他人行、身份配常量、空白当0均拒绝。外部考勤源变化只使SHA失效，须重新导入重核，不承诺静默联动，更不得自动覆盖既有人工确认额。

免打卡不另要求伪造整月每日行。已有本人本月确认或长期人员名单可通过`raw_attendance.no_punch_refs:{员工主键:[{table,row,column:"attendance_type"}]}`引用实际免打卡字段；同一实际行须有`month`，或有`effective_start/effective_end`覆盖本人本月在职范围（长期有效可留空结束日）。程序核实际文件SHA、同人身份、原始类型标准化及月份/有效期，来源纳入后续确认依赖；不能只凭缺记录豁免，也不能忽略已有需打卡/未知/混合类型冲突。原始观察仍保留，免打卡者事病假月度项不记录，其他缺资料项目分别待核。
