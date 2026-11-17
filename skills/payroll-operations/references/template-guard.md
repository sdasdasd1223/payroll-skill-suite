# 指定模板、公式与样式保护

入口为 `scripts/template_guard.py`；`scripts/template_fill.py` 委托同一实现。Python依赖 openpyxl、lxml；政府原始 `.xls` 的只读入口 `scripts/legacy_template_guard.py` 另需 xlrd。`.xlsx` 制作不依赖 Windows，原生重算须另外具备 Excel 及实际运行证据。

## 固定来源与实际调用

先在运行输入清单固定用户指定原件的绝对路径与 SHA256，不能制作输出后拿重建表登记为模板。登记成功后把 `registry_sha256` 也固定在运行记录；后续不得静默修改政策或重签旧凭据。登记文件、原件和封存副本都必须保持 SHA 一致。

```text
python scripts/template_guard.py register --template 指定模板.xlsx --expected-sha256 清单固定SHA --policy 模板政策.json --registry 新登记.json
python scripts/template_guard.py build --registry 新登记.json --registry-sha256 固定登记SHA --plan 核定计划.json --output 新结果.xlsx --receipt 新来源凭据.json
python scripts/template_guard.py verify --registry 新登记.json --registry-sha256 固定登记SHA --plan 核定计划.json --output 新结果.xlsx --receipt 新来源凭据.json --report 新实物检查.json
```

写入一律另存，不覆盖原件和旧结果。`build` 先逐字节复制封存模板，再修改副本；样式从模板继承，不按表头重建。模板原有 Excel 扩展/附件保留，样式索引取自模板规范化副本；不连续日期格式编号会一致转换，避免日期变数字。`verify` 重放模板和计划后读取实际成品，外部自报 `passed:true` 或给新 SHA 不能代替检查。

登记/实物检查涵盖全部工作表及表头和值、公式、字体/字号/粗体/对齐/边框/填充/数字格式/保护、行列尺寸、合并、筛选、冻结、批注和超链接，以及打印范围/标题/页边距/页眉页脚/分页。不把 `max_column` 当业务字段数：只含格式的空列照模板保留，业务映射以真正非空表头和登记字段为准。

## 政策与来源计划

原样复制时政策和计划可用 `{}`。实际写入先登记可写区和人员区，人员区不含表头或合计。例如：

```json
{"sheet_policy":{"keep":["工资表"],"copies":{}},"sheets":{"工资表":{"personnel":{"start_row":4,"end_row":4,"identity_columns":["B"]},"writable_ranges":["A4:AK4"]}}}
```

这是桌面模板业务页的结构示例，不能直接当当月人员、金额或核定依据。该模板其余历史附页不视为本月数据。`sheet_policy.keep` 预先选定保留页及顺序，未配置则全部保留；注册时扫描公式和命名范围依赖，删除仍被使用的内部页会拒绝。必要新页须预先登记 `sheet_policy.copies={"内部核对":"工资表"}`，按指定来源页复制，不能任意新建页面；个税原始 `.xls` 说明页始终必须保留。

每个真实来源为 `{"path":"绝对路径或相对计划目录的路径","sha256":"预先固定SHA","locator":"工作表/行号/条款定位"}`。政策的公式来源相对登记目录解析。写入形如：

```json
{"writes":[{"sheet":"工资表","cell":"B4","expected_old":null,"value":"实际来源内姓名","reason":"本月核定身份","source":{"path":"人员核定文件.xlsx","sha256":"真实SHA","locator":"人员核定第3行"}}]}
```

非空人工值及公式默认保护；确需替换，还要给同结构的 `authorization`，定位本次确认依据。`replace_confirmed:true` 或文字 `source_id` 不足。脚本证明文件与定位可追溯，不能证明文件内人工授权真实；后续业务审计仍须核实。

`template_fill.py` 还须传 `--template --registry --registry-sha256 --receipt`，旧版只有 `sheet/cells/source_id` 的计划被拒绝。`complete_total_formula` 函数保持原接口。

## 合法人员变化与公式

无 `row_maps` 的表自动保持原序；只为有变动的表提供完整来源行数组。整数为原员工整行，包含其人工值、样式及批注；新增写 `{"source_row":4,"mode":"new"}`，只继承公式和样式、清掉原姓名/人工值/批注，再从真实来源填写新身份。不能把旧员工姓名抄进别人样式行，也不能将同一旧员工以 `copy` 重复复制。真实工作优先用员工主键，只有姓名时先处理同名。

例如原表第4行是人员、第5到8行是尾部，新增一人可用 `row_maps={"工资表":[1,2,3,4,{"source_row":4,"mode":"new"},5,6,7,8]}`。每张变化表给 `row_change_sources[表名]` 指向已核定增删/组织调整依据。大量人员行数组可用 Python 的 `list(range(...))` 生成，无须手填全部坐标。表头和尾部不能被当人员行复制。

公式本行相对引用跟随员工，跨行/跨页引用按来源行追踪，合计包含复制自该范围的新人；删除被单独引用的行、形成不连续引用时停下，不能猜测公式。跨人员的纵向组标题合并要给 `merge_maps[表名]=[{"source":"A3:A10","target":"A3:A8"},...]`，来源须为模板真实合并范围且保持合并列，并有该表 `row_change_sources`。组标题沿用原组标题来源，不能用员工数据行代替。

需要补入/改写公式，先在政策 `formula_rules[模板来源坐标]` 登记 `expression`、`required_row_columns` 和依据 `source`；实际写入仍须理由与改写确认。表达式必须与登记公式按本次映射平移后的结果一致，且含储存格依赖，拒绝 `=100`、`=0` 等伪联动。静态依赖不能证明方程业务正确或实际可联动，仍需改变输入值后的原生测试。

复杂合并产生模板外边框样式，或人员重排涉及图表、条件格式、数据验证时，工具明确停止，须原生保留格式处理并复核；无法可靠保留的ActiveX、透视等扩展也不静默另存。标准Comment批注已支持跟随核定人员行移动及从实际工资原件复制，XML/VML关系与对象都会重放验证；含非批注控件的VML不能替换。超链接移动仍须另行原生核对。

## 工资原件批注必须保留对象

工资登记必须设 `require_source_comments:true`，并在 `source_comments` 覆盖本次实际使用的每份工资原件，固定原件路径/SHA。没有批注也登记，程序实际读取后记0条；不能因为空白正式模板没有员工批注就忽略工资原件，也不能仅把原批注文字写进备注后称已保留批注对象。批注涉及工资口径，原文复制不表示该历史口径已经本月人工确认。

先可独立盘点（只读，全部工作表）：

```text
python scripts/template_comments.py --source 工资原件.xlsx --sha256 清单固定SHA --output 新批注盘点.json
```

登记政策示例，坐标从本次实际工资原件和模板取得：

```json
{"require_source_comments":true,"source_comments":[{"id":"工资原件","source":{"path":"工资原件.xlsx","sha256":"真实固定SHA","locator":"完整工作簿全部批注盘点"},"regions":[{"sheet":"本月工资业务页","first_row":4,"last_row":27,"identity_columns":["B"],"target_sheet":"工资表","target_identity_columns":["B"],"header_rows":[2,3],"target_header_rows":[2,3],"fallback_column":"AG"}],"ignored_sheets":{"历史附页":"实际标题及归属期明确不是本次工资来源，说明具体依据"}}]}
```

源与目标都从实际单元格读取身份，同一姓名不唯一时须选可核身份组合，不能猜。先按真实多级表头找同一字段（公司/个人同名子栏分别定位）；默认直接复制 `Comment.text` 与 `Comment.author`，不改业务值和单元格样式。原字段在正式模板没有对应栏，或目标已有不同人工注时，改在同一员工原有备注栏的Comment内汇总，注明源文件/页/格/原字段、原作者及完整原文，已有人工备注值、批注正文和作者保留，不新建业务列。表头批注按真实字段对应保留。

确需明确坐标可在计划加入 `comment_copies:[{"source_id":"工资原件","source_sheet":"本月工资业务页","source_cell":"U20","target_sheet":"工资表","target_cell":"S18"}]`；这不是“强行通过”入口，工具会重新读取源/目标身份，错人即拒绝。大部分正常源不需要逐条手填映射。

全量盘点包含历史附页，只有已经核定非本次业务来源的附页才可在 `ignored_sheets` 说明依据；不能笼统忽略当前工资页。留存人员每条原注都有复制目标。源人不在实际目标主表时记录 `not_in_target`，不据此证明该人业务上应被排除，人员完整去向仍由跨阶段审计核对。未归属批注、源SHA变化或没有登记身份区会停止，不静默当0。

登记内 `comment_source_inventories` 固定全量真实盘点；来源凭据与实物检查内 `source_comment_inventory` 保存每条对象去向。`build/verify` 每次重读真实源及实际成品Comment，不信“已保存文字”或自报copied数量。`template_comments.inventory_sources(configs,base_dir)` 为只读盘点接口；`verify(...,report=None)` 是只读完整重放接口。工资流程的 `payroll_original_source_ids` 应覆盖所有实际工资原件，并由跨阶段门槛核对其路径/SHA都在登记清单。

当前源批注支持xlsx/xlsm，不把无法读取的旧xls批注谎称0条；旧xls工资批注须另用隔离副本原生提取并完成对象核对。本能力不等于政府xls格式检查，也不等于新增人工确认。

人员原型必须来自实际正常空人员行或合法样例，不得选明确说明行。工具从原件登记身份字段识别“从《花名册》提取”“请填写”“示例：”等指引，不按固定行号或F列判断；这些行可原样保留，但不得用new造员工，也不能仅改姓名将其作为人员记录。本例考勤第2行是132点高的说明，正常第3行起才是人员原型；若移除说明并让首人输出在第2行，应映射原第3行，不能沿用说明样式。正常空行、合法样例和不同合法人员原型仍允许。登记含personnel_instruction_rows供定位，build仍实读原件判断，不把自报人员区当豁免。

## 本月考勤来源值页与真实扣款引用

不要用姓名条件包住固定扣款、迟到分钟或职级倍率。工资母版原附页“考勤”是历史制度示例，不能强套为本月字段。登记时可明确授权新增必要内部值页，主工资字段和样式仍来自指定母版：

```json
{"sheet_policy":{"keep":["工资表"],"source_views":[{"id":"attendance","sheet":"考勤依据","source":{"path":"本月已核考勤结果.xlsx","sha256":"预先固定的真实SHA","locator":"模板表F2:F25为姓名，K2:K25为已算扣款"},"source_sheet":"模板表","first_row":2,"last_row":25,"fields":{"name":"F","deduction":"K"},"link":{"target_sheet":"工资表","identity_column":"B","target_column":"T"}}]},"sheets":{"工资表":{"personnel":{"start_row":4,"end_row":4,"identity_columns":["B"]},"writable_ranges":["A4:AK4"]}}}
```

这是合并入完整登记政策的示例；工资原件批注登记等原有政策照常保留。源须是本月已经核算/人工修正的实际结果，固定其文件SHA、页、实际行区及列映射，不能由执行者另造姓名金额表假称考勤确认来源。此功能只读取值，绝不重新计算或覆盖源人工已确认扣款。原源公式须有真实缓存，缺失/错误缓存停止，不当业务空白。

工具生成额外“考勤依据”页：第1行字段名，A列name、B列deduction，第2行起实际人员；有其他简单字段可继续登记在fields。源空金额保持空，姓名重复、无姓名却有金额、源SHA变化或实际目标姓名在源不存在都会拒绝。目前只支持无图表/图片/表格对象/合并区域的简单xlsx来源页，不猜同名身份。

`link`让工具在主表身份写完后自动逐人生成T引用，例如 `=IF(考勤依据!B2="","",考勤依据!B2)`；无需逐人手填来源坐标或T公式规则。计划不得再手工写同一T格。原目标非空人工值与实读源同额时，可建立不改金额的结构引用并记录previous_manual_value；不同金额即停止，不用新源自动覆盖。有实际核定例外才同时在link.reason说明理由、link.authorization登记原授权文件路径/SHA/定位，原值与授权仍留痕，不能靠自报“已核”放行。

修改内部快照金额会通过真实单元格引用联动T；修改过的快照不能继续冒用原来源验证结果，正式复核仍须核定该修订和来源。外部确认原件一旦变化，原SHA校验失效，必须重新固定来源、登记/导入和复核；这不是跨文件静默实时联动。缺项仍留空，不把空扣款变0。

```text
python scripts/template_guard.py source-views --registry 登记.json --registry-sha256 固定登记SHA --view-id attendance
```

公式生成统一使用 `template_source_views.source_view_formula(sheet,cell)`；普通中文/合法标识页名不加多余引号，含空格等页名安全单引号转义，生成器与T审计直接共用，不放宽缓存搬运中的公式一致要求。

只读接口 `source_view_mapping(registry,registry_sha256,view_id=None)` 重读实际源，返回自动值页的人员行映射。`registry.source_view_inventories`固定登记时盘点；build/verify的`source_view_inventory.sources`给出id/source.path/source.sha256/source_sheet/source_columns/sheet/columns/rows，rows含source_row/view_row/identity/values/source_cells；`source_view_inventory.links`给出target_sheet/target_cell/identity/view_sheet/value_cell/source_row/formula。每次verify实际读取固定源并重建快照与同人公式，另实读成品内容/样式及封装，不采信回执passed。

## 原生重算后绑定同一实际文件

普通 `verify` 严格绑定生成 SHA。Excel另存后 SHA 改变，用以下入口重新检查实际公式/样式/打印/扩展并绑定新 SHA：

```text
python scripts/template_guard.py finalize --registry 新登记.json --registry-sha256 固定登记SHA --plan 核定计划.json --output 实际重算结果.xlsx --receipt 原生成凭据.json --recalculation-log 实际重算日志.json --recalculation-log-sha256 日志固定SHA --report 新交付实物检查.json
```

日志须含 `engine,started_at,finished_at,source_sha256,output_sha256,events`，绑定原生成 SHA 及实际新 SHA。报告 `output.path/output.sha256/finalized_output_sha256` 绑定同一文件，供跨环节 `template_receipts` 和 `sources` 实读核对。日志本身不证明重算真实，所以本工具始终 `native_recalculation:false,formal_ready:false`；原生运行、公式缓存、输入联动和人工业务确认由对应审计完成，不因结构通过自动放行。

任何 Excel/WPS 操作都先 `shutil.copy2` 原件到隔离目录，然后只打开隔离副本。不可直接打开 Desktop 或 cases 输入，设置 `ReadOnly=True` 也不能取代隔离复制。本机兼容接口虽然回报 `Name=Microsoft Excel,Version=12.0`，实际 `app.Path` 为 `D:\Program Files (x86)\Kingsoft\WPS Office\12.1.0.29690\office6`；应用名称不能证明是真正 Microsoft Excel。WPS 的实际重算/缓存/扰动结果可以如实记录为兼容应用证据，不能冒称 Microsoft Excel 原生验证。

重算应用可能真实改变行高或舍入列宽，不能放宽样式检查。可用下列工具核对实际计算副本与原guard产物的所有公式、所有非公式输入一致，然后只搬运公式缓存 `v/t` 到保持原模板样式/封装的另一副本，其他 XML/样式/打印都取原guard产物。缺缓存、错误或不支持的缓存类型即停止。

```text
python scripts/template_cache.py --base 原guard结果.xlsx --calculated 实际兼容应用计算副本.xlsx --output 原模板样式与实际缓存.xlsx --report 新缓存来源报告.json
```

缓存报告绑定两份来源与输出 SHA；随后 `finalize` 的日志 `source_sha256` 仍指原guard结果，`output_sha256` 指搬运后的实际输出，同时在事件中保留真实计算副本的 SHA 与应用实际路径。搬运不执行计算，也不证明数值正确；独立重算/缓存/联动审计仍必须执行。

## 政府原始 .xls

```text
python scripts/legacy_template_guard.py --template 指定个税模板.xls --expected-sha256 清单固定SHA --report 新原件检查.json
python scripts/legacy_template_guard.py --template 指定个税模板.xls --expected-sha256 清单固定SHA --output 原格式结果.xls --clear-all-fills --report 新政府格式检查.json
```

制度是先识别原始黄色必填字段，再使导出文件全部单元格不留任何底色，包含说明页。`--clear-all-fills` 才代表这项例外；旧 `--clear-original-yellow` 只兼容仅清黄色字段的旧测试，不能代表正式全清。字体/边框/数字格式等其他样式仍保护。检查原始22栏、字段次序、完整说明页、逐格样式、列尺寸/合并、原始打印/分页记录；原件识别的黄色字段/格子写入报告，不用以往输出猜测。

WPS 在本例另存旧 `.xls` 时曾把原默认行高315/20=15.75点改为312/20=15.6点，检查正确拒绝。已实测的恢复方法：在隔离输出对每张工作表设置 `sheet.Rows.RowHeight=源表.default_row_height/20`，再按源表 `rowinfo_map` 恢复每个特殊行高，最后清全表底色、另存并重读复核；源件只用 xlrd 只读取得行高，不用应用打开。不能只修主表而漏掉说明页，也不能把行高漂移当误差放行。

默认行高并不能证明新增人员行正确。本例原业务页第2至17行有效行高是15点，默认15.75点；新增行若没有ROW记录而继承15.75点仍须拒绝。`--personnel-prototype-rows 2,3`可预先限制合法人员原型，1起算、不允许表头第1行或说明页；默认原业务页第2行起。已有行按原行比较；新增行须整行字体/边框/数字格式等与某个完整合法原型一致，并且行高/隐藏/分组也与该同一原型一致，不把“末行样式”强加给所有新增行。报告personnel_row_mapping留存实际输出行、选用原型行及双方有效行设置。只修默认行高不能取代逐行恢复；原生处理或保留原格式的BIFF写入都必须给新增行正确的实际人员行高，再用此检查读回。

判断旧xls是否有实际人员样式行，必须用`xlrd.open_workbook(...,formatting_info=True)`读取rowinfo和单元格样式；默认读取nrows=1可能只表示仅表头有值，不等于其后没有已格式化的合法人员原型。不得以此误判“缺模板”而放弃可制作的人员行。

此入口只读检查、不制作或转换 `.xls`；xlrd只能取得公式缓存，所以仍须另用真正原生应用在隔离副本核对实际公式及人员/金额，不能把其他 `.xlsx` 通过当此政府格式也通过。可导入 `verify(...,report=None)` 与 `legacy_template_guard.inspect(...,report=None)`，跨环节审计可只读重放，不采信外部回执自报通过。
