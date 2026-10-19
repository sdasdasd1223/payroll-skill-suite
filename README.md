# 工资核算 Skill 套件

这是一套用于公司月度工资核算的 Codex Skills，覆盖考勤初算、人工确认、五险一金并表、工资计算、个税模板和飞书工资单生成。

## 在 Codex 中安装（Windows / macOS）

打开 Codex 新对话，把下面这段话完整发送。Codex 会识别当前系统，使用仓库中的对应脚本；人资无需自行打开终端：

> 请从 https://github.com/sdasdasd1223/payroll-skill-suite 安装或更新四个工资核算 Skill。请先核对仓库与安装脚本，在 Windows 上使用 install.ps1，在 macOS 上使用 install.sh；如有旧版先备份。完成后核对四个 Skill 均已安装，告诉我备份位置和是否需要重启 Codex。不要修改 GitHub 仓库，不要把我的个人资料路径写入共用 Skill；这一步不开始算工资。若权限或安装失败，请给我具体错误。

安装完成后完全退出并重启 Codex。安装文件默认位于各自用户目录下的 `.codex/skills`；设置了 `CODEX_HOME` 时，以其指定位置为准。

## 手动安装（备用）

### Windows 10 / 11

1. 在电脑左下角搜索并打开 **PowerShell**。
2. 复制下面整条命令，粘贴后按回车。
3. 看到“安装成功”后，完全退出并重新启动 Codex。

```powershell
Set-ExecutionPolicy -Scope Process Bypass -Force; $z=Join-Path $env:TEMP 'payroll-skill-suite.zip'; $d=Join-Path $env:TEMP 'payroll-skill-suite-online'; Invoke-WebRequest 'https://github.com/sdasdasd1223/payroll-skill-suite/archive/refs/heads/main.zip' -OutFile $z; if(Test-Path -LiteralPath $d){Remove-Item -LiteralPath $d -Recurse -Force}; Expand-Archive -LiteralPath $z -DestinationPath $d -Force; $i=Get-ChildItem -LiteralPath $d -Filter install.ps1 -Recurse | Select-Object -First 1; if(-not $i){throw '下载包中没有找到 install.ps1'}; & $i.FullName
```

以后更新到最新版时，重新运行同一条命令即可。安装程序会先把电脑上的旧 Skill 备份到 `.codex\skill-backups`，再安装新版本。

### macOS

从本仓库下载或克隆完整安装包后，在“终端”进入仓库目录，执行：

```bash
bash install.sh
```

脚本会先检查安装包、备份旧版，再安装到当前用户的 `~/.codex/skills`（如果设置了 `CODEX_HOME`，则使用该位置），最后核对安装结果。无需管理员权限。不要只下载单个脚本；它需要同目录的 `skills/` 文件夹。

## 安装内容

- `payroll-operations`：工资发放总流程、材料引导和人工确认节点。
- `attendance-calculator`：月度考勤计算与异常审计。
- `attendance-settlement`：考勤确认结果与工资结算衔接。
- `social-insurance-fund`：多公司五险一金汇总。

## 开始使用

重启 Codex 后输入：

> 请使用 payroll-operations，开始核算 YYYY 年 MM 月工资。月度资料在“<实际文件夹路径>”，正式模板在“<实际模板路径>”。请先盘点现有文件，告诉我当前阶段缺什么；名称或归属看不懂时直接问我。

建议将同月资料尽量放在一个总文件夹，文件名标明所属月份、公司或负责人、资料类型及确认状态；已有资料不必为了使用 Skill 重新搬运或改名。Codex 会先盘点当前阶段的文件，不会一次性要求提供全部工资资料。若从后续阶段开始，直接说明当前阶段即可。

## 文档

- [三位人资部署与月度操作教程（当前规则）](skills/payroll-operations/references/hr-deployment-guide.md)
- `docs/hr-payroll-tutorial.docx` 是旧版参考，使用前以当前 Skill 中的教程为准。
- [工资核算 Skill 完整内容](工资核算Skill完整内容.md)
- [在线安装命令](在线安装命令模板.txt)
- [发布说明](发布说明.md)

## 更新纪律

通用规则进入正式 Skill；员工姓名、固定金额、某个月工时和一次性例外只进入当月运行说明，避免影响以后月份。已人工确认的工资表、个税表和工资单是当月正式基准，未经明确要求不得重新计算覆盖。
