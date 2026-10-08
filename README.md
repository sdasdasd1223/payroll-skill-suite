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

- [三位人资部署与月度操作教程](docs/hr-payroll-tutorial.docx)
- [工资核算 Skill 完整内容](工资核算Skill完整内容.md)
- [在线安装命令](在线安装命令模板.txt)
- [发布说明](发布说明.md)

## 更新纪律

通用规则进入正式 Skill；员工姓名、固定金额、某个月工时和一次性例外只进入当月运行说明，避免影响以后月份。已人工确认的工资表、个税表和工资单是当月正式基准，未经明确要求不得重新计算覆盖。
