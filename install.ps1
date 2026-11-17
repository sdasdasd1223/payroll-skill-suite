[CmdletBinding()]
param(
    [string]$CodexHome = $(if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $HOME '.codex' }),
    [string]$DeviceProfile,
    [switch]$SkipDiagnosticSchedule,
    [switch]$CheckOnly
)
$ErrorActionPreference = 'Stop'
$packageRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonCommand = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCommand) { throw '未找到Python，请让Codex先配置Python 3.10或以上再执行安装。' }
$arguments = @((Join-Path $packageRoot 'install_suite.py'), '--codex-home', $CodexHome)
if ($DeviceProfile) { $arguments += @('--device-profile', $DeviceProfile) }
if ($SkipDiagnosticSchedule) { $arguments += '--skip-diagnostic-schedule' }
if ($CheckOnly) { $arguments += '--check-only' }
$env:PYTHONUTF8 = '1'
& $pythonCommand.Source @arguments
if ($LASTEXITCODE -ne 0) { throw '安装未完成，未把诊断组件或云端同步记为成功；请管理员检查安装状态。' }
