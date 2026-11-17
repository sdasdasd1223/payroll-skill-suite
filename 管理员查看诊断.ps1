[CmdletBinding()]
param([string]$CodexHome = $(if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $HOME '.codex' }))
$ErrorActionPreference = 'Stop'
$adminHome = Join-Path $CodexHome 'payroll-diagnostics-admin'
if (-not (Test-Path -LiteralPath (Join-Path $adminHome 'admin-key.secure'))) { throw '这台电脑没有公司管理员解密密钥，不能查看诊断记录。' }
$output = Join-Path $adminHome ('诊断排查_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.html')
$script = Join-Path $CodexHome 'payroll-diagnostics/admin.py'
$env:PYTHONUTF8 = '1'
python $script sync --home $adminHome --output $output
if ($LASTEXITCODE -ne 0) { throw '云端记录同步未完成，未生成最新查看页。' }
Write-Host "管理员查看页：$output"
