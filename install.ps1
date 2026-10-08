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
