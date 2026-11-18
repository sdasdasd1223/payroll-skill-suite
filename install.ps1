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

    $hashFile = Join-Path $SkillPath "checksums.sha256"
    if (-not (Test-Path -LiteralPath $hashFile -PathType Leaf)) { throw "安装包缺少完整性清单：$ExpectedName" }
    foreach ($line in (Get-Content -LiteralPath $hashFile -Encoding UTF8)) {
        if ($line.Length -lt 67) { throw "完整性清单格式错误" }
        $relative = $line.Substring(66)
        $checkedPath = [IO.Path]::GetFullPath((Join-Path $SkillPath $relative))
        $allowedRoot = [IO.Path]::GetFullPath($SkillPath).TrimEnd('\') + '\'
        if (-not $checkedPath.StartsWith($allowedRoot, [StringComparison]::OrdinalIgnoreCase)) { throw "完整性清单路径越界" }
        if (-not (Test-Path -LiteralPath $checkedPath -PathType Leaf)) { throw "安装包缺文件：$relative" }
        if ((Get-FileHash -LiteralPath $checkedPath -Algorithm SHA256).Hash -ne $line.Substring(0,64)) { throw "安装包文件摘要不一致：$relative" }
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
        $resolvedDestination = [IO.Path]::GetFullPath($destination)
        $allowedSkillsRoot = [IO.Path]::GetFullPath($skillsRoot).TrimEnd('\') + '\'
        if (-not $resolvedDestination.StartsWith($allowedSkillsRoot, [StringComparison]::OrdinalIgnoreCase)) { throw "安装目标路径越界" }
        Remove-Item -LiteralPath $destination -Recurse -Force
    }
    Copy-Item -LiteralPath (Join-Path $sourceRoot $name) -Destination $destination -Recurse -Force
    Write-Host "已安装：$name"
}

Write-Step "核对安装结果"
foreach ($name in $skillNames) {
    Assert-SkillPackage -SkillPath (Join-Path $skillsRoot $name) -ExpectedName $name
}

& python -X utf8 (Join-Path $packageRoot 'remove_automatic_records.py') --codex-home $CodexHome
if ($LASTEXITCODE -ne 0) { throw '舊記錄元件尚未完整停用' }

Write-Host "`n安装成功。" -ForegroundColor Green
if ($existing.Count -gt 0) {
    Write-Host "旧版本备份位置：$backupRoot"
}
Write-Host "请完全退出并重新启动 Codex，重启后 Skill 才会生效。" -ForegroundColor Yellow
