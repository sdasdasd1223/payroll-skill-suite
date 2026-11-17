param([string]$ArchiveFolder,[string]$CodexHome)
$ErrorActionPreference='Stop'
if(-not $CodexHome){$CodexHome=Join-Path $env:USERPROFILE '.codex'}
$adminHome=Join-Path $CodexHome 'payroll-diagnostics-admin'
if(-not $ArchiveFolder){
    $config=Get-Content -LiteralPath (Join-Path $adminHome 'archive-config.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $ArchiveFolder=$config.case_root
}
$entry=Join-Path $CodexHome 'payroll-diagnostics/case_admin.py'
if(-not (Test-Path -LiteralPath $entry)){throw '尚未安裝案件查看元件'}
$python=(Get-Command python -ErrorAction Stop).Source
$start=Get-Date
$arguments=@('-X','utf8',('"'+$entry+'"'),'view','--folder',('"'+$ArchiveFolder+'"'),'--home',('"'+$adminHome+'"'))
$process=Start-Process -FilePath $python -ArgumentList $arguments -WindowStyle Hidden -PassThru
$link=Join-Path $adminHome '案件查看入口.txt'
for($attempt=0;$attempt -lt 60;$attempt++){
    Start-Sleep -Milliseconds 250
    $process.Refresh()
    if($process.HasExited){throw '案件頁面未開啟，請先確認歸檔與管理員密鑰'}
    if((Test-Path -LiteralPath $link) -and (Get-Item -LiteralPath $link).LastWriteTime -ge $start){
        $url=(Get-Content -LiteralPath $link -Raw -Encoding UTF8).Trim()
        if($url -notmatch '^http://127\.0\.0\.1:\d+/session/[A-Za-z0-9_-]+/$'){throw '查看入口格式異常'}
        Start-Process -FilePath $url
        Write-Output '已開啟管理員本機案件頁面；入口一小時後失效。'
        exit 0
    }
}
throw '案件頁面尚未就緒，不能當成已開啟'
