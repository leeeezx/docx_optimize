param(
    [Parameter(Mandatory = $false)]
    [string]$TemplateName = "DocxOptimize.dotm"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-WordStartupPath {
    return Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
}

function Test-WordRunning {
    return $null -ne (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
}

$startupPath = Get-WordStartupPath
$installedPath = Join-Path $startupPath $TemplateName

if (Test-WordRunning) {
    throw "检测到 Word 正在运行，请先关闭所有 Word 窗口后再卸载模板。"
}

if (-not (Test-Path -LiteralPath $installedPath)) {
    Write-Host "未发现已安装模板：$installedPath"
    exit 0
}

$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupPath = "$installedPath.uninstalled.$stamp"
Move-Item -LiteralPath $installedPath -Destination $backupPath

Write-Host "已从 Startup 目录移除模板。"
Write-Host "备份文件：$backupPath"
Write-Host "如需恢复，手动将该文件改回 $TemplateName 即可。"
