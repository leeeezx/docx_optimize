param(
    [Parameter(Mandatory = $false)]
    [string]$TemplatePath,

    [Parameter(Mandatory = $false)]
    [string]$TemplateName = "DocxOptimize.dotm",

    [Parameter(Mandatory = $false)]
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-WordStartupPath {
    return Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
}

function Test-WordRunning {
    return $null -ne (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
}

function Resolve-TemplateSourcePath {
    param(
        [string]$UserTemplatePath,
        [string]$Name
    )

    if (-not [string]::IsNullOrWhiteSpace($UserTemplatePath)) {
        if (-not (Test-Path -LiteralPath $UserTemplatePath)) {
            throw "指定的模板不存在：$UserTemplatePath"
        }
        return (Resolve-Path -LiteralPath $UserTemplatePath).Path
    }

    $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path
    $candidates = @(
        (Join-Path $repoRoot "dist\$Name"),
        (Join-Path $repoRoot $Name),
        (Join-Path $PSScriptRoot $Name)
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    throw @"
未找到可安装的模板文件 $Name。
请将模板放到以下任一位置，或显式传入 -TemplatePath：
1. dist\$Name
2. 仓库根目录\$Name
3. scripts\packaging\$Name
"@
}

$startupPath = Get-WordStartupPath
$sourcePath = Resolve-TemplateSourcePath -UserTemplatePath $TemplatePath -Name $TemplateName
$destinationPath = Join-Path $startupPath $TemplateName

if (Test-WordRunning) {
    throw "检测到 Word 正在运行，请先关闭所有 Word 窗口后再安装模板。"
}

if (-not (Test-Path -LiteralPath $startupPath)) {
    New-Item -ItemType Directory -Path $startupPath | Out-Null
}

try {
    Unblock-File -LiteralPath $sourcePath -ErrorAction SilentlyContinue
}
catch {}

if ((Test-Path -LiteralPath $destinationPath) -and -not $Force) {
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupPath = "$destinationPath.bak.$stamp"
    Move-Item -LiteralPath $destinationPath -Destination $backupPath
    Write-Host "已备份旧模板：$backupPath"
}

Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force

try {
    Unblock-File -LiteralPath $destinationPath -ErrorAction SilentlyContinue
}
catch {}

Write-Host "安装完成：$destinationPath"
Write-Host "下一步："
Write-Host "1. 启动 Word 并按提示启用宏内容。"
Write-Host "2. 运行自检脚本：pwsh -ExecutionPolicy Bypass -File scripts\packaging\test_dotm.ps1"
Write-Host "3. 若需要完整流程，再运行 scripts\math_ops\03run_equation_pipeline.py"
