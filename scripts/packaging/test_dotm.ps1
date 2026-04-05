param(
    [Parameter(Mandatory = $false)]
    [string]$TemplateName = "DocxOptimize.dotm",

    [Parameter(Mandatory = $false)]
    [string]$MacroName = "PlainMathMLToEquation"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-WordStartupPath {
    return Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
}

function Resolve-MacroQualifiedName([string]$Name, [string]$Template) {
    if ($Name -match "!") {
        return $Name
    }
    return "$Template!$Name"
}

function Test-WordRunning {
    return $null -ne (Get-Process -Name WINWORD -ErrorAction SilentlyContinue)
}

function Test-TemplateLoaded {
    param(
        [Parameter(Mandatory = $true)]$WordApp,
        [Parameter(Mandatory = $true)][string]$TemplatePath
    )

    $normalizedTemplatePath = [System.IO.Path]::GetFullPath($TemplatePath).Trim().ToLowerInvariant()

    try {
        foreach ($addIn in $WordApp.AddIns) {
            try {
                if (-not [string]::IsNullOrWhiteSpace($addIn.Path) -and -not [string]::IsNullOrWhiteSpace($addIn.Name)) {
                    $candidatePath = [System.IO.Path]::Combine($addIn.Path, $addIn.Name)
                    if ([System.IO.Path]::GetFullPath($candidatePath).Trim().ToLowerInvariant() -eq $normalizedTemplatePath) {
                        return [bool]$addIn.Installed
                    }
                }
            }
            catch {}
        }
    }
    catch {}

    try {
        foreach ($template in $WordApp.Templates) {
            try {
                if (-not [string]::IsNullOrWhiteSpace($template.FullName)) {
                    if ([System.IO.Path]::GetFullPath($template.FullName).Trim().ToLowerInvariant() -eq $normalizedTemplatePath) {
                        return $true
                    }
                }
            }
            catch {}
        }
    }
    catch {}

    return $false
}

$startupPath = Get-WordStartupPath
$installedPath = Join-Path $startupPath $TemplateName
$macroQualifiedName = Resolve-MacroQualifiedName -Name $MacroName -Template $TemplateName

if (-not (Test-Path -LiteralPath $installedPath)) {
    Write-Error @"
未发现模板：$installedPath
请先执行安装脚本：
pwsh -ExecutionPolicy Bypass -File scripts\packaging\install_dotm.ps1
"@
    exit 1
}

if (Test-WordRunning) {
    Write-Warning "检测到 Word 已在运行。为避免受现有会话影响，建议先关闭 Word 再执行自检。"
}

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    Start-Sleep -Milliseconds 300

    if (Test-TemplateLoaded -WordApp $word -TemplatePath $installedPath) {
        Write-Host "已确认 Word 加载模板：$installedPath"
    }
    else {
        Write-Warning "未能在加载项列表中确认模板，接下来将直接尝试执行宏。"
    }

    $doc = $word.Documents.Add()
    $null = $word.Run($macroQualifiedName)

    Write-Host "自检通过：宏可被 Word 调用。"
    Write-Host "宏名：$macroQualifiedName"
    exit 0
}
catch {
    Write-Error @"
自检失败：$($_.Exception.Message)

请检查：
1. Word 是否启用了宏内容
2. 模板文件是否已解除锁定
3. Startup 目录中的模板是否为最新版本
4. 宏名是否存在：$macroQualifiedName
"@
    exit 2
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close($false) | Out-Null } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    }
    if ($word -ne $null) {
        try { $word.Quit() | Out-Null } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
