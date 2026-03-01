param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$DocxPath,

    [Parameter(Mandatory = $false)]
    [string]$OutPath,

    [Parameter(Mandatory = $false)]
    [string]$MacroName = "PlainMathMLToEquation",

    [Parameter(Mandatory = $false)]
    [switch]$ShowWord
)

function Select-WordFile {
    try {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = "选择要处理的 Word 文档"
        $dialog.Filter = "Word 文档 (*.docx;*.docm)|*.docx;*.docm|所有文件 (*.*)|*.*"
        $dialog.Multiselect = $false
        $result = $dialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.FileName
        }
        return $null
    }
    catch {
        Write-Host "无法弹出文件选择窗口，请手动输入文件路径："
        $manualPath = Read-Host "DocxPath"
        if ([string]::IsNullOrWhiteSpace($manualPath)) {
            return $null
        }
        return $manualPath
    }
}

function Resolve-InputPath([string]$InputPath) {
    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        return $null
    }
    try {
        return (Resolve-Path -LiteralPath $InputPath).Path
    }
    catch {
        return $null
    }
}

function Get-OutputPath([string]$SourcePath, [string]$UserOutPath) {
    if (-not [string]::IsNullOrWhiteSpace($UserOutPath)) {
        return [System.IO.Path]::GetFullPath($UserOutPath)
    }

    $dir = [System.IO.Path]::GetDirectoryName($SourcePath)
    $name = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $ext = [System.IO.Path]::GetExtension($SourcePath)

    $candidate = Join-Path $dir ($name + "_OMML" + $ext)
    if (-not (Test-Path -LiteralPath $candidate)) {
        return $candidate
    }

    $i = 1
    do {
        $candidate = Join-Path $dir ($name + "_OMML(" + $i + ")" + $ext)
        $i++
    } while (Test-Path -LiteralPath $candidate)

    return $candidate
}

if ([string]::IsNullOrWhiteSpace($DocxPath)) {
    $DocxPath = Select-WordFile
}

$fullPath = Resolve-InputPath $DocxPath
if ([string]::IsNullOrWhiteSpace($fullPath)) {
    Write-Error "未找到有效的文档路径，脚本已退出。"
    exit 1
}

$ext = [System.IO.Path]::GetExtension($fullPath).ToLowerInvariant()
if ($ext -ne ".docx" -and $ext -ne ".docm") {
    Write-Error "仅支持 .docx 或 .docm 文件：$fullPath"
    exit 1
}

$targetPath = Get-OutputPath -SourcePath $fullPath -UserOutPath $OutPath
$targetExt = [System.IO.Path]::GetExtension($targetPath).ToLowerInvariant()
if ($targetExt -ne ".docx" -and $targetExt -ne ".docm") {
    Write-Error "输出文件仅支持 .docx 或 .docm：$targetPath"
    exit 1
}

if ([System.IO.Path]::GetFullPath($targetPath).Trim().ToLowerInvariant() -eq $fullPath.Trim().ToLowerInvariant()) {
    Write-Error "输出路径不能与原文件相同，请修改 -OutPath。"
    exit 1
}

$word = $null
$doc = $null

try {
    Copy-Item -LiteralPath $fullPath -Destination $targetPath -Force
    Write-Host "已创建副本：$targetPath"

    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$ShowWord
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Open($targetPath)

    Write-Host "正在执行宏：$MacroName"
    $null = $word.Run($MacroName)

    $doc.Save()
    Write-Host "处理完成（原文件未修改）：$targetPath"
    exit 0
}
catch {
    Write-Error ("处理失败：{0}" -f $_.Exception.Message)
    exit 2
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close($true) | Out-Null } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
    }
    if ($word -ne $null) {
        try { $word.Quit() | Out-Null } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
