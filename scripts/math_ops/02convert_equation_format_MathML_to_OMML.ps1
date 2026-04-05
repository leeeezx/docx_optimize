param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string]$DocxPath,

    [Parameter(Mandatory = $false)]
    [string]$OutPath,

    [Parameter(Mandatory = $false)]
    [string]$MacroName = "PlainMathMLToEquation",

    [Parameter(Mandatory = $false)]
    [string]$TemplateName = "DocxOptimize.dotm",

    [Parameter(Mandatory = $false)]
    [switch]$ShowWord
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-InfoMessage([string]$Message) {
    Write-Host $Message -ForegroundColor Cyan
}

function Write-SuccessMessage([string]$Message) {
    Write-Host $Message -ForegroundColor Green
}

function Write-NoticeMessage([string]$Message) {
    Write-Host $Message -ForegroundColor Yellow
}

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
        Write-NoticeMessage "无法弹出文件选择窗口，请手动输入文件路径："
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

function Get-WordStartupPath {
    return Join-Path $env:APPDATA "Microsoft\Word\STARTUP"
}

function Get-InstalledTemplatePath([string]$Name) {
    return Join-Path (Get-WordStartupPath) $Name
}

function Resolve-MacroQualifiedName([string]$Name, [string]$Template) {
    if ($Name -match "!") {
        return $Name
    }
    return "$Template!$Name"
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

function Get-InstallHint([string]$Template) {
    return (
        @(
            "未检测到已安装的全局宏模板 $Template。"
            "请先执行安装脚本："
            "pwsh -ExecutionPolicy Bypass -File scripts\packaging\install_dotm.ps1"
            ""
            "若已安装但仍失败，请检查："
            "1. Word 是否已启用宏内容"
            "2. 文件是否已解除锁定"
            "3. Startup 目录中的模板是否为最新版本"
        ) -join [Environment]::NewLine
    )
}

function Get-MathMLMatchCount {
    param(
        [Parameter(Mandatory = $true)]$Document
    )

    $range = $Document.Content.Duplicate
    $count = 0
    $find = $range.Find

    $find.ClearFormatting()
    $find.Text = "\<math?*\</math\>^13"
    $find.Replacement.Text = ""
    $find.Forward = $true
    $find.Wrap = 0
    $find.Format = $false
    $find.MatchCase = $false
    $find.MatchWholeWord = $false
    $find.MatchByte = $false
    $find.MatchAllWordForms = $false
    $find.MatchSoundsLike = $false
    $find.MatchWildcards = $true

    while ($find.Execute()) {
        $count++
        $range.Collapse(0)
    }

    return $count
}

function Test-IsLikelyMathTypeProgId([string]$ProgId) {
    if ([string]::IsNullOrWhiteSpace($ProgId)) {
        return $false
    }
    return ($ProgId -match '(?i)equation\.dsmt\d*') -or ($ProgId -match '(?i)mathtype')
}

function Get-SafeProgId($ShapeObject) {
    try {
        return [string]$ShapeObject.OLEFormat.ProgID
    }
    catch {
        return ""
    }
}

function Get-LikelyMathTypeOleCount {
    param(
        [Parameter(Mandatory = $true)]$Document
    )

    $count = 0

    try {
        for ($i = 1; $i -le $Document.InlineShapes.Count; $i++) {
            $progId = Get-SafeProgId -ShapeObject $Document.InlineShapes.Item($i)
            if (Test-IsLikelyMathTypeProgId -ProgId $progId) {
                $count++
            }
        }
    }
    catch {}

    try {
        for ($i = 1; $i -le $Document.Shapes.Count; $i++) {
            $progId = Get-SafeProgId -ShapeObject $Document.Shapes.Item($i)
            if (Test-IsLikelyMathTypeProgId -ProgId $progId) {
                $count++
            }
        }
    }
    catch {}

    return $count
}

function Get-MathMLMissingHint([int]$LikelyOleCount) {
    $lines = @(
        '当前文档中未检测到可转换的 MathML 文本片段。'
        '`02convert_equation_format_MathML_to_OMML.ps1` 只处理形如 `<math ...>...</math>` 的文本，不处理 MathType OLE 对象。'
        ''
    )

    if ($LikelyOleCount -gt 0) {
        $lines += "另外，当前文档中检测到疑似 MathType OLE 对象，数量=$LikelyOleCount。"
        $lines += '这说明你传入的很可能仍是 OLE 原文档，而不是已经转成 MathML 的中间文档。'
        $lines += ''
    }

    $lines += @(
        '若你的输入文档仍是 OLE 公式，请改用以下任一方式：'
        '1. 先执行 `scripts\math_ops\01convert_equation_format_OLE_to_MathML.ps1`'
        '2. 直接执行 `python scripts\math_ops\03run_equation_pipeline.py <docx>`'
    )

    return $lines -join [Environment]::NewLine
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

$templatePath = Get-InstalledTemplatePath -Name $TemplateName
if (-not (Test-Path -LiteralPath $templatePath)) {
    Write-Error ((Get-InstallHint -Template $TemplateName).Trim())
    exit 1
}

$macroQualifiedName = Resolve-MacroQualifiedName -Name $MacroName -Template $TemplateName

$word = $null
$doc = $null

try {
    Copy-Item -LiteralPath $fullPath -Destination $targetPath -Force
    Write-InfoMessage "已创建副本：$targetPath"
    Write-InfoMessage "检测到全局模板：$templatePath"

    $word = New-Object -ComObject Word.Application
    $word.Visible = [bool]$ShowWord
    $word.DisplayAlerts = 0
    Start-Sleep -Milliseconds 300

    if (Test-TemplateLoaded -WordApp $word -TemplatePath $templatePath) {
        Write-SuccessMessage "Word 已加载模板：$TemplateName"
    }
    else {
        Write-Warning "未能在 Word 加载项列表中确认 $TemplateName，后续将继续尝试执行宏。若失败，请重新运行安装脚本。"
    }

    $doc = $word.Documents.Open($targetPath)
    $mathMLMatchCount = Get-MathMLMatchCount -Document $doc
    if ($mathMLMatchCount -le 0) {
        $likelyOleCount = Get-LikelyMathTypeOleCount -Document $doc
        throw (Get-MathMLMissingHint -LikelyOleCount $likelyOleCount).Trim()
    }

    Write-InfoMessage "检测到可转换 MathML 段数：$mathMLMatchCount"

    Write-NoticeMessage "正在执行宏：$macroQualifiedName"
    try {
        $null = $word.Run($macroQualifiedName)
    }
    catch {
        $detail = $_.Exception.Message
        throw "执行宏失败：$detail`n`n$((Get-InstallHint -Template $TemplateName).Trim())"
    }

    $doc.Save()
    Write-SuccessMessage "处理完成（原文件未修改）：$targetPath"
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
