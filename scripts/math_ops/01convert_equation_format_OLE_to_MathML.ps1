param(
    [string]$DocPath,
    [string]$OutPath,
    [string]$MacroName = "MTCommand_ConvertEqns",
    [switch]$ShowWord
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Select-WordFile {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Word 文件 (*.docx;*.docm)|*.docx;*.docm"
    $dlg.Title = "选择要处理的 Word 文件"
    $dlg.Multiselect = $false
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dlg.FileName
    }
    throw "未选择文件。"
}

function Get-OutputPath {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [string]$CustomOutPath
    )

    if ($CustomOutPath) {
        $fullOut = [System.IO.Path]::GetFullPath($CustomOutPath)
        $parent = Split-Path -Parent $fullOut
        if ($parent -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent | Out-Null
        }
        return $fullOut
    }

    $dir = Split-Path -Parent $SourcePath
    $name = [System.IO.Path]::GetFileNameWithoutExtension($SourcePath)
    $ext = [System.IO.Path]::GetExtension($SourcePath)

    $candidate = Join-Path $dir ($name + "_MTConvert" + $ext)
    $i = 1
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $dir ($name + "_MTConvert(" + $i + ")" + $ext)
        $i++
    }
    return $candidate
}

function Ensure-WinApi {
    if (-not ("WinApi.User32" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace WinApi {
    public static class User32 {
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll", SetLastError=true)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        public static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
    }

    public static class Kernel32 {
        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();
    }
}
"@
    }
}

function Get-WordWindowHandle {
    param(
        [Parameter(Mandatory = $true)]$WordApp
    )

    # 不同 Office 版本/COM 包装对句柄属性暴露不一致，逐个尝试
    try {
        $v = $WordApp.Hwnd
        if ($null -ne $v -and [int64]$v -ne 0) { return [IntPtr]$v }
    }
    catch {}

    try {
        $v = $WordApp.ActiveWindow.Hwnd
        if ($null -ne $v -and [int64]$v -ne 0) { return [IntPtr]$v }
    }
    catch {}

    try {
        $v = $WordApp.Windows.Item(1).Hwnd
        if ($null -ne $v -and [int64]$v -ne 0) { return [IntPtr]$v }
    }
    catch {}

    return [IntPtr]::Zero
}

function Bring-WordToFrontStrong {
    param(
        [Parameter(Mandatory = $true)]$WordApp
    )

    Ensure-WinApi

    $hWnd = Get-WordWindowHandle -WordApp $WordApp
    if ($hWnd -eq [IntPtr]::Zero) {
        # 回退：不依赖句柄，按窗口标题激活
        try {
            $ws = New-Object -ComObject WScript.Shell
            [void]$ws.AppActivate("Microsoft Word")
            Start-Sleep -Milliseconds 120
            return $true
        }
        catch {
            return $false
        }
    }

    $SW_MAXIMIZE = 3
    $SW_RESTORE = 9
    $SWP_NOSIZE = 0x0001
    $SWP_NOMOVE = 0x0002

    [void][WinApi.User32]::ShowWindowAsync($hWnd, $SW_RESTORE)
    Start-Sleep -Milliseconds 120
    [void][WinApi.User32]::ShowWindowAsync($hWnd, $SW_MAXIMIZE)
    Start-Sleep -Milliseconds 120

    $currentThreadId = [WinApi.Kernel32]::GetCurrentThreadId()
    $fgHwnd = [WinApi.User32]::GetForegroundWindow()
    $fgThreadId = 0
    if ($fgHwnd -ne [IntPtr]::Zero) {
        [uint32]$tmpPid = 0
        $fgThreadId = [WinApi.User32]::GetWindowThreadProcessId($fgHwnd, [ref]$tmpPid)
    }

    $attached = $false
    if ($fgThreadId -ne 0 -and $fgThreadId -ne $currentThreadId) {
        $attached = [WinApi.User32]::AttachThreadInput($currentThreadId, $fgThreadId, $true)
    }

    try {
        [void][WinApi.User32]::BringWindowToTop($hWnd)
        [void][WinApi.User32]::SetForegroundWindow($hWnd)

        # 临时置顶再取消，提升“可见且在最前”的成功率
        [void][WinApi.User32]::SetWindowPos($hWnd, [WinApi.User32]::HWND_TOPMOST, 0, 0, 0, 0, ($SWP_NOSIZE -bor $SWP_NOMOVE))
        [void][WinApi.User32]::SetWindowPos($hWnd, [WinApi.User32]::HWND_NOTOPMOST, 0, 0, 0, 0, ($SWP_NOSIZE -bor $SWP_NOMOVE))
    }
    finally {
        if ($attached) {
            [void][WinApi.User32]::AttachThreadInput($currentThreadId, $fgThreadId, $false)
        }
    }

    Start-Sleep -Milliseconds 120
    $nowFg = [WinApi.User32]::GetForegroundWindow()
    if ($nowFg -eq $hWnd) {
        return $true
    }

    # 回退：按 Word 进程 PID 激活
    [uint32]$wordPid = 0
    [void][WinApi.User32]::GetWindowThreadProcessId($hWnd, [ref]$wordPid)
    if ($wordPid -ne 0) {
        try {
            $ws = New-Object -ComObject WScript.Shell
            [void]$ws.AppActivate([int]$wordPid)
            Start-Sleep -Milliseconds 120
        }
        catch {}
    }

    return ([WinApi.User32]::GetForegroundWindow() -eq $hWnd)
}

function Release-ComObject {
    param($ComObject)
    if ($null -ne $ComObject) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
    }
}

if (-not $DocPath) {
    try {
        $DocPath = Select-WordFile
    }
    catch {
        $DocPath = Read-Host "未弹出文件选择器，请输入 Word 文件完整路径"
    }
}

if (-not (Test-Path -LiteralPath $DocPath)) {
    throw "文件不存在：$DocPath"
}

$extCheck = [System.IO.Path]::GetExtension($DocPath).ToLowerInvariant()
if ($extCheck -notin @('.docx', '.docm')) {
    throw "仅支持 .docx 或 .docm，当前文件：$DocPath"
}

$source = (Resolve-Path -LiteralPath $DocPath).Path
$target = Get-OutputPath -SourcePath $source -CustomOutPath $OutPath

Copy-Item -LiteralPath $source -Destination $target -Force
Write-Host "已创建副本：$target"

$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Open($target, $false, $false)
    [void]$doc.Activate()

    $toFrontOk = Bring-WordToFrontStrong -WordApp $word
    if ($toFrontOk) {
        Write-Host "Word 已尝试置前。"
    }
    else {
        Write-Host "提示：Windows 前台策略仍可能阻止置前，请手动点击任务栏 Word。"
    }

    Write-Host "开始执行宏：$MacroName"
    Write-Host "请在 Word 弹窗中手动确认。"
    [void]$word.Run($MacroName)

    $doc.Save()
    Write-Host "执行完成，输出文件：$target"
}
catch {
    throw ("执行失败：" + $_.Exception.Message)
}
finally {
    if ($doc -ne $null) {
        try { [void]$doc.Close($true) } catch {}
    }
    if ($word -ne $null) {
        try { [void]$word.Quit() } catch {}
    }

    Release-ComObject -ComObject $doc
    Release-ComObject -ComObject $word

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
