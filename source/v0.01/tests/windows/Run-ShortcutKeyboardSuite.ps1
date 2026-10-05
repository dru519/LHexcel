param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RunId = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }
if (-not [Environment]::UserInteractive) { throw 'Shortcut keyboard suite requires an interactive Windows desktop' }

$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$ProductXlam = [IO.Path]::GetFullPath($ProductXlam)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not (Test-Path -LiteralPath $ProductXlam -PathType Leaf)) { throw 'Product.xlam not found' }
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing ShortcutKeyboard evidence root is rejected' }
if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -ne 0) { throw 'Shortcut keyboard suite requires zero pre-existing Excel processes' }
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
Add-Type -AssemblyName System.Windows.Forms

if (-not ('NxShortcutKeyboardWin32' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class NxShortcutKeyboardWin32 {
  public delegate bool EnumProc(IntPtr hwnd, IntPtr state);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, IntPtr state);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder value, int maximum);
  [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
  [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr hwnd);
  [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hwnd, int command);
  [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
  public static long FindVisibleTopLevelWindow(int wantedPid, string wantedTitle) {
    long found=0;
    EnumWindows((hwnd,state)=>{uint pid;GetWindowThreadProcessId(hwnd,out pid);if(pid!=(uint)wantedPid||!IsWindowVisible(hwnd))return true;var text=new StringBuilder(1024);GetWindowText(hwnd,text,text.Capacity);if(String.Equals(text.ToString(),wantedTitle,StringComparison.Ordinal)){found=hwnd.ToInt64();return false;}return true;},IntPtr.Zero);
    return found;
  }
  public static bool Activate(long hwnd) { var target=new IntPtr(hwnd); ShowWindow(target,5); BringWindowToTop(target); return SetForegroundWindow(target); }
  public static long Foreground() { return GetForegroundWindow().ToInt64(); }
}
'@
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}

function Write-AtomicJson([string]$Path, [object]$Value) {
    $temporary = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    $backup = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.bak'
    $encoding = New-Object Text.UTF8Encoding($false)
    try {
        [IO.File]::WriteAllText($temporary, (($Value | ConvertTo-Json -Depth 20) + "`n"), $encoding)
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            [IO.File]::Replace($temporary, $Path, $backup)
        } else {
            [IO.File]::Move($temporary, $Path)
        }
    } finally {
        if (Test-Path -LiteralPath $temporary -PathType Leaf) { [IO.File]::Delete($temporary) }
        if (Test-Path -LiteralPath $backup -PathType Leaf) {
            if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { [IO.File]::Move($backup, $Path) }
            else { [IO.File]::Delete($backup) }
        }
    }
}

function Flush-CaseProgress([string]$Stage) {
    if ([string]::IsNullOrWhiteSpace([string]$script:caseProgressPath)) { return }
    Write-AtomicJson $script:caseProgressPath ([ordered]@{
        schema='lhexcel-shortcut-keyboard-progress-v1';run_id=$RunId;stage=$Stage;failure_stage=$script:failureStage
        cases=$script:cases.ToArray();written_utc=[DateTime]::UtcNow.ToString('o')
    })
}

function Get-FreshExcelProcessIds {
    $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    try { return @($processes | ForEach-Object { [int]$_.Id }) }
    finally { foreach ($process in $processes) { try { $process.Dispose() } catch {} } }
}

function Wait-FreshExcelProcessExit([int]$TimeoutMs = 2000) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $ids = @()
    do {
        $ids = @(Get-FreshExcelProcessIds)
        if ($ids.Count -eq 0) { return [pscustomobject][ordered]@{status='CLEARED';elapsed_ms=$watch.ElapsedMilliseconds;remaining_pids=@()} }
        if ($watch.ElapsedMilliseconds -lt $TimeoutMs) { Start-Sleep -Milliseconds 100 }
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    return [pscustomobject][ordered]@{status='REMAINING';elapsed_ms=$watch.ElapsedMilliseconds;remaining_pids=$ids}
}

function Wait-OwnedFormWindow([int]$ProcessId, [string]$Caption, [int]$TimeoutMs = 5000) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $handle = [NxShortcutKeyboardWin32]::FindVisibleTopLevelWindow($ProcessId, $Caption)
        if ($handle -ne 0) { return $handle }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    throw ('Shortcut manager form did not appear for owned EXCEL PID ' + $ProcessId)
}

function Assert-ActivateWindow([Int64]$Handle, [string]$Stage) {
    if (-not [NxShortcutKeyboardWin32]::Activate($Handle)) { throw ($Stage + ': owned window activation rejected') }
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        if ([NxShortcutKeyboardWin32]::Foreground() -eq $Handle) { return }
        Start-Sleep -Milliseconds 50
    } while ($watch.ElapsedMilliseconds -lt 2000)
    throw ($Stage + ': owned window did not become foreground')
}

function Send-PhysicalKeys([string]$Keys, [Int64]$ExpectedForeground, [string]$Stage) {
    if ([NxShortcutKeyboardWin32]::Foreground() -ne $ExpectedForeground) {
        throw ($Stage + ': foreground ownership changed immediately before keyboard delivery')
    }
    try { [Windows.Forms.SendKeys]::SendWait($Keys) }
    catch { throw ($Stage + ': keyboard delivery failed: ' + $_.Exception.Message) }
    Start-Sleep -Milliseconds 150
}

function Focus-FormControlByTab([object]$Excel, [string]$Macro, [Int64]$FormHandle, [string]$TargetControl, [string]$Stage) {
    for ($attempt = 0; $attempt -lt 7; $attempt++) {
        $active = [string]$Excel.Run($Macro + 'NxKeyboardBridgeActiveControlName')
        if ($active -ceq $TargetControl) { return $active }
        Send-PhysicalKeys '{TAB}' $FormHandle $Stage
    }
    throw ($Stage + ': physical Tab navigation did not reach ' + $TargetControl + '; active=' + [string]$Excel.Run($Macro + 'NxKeyboardBridgeActiveControlName'))
}

function Invoke-ShortcutKeyboardVbaCompile([object]$Excel, [object]$Binding, [object]$Addin, [string]$CopyPath) {
    $vbe = $null; $bars = $null; $compileControl = $null; $project = $null; $components = $null; $compileTarget = $null; $activeProject = $null; $mainWindow = $null; $compileWatcher = $null
    $script:compileEvidence = [ordered]@{status='RUNNING';control_id=578;active_project=$null;dialog=$null;failure=$null}
    try {
        $vbe = $Excel.VBE
        $bars = $vbe.CommandBars
        $compileControl = Find-VbeCompileControl $bars
        if ($null -eq $compileControl -or [int]$compileControl.Id -ne 578) { throw 'Shortcut keyboard VBA compile control 578 unavailable' }
        $project = $Addin.VBProject
        $components = $project.VBComponents
        $compileTarget = $components.Item('NxShortcutKeyboardBridge')
        [void]$compileTarget.Activate()
        $activeProject = $vbe.ActiveVBProject
        $activeProjectPath = [IO.Path]::GetFullPath([string]$activeProject.FileName)
        $script:compileEvidence.active_project = $activeProjectPath
        if ($activeProjectPath -cne [IO.Path]::GetFullPath($CopyPath)) { throw 'Shortcut keyboard VBA compile active project identity mismatch' }
        if (-not [bool]$compileControl.Enabled) { throw 'Shortcut keyboard VBA compile control was disabled before explicit compile' }
        $mainWindow = $vbe.MainWindow
        $compileWatcher = Start-ExactOwnedCompileDialogWatcher -ExcelPid ([int]$Binding.pid) -ExcelHwnd ([Int64]$Binding.hwnd) -VbeHwnd ([Int64]$mainWindow.HWnd)
        $compileExecuteError = $null
        try { [void]$compileControl.Execute() } catch { $compileExecuteError = $_.Exception.Message }
        $watchCompleted = Wait-Job -Job $compileWatcher -Timeout 30
        if ($null -eq $watchCompleted) { Stop-Job -Job $compileWatcher -ErrorAction SilentlyContinue; throw 'Shortcut keyboard VBA compile dialog watcher timed out' }
        $watchResult = @(Receive-Job -Job $compileWatcher -Wait -AutoRemoveJob -ErrorAction Stop | Select-Object -First 1)
        $compileWatcher = $null
        if (-not [string]::IsNullOrWhiteSpace([string]$compileExecuteError)) { throw ('Shortcut keyboard VBA compile 578 execution failed: ' + $compileExecuteError) }
        if ($watchResult.Count -ne 1 -or [string]$watchResult[0].status -ne 'NO_DIALOG') {
            $dialogText = if ($watchResult.Count -eq 1) { [string]$watchResult[0].dialog_text } else { 'watcher returned no result' }
            $script:compileEvidence.dialog = $dialogText
            throw ('Shortcut keyboard VBA compile dialog detected: ' + $dialogText)
        }
        if ([bool]$compileControl.Enabled) { throw 'Shortcut keyboard VBA compile did not reach saved disabled state' }
        $script:compileEvidence.status = 'PASS'
    } catch {
        $script:compileEvidence.status = 'FAIL'
        $script:compileEvidence.failure = $_.Exception.Message
        throw
    } finally {
        if ($null -ne $compileWatcher) { try { Stop-Job -Job $compileWatcher -ErrorAction SilentlyContinue; Remove-Job -Job $compileWatcher -Force -ErrorAction SilentlyContinue } catch {} }
        Release-ComObject $mainWindow; Release-ComObject $activeProject; Release-ComObject $compileTarget; Release-ComObject $components; Release-ComObject $project; Release-ComObject $compileControl; Release-ComObject $bars; Release-ComObject $vbe
    }
}

function Cleanup-OwnedBridgeForm([object]$Excel, [string]$Macro, [object]$Binding, [string]$Caption) {
    $result = [ordered]@{status='NOT_VISIBLE';failure=$null;handle=0}
    try {
        if ($null -eq $Excel -or [string]::IsNullOrWhiteSpace($Macro) -or $null -eq $Binding -or -not [bool]$Binding.owned -or [string]::IsNullOrWhiteSpace($Caption)) { return [pscustomobject]$result }
        $handle = [NxShortcutKeyboardWin32]::FindVisibleTopLevelWindow([int]$Binding.pid, $Caption)
        if ($handle -eq 0) { return [pscustomobject]$result }
        $result.handle = $handle
        Assert-ActivateWindow $handle 'cleanup_owned_bridge_form'
        [void]$Excel.Run($Macro + 'NxKeyboardBridgeCleanupUnload')
        $watch = [Diagnostics.Stopwatch]::StartNew()
        do {
            if ([NxShortcutKeyboardWin32]::FindVisibleTopLevelWindow([int]$Binding.pid, $Caption) -eq 0) { $result.status = 'UNLOADED'; return [pscustomobject]$result }
            Start-Sleep -Milliseconds 100
        } while ($watch.ElapsedMilliseconds -lt 3000)
        throw 'Owned shortcut manager form survived bridge cleanup unload'
    } catch {
        $result.status = 'FAILED'
        $result.failure = $_.Exception.Message
        return [pscustomobject]$result
    }
}

function Wait-NumberFormat([object]$Range, [string]$Expected, [int]$TimeoutMs = 3000) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $actual = ''
    do {
        $actual = [string]$Range.NumberFormat
        if ($actual -ceq $Expected) { return $actual }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    return $actual
}

function Add-Case([string]$Name, [bool]$Passed, [object]$Details, [bool]$AbortOnFailure = $true) {
    $script:cases.Add([pscustomobject][ordered]@{name=$Name;status=if($Passed){'PASS'}else{'FAIL'};details=$Details})
    Flush-CaseProgress $Name
    if (-not $Passed -and $AbortOnFailure) { throw ('Native keyboard case failed: ' + $Name) }
}

function Complete-ExcelPhase([object]$Excel, [object]$Addin, [object]$Book, [object]$Sheet, [object]$Range, [object]$Binding, [string]$Label) {
    $result = [ordered]@{label=$Label;exit_mode='NOT_STARTED';failure=$null;owned_pid=0}
    try {
        if ($null -ne $Book) { try { $Book.Close($false) } catch {} }
        if ($null -ne $Addin) { try { $Addin.Close($false) } catch {} }
        Release-ComObject $Range
        Release-ComObject $Sheet
        Release-ComObject $Book
        Release-ComObject $Addin
        if ($null -ne $Excel) { try { $Excel.Quit() } catch {} }
        Release-ComObject $Excel
        [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        if ($null -eq $Binding -or -not [bool]$Binding.owned -or $null -eq $Binding.process) { throw ($Label + ' has no exact owned Excel process for lifecycle verification') }
        $result.owned_pid = [int]$Binding.pid
        $stop = Stop-ExactProcessAfterGrace $Binding.process $Label 15000 -Detailed
        $result.exit_mode = [string]$stop.exit_mode
        $result.failure = $stop.failure
        if ($result.exit_mode -cne 'NATURAL' -or -not [string]::IsNullOrWhiteSpace([string]$result.failure)) { throw ($Label + ' did not exit naturally: ' + $result.exit_mode + '; ' + [string]$result.failure) }
    } catch {
        if ([string]::IsNullOrWhiteSpace([string]$result.failure)) { $result.failure = $_.Exception.Message }
    } finally {
        if ($null -ne $Binding -and $null -ne $Binding.process) { try { $Binding.process.Dispose() } catch {} }
    }
    return [pscustomobject]$result
}

$bridgeCode = @'
Option Explicit

Private gManager As Object

Private Function Manager() As Object
    If gManager Is Nothing Then Err.Raise vbObjectError + 481, "NxKeyboardBridge", "Shortcut manager is not open"
    Set Manager = gManager
End Function

Public Sub NxKeyboardBridgeShow()
    Set gManager = New FNxShortcutManager
    gManager.Show vbModeless
End Sub

Public Function NxKeyboardBridgeCaption() As String
    NxKeyboardBridgeCaption = CStr(Manager.Caption)
End Function

Public Function NxKeyboardBridgeControlNames() As String
    Dim control As Object
    Dim result As String
    For Each control In Manager.Controls
        If Len(result) > 0 Then result = result & "|"
        result = result & CStr(control.Name)
    Next control
    NxKeyboardBridgeControlNames = result
End Function

Public Function NxKeyboardBridgeSelectRoute(ByVal routeKey As String) As Long
    Dim list As Object
    Dim index As Long
    Set list = Manager.Controls.Item("lstCommands")
    For index = 0 To list.ListCount - 1
        If StrComp(CStr(list.List(index, 0)), routeKey, vbBinaryCompare) = 0 Then
            list.ListIndex = index
            NxKeyboardBridgeSelectRoute = index
            Exit Function
        End If
    Next index
    Err.Raise vbObjectError + 482, "NxKeyboardBridge", "Requested route was not visible"
End Function

Public Function NxKeyboardBridgeSelectedRoute() As String
    Dim list As Object
    Set list = Manager.Controls.Item("lstCommands")
    If list.ListIndex < 0 Then Exit Function
    NxKeyboardBridgeSelectedRoute = CStr(list.List(list.ListIndex, 0))
End Function

Public Function NxKeyboardBridgeSelectedListIndex() As Long
    NxKeyboardBridgeSelectedListIndex = CLng(Manager.Controls.Item("lstCommands").ListIndex)
End Function

Public Function NxKeyboardBridgeRouteThirdColumn(ByVal routeKey As String) As String
    Dim list As Object
    Dim index As Long
    Set list = Manager.Controls.Item("lstCommands")
    For index = 0 To list.ListCount - 1
        If StrComp(CStr(list.List(index, 0)), routeKey, vbBinaryCompare) = 0 Then
            NxKeyboardBridgeRouteThirdColumn = CStr(list.List(index, 2))
            Exit Function
        End If
    Next index
    Err.Raise vbObjectError + 483, "NxKeyboardBridge", "Requested route was not visible"
End Function

Public Function NxKeyboardBridgeShortcutText() As String
    NxKeyboardBridgeShortcutText = CStr(Manager.Controls.Item("txtShortcut").Value)
End Function

Public Function NxKeyboardBridgeAssignEnabled() As Boolean
    NxKeyboardBridgeAssignEnabled = CBool(Manager.Controls.Item("cmdAssign").Enabled)
End Function

Public Function NxKeyboardBridgePreviewCaption() As String
    NxKeyboardBridgePreviewCaption = CStr(Manager.Controls.Item("lblPreview").Caption)
End Function

Public Function NxKeyboardBridgeStatusCaption() As String
    NxKeyboardBridgeStatusCaption = CStr(Manager.Controls.Item("lblStatus").Caption)
End Function

Public Function NxKeyboardBridgeActiveControlName() As String
    On Error Resume Next
    NxKeyboardBridgeActiveControlName = CStr(Manager.ActiveControl.Name)
End Function

Public Function NxKeyboardBridgeControlState(ByVal controlName As String) As String
    Dim control As Object
    Set control = Manager.Controls.Item(controlName)
    NxKeyboardBridgeControlState = CStr(control.Visible) & "|" & CStr(control.Enabled) & "|" & CStr(control.TabStop) & "|" & CStr(control.Locked)
End Function

Public Sub NxKeyboardBridgeCleanupUnload()
    On Error Resume Next
    If Not gManager Is Nothing Then Unload gManager
    Set gManager = Nothing
End Sub
'@

$excel = $null; $addin = $null; $testBook = $null; $testSheet = $null; $targetRange = $null; $binding = $null
$firstLifecycle = [ordered]@{label='Shortcut keyboard first Excel';exit_mode='NOT_STARTED';failure=$null;owned_pid=0}
$restartLifecycle = [ordered]@{label='Shortcut keyboard restarted Excel';exit_mode='NOT_STARTED';failure=$null;owned_pid=0}
$failure = $null; $failureStage = 'initialization'
$cases = New-Object 'Collections.Generic.List[object]'
$caseProgressPath = Join-Path $EvidenceRoot 'ShortcutKeyboard.case-progress.json'
$profileRoot = [IO.Path]::GetFullPath((Join-Path $EvidenceRoot 'Profile'))
$previousProfileRoot = [Environment]::GetEnvironmentVariable('LHEXCEL_PROFILE_ROOT', 'Process')
$copyPath = Join-Path $EvidenceRoot 'Product.ShortcutKeyboard.xlam'
$sourceHash = (Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash.ToLowerInvariant()
$compileEvidence = [ordered]@{status='NOT_RUN';control_id=578;active_project=$null;dialog=$null;failure=$null}
$ownedProcessReceipts = New-Object 'Collections.Generic.List[object]'
$bridgeMacro = $null
$bridgeCaption = $null
$bridgeCleanup = [ordered]@{status='NOT_RUN';failure=$null;handle=0}
$restartExit = [pscustomobject][ordered]@{status='NOT_RUN';elapsed_ms=0;remaining_pids=@()}
$finalExit = [pscustomobject][ordered]@{status='NOT_RUN';elapsed_ms=0;remaining_pids=@()}
$remainingImmediate = @()

try {
    [Environment]::SetEnvironmentVariable('LHEXCEL_PROFILE_ROOT', $profileRoot, 'Process')
    Copy-Item -LiteralPath $ProductXlam -Destination $copyPath -ErrorAction Stop
    if ((Get-FileHash -LiteralPath $copyPath -Algorithm SHA256).Hash.ToLowerInvariant() -cne $sourceHash) { throw 'Copy-only keyboard bridge artifact hash mismatch before modification' }

    $failureStage = 'first_excel_launch'
    $baseline = Get-ExcelProcessBaseline
    $excel = New-Object -ComObject Excel.Application
    $binding = Get-ExactExcelProcessOwnership $excel $baseline 'Shortcut keyboard first Excel'
    if (-not [bool]$binding.owned) { throw ('Excel ownership rejected: ' + [string]$binding.failure) }
    $firstOwnerReceipt = [ordered]@{schema='lhexcel-shortcut-keyboard-owned-process-v1';run_id=$RunId;phase='first';owned_pid=[int]$binding.pid;owned_hwnd=[Int64]$binding.hwnd;started_utc=[string]$binding.started_utc;executable=[string]$binding.executable;written_utc=[DateTime]::UtcNow.ToString('o')}
    $firstOwnerReceiptPath = Join-Path $EvidenceRoot 'ShortcutKeyboard.first-owned-process.json'
    Write-AtomicJson $firstOwnerReceiptPath $firstOwnerReceipt
    [void]$ownedProcessReceipts.Add([ordered]@{path=$firstOwnerReceiptPath;receipt=$firstOwnerReceipt})
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $addin = $excel.Workbooks.Open($copyPath, $false, $false)
    $macro = "'" + ([string]$addin.Name).Replace("'", "''") + "'!"
    $failureStage = 'first_profile_root_assertion'
    $actualProfileRoot = [string]$excel.Run($macro + 'NxLHexcelProfileRoot')
    if ($actualProfileRoot -cne $profileRoot) { throw ('Profile isolation rejected before keyboard assignment: expected=' + $profileRoot + '; actual=' + $actualProfileRoot) }
    Add-Case 'profile_root_first_open' $true ([ordered]@{expected=$profileRoot;actual=$actualProfileRoot})

    $failureStage = 'copy_only_bridge_injection'
    $project = $null; $components = $null; $component = $null; $codeModule = $null
    try {
        $project = $addin.VBProject
        $components = $project.VBComponents
        $component = $components.Add(1)
        $component.Name = 'NxShortcutKeyboardBridge'
        $codeModule = $component.CodeModule
        $normalizedBridgeCode = $bridgeCode -replace '\r\n|\r|\n', "`r`n"
        [void]$codeModule.AddFromString($normalizedBridgeCode)
        $addin.Save()
    } finally {
        Release-ComObject $codeModule; Release-ComObject $component; Release-ComObject $components; Release-ComObject $project
    }

    $failureStage = 'bridge_save_reopen_verification'
    $addin.Close($false)
    Release-ComObject $addin
    $addin = $excel.Workbooks.Open($copyPath, $false, $false)
    $macro = "'" + ([string]$addin.Name).Replace("'", "''") + "'!"
    $project = $null; $components = $null; $component = $null; $codeModule = $null
    try {
        $project = $addin.VBProject
        $components = $project.VBComponents
        $component = $components.Item('NxShortcutKeyboardBridge')
        $codeModule = $component.CodeModule
        $bridgeBody = [string]$codeModule.Lines(1, [int]$codeModule.CountOfLines)
        if ($bridgeBody -notmatch '(?m)^Public Sub NxKeyboardBridgeShow\(\)' -or $bridgeBody -notmatch '(?m)^Public Sub NxKeyboardBridgeCleanupUnload\(\)') { throw 'Copy-only keyboard bridge body did not survive save/reopen' }
    } finally {
        Release-ComObject $codeModule; Release-ComObject $component; Release-ComObject $components; Release-ComObject $project
    }
    $failureStage = 'copy_only_bridge_compile_578'
    Invoke-ShortcutKeyboardVbaCompile $excel $binding $addin $copyPath
    Add-Case 'copy_only_bridge_survives_reopen_and_compile' ($compileEvidence.status -ceq 'PASS') ([ordered]@{component='NxShortcutKeyboardBridge';compile=$compileEvidence})
    $bridgeMacro = $macro

    $failureStage = 'form_modeless_show'
    [void]$excel.Run($macro + 'NxKeyboardBridgeShow')
    $caption = [string]$excel.Run($macro + 'NxKeyboardBridgeCaption')
    $bridgeCaption = $caption
    $formHandle = Wait-OwnedFormWindow ([int]$binding.pid) $caption
    Assert-ActivateWindow $formHandle 'form_modeless_show'
    $controlNames = [string]$excel.Run($macro + 'NxKeyboardBridgeControlNames')
    $requiredControls = @('lstCommands','txtShortcut','cmdAssign','cmdClose')
    $controlsVisible = @($requiredControls | Where-Object { $controlNames -split '\|' -cnotcontains $_ }).Count -eq 0
    Add-Case 'actual_form_public_controls' $controlsVisible ([ordered]@{controls=$controlNames;required=$requiredControls})
    $shortcutControlState = [string]$excel.Run($macro + 'NxKeyboardBridgeControlState', 'txtShortcut')
    $shortcutControlFocusable = ($shortcutControlState -ceq 'True|True|True|True')
    Add-Case 'actual_form_shortcut_control_focusable' $shortcutControlFocusable ([ordered]@{control='txtShortcut';state=$shortcutControlState;expected='True|True|True|True'})

    $routeKey = 'command:NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-NUMBER'
    $failureStage = 'form_observed_route_selection'
    $selectedIndex = [int]$excel.Run($macro + 'NxKeyboardBridgeSelectRoute', $routeKey)
    Assert-ActivateWindow $formHandle 'form_observed_route_selection'
    [void](Focus-FormControlByTab $excel $macro $formHandle 'lstCommands' 'form_observed_route_selection')
    if ($selectedIndex -gt 0) { Send-PhysicalKeys '{UP}{DOWN}' $formHandle 'form_observed_route_selection' } else { Send-PhysicalKeys '{DOWN}{UP}' $formHandle 'form_observed_route_selection' }
    $selectedRoute = [string]$excel.Run($macro + 'NxKeyboardBridgeSelectedRoute')
    if ($selectedRoute -cne $routeKey) { throw ('Observed route selection changed: expected=' + $routeKey + '; actual=' + $selectedRoute) }
    Add-Case 'actual_form_route_selected' $true ([ordered]@{route=$selectedRoute;index=$selectedIndex})

    $failureStage = 'form_reserved_ctrl_c_input'
    Assert-ActivateWindow $formHandle 'form_reserved_ctrl_c_input'
    [void](Focus-FormControlByTab $excel $macro $formHandle 'txtShortcut' 'form_reserved_ctrl_c_input')
    Send-PhysicalKeys '^c' $formHandle 'form_reserved_ctrl_c_input'
    $reservedCapture = [string]$excel.Run($macro + 'NxKeyboardBridgeShortcutText')
    $reservedAssignEnabled = [bool]$excel.Run($macro + 'NxKeyboardBridgeAssignEnabled')
    Add-Case 'actual_form_reserved_ctrl_c_not_assignable' ($reservedCapture.Length -eq 0 -and -not $reservedAssignEnabled) ([ordered]@{capture=$reservedCapture;assign_enabled=$reservedAssignEnabled})

    $failureStage = 'form_ctrl_alt_k_input'
    Assert-ActivateWindow $formHandle 'form_ctrl_alt_k_input'
    [void](Focus-FormControlByTab $excel $macro $formHandle 'txtShortcut' 'form_ctrl_alt_k_input')
    $activeControlBeforeChord = [string]$excel.Run($macro + 'NxKeyboardBridgeActiveControlName')
    Send-PhysicalKeys '^%k' $formHandle 'form_ctrl_alt_k_input'
    $capture = [string]$excel.Run($macro + 'NxKeyboardBridgeShortcutText')
    $assignEnabled = [bool]$excel.Run($macro + 'NxKeyboardBridgeAssignEnabled')
    $previewCaption = [string]$excel.Run($macro + 'NxKeyboardBridgePreviewCaption')
    $statusCaption = [string]$excel.Run($macro + 'NxKeyboardBridgeStatusCaption')
    $selectedRouteAfterChord = [string]$excel.Run($macro + 'NxKeyboardBridgeSelectedRoute')
    $selectedListIndexAfterChord = [int]$excel.Run($macro + 'NxKeyboardBridgeSelectedListIndex')
    $activeControlAfterChord = [string]$excel.Run($macro + 'NxKeyboardBridgeActiveControlName')
    $conflictReason = [string]$excel.Run($macro + 'NxShortcutConflictReason', 'Ctrl+Alt', 'K', $routeKey)
    Add-Case 'actual_form_ctrl_alt_k_capture' ($capture -ceq 'Ctrl+Alt+K' -and $assignEnabled) ([ordered]@{
        capture=$capture;assign_enabled=$assignEnabled;preview_caption=$previewCaption;status_caption=$statusCaption
        selected_route=$selectedRouteAfterChord;selected_list_index=$selectedListIndexAfterChord
        active_control_before_chord=$activeControlBeforeChord;active_control_after_chord=$activeControlAfterChord
        diagnostic_conflict_reason=$conflictReason
    })

    $failureStage = 'form_assign_button_keyboard_activation'
    Assert-ActivateWindow $formHandle 'form_assign_button_keyboard_activation'
    [void](Focus-FormControlByTab $excel $macro $formHandle 'cmdAssign' 'form_assign_button_keyboard_activation')
    Send-PhysicalKeys '{ENTER}' $formHandle 'form_assign_button_keyboard_activation'
    $visibleBinding = [string]$excel.Run($macro + 'NxKeyboardBridgeRouteThirdColumn', $routeKey)
    Add-Case 'actual_form_assign_updates_third_column' ($visibleBinding -ceq 'Ctrl+Alt+K') ([ordered]@{route=$routeKey;third_column=$visibleBinding})

    $failureStage = 'form_close_button_keyboard_activation'
    Assert-ActivateWindow $formHandle 'form_close_button_keyboard_activation'
    [void](Focus-FormControlByTab $excel $macro $formHandle 'cmdClose' 'form_close_button_keyboard_activation')
    Send-PhysicalKeys '{ENTER}' $formHandle 'form_close_button_keyboard_activation'
    $formGoneWatch = [Diagnostics.Stopwatch]::StartNew()
    do { Start-Sleep -Milliseconds 100 } while ([NxShortcutKeyboardWin32]::FindVisibleTopLevelWindow([int]$binding.pid, $caption) -ne 0 -and $formGoneWatch.ElapsedMilliseconds -lt 3000)
    if ([NxShortcutKeyboardWin32]::FindVisibleTopLevelWindow([int]$binding.pid, $caption) -ne 0) { throw 'Shortcut manager did not close after actual close-button keyboard activation' }
    Add-Case 'actual_form_closed_before_onkey_dispatch' $true ([ordered]@{caption=$caption})

    $failureStage = 'first_onkey_dispatch_target_seed'
    $testBook = $excel.Workbooks.Add()
    $testSheet = $testBook.Worksheets.Item(1)
    [void]$testBook.Activate(); [void]$testSheet.Activate()
    $targetRange = $testSheet.Range('A1')
    $targetRange.NumberFormat = '0.00'
    [void]$targetRange.Select()
    Assert-ActivateWindow ([Int64]$binding.hwnd) 'first_onkey_dispatch_target_seed'
    $failureStage = 'first_actual_onkey_dispatch'
    Send-PhysicalKeys '^%k' ([Int64]$binding.hwnd) 'first_actual_onkey_dispatch'
    $firstFormat = Wait-NumberFormat $targetRange '#,##0'
    Add-Case 'actual_onkey_dispatch_executes_target' ($firstFormat -ceq '#,##0') ([ordered]@{expected='#,##0';actual=$firstFormat})

    $failureStage = 'first_excel_lifecycle'
    $firstLifecycle = Complete-ExcelPhase $excel $addin $testBook $testSheet $targetRange $binding 'Shortcut keyboard first Excel'
    $excel = $null; $addin = $null; $testBook = $null; $testSheet = $null; $targetRange = $null; $binding = $null
    if ($firstLifecycle.exit_mode -cne 'NATURAL' -or -not [string]::IsNullOrWhiteSpace([string]$firstLifecycle.failure)) { throw ('First Excel lifecycle rejected: ' + [string]$firstLifecycle.failure) }

    $failureStage = 'restart_excel_exit_convergence'
    $restartExit = Wait-FreshExcelProcessExit 2000
    if ($restartExit.status -cne 'CLEARED') { throw ('Restart blocked by remaining EXCEL PIDs: ' + ($restartExit.remaining_pids -join ',')) }
    $failureStage = 'restart_excel_launch'
    $restartBaseline = Get-ExcelProcessBaseline
    $excel = New-Object -ComObject Excel.Application
    $binding = Get-ExactExcelProcessOwnership $excel $restartBaseline 'Shortcut keyboard restarted Excel'
    if (-not [bool]$binding.owned) { throw ('Excel restart ownership rejected: ' + [string]$binding.failure) }
    $restartOwnerReceipt = [ordered]@{schema='lhexcel-shortcut-keyboard-owned-process-v1';run_id=$RunId;phase='restart';owned_pid=[int]$binding.pid;owned_hwnd=[Int64]$binding.hwnd;started_utc=[string]$binding.started_utc;executable=[string]$binding.executable;written_utc=[DateTime]::UtcNow.ToString('o')}
    $restartOwnerReceiptPath = Join-Path $EvidenceRoot 'ShortcutKeyboard.restart-owned-process.json'
    Write-AtomicJson $restartOwnerReceiptPath $restartOwnerReceipt
    [void]$ownedProcessReceipts.Add([ordered]@{path=$restartOwnerReceiptPath;receipt=$restartOwnerReceipt})
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $addin = $excel.Workbooks.Open($ProductXlam, $false, $true)
    $macro = "'" + ([string]$addin.Name).Replace("'", "''") + "'!"
    $failureStage = 'restart_profile_root_assertion'
    $actualProfileRoot = [string]$excel.Run($macro + 'NxLHexcelProfileRoot')
    if ($actualProfileRoot -cne $profileRoot) { throw ('Profile isolation rejected after restart: expected=' + $profileRoot + '; actual=' + $actualProfileRoot) }
    Add-Case 'profile_root_restart' $true ([ordered]@{expected=$profileRoot;actual=$actualProfileRoot})

    $failureStage = 'restart_onkey_dispatch_target_seed'
    $testBook = $excel.Workbooks.Add()
    $testSheet = $testBook.Worksheets.Item(1)
    [void]$testBook.Activate(); [void]$testSheet.Activate()
    $targetRange = $testSheet.Range('A1')
    $targetRange.NumberFormat = '0.00'
    [void]$targetRange.Select()
    Assert-ActivateWindow ([Int64]$binding.hwnd) 'restart_onkey_dispatch_target_seed'
    $failureStage = 'restart_actual_onkey_dispatch'
    Send-PhysicalKeys '^%k' ([Int64]$binding.hwnd) 'restart_actual_onkey_dispatch'
    $restartFormat = Wait-NumberFormat $targetRange '#,##0'
    Add-Case 'reloaded_actual_onkey_dispatch_executes_target' ($restartFormat -ceq '#,##0') ([ordered]@{expected='#,##0';actual=$restartFormat})
} catch {
    $failure = 'stage=' + $failureStage + '; ' + $_.Exception.Message
} finally {
    if ($null -ne $binding) {
        $addinIsBridgeCopy = $false
        if ($null -ne $addin) {
            try { $addinIsBridgeCopy = [string]::Equals([string]$addin.FullName, $copyPath, [StringComparison]::OrdinalIgnoreCase) } catch {}
        }
        if ($addinIsBridgeCopy) {
            $bridgeCleanup = Cleanup-OwnedBridgeForm $excel $bridgeMacro $binding $bridgeCaption
        }
        if ($firstLifecycle.exit_mode -ceq 'NOT_STARTED') { $firstLifecycle = Complete-ExcelPhase $excel $addin $testBook $testSheet $targetRange $binding 'Shortcut keyboard first Excel' }
        else { $restartLifecycle = Complete-ExcelPhase $excel $addin $testBook $testSheet $targetRange $binding 'Shortcut keyboard restarted Excel' }
        $excel = $null; $addin = $null; $testBook = $null; $testSheet = $null; $targetRange = $null; $binding = $null
    } else {
        Release-ComObject $targetRange; Release-ComObject $testSheet; Release-ComObject $testBook; Release-ComObject $addin; Release-ComObject $excel
    }
    [Environment]::SetEnvironmentVariable('LHEXCEL_PROFILE_ROOT', $previousProfileRoot, 'Process')
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

$remainingImmediate = @(Get-FreshExcelProcessIds)
$finalExit = Wait-FreshExcelProcessExit 2000
$remaining = @($finalExit.remaining_pids)
$sourceHashAfter = (Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash.ToLowerInvariant()
$lifecyclePassed = ($firstLifecycle.exit_mode -ceq 'NATURAL' -and [string]::IsNullOrWhiteSpace([string]$firstLifecycle.failure) -and $restartLifecycle.exit_mode -ceq 'NATURAL' -and [string]::IsNullOrWhiteSpace([string]$restartLifecycle.failure) -and $remaining.Count -eq 0)
Add-Case 'cleanup_exact_owned_excel_natural_exit' $lifecyclePassed ([ordered]@{first=$firstLifecycle;restart=$restartLifecycle;immediate_pids=$remainingImmediate;recheck=$finalExit;remaining_pids=$remaining}) $false
$failed = @($cases | Where-Object status -ne 'PASS').Count
$status = if ($null -eq $failure -and $sourceHashAfter -ceq $sourceHash -and $failed -eq 0 -and $cases.Count -eq 13) { 'PASS' } else { 'FAIL' }
$receipt = [ordered]@{
    schema='lhexcel-shortcut-keyboard-v1';run_id=$RunId;generated_utc=[DateTime]::UtcNow.ToString('o');status=$status
    product=$ProductXlam;product_sha256=$sourceHash;product_sha256_after=$sourceHashAfter;bridge_copy=$copyPath
    profile_root=$profileRoot;failure=$failure;failure_stage=$failureStage;cases=$cases.ToArray()
    owned_process_receipts=$ownedProcessReceipts.ToArray();compile=$compileEvidence;bridge_cleanup=$bridgeCleanup
    lifecycle=[ordered]@{first=$firstLifecycle;restart=$restartLifecycle;restart_exit_recheck=$restartExit;remaining_immediate=$remainingImmediate;remaining_recheck=$finalExit;remaining_process_ids=$remaining}
}
$receiptPath = Join-Path $EvidenceRoot 'ShortcutKeyboard.json'
Write-AtomicJson $receiptPath $receipt
if ($status -eq 'PASS') { Write-Output ('PASS|ShortcutKeyboard|' + ($cases.Count - $failed) + '/' + $cases.Count); Write-Output $receiptPath; exit 0 }
Write-Output ('FAIL|ShortcutKeyboard|' + ($cases.Count - $failed) + '/' + $cases.Count); Write-Output $receiptPath; exit 24
