param(
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RunId = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }

$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$DataRoot = [IO.Path]::GetFullPath($DataRoot)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
if ($RunId -notmatch '^[0-9a-f-]{36}$') { throw 'RunId must be a lowercase GUID' }
[void](New-Item -ItemType Directory -Path $EvidenceRoot -Force)
$originalProfileOverride = [string]$env:LHEXCEL_PROFILE_ROOT
$isolatedProfileRoot = Join-Path $EvidenceRoot 'profile/LHexcel'
[void](New-Item -ItemType Directory -Path $isolatedProfileRoot -Force)
$env:LHEXCEL_PROFILE_ROOT = $isolatedProfileRoot

. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
Add-Type -AssemblyName System.Windows.Forms

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Get-HexDigest([string]$Path) {
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
}

function Write-Json([string]$Path, [object]$Value) {
    $utf8 = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 30), $utf8)
}

function Get-Percentile95([double[]]$Values) {
    if ($null -eq $Values -or $Values.Count -eq 0) { return 0.0 }
    $ordered = @($Values | Sort-Object)
    $index = [Math]::Ceiling(0.95 * $ordered.Count) - 1
    if ($index -lt 0) { $index = 0 }
    return [Math]::Round([double]$ordered[$index], 3)
}

$script:Cases = New-Object 'Collections.Generic.List[object]'
function Add-FocusCase([string]$Name, [bool]$Passed, [object]$Details) {
    $script:Cases.Add([pscustomobject][ordered]@{
        ordinal = $script:Cases.Count + 1
        name = $Name
        status = if ($Passed) { 'PASS' } else { 'FAIL' }
        details = $Details
    })
}

function Invoke-ExcelCleanup([object]$Excel, [object]$Binding, [object[]]$Objects, [string]$Label) {
    $failures = New-Object 'Collections.Generic.List[string]'
    if ($null -ne $Excel) {
        try { $Excel.DisplayAlerts = $false; $Excel.Quit() } catch { $failures.Add($_.Exception.Message) }
    }
    foreach ($item in $Objects) {
        try { Release-ComObject $item } catch { $failures.Add($_.Exception.Message) }
    }
    try { Release-ComObject $Excel } catch { $failures.Add($_.Exception.Message) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $exit = Stop-ExactProcessAfterGrace $Binding.process $Label 15000 -Detailed
    return [pscustomobject][ordered]@{
        exit_mode = [string]$exit.exit_mode
        failure = if ($failures.Count -gt 0) { $failures -join ' | ' } else { [string]$exit.failure }
        owned_pid = if ($null -eq $Binding) { 0 } else { [int]$Binding.pid }
    }
}

function Wait-FocusRefresh([object]$Excel, [string]$MacroPrefix) {
    Start-Sleep -Milliseconds 120
    [void]$Excel.Run($MacroPrefix + 'NxFocusOverlayRefresh')
    Start-Sleep -Milliseconds 60
}

function Initialize-FocusWindowApi {
    if ('NxFocusProductWindowApi' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Threading;
public static class NxFocusProductWindowApi {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int command);
  [DllImport("user32.dll")] public static extern void SwitchToThisWindow(IntPtr hWnd, bool altTab);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
  [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint fromThread, uint toThread, bool attach);
  [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
  public static bool Activate(long hWnd) {
    var target = new IntPtr(hWnd);
    ShowWindow(target, 9);
    SwitchToThisWindow(target, true);
    for (var attempt = 0; attempt < 40; attempt++) {
      if (GetForegroundWindow() == target) return true;
      var foreground = GetForegroundWindow();
      uint ignored;
      uint targetThread = GetWindowThreadProcessId(target, out ignored);
      uint foregroundThread = foreground == IntPtr.Zero ? 0 : GetWindowThreadProcessId(foreground, out ignored);
      uint currentThread = GetCurrentThreadId();
      bool attachedTarget = targetThread != 0 && targetThread != currentThread && AttachThreadInput(currentThread, targetThread, true);
      bool attachedForeground = foregroundThread != 0 && foregroundThread != currentThread && foregroundThread != targetThread && AttachThreadInput(currentThread, foregroundThread, true);
      try {
        ShowWindow(target, 9);
        BringWindowToTop(target);
        SetForegroundWindow(target);
        SetFocus(target);
      } finally {
        if (attachedForeground) AttachThreadInput(currentThread, foregroundThread, false);
        if (attachedTarget) AttachThreadInput(currentThread, targetThread, false);
      }
      Thread.Sleep(50);
    }
    return GetForegroundWindow() == target;
  }
  public static long Foreground() { return GetForegroundWindow().ToInt64(); }
}
'@
}

function Invoke-FocusWindowActivation([object]$Workbook, [object]$ExcelWindow) {
    [void](Initialize-FocusWindowApi)
    [void]$ExcelWindow.Activate()
    [void]$Workbook.Activate()
    $windowHandle = [Int64]$ExcelWindow.Hwnd
    if ($windowHandle -le 0 -or -not [NxFocusProductWindowApi]::Activate($windowHandle)) {
        throw ('Exact Excel window activation failed: hwnd=' + $windowHandle)
    }
}

function Invoke-NativeUndoProbe([object]$Excel, [object]$Workbook, [object]$Sheet, [object]$ExcelWindow, [string]$MacroPrefix) {
    $target = $null
    [void](Initialize-FocusWindowApi)
    $result = [ordered]@{stage='initialize';excel_hwnd=[Int64]$ExcelWindow.Hwnd;activated=$false;foreground_before=0;foreground_after=0;typed=$null;undo_enabled=$false;undo_after_refresh=$false;after_undo=$null;failure=$null;passed=$false}
    try {
        $result.stage = 'select-target'
        Invoke-FocusWindowActivation $Workbook $ExcelWindow
        [void]$Sheet.Activate()
        $target = $Sheet.Range('Z205')
        [void]$target.ClearContents()
        [void]$target.Select()
        $result.foreground_before = [Int64][NxFocusProductWindowApi]::Foreground()
        $result.activated = ([Int64]$Excel.ActiveWindow.Hwnd -eq [Int64]$ExcelWindow.Hwnd)
        $result.foreground_after = [Int64][NxFocusProductWindowApi]::Foreground()
        if (-not $result.activated) { return [pscustomobject]$result }
        $result.stage = 'execute-paste'
        [Windows.Forms.Clipboard]::SetText('FOCUS_OVERLAY_UNDO')
        [void]$Excel.CommandBars.ExecuteMso('Paste')
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            Start-Sleep -Milliseconds 200
            $result.typed = [string]$target.Value2
            $result.undo_enabled = [bool]$Excel.CommandBars.GetEnabledMso('Undo')
            if (-not [string]::IsNullOrEmpty($result.typed) -and $result.undo_enabled) { break }
        }
        if ([string]::IsNullOrEmpty($result.typed) -or -not $result.undo_enabled) { return [pscustomobject]$result }
        [void]$Excel.Run($MacroPrefix + 'NxFocusOverlayRefresh')
        $result.undo_after_refresh = [bool]$Excel.CommandBars.GetEnabledMso('Undo')
        if (-not $result.undo_after_refresh) { return [pscustomobject]$result }
        $result.stage = 'undo'
        [void]$Excel.Undo()
        Start-Sleep -Milliseconds 500
        $result.after_undo = [string]$target.Value2
        $result.passed = [string]::IsNullOrEmpty($result.after_undo)
        $result.stage = 'complete'
        return [pscustomobject]$result
    } catch {
        $result.failure = $_.Exception.Message
        return [pscustomobject]$result
    } finally {
        Release-ComObject $target
    }
}

function Get-OverlaySnapshot([object]$Excel, [string]$MacroPrefix) {
    [ordered]@{
        enabled = [bool]$Excel.Run($MacroPrefix + 'NxFocusIsEnabled')
        windows = [int]$Excel.Run($MacroPrefix + 'NxFocusOverlayWindowCount')
        visible = [int]$Excel.Run($MacroPrefix + 'NxFocusOverlayVisibleWindowCount')
        style = [bool]$Excel.Run($MacroPrefix + 'NxFocusOverlayStyleContractSatisfied')
        bounds_match = [bool]$Excel.Run($MacroPrefix + 'NxFocusOverlayExpectedBoundsMatch')
        selection = [string]$Excel.Run($MacroPrefix + 'NxFocusOverlaySelectionSignature')
        bounds = [string]$Excel.Run($MacroPrefix + 'NxFocusOverlayBoundsSignature')
        last_error = [string]$Excel.Run($MacroPrefix + 'NxFocusOverlayLastError')
    }
}

function Get-UserFormatFingerprint([object]$Range) {
    $conditions = $null; $rule = $null; $applies = $null
    try {
        $conditions = $Range.FormatConditions
        if ([int]$conditions.Count -ne 1) { return 'count=' + [string]$conditions.Count }
        $rule = $conditions.Item(1)
        $applies = $rule.AppliesTo
        return ('count=1|type={0}|formula={1}|priority={2}|stop={3}|applies={4}|color={5}' -f `
            [int]$rule.Type,[string]$rule.Formula1,[int]$rule.Priority,[bool]$rule.StopIfTrue,[string]$applies.Address(),[int]$rule.Interior.Color)
    } finally {
        Release-ComObject $applies; Release-ComObject $rule; Release-ComObject $conditions
    }
}

$artifact = Join-Path $EvidenceRoot 'Product.xlam'
$manifestPath = Join-Path $SourceRoot 'build/manifests/Product.json'
$receiptPath = Join-Path $EvidenceRoot 'FocusOverlayProduct.json'
if (Test-Path -LiteralPath $artifact) { throw 'Existing Product.xlam evidence artifact is rejected' }

# Current-tree Product build: build/Build-Xlam.ps1
& (Join-Path $SourceRoot 'build/Build-Xlam.ps1') -DataRoot $DataRoot -OutputPath $artifact `
    -ManifestPath $manifestPath -RibbonPath (Join-Path $SourceRoot 'src/ribbon/customUI14.xml')
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $artifact -PathType Leaf)) { throw 'Current Product.xlam build failed' }

$artifactSha = Get-HexDigest $artifact
$manifestSha = Get-HexDigest $manifestPath
$baseline = Get-ExcelProcessBaseline
if (@($baseline.process_ids).Count -ne 0) { throw 'Focus native suite requires zero pre-existing Excel processes' }

$excel = $null; $binding = $null; $books = $null; $addin = $null; $workbook = $null; $secondWorkbook = $null
$firstBookWindow = $null; $secondBookWindow = $null; $undoWindow = $null
$sheet = $null; $secondSheet = $null; $secondBookSheet = $null; $userRange = $null; $userRule = $null; $userShape = $null; $userName = $null
$cleanup = [pscustomobject][ordered]@{exit_mode='NOT_STARTED';failure=$null;owned_pid=0}
$diagnostic = $null
$refreshP95 = 0.0
$excelVersion = $null
$excelBuild = $null
$MoveCount = 500
$WarmupCount = 20

try {
    $excel = New-Object -ComObject Excel.Application
    $binding = Get-ExactExcelProcessOwnership $excel $baseline 'FocusOverlayProduct Excel'
    if (-not [bool]$binding.owned) { throw ('Excel ownership rejected: ' + [string]$binding.failure) }
    $excelVersion = [string]$excel.Version
    $excelBuild = [string]$excel.Build
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $books = $excel.Workbooks
    $addin = $books.Open($artifact, $false, $true)
    $workbook = $books.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'FocusPrimary'
    $secondSheet = $workbook.Worksheets.Add()
    $secondSheet.Name = 'FocusSecondary'
    $sheet.Activate()
    $excel.WindowState = -4137
    $initialWindow = $excel.ActiveWindow
    try { Invoke-FocusWindowActivation $workbook $initialWindow }
    finally { Release-ComObject $initialWindow }
    Start-Sleep -Milliseconds 500
    $excel.ActiveWindow.ScrollRow = 1
    $excel.ActiveWindow.ScrollColumn = 1

    $userRange = $sheet.Range('A1:Z100')
    $userRule = $userRange.FormatConditions.Add(2, $null, '=MOD(ROW()+COLUMN(),3)=0')
    $userRule.StopIfTrue = $false
    $userRule.Interior.Color = 65535
    $userShape = $sheet.Shapes.AddShape(1, 10, 10, 30, 20)
    $userShape.Name = 'UserKeepShape'
    $userName = $workbook.Names.Add('UserKeepName', '=FocusPrimary!$A$1')
    $fixturePath = Join-Path $EvidenceRoot 'FocusFixture.xlsx'
    $workbook.SaveAs($fixturePath, 51)
    $userFormatBefore = Get-UserFormatFingerprint $userRange
    $userShapeCountBefore = [int]$sheet.Shapes.Count
    $userNameCountBefore = [int]$workbook.Names.Count
    $macro = "'Product.xlam'!"

    Add-FocusCase 'session_starts_off' (-not [bool]$excel.Run($macro + 'NxFocusIsEnabled')) @{enabled=$false}

    $sheet.Range('A1:C3').Select() | Out-Null
    [void]$excel.Run($macro + 'NxFocusEnable', 12632256)
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    $backend = [string]$excel.Run($macro + 'NxFocusControllerBackend')
    Add-FocusCase 'vba_fallback_backend_selected' ($backend -eq 'vba') @{backend=$backend}
    Add-FocusCase 'enable_creates_four_owned_windows' ($state.enabled -and $state.windows -eq 4 -and $state.style -and -not $state.last_error) $state
    $settingsPath = [string]$excel.Run($macro + 'NxFocusSettingsPath')
    $settingsPayload = if (Test-Path -LiteralPath $settingsPath -PathType Leaf) { Get-Content -LiteralPath $settingsPath -Raw -Encoding UTF8 } else { '' }
    $settingsReloaded = [string]$excel.Run($macro + 'NxFocusLoadColors')
    Add-FocusCase 'settings_saved_and_reloaded' (
        $settingsPath.StartsWith($isolatedProfileRoot, [StringComparison]::OrdinalIgnoreCase) -and
        $settingsPayload -match 'schema_version=2' -and
        $settingsPayload -match 'color=#C0C0C0' -and
        $settingsReloaded.StartsWith('#C0C0C0', [StringComparison]::OrdinalIgnoreCase)
    ) @{path=$settingsPath;payload_sha256=if(Test-Path -LiteralPath $settingsPath -PathType Leaf){Get-HexDigest $settingsPath}else{''};reloaded=$settingsReloaded}

    $sheet.Range('B2').Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'single_cell_two_axes' ($state.visible -eq 2 -and $state.selection -eq '2:2|2:2' -and $state.bounds_match) $state

    $sheet.Range('A1:C3').Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'a1_c3_four_endpoint_axes' ($state.visible -eq 4 -and $state.selection -eq '1:3|1:3' -and $state.bounds_match) $state

    $sheet.Range('C3:A1').Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'reverse_selection_normalized' ($state.selection -eq '1:3|1:3' -and $state.visible -eq 4 -and $state.bounds_match) $state

    $union = $null
    try { $union = $sheet.Range('A1,C3'); $union.Select() | Out-Null; Wait-FocusRefresh $excel $macro }
    finally { Release-ComObject $union }
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'noncontiguous_hidden' ($state.visible -eq 0) $state

    $sheet.Range('D4:F6').Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'supported_selection_resumes' ($state.visible -eq 4 -and $state.selection -eq '4:6|4:6' -and $state.bounds_match) $state

    $sheet.Rows.Item(1).Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'full_row_hidden' ($state.visible -eq 0) $state

    $sheet.Columns.Item(1).Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'full_column_hidden' ($state.visible -eq 0) $state

    $sheet.Range('A1:A100001').Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'oversized_selection_hidden' ($state.visible -eq 0) $state

    $sheet.Protect()
    $sheet.Range('E5:G7').Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'protected_sheet_supported' ($state.visible -eq 4 -and $state.selection -eq '5:7|5:7' -and $state.bounds_match) $state
    $sheet.Unprotect()

    $workbook.Saved = $true
    $sheet.Range('H8:J10').Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    Add-FocusCase 'saved_state_preserved' ([bool]$workbook.Saved) @{saved=[bool]$workbook.Saved}

    $userFormatAfter = Get-UserFormatFingerprint $userRange
    Add-FocusCase 'user_conditional_format_preserved' ($userFormatAfter -ceq $userFormatBefore) @{before=$userFormatBefore;after=$userFormatAfter}
    Add-FocusCase 'user_name_preserved' ([int]$workbook.Names.Count -eq $userNameCountBefore) @{before=$userNameCountBefore;after=[int]$workbook.Names.Count}
    Add-FocusCase 'user_shape_preserved' ([int]$sheet.Shapes.Count -eq $userShapeCountBefore) @{before=$userShapeCountBefore;after=[int]$sheet.Shapes.Count}

    $excel.ActiveWindow.Zoom = 75
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'zoom_75_tracking' ($state.visible -eq 4 -and $state.bounds_match) $state

    $excel.ActiveWindow.Zoom = 125
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'zoom_125_tracking' ($state.visible -eq 4 -and $state.bounds_match) $state

    $excel.ActiveWindow.ScrollRow = 20
    $excel.ActiveWindow.ScrollColumn = 5
    $sheet.Range('H25:J27').Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'scroll_tracking' ($state.visible -eq 4 -and $state.selection -eq '25:27|8:10' -and $state.bounds_match) $state

    $excel.WindowState = -4143
    $excel.Top = 40; $excel.Left = 40; $excel.Width = 1000; $excel.Height = 700
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'resize_tracking' ($state.visible -eq 4 -and $state.bounds_match) $state
    $excel.WindowState = -4137

    [void](Initialize-FocusWindowApi)
    $secondWorkbook = $books.Add()
    $secondBookSheet = $secondWorkbook.Worksheets.Item(1)
    $secondBookWindow = $secondWorkbook.Windows.Item(1)
    $secondTargetHwnd = [Int64]$secondBookWindow.Hwnd
    $secondWorkbookName = [string]$secondWorkbook.Name
    Invoke-FocusWindowActivation $secondWorkbook $secondBookWindow
    [void]$secondBookSheet.Activate()
    $secondActiveWorkbookName = [string]$excel.ActiveWorkbook.Name
    $secondActiveSheetName = [string]$excel.ActiveSheet.Name
    $secondBookTarget = $secondBookSheet.Range('B2:D4')
    [void]$secondBookTarget.Select()
    Wait-FocusRefresh $excel $macro
    $secondState = Get-OverlaySnapshot $excel $macro
    $secondActiveHwnd = [Int64]$excel.ActiveWindow.Hwnd
    $secondForegroundHwnd = [Int64][NxFocusProductWindowApi]::Foreground()
    $firstBookWindow = $workbook.Windows.Item(1)
    $firstTargetHwnd = [Int64]$firstBookWindow.Hwnd
    $firstWorkbookName = [string]$workbook.Name
    Invoke-FocusWindowActivation $workbook $firstBookWindow
    [void]$sheet.Activate()
    $firstActiveWorkbookName = [string]$excel.ActiveWorkbook.Name
    $firstActiveSheetName = [string]$excel.ActiveSheet.Name
    $firstBookTarget = $sheet.Range('A1:C3')
    [void]$firstBookTarget.Select()
    Wait-FocusRefresh $excel $macro
    $firstState = Get-OverlaySnapshot $excel $macro
    $firstActiveHwnd = [Int64]$excel.ActiveWindow.Hwnd
    $firstForegroundHwnd = [Int64][NxFocusProductWindowApi]::Foreground()
    Add-FocusCase 'second_workbook_isolation' ($secondActiveHwnd -eq $secondTargetHwnd -and $firstActiveHwnd -eq $firstTargetHwnd -and $secondState.selection -eq '2:4|2:4' -and $firstState.selection -eq '1:3|1:3' -and $firstState.bounds_match) @{second=$secondState;first=$firstState;second_target_hwnd=$secondTargetHwnd;second_active_hwnd=$secondActiveHwnd;second_foreground_hwnd=$secondForegroundHwnd;second_workbook=$secondWorkbookName;second_active_workbook=$secondActiveWorkbookName;second_active_sheet=$secondActiveSheetName;first_target_hwnd=$firstTargetHwnd;first_active_hwnd=$firstActiveHwnd;first_foreground_hwnd=$firstForegroundHwnd;first_workbook=$firstWorkbookName;first_active_workbook=$firstActiveWorkbookName;first_active_sheet=$firstActiveSheetName}
    Release-ComObject $firstBookTarget; $firstBookTarget = $null
    Release-ComObject $secondBookTarget; $secondBookTarget = $null

    $secondSheet.Activate(); $secondSheet.Range('D4:F6').Select() | Out-Null
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'second_sheet_tracking' ($state.selection -eq '4:6|4:6' -and $state.visible -eq 4 -and $state.bounds_match) $state

    [void]$excel.Run($macro + 'NxFocusSuspendForOperation')
    Start-Sleep -Milliseconds 150
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'suspend_hides' ($state.visible -eq 0 -and $state.windows -eq 4) $state

    [void]$excel.Run($macro + 'NxFocusResumeAfterOperation')
    Wait-FocusRefresh $excel $macro
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'resume_restores' ($state.visible -eq 4 -and $state.bounds_match) $state

    [void]$firstBookWindow.Activate(); [void]$sheet.Activate(); $excel.ActiveWindow.ScrollRow = 1; $excel.ActiveWindow.ScrollColumn = 1
    $undoWindow = $excel.ActiveWindow
    $undoProbe = @(Invoke-NativeUndoProbe $excel $workbook $sheet $undoWindow $macro)[-1]
    Add-FocusCase 'native_edit_undo_preserved' ([bool]$undoProbe.passed) $undoProbe

    $timings = New-Object 'Collections.Generic.List[double]'
    for ($index = 0; $index -lt ($MoveCount + $WarmupCount); $index++) {
        $row = 2 + ($index % 37); $column = 2 + (($index * 7) % 19)
        $sheet.Cells.Item($row, $column).Select() | Out-Null
        $timer = [Diagnostics.Stopwatch]::StartNew()
        [void]$excel.Run($macro + 'NxFocusOverlayRefresh')
        $timer.Stop()
        if ($index -ge $WarmupCount) { $timings.Add($timer.Elapsed.TotalMilliseconds) }
    }
    $refreshP95 = Get-Percentile95 $timings.ToArray()
    $state = Get-OverlaySnapshot $excel $macro
    Add-FocusCase 'five_hundred_moves_p95' ($timings.Count -eq $MoveCount -and $refreshP95 -le 50.0 -and $state.bounds_match) @{move_count=$MoveCount;warmup_count=$WarmupCount;refresh_p95_ms=$refreshP95;state=$state}

    [void]$excel.Run($macro + 'NxFocusDisable')
    Start-Sleep -Milliseconds 200
    $remainingWindows = [int]$excel.Run($macro + 'NxFocusOverlayWindowCount')
    Add-FocusCase 'disable_zero_windows' ($remainingWindows -eq 0 -and -not [bool]$excel.Run($macro + 'NxFocusIsEnabled')) @{windows=$remainingWindows}
} catch {
    $diagnostic = $_.Exception.Message
} finally {
    if ($null -ne $secondWorkbook) { try { $secondWorkbook.Close($false) } catch {} }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $addin) { try { $addin.Close($false) } catch {} }
    if ($null -ne $binding) {
        $cleanup = Invoke-ExcelCleanup $excel $binding @($undoWindow,$secondBookWindow,$firstBookWindow,$userName,$userShape,$userRule,$userRange,$secondBookSheet,$secondSheet,$sheet,$secondWorkbook,$workbook,$addin,$books) 'FocusOverlayProduct Excel'
    }
}

$env:LHEXCEL_PROFILE_ROOT = $originalProfileOverride
Start-Sleep -Milliseconds 750
$excelResidue = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
Add-FocusCase 'excel_natural_exit_zero_residue' ($cleanup.exit_mode -eq 'NATURAL' -and [string]::IsNullOrEmpty($cleanup.failure) -and $excelResidue -eq 0) @{exit_mode=$cleanup.exit_mode;failure=$cleanup.failure;owned_pid=$cleanup.owned_pid;excel_residue=$excelResidue}

$expectedNames = @(
    'session_starts_off','vba_fallback_backend_selected','enable_creates_four_owned_windows','settings_saved_and_reloaded','single_cell_two_axes','a1_c3_four_endpoint_axes',
    'reverse_selection_normalized','noncontiguous_hidden','supported_selection_resumes','full_row_hidden',
    'full_column_hidden','oversized_selection_hidden','protected_sheet_supported','saved_state_preserved',
    'user_conditional_format_preserved','user_name_preserved','user_shape_preserved','zoom_75_tracking',
    'zoom_125_tracking','scroll_tracking','resize_tracking','second_workbook_isolation','second_sheet_tracking',
    'suspend_hides','resume_restores','native_edit_undo_preserved','five_hundred_moves_p95',
    'disable_zero_windows','excel_natural_exit_zero_residue'
)
foreach ($expected in $expectedNames) {
    if (@($Cases | Where-Object { $_.name -eq $expected }).Count -eq 0) {
        $Cases.Add([pscustomobject][ordered]@{ordinal=$Cases.Count+1;name=$expected;status='FAIL';details=@{not_reached=$true;diagnostic=$diagnostic}})
    }
}
$Cases = @($Cases | Sort-Object { [Array]::IndexOf($expectedNames, [string]$_.name) })
for ($index = 0; $index -lt $Cases.Count; $index++) { $Cases[$index].ordinal = $index + 1 }
$passed = @($Cases | Where-Object { $_.status -eq 'PASS' }).Count
$allPass = ($Cases.Count -eq 29 -and $passed -eq 29 -and [string]::IsNullOrEmpty($diagnostic))
$receipt = [ordered]@{
    schema_version = 1
    suite = 'FocusOverlayProduct'
    status = if ($allPass) { 'PASS' } else { 'DIAGNOSTIC' }
    run_id = $RunId
    product_manifest_sha256 = $manifestSha
    artifact_sha256 = $artifactSha
    excel_version = $excelVersion
    excel_build = $excelBuild
    office_bitness = 'x64'
    cases = $Cases
    measurements = @{move_count=$MoveCount;warmup_count=$WarmupCount;refresh_p95_ms=$refreshP95}
    cleanup = $cleanup
    diagnostic = $diagnostic
}
Write-Json $receiptPath $receipt

if ($allPass) {
    Write-Output 'PASS|FocusOverlayProduct|29/29'
    Write-Output $receiptPath
    exit 0
}
Write-Output ('DIAGNOSTIC|FocusOverlayProduct|' + $passed + '/29')
Write-Output $receiptPath
exit 24
