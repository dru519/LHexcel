param(
    [string]$EvidenceRoot = '',
    [string]$RunId = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$DataRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) {
    $EvidenceRoot = Join-Path $DataRoot 'out/evidence/focus-overlay-feasibility'
}
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString() }
if ($RunId -notmatch '^[0-9a-fA-F-]{36}$') { throw 'RunId rejected' }

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

. (Join-Path $DataRoot 'build/Excel-ProcessLifecycle.ps1')

$RefreshCount = 200
$WarmupCount = 20
$FocusColor = 13434879
$ProbeSources = @(
    (Join-Path $DataRoot 'tests/vba/focus/FxFocusOverlayProbe.bas'),
    (Join-Path $DataRoot 'tests/vba/focus/CFxFocusOverlayEvents.cls'),
    (Join-Path $DataRoot 'tests/vba/focus/FNxFocusOverlayProbe.vba')
)
$ProbePath = Join-Path $EvidenceRoot 'FocusOverlayFeasibilityProbe.xlam'
$NormalPath = Join-Path $EvidenceRoot 'FocusOverlayNormal.xlsx'
$EvidencePath = Join-Path $EvidenceRoot 'FocusOverlayFeasibility.json'

function Get-Sha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-TextSha256([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').ToLowerInvariant()
    } finally { $sha.Dispose() }
}

function Write-AtomicJson([string]$Path, [object]$Value) {
    $directory = Split-Path -Parent $Path
    [void](New-Item -ItemType Directory -Path $directory -Force)
    $temporary = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    $writer = New-Object IO.StreamWriter($temporary, $false, (New-Object Text.UTF8Encoding($false)))
    try {
        $writer.Write(($Value | ConvertTo-Json -Depth 12) + "`n")
        $writer.Flush()
        $writer.BaseStream.Flush($true)
    } finally { $writer.Dispose() }
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Get-Percentile95([double[]]$Samples) {
    if ($null -eq $Samples -or $Samples.Count -eq 0) { return [double]::PositiveInfinity }
    $ordered = @($Samples | Sort-Object)
    $index = [Math]::Ceiling($ordered.Count * 0.95) - 1
    if ($index -lt 0) { $index = 0 }
    return [Math]::Round([double]$ordered[$index], 3)
}

function Set-VbComponentProperty([object]$Component, [string]$Name, [object]$Value) {
    $properties = $null; $property = $null
    try {
        $properties = $Component.Properties
        $property = $properties.Item($Name)
        $property.Value = [Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture)
    } finally {
        Release-ComObject $property
        Release-ComObject $properties
    }
}

function Invoke-ExcelCleanup([object]$Excel, [object]$Binding, [object[]]$ComObjects, [string]$Label) {
    $failures = New-Object Collections.Generic.List[string]
    $owned = $false
    try { $owned = Assert-ExactExcelProcessOwnership $Binding } catch { [void]$failures.Add($_.Exception.Message) }
    if ($owned -and $null -ne $Excel) {
        try { $Excel.DisplayAlerts = $false; $Excel.Quit() } catch { [void]$failures.Add($_.Exception.Message) }
    }
    foreach ($item in @($ComObjects)) {
        try { Release-ComObject $item } catch { [void]$failures.Add($_.Exception.Message) }
    }
    try { Release-ComObject $Excel } catch { [void]$failures.Add($_.Exception.Message) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($owned -and $null -ne $Binding.process) {
        $cleanup = Stop-ExactProcessAfterGrace $Binding.process $Label 30000
        if (-not [string]::IsNullOrWhiteSpace($cleanup)) { [void]$failures.Add($cleanup) }
    }
    if ($null -ne $Binding -and $null -ne $Binding.process) {
        try { $Binding.process.Dispose() } catch { [void]$failures.Add($_.Exception.Message) }
    }
    return @($failures)
}

function Initialize-OverlayWindowApi {
    if ('NxFocusOverlayWindowApi' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class NxFocusOverlayWindowApi {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr state);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int command);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int max);
  public static bool Activate(long hWnd) {
    var target = new IntPtr(hWnd); ShowWindow(target,5); return SetForegroundWindow(target);
  }
  public static long Foreground() { return GetForegroundWindow().ToInt64(); }
  public static int Count(uint wantedPid) {
    int count = 0;
    EnumWindows((h,s) => {
      uint pid; GetWindowThreadProcessId(h,out pid);
      if (pid == wantedPid) {
        var text = new StringBuilder(512); GetWindowText(h,text,text.Capacity);
        if (text.ToString().StartsWith("NX_FOCUS_OVERLAY_", StringComparison.Ordinal)) count++;
      }
      return true;
    }, IntPtr.Zero);
    return count;
  }
}
'@
}

function Get-OverlayWindowCount([int]$ProcessId) {
    Initialize-OverlayWindowApi
    return [int][NxFocusOverlayWindowApi]::Count([uint32]$ProcessId)
}

function New-FocusOverlayProbeXlam([string]$OutputPath) {
    if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
    foreach ($source in $ProbeSources) {
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw 'Overlay probe source missing' }
    }
    $excel = $null; $binding = $null; $books = $null; $book = $null
    $project = $null; $components = $null; $component = $null; $designer = $null; $codeModule = $null
    $importRoot = Join-Path $EvidenceRoot ('import-' + [Guid]::NewGuid().ToString('N'))
    try {
        $baseline = Get-ExcelProcessBaseline
        $excel = New-Object -ComObject Excel.Application
        $binding = Get-ExactExcelProcessOwnership $excel $baseline 'focus overlay probe build Excel'
        if (-not [bool]$binding.owned) { throw ('focus overlay build ownership rejected: ' + [string]$binding.failure) }
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $books = $excel.Workbooks
        $book = $books.Add()
        $project = $book.VBProject
        $components = $project.VBComponents
        [void](New-Item -ItemType Directory -Path $importRoot -Force)

        foreach ($source in @($ProbeSources | Where-Object { -not $_.EndsWith('.vba') })) {
            $target = Join-Path $importRoot ([IO.Path]::GetFileName($source))
            $text = [IO.File]::ReadAllText($source, [Text.Encoding]::UTF8) -replace '\r\n|\r|\n', "`r`n"
            [IO.File]::WriteAllText($target, $text, [Text.Encoding]::Default)
            try { $component = $components.Import($target) }
            finally { Release-ComObject $component; $component = $null }
        }

        $component = $components.Add(3)
        $component.Name = 'FNxFocusOverlayProbe'
        Set-VbComponentProperty $component 'Caption' 'NX_FOCUS_OVERLAY_PROBE'
        Set-VbComponentProperty $component 'Width' 100
        Set-VbComponentProperty $component 'Height' 100
        Set-VbComponentProperty $component 'StartUpPosition' 0
        Set-VbComponentProperty $component 'BorderStyle' 0
        Set-VbComponentProperty $component 'ShowModal' $false
        $designer = $component.Designer
        if ($null -eq $designer) { throw 'Overlay UserForm designer unavailable' }
        $codeModule = $component.CodeModule
        $formCode = [IO.File]::ReadAllText($ProbeSources[2], [Text.Encoding]::UTF8) -replace '\r\n|\r|\n', "`r`n"
        [void]$codeModule.AddFromString($formCode)
        [void]$book.SaveAs($OutputPath, 55)
        [void]$book.Close($false)
        Release-ComObject $book
        $book = $null
    } finally {
        if ($null -ne $book) { try { $book.Close($false) } catch {} }
        $cleanup = @(Invoke-ExcelCleanup $excel $binding @($codeModule,$designer,$component,$components,$project,$book,$books) 'focus overlay probe build Excel')
        Remove-Item -LiteralPath $importRoot -Recurse -Force -ErrorAction SilentlyContinue
        if ($cleanup.Count -ne 0) { throw ('focus overlay build cleanup failed: ' + ($cleanup -join ' | ')) }
    }
    if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) { throw 'Overlay probe XLAM was not created' }
}

function Add-WorkbookSentinels([object]$Book, [object]$Sheet) {
    $range = $null; $conditions = $null; $rule = $null; $shape = $null; $name = $null
    try {
        $range = $Sheet.Range('A1:Z100')
        $conditions = $range.FormatConditions
        $rule = $conditions.Add(2, $null, '=MOD(ROW()+COLUMN(),3)=0')
        $rule.StopIfTrue = $false
        $shape = $Sheet.Shapes.AddShape(1, 10, 10, 30, 20)
        $shape.Name = 'UserKeepShape'
        $name = $Book.Names.Add('UserKeepName', '=Normal!$A$1')
    } finally {
        Release-ComObject $name; Release-ComObject $shape; Release-ComObject $rule
        Release-ComObject $conditions; Release-ComObject $range
    }
}

function Get-WorkbookAssetFingerprint([object]$Book, [object]$Sheet) {
    # Book and Sheet are borrowed RCWs. Never FinalRelease them here because
    # PowerShell may return the same RCW identity to the caller.
    $names = $null; $sheets = $null; $cells = $null; $conditions = $null; $shapes = $null
    try {
        $names = $Book.Names
        $sheets = $Book.Worksheets
        $cells = $Sheet.Cells
        $conditions = $cells.FormatConditions
        $shapes = $Sheet.Shapes
        $signature = 'NAMES=' + [string]$names.Count + ';SHEETS=' + [string]$sheets.Count +
            ';CF=' + [string]$conditions.Count + ';SHAPES=' + [string]$shapes.Count
        return Get-TextSha256 $signature
    } finally {
        Release-ComObject $shapes
        Release-ComObject $conditions
        Release-ComObject $cells
        Release-ComObject $sheets
        Release-ComObject $names
    }
}

function Invoke-NativeUndoProbe([object]$Excel, [object]$Sheet, [int]$ExcelPid, [Int64]$ExcelHwnd) {
    $target = $null; $shell = $null
    Initialize-OverlayWindowApi
    $script:UndoDiagnostic = [ordered]@{stage='initialize';excel_hwnd=$ExcelHwnd;app_activated=$false;foreground_before=0;foreground_after=0;value_after_type=$null;undo_enabled=$false;value_after_undo=$null;failure=$null}
    try {
        $script:UndoDiagnostic.stage = 'select-target'
        [void]$Sheet.Activate()
        $target = $Sheet.Range('Z205')
        [void]$Excel.Goto($target, $true)
        $shell = New-Object -ComObject WScript.Shell
        $script:UndoDiagnostic.foreground_before = [Int64][NxFocusOverlayWindowApi]::Foreground()
        $script:UndoDiagnostic.app_activated = [bool][NxFocusOverlayWindowApi]::Activate($ExcelHwnd)
        if (-not $script:UndoDiagnostic.app_activated) { return $false }
        Start-Sleep -Milliseconds 300
        $script:UndoDiagnostic.foreground_after = [Int64][NxFocusOverlayWindowApi]::Foreground()
        $script:UndoDiagnostic.stage = 'send-keys'
        $shell.SendKeys('{F2}FOCUS_OVERLAY_UNDO{ENTER}')
        Start-Sleep -Milliseconds 900
        $script:UndoDiagnostic.stage = 'read-after-type'
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            try {
                $script:UndoDiagnostic.value_after_type = [string]$target.Value2
                $script:UndoDiagnostic.undo_enabled = [bool]$Excel.CommandBars.GetEnabledMso('Undo')
                if (-not [string]::IsNullOrEmpty($script:UndoDiagnostic.value_after_type) -or $script:UndoDiagnostic.undo_enabled) { break }
            } catch {
                $script:UndoDiagnostic.failure = $_.Exception.Message
            }
            Start-Sleep -Milliseconds 200
        }
        if ([string]::IsNullOrEmpty($script:UndoDiagnostic.value_after_type)) { return $false }
        if (-not $script:UndoDiagnostic.undo_enabled) { return $false }
        $script:UndoDiagnostic.stage = 'undo'
        [void]$Excel.Undo()
        Start-Sleep -Milliseconds 500
        $script:UndoDiagnostic.stage = 'read-after-undo'
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            try {
                $script:UndoDiagnostic.value_after_undo = [string]$target.Value2
                break
            } catch {
                $script:UndoDiagnostic.failure = $_.Exception.Message
                Start-Sleep -Milliseconds 200
            }
        }
        $script:UndoDiagnostic.stage = 'complete'
        return [string]::IsNullOrEmpty($script:UndoDiagnostic.value_after_undo)
    } catch { $script:UndoDiagnostic.failure = $_.Exception.Message; return $false }
    finally { Release-ComObject $shell; Release-ComObject $target }
}

function Test-ProcessGone([int]$ProcessId) {
    if ($ProcessId -le 0) { return $false }
    for ($attempt = 0; $attempt -lt 25; $attempt++) {
        $candidate = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        if ($null -eq $candidate) { return $true }
        try { $candidate.Dispose() } catch {}
        Start-Sleep -Milliseconds 200
    }
    return $false
}

function Invoke-RefreshMeasurements([object]$Excel, [string]$ProbeName) {
    $samples = New-Object Collections.Generic.List[double]
    for ($index = 0; $index -lt ($RefreshCount + $WarmupCount); $index++) {
        $watch = [Diagnostics.Stopwatch]::StartNew()
        [void]$Excel.Run("'" + $ProbeName + "'!FxOverlayRefresh")
        $watch.Stop()
        if ($index -ge $WarmupCount) { [void]$samples.Add($watch.Elapsed.TotalMilliseconds) }
    }
    return Get-Percentile95 $samples.ToArray()
}

[void](New-Item -ItemType Directory -Path $EvidenceRoot -Force)
foreach ($path in @($ProbePath,$NormalPath,$EvidencePath)) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }

$gateNames = @(
    'saved_state_preserved','undo_preserved','workbook_assets_unchanged','endpoint_axes_exact',
    'click_through_no_activate','scroll_zoom_resize_tracking','refresh_p95_within_50ms',
    'multi_document_isolated','cleanup_zero_windows','excel_process_zero_residue'
)
$gates = [ordered]@{}
foreach ($name in $gateNames) { $gates[$name] = $false }
$measurements = [ordered]@{refresh_count=$RefreshCount;warmup_count=$WarmupCount;refresh_p95_ms=$null;timer_refresh_delta=0}
$diagnostic = [ordered]@{}
$failure = $null
$excelVersion = 'unavailable'; $excelBuild = 'unavailable'; $probeSha = '0' * 64
$sourceTreeSha = Get-TextSha256 (($ProbeSources | ForEach-Object { Get-Sha256 $_ }) -join "`n")
$runPid = 0
$stage = 'initialize'

$excel = $null; $binding = $null; $books = $null; $probeBook = $null
$normalBook = $null; $normalSheet = $null; $secondBook = $null; $secondSheet = $null
try {
    $stage = 'build-probe'
    New-FocusOverlayProbeXlam $ProbePath
    $probeSha = Get-Sha256 $ProbePath
    $stage = 'create-runtime-excel'
    $baseline = Get-ExcelProcessBaseline
    $excel = New-Object -ComObject Excel.Application
    $binding = Get-ExactExcelProcessOwnership $excel $baseline 'focus overlay feasibility Excel'
    if (-not [bool]$binding.owned) { throw ('focus overlay ownership rejected: ' + [string]$binding.failure) }
    $runPid = [int]$binding.pid
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excelVersion = [string]$excel.Version
    try { $excelBuild = [string]$excel.Build } catch { $excelBuild = $excelVersion }
    $books = $excel.Workbooks
    $stage = 'open-probe'
    $probeBook = $books.Open($ProbePath)
    $probeBook.IsAddin = $true
    $probeName = [string]$probeBook.Name

    $stage = 'create-normal-workbook'
    $normalBook = $books.Add()
    $normalSheet = $normalBook.Worksheets.Item(1)
    $normalSheet.Name = 'Normal'
    $normalSheet.Range('A1:AZ200').Value2 = 1
    $stage = 'add-normal-sentinels'
    Add-WorkbookSentinels $normalBook $normalSheet
    $stage = 'save-normal-workbook'
    [void]$normalBook.SaveAs($NormalPath, 51)
    $stage = 'fingerprint-normal-before'
    $normalFingerprintBefore = Get-WorkbookAssetFingerprint $normalBook $normalSheet
    [void]$normalSheet.Activate()
    [void]$normalSheet.Range('A1:C3').Select()
    $stage = 'start-overlay'
    if (-not [bool]$excel.Run("'" + $probeName + "'!FxOverlayStart", $excel, $FocusColor)) { throw 'Overlay start rejected' }
    Start-Sleep -Milliseconds 500

    $stage = 'inspect-initial-overlay'
    $selectionSignature = [string]$excel.Run("'" + $probeName + "'!FxOverlaySelectionSignature")
    $boundsInitial = [string]$excel.Run("'" + $probeName + "'!FxOverlayBoundsSignature")
    $gates.endpoint_axes_exact = (
        $selectionSignature -eq '1:3|1:3' -and
        [int]$excel.Run("'" + $probeName + "'!FxOverlayWindowCount") -eq 4 -and
        [int]$excel.Run("'" + $probeName + "'!FxOverlayVisibleWindowCount") -eq 4 -and
        [bool]$excel.Run("'" + $probeName + "'!FxOverlayExpectedBoundsMatch", 3)
    )
    $gates.click_through_no_activate = [bool]$excel.Run("'" + $probeName + "'!FxOverlayStyleContractSatisfied")
    $gates.saved_state_preserved = [bool]$normalBook.Saved
    $stage = 'native-undo'
    $gates.undo_preserved = Invoke-NativeUndoProbe $excel $normalSheet $runPid ([Int64]$binding.hwnd)

    $stage = 'tracking-probe'
    [void]$normalSheet.Range('J20:L22').Select()
    $refreshBefore = [int]$excel.Run("'" + $probeName + "'!FxOverlayRefreshCount")
    $window = $excel.ActiveWindow
    $originalZoom = [int]$window.Zoom
    $originalScrollRow = [int]$window.ScrollRow
    $originalScrollColumn = [int]$window.ScrollColumn
    $window.ScrollRow = 10
    $window.ScrollColumn = 5
    $window.Zoom = 125
    Start-Sleep -Milliseconds 700
    $refreshAfter = [int]$excel.Run("'" + $probeName + "'!FxOverlayRefreshCount")
    $trackingMatch = [bool]$excel.Run("'" + $probeName + "'!FxOverlayExpectedBoundsMatch", 3)
    $window.Zoom = $originalZoom
    $window.ScrollRow = $originalScrollRow
    $window.ScrollColumn = $originalScrollColumn
    Start-Sleep -Milliseconds 400
    $trackingRestored = [bool]$excel.Run("'" + $probeName + "'!FxOverlayExpectedBoundsMatch", 3)
    $measurements.timer_refresh_delta = $refreshAfter - $refreshBefore
    $gates.scroll_zoom_resize_tracking = ($refreshAfter -gt $refreshBefore -and $trackingMatch -and $trackingRestored)
    Release-ComObject $window

    $stage = 'measure-refresh'
    $measurements.refresh_p95_ms = Invoke-RefreshMeasurements $excel $probeName
    $gates.refresh_p95_within_50ms = ([double]$measurements.refresh_p95_ms -le 50.0)

    $stage = 'create-second-workbook'
    $secondBook = $books.Add()
    $secondSheet = $secondBook.Worksheets.Item(1)
    $secondSheet.Name = 'Second'
    $secondSheet.Range('A1:Z100').Value2 = 2
    $stage = 'fingerprint-second-before'
    $secondFingerprintBefore = Get-WorkbookAssetFingerprint $secondBook $secondSheet
    [void]$secondSheet.Activate()
    [void]$secondSheet.Range('B2:D4').Select()
    Start-Sleep -Milliseconds 400
    $secondSignature = [string]$excel.Run("'" + $probeName + "'!FxOverlaySelectionSignature")
    $stage = 'fingerprint-after'
    $secondFingerprintAfter = Get-WorkbookAssetFingerprint $secondBook $secondSheet
    $normalFingerprintAfter = Get-WorkbookAssetFingerprint $normalBook $normalSheet
    $gates.multi_document_isolated = ($secondSignature -eq '2:4|2:4' -and [int]$excel.Run("'" + $probeName + "'!FxOverlayWindowCount") -eq 4)
    $gates.workbook_assets_unchanged = ($normalFingerprintBefore -eq $normalFingerprintAfter -and $secondFingerprintBefore -eq $secondFingerprintAfter)

    $stage = 'stop-overlay'
    [void]$excel.Run("'" + $probeName + "'!FxOverlayStop")
    Start-Sleep -Milliseconds 300
    $engineWindowsAfterStop = [int]$excel.Run("'" + $probeName + "'!FxOverlayWindowCount")
    $nativeWindowsAfterStop = Get-OverlayWindowCount $runPid
    $gates.cleanup_zero_windows = ($engineWindowsAfterStop -eq 0 -and $nativeWindowsAfterStop -eq 0)
    $diagnostic.selection_signature = $selectionSignature
    $diagnostic.second_selection_signature = $secondSignature
    $diagnostic.initial_bounds = $boundsInitial
    $diagnostic.engine_windows_after_stop = $engineWindowsAfterStop
    $diagnostic.native_windows_after_stop = $nativeWindowsAfterStop
    $diagnostic.last_error = [string]$excel.Run("'" + $probeName + "'!FxOverlayLastError")
    $diagnostic.undo = $script:UndoDiagnostic
} catch {
    $failure = $stage + ': ' + $_.Exception.Message
} finally {
    if ($null -ne $excel -and $null -ne $probeBook) {
        try { [void]$excel.Run("'" + [string]$probeBook.Name + "'!FxOverlayStop") } catch {}
    }
    foreach ($book in @($secondBook,$normalBook,$probeBook)) {
        if ($null -ne $book) { try { [void]$book.Close($false) } catch {} }
    }
    $cleanup = @(Invoke-ExcelCleanup $excel $binding @($secondSheet,$secondBook,$normalSheet,$normalBook,$probeBook,$books) 'focus overlay feasibility Excel')
    if ($cleanup.Count -ne 0) {
        if ([string]::IsNullOrWhiteSpace($failure)) { $failure = 'cleanup: ' + ($cleanup -join ' | ') }
        else { $failure += ' | cleanup: ' + ($cleanup -join ' | ') }
    }
}

if ($runPid -gt 0) {
    $gates.excel_process_zero_residue = Test-ProcessGone $runPid
}
$passed = @($gateNames | Where-Object { [bool]$gates[$_] }).Count
$status = if ($passed -eq $gateNames.Count -and [string]::IsNullOrWhiteSpace($failure)) { 'PASS' } else { 'HOLD' }
$receipt = [ordered]@{
    schema_version = 1
    status = $status
    run_id = $RunId
    source_tree_sha256 = $sourceTreeSha
    probe_xlam_sha256 = $probeSha
    excel_version = $excelVersion
    excel_build = $excelBuild
    office_bitness = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
    selected_strategy = 'owned_modeless_overlay'
    measurements = $measurements
    gates = $gates
    diagnostic = [ordered]@{details=$diagnostic;failure=$failure}
}
Write-AtomicJson $EvidencePath $receipt
Write-Output ($status + '|FocusOverlayFeasibility|' + [string]$passed + '/' + [string]$gateNames.Count)
Write-Output $EvidencePath
if ($status -ne 'PASS') { exit 24 }
exit 0
