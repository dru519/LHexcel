param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RunId = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }

$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$ProductXlam = [IO.Path]::GetFullPath($ProductXlam)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not (Test-Path -LiteralPath $ProductXlam -PathType Leaf)) { throw 'Product.xlam not found' }
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing ShortcutManager evidence root is rejected' }
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { throw 'Shortcut manager suite requires zero pre-existing Excel processes' }
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}

function Write-AtomicJson([string]$Path, [object]$Value) {
    $temporary = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    $encoding = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($temporary, (($Value | ConvertTo-Json -Depth 12) + "`n"), $encoding)
    [IO.File]::Move($temporary, $Path)
}

function Complete-ExcelPhase([object]$Excel, [object]$Addin, [object]$Book, [object]$Sheet, [object]$TargetRange, [object]$Binding, [string]$Label) {
    $result = [ordered]@{label=$Label;exit_mode='NOT_STARTED';failure=$null;owned_pid=0}
    try {
        if ($null -ne $Book) { try { $Book.Close($false) } catch {} }
        if ($null -ne $Addin) { try { $Addin.Close($false) } catch {} }
        Release-ComObject $TargetRange
        Release-ComObject $Sheet
        Release-ComObject $Book
        Release-ComObject $Addin
        if ($null -ne $Excel) { try { $Excel.Quit() } catch {} }
        Release-ComObject $Excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        if ($null -eq $Binding -or -not [bool]$Binding.owned -or $null -eq $Binding.process) {
            throw ($Label + ' has no exact owned Excel process for lifecycle verification')
        }
        $result.owned_pid = [int]$Binding.pid
        $stop = Stop-ExactProcessAfterGrace $Binding.process $Label 15000 -Detailed
        $result.exit_mode = [string]$stop.exit_mode
        $result.failure = $stop.failure
        if ($result.exit_mode -cne 'NATURAL' -or -not [string]::IsNullOrWhiteSpace([string]$result.failure)) {
            throw ($Label + ' did not exit naturally: ' + $result.exit_mode + '; ' + [string]$result.failure)
        }
    } catch {
        if ([string]::IsNullOrWhiteSpace([string]$result.failure)) { $result.failure = $_.Exception.Message }
    } finally {
        if ($null -ne $Binding -and $null -ne $Binding.process) { try { $Binding.process.Dispose() } catch {} }
    }
    return [pscustomobject]$result
}

function Get-ExcelProcessIds {
    $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
    try { return @($processes | ForEach-Object { [int]$_.Id }) }
    finally { foreach ($process in $processes) { try { $process.Dispose() } catch {} } }
}

function Wait-FreshExcelProcessExit([int]$TimeoutMs = 2000) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $ids = @()
    do {
        $ids = @(Get-ExcelProcessIds)
        if ($ids.Count -eq 0) {
            return [pscustomobject][ordered]@{status='CLEARED';elapsed_ms=$watch.ElapsedMilliseconds;remaining_pids=@()}
        }
        if ($watch.ElapsedMilliseconds -lt $TimeoutMs) { Start-Sleep -Milliseconds 100 }
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    return [pscustomobject][ordered]@{status='REMAINING';elapsed_ms=$watch.ElapsedMilliseconds;remaining_pids=$ids}
}

$excel = $null
$addin = $null
$testBook = $null
$testSheet = $null
$targetRange = $null
$firstBinding = $null
$restartBinding = $null
$firstLifecycle = [ordered]@{label='Shortcut manager first Excel';exit_mode='NOT_STARTED';failure=$null;owned_pid=0}
$restartLifecycle = [ordered]@{label='Shortcut manager restarted Excel';exit_mode='NOT_STARTED';failure=$null;owned_pid=0}
$failure = $null
$failureStage = 'initialization'
$cases = New-Object 'Collections.Generic.List[object]'
$restartExitImmediate = @()
$restartExitConvergence = [pscustomobject][ordered]@{status='NOT_RUN';elapsed_ms=0;remaining_pids=@()}
$profileRoot = [IO.Path]::GetFullPath((Join-Path $EvidenceRoot 'Profile'))
$previousProfileRoot = [Environment]::GetEnvironmentVariable('LHEXCEL_PROFILE_ROOT', 'Process')
try {
    [Environment]::SetEnvironmentVariable('LHEXCEL_PROFILE_ROOT', $profileRoot, 'Process')
    $firstBaseline = Get-ExcelProcessBaseline
    $excel = New-Object -ComObject Excel.Application
    $firstBinding = Get-ExactExcelProcessOwnership $excel $firstBaseline 'Shortcut manager first Excel'
    if (-not [bool]$firstBinding.owned) { throw ('Excel ownership rejected: ' + [string]$firstBinding.failure) }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $addin = $excel.Workbooks.Open($ProductXlam, 0, $true)
    $macro = "'" + [string]$addin.Name + "'!"
    $failureStage = 'first_profile_root_assertion'
    $actualProfileRoot = [string]$excel.Run($macro + 'NxLHexcelProfileRoot')
    if ($actualProfileRoot -cne $profileRoot) {
        throw ('Profile isolation rejected before shortcut assignment: expected=' + $profileRoot + '; actual=' + $actualProfileRoot)
    }
    $cases.Add([pscustomobject][ordered]@{name='profile_root_first_open';status='PASS';expected=$profileRoot;actual=$actualProfileRoot})

    $failureStage = 'first_capture_and_reservation'
    $observations = @(
        @{name='capture_ctrl_shift_k'; actual=[string]$excel.Run($macro + 'NxShortcutCaptureDisplay', 75, 3); expected='Ctrl+Shift+K'},
        @{name='capture_ctrl_alt_f8'; actual=[string]$excel.Run($macro + 'NxShortcutCaptureDisplay', 119, 6); expected='Ctrl+Alt+F8'},
        @{name='reject_ctrl_only'; actual=[string]$excel.Run($macro + 'NxShortcutCaptureDisplay', 75, 2); expected=''},
        @{name='reject_without_ctrl'; actual=[string]$excel.Run($macro + 'NxShortcutCaptureDisplay', 75, 1); expected=''}
    )
    foreach ($row in $observations) {
        $cases.Add([pscustomobject][ordered]@{
            name = $row.name
            status = if ($row.actual -ceq $row.expected) { 'PASS' } else { 'FAIL' }
            expected = $row.expected
            actual = $row.actual
        })
    }

    foreach ($reserved in @(
        @{modifier='Ctrl';key='C'},
        @{modifier='Ctrl+Shift';key='L'},
        @{modifier='Ctrl+Alt';key='V'},
        @{modifier='Ctrl+Alt+Shift';key='F2'},
        @{modifier='Ctrl+Alt+Shift';key='F9'},
        @{modifier='Ctrl+Alt+Shift';key='H'},
        @{modifier='Ctrl+Alt+Shift';key='T'}
    )) {
        $actual = [bool]$excel.Run($macro + 'NxShortcutIsReservedCombination', $reserved.modifier, $reserved.key)
        $cases.Add([pscustomobject][ordered]@{
            name = 'reserved_' + $reserved.modifier.Replace('+','_').ToLowerInvariant() + '_' + $reserved.key.ToLowerInvariant()
            status = if ($actual) { 'PASS' } else { 'FAIL' }
            expected = $true
            actual = $actual
        })
    }

    $routeKey = 'command:NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-NUMBER'
    $failureStage = 'first_fixture_target_seed'
    $testBook = $excel.Workbooks.Add()
    $testSheet = $testBook.Worksheets.Item(1)
    [void]$testBook.Activate()
    [void]$testSheet.Activate()
    $targetRange = $testSheet.Range('A1')
    $targetRange.NumberFormat = '0.00'
    [void]$targetRange.Select()
    $failureStage = 'unused_chord_conflict_reason'
    $freeChordReason = [string]$excel.Run($macro + 'NxShortcutConflictReason', 'Ctrl+Alt', 'K', $routeKey)
    $cases.Add([pscustomobject][ordered]@{
        name = 'unused_chord_has_no_conflict'
        status = if ($freeChordReason.Length -eq 0) { 'PASS' } else { 'FAIL' }
        expected = ''
        actual = $freeChordReason
    })
    $failureStage = 'first_shortcut_assignment'
    $excel.Run($macro + 'NxShortcutsAssign', $routeKey, 'Ctrl+Alt', 'K')
    $actual = [string]$excel.Run($macro + 'NxShortcutsBindingDisplayKey', 1)
    $cases.Add([pscustomobject][ordered]@{
        name = 'assigns_safe_chord'
        status = if ($actual -ceq 'Ctrl+Alt+K') { 'PASS' } else { 'FAIL' }
        expected = 'Ctrl+Alt+K'
        actual = $actual
    })
    $failureStage = 'first_shortcut_dispatch'
    $excel.Run($macro + 'NxShortcutDispatch01')
    $actual = [string]$targetRange.NumberFormat
    $cases.Add([pscustomobject][ordered]@{
        name = 'dispatch_executes_assigned_command'
        status = if ($actual -ceq '#,##0') { 'PASS' } else { 'FAIL' }
        expected = '#,##0'
        actual = $actual
    })
    $failureStage = 'first_persistence_read'
    $shortcutFile = Join-Path $profileRoot 'Settings\shortcuts-v2.cfg'
    $actual = if (Test-Path -LiteralPath $shortcutFile -PathType Leaf) { [IO.File]::ReadAllText($shortcutFile) } else { '' }
    $expected = "NXSHORT2`r`n1|Ctrl+Alt|K|$routeKey"
    $cases.Add([pscustomobject][ordered]@{
        name = 'assignment_persists_route_binding'
        status = if ($actual -like "*$expected*") { 'PASS' } else { 'FAIL' }
        expected = $expected
        actual = $actual
    })

    $failureStage = 'first_excel_lifecycle'
    $firstLifecycle = Complete-ExcelPhase $excel $addin $testBook $testSheet $targetRange $firstBinding 'Shortcut manager first Excel'
    $targetRange = $null
    $testSheet = $null
    $testBook = $null
    $addin = $null
    $excel = $null
    $firstBinding = $null
    if ($firstLifecycle.exit_mode -cne 'NATURAL' -or -not [string]::IsNullOrWhiteSpace([string]$firstLifecycle.failure)) {
        throw ('First Excel lifecycle rejected: ' + [string]$firstLifecycle.failure)
    }

    $failureStage = 'restart_excel_exit_convergence'
    $restartExitImmediate = @(Get-ExcelProcessIds)
    $restartExitConvergence = Wait-FreshExcelProcessExit 2000
    if ($restartExitConvergence.status -cne 'CLEARED') {
        throw ('Shortcut manager restart blocked by remaining EXCEL PIDs: ' + ($restartExitConvergence.remaining_pids -join ','))
    }
    $failureStage = 'restart_excel_launch'
    $restartBaseline = Get-ExcelProcessBaseline
    if (@($restartBaseline.process_ids).Count -ne 0) { throw 'Shortcut manager restart requires zero Excel processes after first lifecycle' }
    $excel = New-Object -ComObject Excel.Application
    $restartBinding = Get-ExactExcelProcessOwnership $excel $restartBaseline 'Shortcut manager restarted Excel'
    if (-not [bool]$restartBinding.owned) { throw ('Excel restart ownership rejected: ' + [string]$restartBinding.failure) }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $addin = $excel.Workbooks.Open($ProductXlam, 0, $true)
    $macro = "'" + [string]$addin.Name + "'!"
    $failureStage = 'restart_profile_root_assertion'
    $actualProfileRoot = [string]$excel.Run($macro + 'NxLHexcelProfileRoot')
    if ($actualProfileRoot -cne $profileRoot) {
        throw ('Profile isolation rejected after restart: expected=' + $profileRoot + '; actual=' + $actualProfileRoot)
    }
    $cases.Add([pscustomobject][ordered]@{name='profile_root_restart';status='PASS';expected=$profileRoot;actual=$actualProfileRoot})
    $failureStage = 'restart_binding_reload'
    $actual = [string]$excel.Run($macro + 'NxShortcutsBindingRouteKey', 1)
    $cases.Add([pscustomobject][ordered]@{
        name = 'reloads_persisted_route_binding'
        status = if ($actual -ceq $routeKey) { 'PASS' } else { 'FAIL' }
        expected = $routeKey
        actual = $actual
    })
    $failureStage = 'restart_fixture_target_seed'
    $testBook = $excel.Workbooks.Add()
    $testSheet = $testBook.Worksheets.Item(1)
    [void]$testBook.Activate()
    [void]$testSheet.Activate()
    $targetRange = $testSheet.Range('A1')
    $targetRange.NumberFormat = '0.00'
    [void]$targetRange.Select()
    $failureStage = 'restart_shortcut_dispatch'
    $excel.Run($macro + 'NxShortcutDispatch01')
    $actual = [string]$targetRange.NumberFormat
    $cases.Add([pscustomobject][ordered]@{
        name = 'reloaded_binding_executes_assigned_command'
        status = if ($actual -ceq '#,##0') { 'PASS' } else { 'FAIL' }
        expected = '#,##0'
        actual = $actual
    })
} catch {
    $failure = 'stage=' + $failureStage + '; ' + $_.Exception.Message
} finally {
    if ($null -ne $firstBinding) {
        $firstLifecycle = Complete-ExcelPhase $excel $addin $testBook $testSheet $targetRange $firstBinding 'Shortcut manager first Excel'
        $excel = $null; $addin = $null; $testBook = $null; $testSheet = $null; $targetRange = $null; $firstBinding = $null
    } elseif ($null -ne $restartBinding) {
        $restartLifecycle = Complete-ExcelPhase $excel $addin $testBook $testSheet $targetRange $restartBinding 'Shortcut manager restarted Excel'
        $excel = $null; $addin = $null; $testBook = $null; $testSheet = $null; $targetRange = $null; $restartBinding = $null
    } else {
        if ($null -ne $testBook) { try { $testBook.Close($false) } catch {} }
        if ($null -ne $addin) { try { $addin.Close($false) } catch {} }
        if ($null -ne $excel) { try { $excel.Quit() } catch {} }
        Release-ComObject $targetRange
        Release-ComObject $testSheet
        Release-ComObject $testBook
        Release-ComObject $addin
        Release-ComObject $excel
    }
    [Environment]::SetEnvironmentVariable('LHEXCEL_PROFILE_ROOT', $previousProfileRoot, 'Process')
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

$remainingExcelImmediate = @(Get-ExcelProcessIds)
$remainingExcelRecheck = Wait-FreshExcelProcessExit 2000
$remainingExcelProcessIds = @($remainingExcelRecheck.remaining_pids)
$lifecyclePassed = ($firstLifecycle.exit_mode -ceq 'NATURAL' -and
    [string]::IsNullOrWhiteSpace([string]$firstLifecycle.failure) -and
    $restartLifecycle.exit_mode -ceq 'NATURAL' -and
    [string]::IsNullOrWhiteSpace([string]$restartLifecycle.failure) -and
    $remainingExcelProcessIds.Count -eq 0)
$cases.Add([pscustomobject][ordered]@{
    name = 'cleanup_exact_owned_excel_natural_exit'
    status = if ($lifecyclePassed) { 'PASS' } else { 'FAIL' }
    expected = 'two natural exits and zero remaining EXCEL processes'
    actual = [ordered]@{first=$firstLifecycle;restart=$restartLifecycle;restart_immediate=$restartExitImmediate;restart_recheck=$restartExitConvergence;remaining_immediate=$remainingExcelImmediate;remaining_recheck=$remainingExcelRecheck;remaining_process_ids=$remainingExcelProcessIds}
})
$failed = @($cases | Where-Object status -ne 'PASS').Count
$receipt = [pscustomobject][ordered]@{
    schema = 'lhexcel-shortcut-manager-v1'
    run_id = $RunId
    generated_utc = [DateTime]::UtcNow.ToString('o')
    product = $ProductXlam
    product_sha256 = (Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash.ToLowerInvariant()
    status = if ($null -eq $failure -and $failed -eq 0 -and $cases.Count -eq 20) { 'PASS' } else { 'FAIL' }
    failure = $failure
    failure_stage = $failureStage
    cases = $cases.ToArray()
    lifecycle = [ordered]@{first=$firstLifecycle;restart=$restartLifecycle;restart_immediate=$restartExitImmediate;restart_recheck=$restartExitConvergence;remaining_immediate=$remainingExcelImmediate;remaining_recheck=$remainingExcelRecheck;remaining_process_ids=$remainingExcelProcessIds}
}
Write-AtomicJson (Join-Path $EvidenceRoot 'ShortcutManager.json') $receipt
if ($receipt.status -ne 'PASS') { throw 'Shortcut manager native suite failed' }
