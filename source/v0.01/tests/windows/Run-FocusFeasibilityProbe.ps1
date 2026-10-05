param(
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$RunId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$MoveCount = 500
$WarmupCount = 20
$SourceRoot = [IO.Path]::GetFullPath($DataRoot)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
$EvidencePath = Join-Path $EvidenceRoot 'FocusFeasibility.json'
$ProbePath = Join-Path $EvidenceRoot 'FocusFeasibilityProbe.xlam'
$NormalPath = Join-Path $EvidenceRoot 'focus-normal.xlsx'
$StressPath = Join-Path $EvidenceRoot 'focus-stress.xlsx'
$ProbeSources = @(
    (Join-Path $SourceRoot 'tests\vba\focus\CFxFocusProbeEvents.cls'),
    (Join-Path $SourceRoot 'tests\vba\focus\FxFocusProbeApi.bas')
)
$MoveAddresses = @('A1:C3','D4:F6','J10:L12','B20:D22','M30:O32','C45:E47','R8:T10','F60:H62')
$FocusColor = 14207071

. (Join-Path $SourceRoot 'build\Excel-ProcessLifecycle.ps1')

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Get-Sha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
}

function Get-TextSha256([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').ToLowerInvariant()
    } finally {
        $sha.Dispose()
    }
}

function Get-Percentile95([double[]]$Values) {
    if ($null -eq $Values -or $Values.Count -eq 0) { throw 'No latency samples' }
    $ordered = @($Values | Sort-Object)
    $index = [Math]::Ceiling($ordered.Count * 0.95) - 1
    return [Math]::Round([double]$ordered[[Math]::Max(0, $index)], 3)
}

function Write-AtomicJson([string]$Path, [object]$Value) {
    $directory = Split-Path -Parent $Path
    [void](New-Item -ItemType Directory -Path $directory -Force)
    $temporary = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    $writer = New-Object IO.StreamWriter($temporary, $false, (New-Object Text.UTF8Encoding($false)))
    try {
        $writer.Write(($Value | ConvertTo-Json -Depth 16 -Compress) + "`n")
        $writer.Flush()
        $writer.BaseStream.Flush($true)
    } finally {
        $writer.Close()
    }
    if (Test-Path -LiteralPath $Path) {
        $backup = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.bak'
        try { [IO.File]::Replace($temporary, $Path, $backup) }
        finally { if (Test-Path -LiteralPath $backup) { Remove-Item -LiteralPath $backup -Force } }
    } else {
        [IO.File]::Move($temporary, $Path)
    }
}

function Get-DisplayScalePercent {
    try {
        $dpi = [int](Get-ItemPropertyValue -LiteralPath 'HKCU:\Control Panel\Desktop\WindowMetrics' -Name AppliedDPI -ErrorAction Stop)
        if ($dpi -gt 0) { return [int][Math]::Round(($dpi / 96.0) * 100.0) }
    } catch {}
    return 100
}

function Get-CalculationModeName([int]$Value) {
    if ($Value -eq -4105) { return 'automatic' }
    if ($Value -eq -4135) { return 'manual' }
    return 'semiautomatic'
}

function Invoke-ExcelCleanup([object]$Excel, [object]$Binding, [object[]]$ComObjects, [string]$Label) {
    $failures = New-Object Collections.Generic.List[string]
    $owned = $false
    if ($null -ne $Binding) {
        try { [void](Assert-ExactExcelProcessOwnership $Binding); $owned = $true }
        catch { [void]$failures.Add($_.Exception.Message) }
    }
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

function New-FocusProbeXlam([string]$OutputPath) {
    if (Test-Path -LiteralPath $OutputPath) { throw 'Existing focus probe artifact rejected' }
    foreach ($source in $ProbeSources) {
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw 'Focus probe source missing' }
    }
    $excel = $null; $binding = $null; $books = $null; $book = $null
    $project = $null; $components = $null
    $importRoot = Join-Path $EvidenceRoot ('import-' + [Guid]::NewGuid().ToString('N'))
    try {
        $baseline = Get-ExcelProcessBaseline
        $excel = New-Object -ComObject Excel.Application
        $binding = Get-ExactExcelProcessOwnership $excel $baseline 'focus probe build Excel'
        if (-not [bool]$binding.owned) { throw ('focus probe build ownership rejected: ' + [string]$binding.failure) }
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $books = $excel.Workbooks
        $book = $books.Add()
        $project = $book.VBProject
        $components = $project.VBComponents
        [void](New-Item -ItemType Directory -Path $importRoot -Force)
        foreach ($source in $ProbeSources) {
            $target = Join-Path $importRoot ([IO.Path]::GetFileName($source))
            $text = [IO.File]::ReadAllText($source, [Text.Encoding]::UTF8) -replace '\r\n|\r|\n', "`r`n"
            [IO.File]::WriteAllText($target, $text, [Text.Encoding]::Default)
            $component = $null
            try { $component = $components.Import($target) }
            finally { Release-ComObject $component }
        }
        $book.SaveAs($OutputPath, 55)
        $book.Close($false)
        Release-ComObject $book
        $book = $null
    } finally {
        if ($null -ne $book) { try { $book.Close($false) } catch {} }
        $cleanup = @(Invoke-ExcelCleanup $excel $binding @($components,$project,$book,$books) 'focus probe build Excel')
        Remove-Item -LiteralPath $importRoot -Recurse -Force -ErrorAction SilentlyContinue
        if ($cleanup.Count -ne 0) { throw ('focus probe build cleanup failed: ' + ($cleanup -join ' | ')) }
    }
    if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) { throw 'Focus probe XLAM was not created' }
}

function Get-UserRuleRecords([object]$Sheet) {
    $records = New-Object Collections.Generic.List[string]
    $conditions = $null
    try {
        $conditions = $Sheet.Cells.FormatConditions
        for ($index = 1; $index -le [int]$conditions.Count; $index++) {
            $rule = $null; $applies = $null; $interior = $null
            try {
                $rule = $conditions.Item($index)
                $formula = ''
                try { $formula = [string]$rule.Formula1 } catch {}
                if ($formula -match 'NX_FOCUS_RULE_V002_(ROW|COLUMN)') { continue }
                try { $applies = $rule.AppliesTo } catch {}
                try { $interior = $rule.Interior } catch {}
                $address = if ($null -eq $applies) { '' } else { [string]$applies.Address($true,$true,1,$false) }
                $color = if ($null -eq $interior) { 0 } else { try { [int]$interior.Color } catch { 0 } }
                $stop = try { [bool]$rule.StopIfTrue } catch { $false }
                [void]$records.Add(([string]$rule.Type + '|' + [string]$rule.Priority + '|' + $formula + '|' + $address + '|' + $stop + '|' + $color))
            } finally {
                Release-ComObject $interior; Release-ComObject $applies; Release-ComObject $rule
            }
        }
    } finally { Release-ComObject $conditions }
    return (($records | Sort-Object) -join "`n")
}

function Get-UserRuleFingerprint([object]$Sheet) {
    return Get-TextSha256 (Get-UserRuleRecords $Sheet)
}

function Add-UserRules([object]$Sheet, [string]$Address, [int]$Count) {
    $range = $null; $conditions = $null
    try {
        $range = $Sheet.Range($Address)
        $conditions = $range.FormatConditions
        for ($index = 1; $index -le $Count; $index++) {
            $rule = $null
            try {
                $modulus = 2 + ($index % 17)
                $formula = '=MOD(ROW()+COLUMN(),' + [string]$modulus + ')=0'
                $rule = $conditions.Add(2, $null, $formula)
                $rule.StopIfTrue = $false
            } finally { Release-ComObject $rule }
        }
    } finally { Release-ComObject $conditions; Release-ComObject $range }
}

function Get-ManagedRuleCount([object]$Excel, [string]$ProbeName, [object]$Sheet) {
    return [int]$Excel.Run("'" + $ProbeName + "'!FxManagedRuleCount", $Sheet)
}

function Invoke-MoveSequence([object]$Sheet, [int]$Count, [int]$Warmups) {
    $samples = New-Object Collections.Generic.List[double]
    for ($index = 0; $index -lt $Count; $index++) {
        $range = $null
        try {
            $range = $Sheet.Range($MoveAddresses[$index % $MoveAddresses.Count])
            $watch = [Diagnostics.Stopwatch]::StartNew()
            [void]$range.Select()
            $watch.Stop()
            if ($index -ge $Warmups) { [void]$samples.Add($watch.Elapsed.TotalMilliseconds) }
        } finally { Release-ComObject $range }
    }
    return Get-Percentile95 $samples.ToArray()
}

function Invoke-NativeUndoProbe([object]$Excel, [object]$Sheet, [int]$ExcelPid) {
    $target = $null; $shell = $null
    try {
        [void]$Sheet.Activate()
        $target = $Sheet.Range('Z9')
        [void]$target.ClearContents()
        [void]$target.Select()
        $Excel.Visible = $true
        $shell = New-Object -ComObject WScript.Shell
        if (-not $shell.AppActivate($ExcelPid)) { return $false }
        Start-Sleep -Milliseconds 300
        $shell.SendKeys('FOCUS_UNDO_PROBE{ENTER}')
        Start-Sleep -Milliseconds 800
        if (-not [bool]$Excel.CommandBars.GetEnabledMso('Undo')) { return $false }
        [void]$Excel.Undo()
        Start-Sleep -Milliseconds 300
        return [string]::IsNullOrEmpty([string]$target.Value2)
    } catch { return $false }
    finally { Release-ComObject $shell; Release-ComObject $target }
}

[void](New-Item -ItemType Directory -Path $EvidenceRoot -Force)
$gateNames = @(
    'saved_state_preserved','undo_preserved','exact_two_rules_after_500_moves',
    'normal_p95_within_50ms','stress_p95_within_150ms','cleanup_zero_assets',
    'user_rules_unchanged','paste_has_no_duplicates','cancel_paths_consistent',
    'multi_document_isolated'
)
$gates = [ordered]@{}
foreach ($name in $gateNames) { $gates[$name] = $false }
$measurements = [ordered]@{move_count=$MoveCount;warmup_count=$WarmupCount;normal_p95_ms=$null;stress_p95_ms=$null}
$diagnostic = [ordered]@{}
$failure = $null
$excelVersion = 'unavailable'; $excelBuild = 'unavailable'; $calculationMode = 'automatic'
$probeSha = '0' * 64
$fixtureSha = Get-TextSha256 ('normal=8000;stress=52000;user_rules=3,64;moves=' + ($MoveAddresses -join ','))
$sourceTreeSha = Get-TextSha256 (($ProbeSources | ForEach-Object { (Get-Sha256 $_) }) -join "`n")

$excel = $null; $binding = $null; $books = $null; $probeBook = $null
$normalBook = $null; $normalSheet = $null; $stressBook = $null; $stressSheet = $null
$isolationBook = $null; $isolationSheet = $null
try {
    New-FocusProbeXlam $ProbePath
    $probeSha = Get-Sha256 $ProbePath

    $baseline = Get-ExcelProcessBaseline
    $excel = New-Object -ComObject Excel.Application
    $binding = Get-ExactExcelProcessOwnership $excel $baseline 'focus feasibility Excel'
    if (-not [bool]$binding.owned) { throw ('focus feasibility ownership rejected: ' + [string]$binding.failure) }
    $excel.DisplayAlerts = $false
    $excel.Visible = $true
    $excelVersion = [string]$excel.Version
    try { $excelBuild = [string]$excel.Build } catch { $excelBuild = $excelVersion }
    $calculationMode = Get-CalculationModeName ([int]$excel.Calculation)
    $books = $excel.Workbooks
    $probeBook = $books.Open($ProbePath)
    $probeBook.IsAddin = $true
    $probeName = [string]$probeBook.Name
    $excel.Run("'" + $probeName + "'!FxFocusStart", $excel)

    $normalBook = $books.Add()
    $normalSheet = $normalBook.Worksheets.Item(1)
    $normalSheet.Name = 'Normal'
    $normalSheet.Range('A1:T400').Value2 = 1
    Add-UserRules $normalSheet 'A1:Z400' 3
    $normalRecordsBefore = Get-UserRuleRecords $normalSheet
    $normalFingerprintBefore = Get-UserRuleFingerprint $normalSheet
    $normalSheet.Activate()
    $normalSheet.Range('A1:C3').Select()
    $normalInstalled = [bool]$excel.Run("'" + $probeName + "'!FxInstallRules", $normalSheet, $normalSheet.Range('A1:Z400'), $FocusColor, $false)
    if (-not $normalInstalled) {
        $probeFailure = [string]$excel.Run("'" + $probeName + "'!FxLastError")
        throw ('addin_owned_name_cache rejected by Excel conditional-format parser: ' + $probeFailure)
    }
    $normalBook.SaveAs($NormalPath, 51)
    $savedBeforeMoves = [bool]$normalBook.Saved
    $measurements.normal_p95_ms = Invoke-MoveSequence $normalSheet $MoveCount $WarmupCount
    $normalSheet.Range('A1:C3').Select()
    $geometry = [string]$excel.Run("'" + $probeName + "'!FxSelectionGeometry")
    $endpointLogic = (
        [bool]$excel.Run("'" + $probeName + "'!FxProbePointMatches", $normalSheet, 1, 10, $true) -and
        [bool]$excel.Run("'" + $probeName + "'!FxProbePointMatches", $normalSheet, 3, 10, $true) -and
        [bool]$excel.Run("'" + $probeName + "'!FxProbePointMatches", $normalSheet, 10, 1, $false) -and
        [bool]$excel.Run("'" + $probeName + "'!FxProbePointMatches", $normalSheet, 10, 3, $false) -and
        -not [bool]$excel.Run("'" + $probeName + "'!FxProbePointMatches", $normalSheet, 2, 10, $true) -and
        -not [bool]$excel.Run("'" + $probeName + "'!FxProbePointMatches", $normalSheet, 10, 2, $false) -and
        -not [bool]$excel.Run("'" + $probeName + "'!FxProbePointMatches", $normalSheet, 1, 1, $true)
    )
    $normalRuleCount = Get-ManagedRuleCount $excel $probeName $normalSheet
    $gates.saved_state_preserved = ($savedBeforeMoves -and [bool]$normalBook.Saved)
    $gates.exact_two_rules_after_500_moves = ($normalRuleCount -eq 2 -and $geometry -eq '1:3|1:3' -and $endpointLogic)
    $gates.normal_p95_within_50ms = ([double]$measurements.normal_p95_ms -le 50.0)
    $gates.undo_preserved = Invoke-NativeUndoProbe $excel $normalSheet ([int]$binding.pid)
    $excel.Run("'" + $probeName + "'!FxFocusResumeAfterEdit")

    $normalSheet.Range('A1:C3').Copy($normalSheet.Range('M20'))
    Start-Sleep -Milliseconds 200
    try { $excel.Undo() } catch {}
    $normalRepaired = [bool]$excel.Run("'" + $probeName + "'!FxRepairRules", $normalSheet, $normalSheet.Range('A1:Z400'), $FocusColor, $false)
    $gates.paste_has_no_duplicates = ($normalRepaired -and (Get-ManagedRuleCount $excel $probeName $normalSheet) -eq 2)

    $normalFingerprintBeforeCancel = Get-UserRuleFingerprint $normalSheet
    if ($normalFingerprintBeforeCancel -eq $normalFingerprintBefore) { $normalBook.Saved = $true }
    $savedBeforeCancel = [bool]$normalBook.Saved
    $excel.Run("'" + $probeName + "'!FxDeleteRules", $normalSheet)
    if ($savedBeforeCancel -and (Get-UserRuleFingerprint $normalSheet) -eq $normalFingerprintBefore) { $normalBook.Saved = $true }
    $cancelRestored = [bool]$excel.Run("'" + $probeName + "'!FxInstallRules", $normalSheet, $normalSheet.Range('A1:Z400'), $FocusColor, $false)
    if ($savedBeforeCancel -and (Get-UserRuleFingerprint $normalSheet) -eq $normalFingerprintBefore) { $normalBook.Saved = $true }
    $gates.cancel_paths_consistent = ($cancelRestored -and [bool]$normalBook.Saved -and (Get-ManagedRuleCount $excel $probeName $normalSheet) -eq 2)

    $stressBook = $books.Add()
    $stressSheet = $stressBook.Worksheets.Item(1)
    $stressSheet.Name = 'Stress'
    $stressSheet.Range('A1:AZ1000').Value2 = 1
    Add-UserRules $stressSheet 'A1:AZ1000' 64
    $stressRecordsBefore = Get-UserRuleRecords $stressSheet
    $stressFingerprintBefore = Get-UserRuleFingerprint $stressSheet
    $stressSheet.Activate()
    $stressSheet.Range('A1:C3').Select()
    $stressInstalled = [bool]$excel.Run("'" + $probeName + "'!FxInstallRules", $stressSheet, $stressSheet.Range('A1:AZ1000'), $FocusColor, $false)
    if (-not $stressInstalled) { throw 'addin_owned_name_cache rejected by Excel stress fixture' }
    $stressBook.SaveAs($StressPath, 51)
    $measurements.stress_p95_ms = Invoke-MoveSequence $stressSheet $MoveCount $WarmupCount
    $stressRuleCount = Get-ManagedRuleCount $excel $probeName $stressSheet
    $gates.stress_p95_within_150ms = ([double]$measurements.stress_p95_ms -le 150.0)
    $normalFingerprintAfter = Get-UserRuleFingerprint $normalSheet
    $stressFingerprintAfter = Get-UserRuleFingerprint $stressSheet
    $gates.user_rules_unchanged = ($normalFingerprintBefore -eq $normalFingerprintAfter -and $stressFingerprintBefore -eq $stressFingerprintAfter)

    $isolationBook = $books.Add()
    $isolationSheet = $isolationBook.Worksheets.Item(1)
    $isolationSheet.Name = 'Protected'
    $isolationSheet.Protect()
    $isolationSheet.Activate()
    $isolationSheet.Range('A1').Select()
    $gates.multi_document_isolated = (
        (Get-ManagedRuleCount $excel $probeName $normalSheet) -eq 2 -and
        (Get-ManagedRuleCount $excel $probeName $stressSheet) -eq 2 -and
        (Get-ManagedRuleCount $excel $probeName $isolationSheet) -eq 0
    )

    $excel.Run("'" + $probeName + "'!FxDeleteRules", $normalSheet)
    $excel.Run("'" + $probeName + "'!FxDeleteRules", $stressSheet)
    $excel.Run("'" + $probeName + "'!FxDeleteRules", $isolationSheet)
    $excel.Run("'" + $probeName + "'!FxFocusStop")
    $gates.cleanup_zero_assets = (
        (Get-ManagedRuleCount $excel $probeName $normalSheet) -eq 0 -and
        (Get-ManagedRuleCount $excel $probeName $stressSheet) -eq 0 -and
        (Get-ManagedRuleCount $excel $probeName $isolationSheet) -eq 0
    )
    $diagnostic.geometry = $geometry
    $diagnostic.endpoint_logic = $endpointLogic
    $diagnostic.normal_rule_count_after_moves = $normalRuleCount
    $diagnostic.stress_rule_count_after_moves = $stressRuleCount
    $diagnostic.normal_user_rules_before = $normalFingerprintBefore
    $diagnostic.normal_user_rules_after = $normalFingerprintAfter
    $diagnostic.stress_user_rules_before = $stressFingerprintBefore
    $diagnostic.stress_user_rules_after = $stressFingerprintAfter
    $diagnostic.normal_user_rule_records_before = $normalRecordsBefore
    $diagnostic.normal_user_rule_records_after = Get-UserRuleRecords $normalSheet
    $diagnostic.stress_user_rule_records_before = $stressRecordsBefore
    $diagnostic.stress_user_rule_records_after = Get-UserRuleRecords $stressSheet
    $diagnostic.saved_before_cancel = $savedBeforeCancel
} catch {
    $failure = $_.Exception.Message
} finally {
    foreach ($book in @($isolationBook,$stressBook,$normalBook,$probeBook)) {
        if ($null -ne $book) { try { $book.Close($false) } catch {} }
    }
    $cleanup = @(Invoke-ExcelCleanup $excel $binding @($isolationSheet,$isolationBook,$stressSheet,$stressBook,$normalSheet,$normalBook,$probeBook,$books) 'focus feasibility Excel')
    if ($cleanup.Count -ne 0) {
        if ([string]::IsNullOrWhiteSpace($failure)) { $failure = 'cleanup: ' + ($cleanup -join ' | ') }
        else { $failure += ' | cleanup: ' + ($cleanup -join ' | ') }
    }
}

$passed = @($gateNames | Where-Object { [bool]$gates[$_] }).Count
$status = if ($passed -eq $gateNames.Count -and [string]::IsNullOrWhiteSpace($failure)) { 'PASS' } else { 'HOLD' }
$receipt = [ordered]@{
    schema_version = 2
    status = $status
    run_id = $RunId
    source_tree_sha256 = $sourceTreeSha
    probe_xlam_sha256 = $probeSha
    fixture_sha256 = $fixtureSha
    excel_version = $excelVersion
    excel_build = $excelBuild
    office_bitness = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
    calculation_mode = $calculationMode
    display_scale_percent = Get-DisplayScalePercent
    selected_strategy = 'addin_owned_name_cache'
    measurements = $measurements
    gates = $gates
    diagnostic = [ordered]@{details=$diagnostic;failure=$failure}
}
Write-AtomicJson $EvidencePath $receipt
Write-Output ($status + '|FocusFeasibility|' + [string]$passed + '/' + [string]$gateNames.Count)
Write-Output $EvidencePath
if ($status -ne 'PASS') { exit 24 }
exit 0
