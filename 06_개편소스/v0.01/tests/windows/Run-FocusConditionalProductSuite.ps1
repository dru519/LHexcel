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
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing FocusConditionalProduct evidence root is rejected' }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)

$required = @(
    'backend_is_cf',
    'single_cell_two_rules',
    'range_a1_c3_outer_axes',
    'shape_modes_and_selection_option',
    'five_hundred_moves_exact_two_rules',
    'move_p95_within_50ms',
    'saved_state_preserved',
    'native_undo_preserved',
    'user_rules_unchanged',
    'paste_has_no_duplicates',
    'protected_sheet_safe_exclusion',
    'disable_cleanup_zero_assets'
)
$MoveCount = 500
$WarmupCount = 20
$MoveAddresses = @('A1:C3','D4:F6','J10:L12','B20:D22','M30:O32','C45:E47','R8:T10','F60:H62')
$originalProfilePresent = Test-Path Env:LHEXCEL_PROFILE_ROOT
$originalProfile = if ($originalProfilePresent) { [string]$env:LHEXCEL_PROFILE_ROOT } else { $null }
$isolatedProfileRoot = Join-Path $EvidenceRoot 'profile/LHexcel'
[void](New-Item -ItemType Directory -Path $isolatedProfileRoot -Force)
$env:LHEXCEL_PROFILE_ROOT = $isolatedProfileRoot

. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

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
    } finally { $sha.Dispose() }
}

function Write-AtomicJson([string]$Path, [object]$Value) {
    $temporary = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    $writer = New-Object IO.StreamWriter($temporary, $false, (New-Object Text.UTF8Encoding($false)))
    try {
        $writer.Write(($Value | ConvertTo-Json -Depth 24) + "`n")
        $writer.Flush()
        $writer.BaseStream.Flush($true)
    } finally { $writer.Close() }
    if (Test-Path -LiteralPath $Path) {
        $backup = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.bak'
        try { [IO.File]::Replace($temporary, $Path, $backup) }
        finally { if (Test-Path -LiteralPath $backup) { Remove-Item -LiteralPath $backup -Force } }
    } else { [IO.File]::Move($temporary, $Path) }
}

function Get-Percentile95([double[]]$Values) {
    if ($null -eq $Values -or $Values.Count -eq 0) { throw 'No latency samples' }
    $ordered = @($Values | Sort-Object)
    $index = [Math]::Ceiling(0.95 * $ordered.Count) - 1
    return [Math]::Round([double]$ordered[[Math]::Max(0, $index)], 3)
}

$script:Cases = New-Object 'Collections.Generic.List[object]'
function Add-Case([string]$Name, [bool]$Passed, [object]$Details) {
    $script:Cases.Add([pscustomobject][ordered]@{
        ordinal = $script:Cases.Count + 1
        name = $Name
        status = if ($Passed) { 'PASS' } else { 'FAIL' }
        details = $Details
    })
}

function Get-ManagedRuleCount([object]$Excel, [string]$Macro, [object]$Sheet) {
    return [int]$Excel.Run($Macro + 'NxFocusManagedRuleCount', $Sheet)
}

function Get-RuleSnapshot([object]$Sheet) {
    $cells = $null; $conditions = $null
    $records = New-Object 'Collections.Generic.List[object]'
    try {
        $cells = $Sheet.Cells
        $conditions = $cells.FormatConditions
        for ($index = 1; $index -le [int]$conditions.Count; $index++) {
            $rule = $null; $interior = $null
            try {
                $rule = $conditions.Item($index)
                $formula = ''
                try { $formula = [string]$rule.Formula1 } catch {}
                if ($formula -notmatch 'NX_FOCUS_RULE_V002_(ROW|COLUMN)') { continue }
                $interior = $rule.Interior
                $records.Add([pscustomobject][ordered]@{
                    formula = $formula
                    color = [int]$interior.Color
                    stop_if_true = [bool]$rule.StopIfTrue
                })
            } finally { Release-ComObject $interior; Release-ComObject $rule }
        }
    } finally { Release-ComObject $conditions; Release-ComObject $cells }
    return $records.ToArray()
}

function Get-UserRuleFingerprint([object]$Sheet) {
    $scanRange = $null; $conditions = $null
    $records = New-Object 'Collections.Generic.List[string]'
    try {
        $scanRange = $Sheet.Range('A1:Z100')
        $conditions = $scanRange.FormatConditions
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
                $color = if ($null -eq $interior) { 0 } else { [int]$interior.Color }
                $records.Add(('{0}|{1}|{2}|{3}|{4}|{5}' -f [int]$rule.Type,[int]$rule.Priority,$formula,$address,[bool]$rule.StopIfTrue,$color))
            } finally { Release-ComObject $interior; Release-ComObject $applies; Release-ComObject $rule }
        }
    } finally { Release-ComObject $conditions; Release-ComObject $scanRange }
    return Get-TextSha256 (($records | Sort-Object) -join "`n")
}

function Get-ManagedNameCount([object]$Names) {
    $count = 0
    for ($index = 1; $index -le [int]$Names.Count; $index++) {
        $item = $null
        try {
            $item = $Names.Item($index)
            if ([string]$item.Name -match 'NX_FOCUS_(TOP|BOTTOM|LEFT|RIGHT)_V2$') { $count++ }
        } finally { Release-ComObject $item }
    }
    return $count
}

function Get-CacheBounds([object]$Addin) {
    $result = [ordered]@{}
    foreach ($name in @('NX_FOCUS_TOP_V2','NX_FOCUS_BOTTOM_V2','NX_FOCUS_LEFT_V2','NX_FOCUS_RIGHT_V2')) {
        $item = $null
        try {
            $item = $Addin.Names.Item($name)
            $value = ([string]$item.RefersTo).TrimStart('=','$')
            $result[$name] = [int][double]$value
        } finally { Release-ComObject $item }
    }
    return [pscustomobject]$result
}

function Get-ConditionalColor([string]$Hex, [int]$Intensity) {
    $red = [Convert]::ToInt32($Hex.Substring(1,2),16)
    $green = [Convert]::ToInt32($Hex.Substring(3,2),16)
    $blue = [Convert]::ToInt32($Hex.Substring(5,2),16)
    $red = 255 - [int][Math]::Floor((((255 - $red) * $Intensity) + 50) / 100)
    $green = 255 - [int][Math]::Floor((((255 - $green) * $Intensity) + 50) / 100)
    $blue = 255 - [int][Math]::Floor((((255 - $blue) * $Intensity) + 50) / 100)
    return $red + ($green * 256) + ($blue * 65536)
}

function Invoke-NativeUndoProbe([object]$Excel, [object]$Sheet) {
    $target = $null; $nextTarget = $null
    $result = [ordered]@{
        stage = 'initialize'
        value_after_clear = $null
        undo_after_clear = $false
        undo_after_selection_change = $false
        value_after_undo = $null
        failure = $null
        passed = $false
    }
    try {
        $result.stage = 'select-target'
        [void]$Sheet.Activate()
        $target = $Sheet.Range('Z205')
        $nextTarget = $Sheet.Range('Z206')
        $target.Value2 = 'FOCUS_CF_UNDO'
        [void]$target.Select()
        $result.stage = 'native-clear'
        [void]$Excel.CommandBars.ExecuteMso('ClearContents')
        for ($attempt = 0; $attempt -lt 20; $attempt++) {
            Start-Sleep -Milliseconds 100
            $result.value_after_clear = [string]$target.Value2
            $result.undo_after_clear = [bool]$Excel.CommandBars.GetEnabledMso('Undo')
            if ([string]::IsNullOrEmpty($result.value_after_clear) -and $result.undo_after_clear) { break }
        }
        if (-not [string]::IsNullOrEmpty($result.value_after_clear) -or -not $result.undo_after_clear) { return [pscustomobject]$result }
        $result.stage = 'selection-change'
        [void]$nextTarget.Select()
        Start-Sleep -Milliseconds 250
        $result.undo_after_selection_change = [bool]$Excel.CommandBars.GetEnabledMso('Undo')
        if (-not $result.undo_after_selection_change) { return [pscustomobject]$result }
        $result.stage = 'undo'
        [void]$Excel.Undo()
        Start-Sleep -Milliseconds 500
        $result.value_after_undo = [string]$target.Value2
        $result.passed = ($result.value_after_undo -ceq 'FOCUS_CF_UNDO')
        $result.stage = 'complete'
        return [pscustomobject]$result
    } catch {
        $result.failure = $_.Exception.Message
        return [pscustomobject]$result
    } finally {
        if ($null -ne $target) { try { [void]$target.ClearContents() } catch {} }
        Release-ComObject $nextTarget
        Release-ComObject $target
    }
}

function Get-ExcelResidueCount([int]$TimeoutMs = 5000) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $count = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
        if ($count -eq 0) { return 0 }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    return @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count
}

function Invoke-ExcelCleanup([object]$Excel, [object]$Binding, [object[]]$Objects, [string]$Label) {
    $failures = New-Object 'Collections.Generic.List[string]'
    if ($null -ne $Excel) {
        try { $Excel.DisplayAlerts = $false; $Excel.Quit() } catch { $failures.Add($_.Exception.Message) }
    }
    foreach ($item in $Objects) { try { Release-ComObject $item } catch { $failures.Add($_.Exception.Message) } }
    try { Release-ComObject $Excel } catch { $failures.Add($_.Exception.Message) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($null -eq $Binding) {
        return [pscustomobject][ordered]@{
            exit_mode = 'NOT_STARTED'
            failure = if ($failures.Count) { $failures -join ' | ' } else { $null }
            owned_pid = 0
        }
    }
    $exit = Stop-ExactProcessAfterGrace $Binding.process $Label 15000 -Detailed
    if (-not [string]::IsNullOrWhiteSpace([string]$exit.failure)) { $failures.Add([string]$exit.failure) }
    try { $Binding.process.Dispose() } catch { $failures.Add($_.Exception.Message) }
    return [pscustomobject][ordered]@{
        exit_mode = [string]$exit.exit_mode
        failure = if ($failures.Count) { $failures -join ' | ' } else { $null }
        owned_pid = [int]$Binding.pid
    }
}

$artifact = Join-Path $EvidenceRoot 'Product.xlam'
$manifestPath = Join-Path $SourceRoot 'build/manifests/Product.json'
$receiptPath = Join-Path $EvidenceRoot 'FocusConditionalProduct.json'
$artifactSha = '0' * 64
$manifestSha = Get-Sha256 $manifestPath
$sourceSha = Get-TextSha256 ((Get-Sha256 (Join-Path $SourceRoot 'src/vba/features/data/focus/NxFocusController.bas')) + "`n" + (Get-Sha256 (Join-Path $SourceRoot 'src/vba/features/data/focus/CNxFocusRuleEngine.cls')))
$excel = $null; $binding = $null; $books = $null; $addin = $null; $workbook = $null
$sheet = $null; $protectedSheet = $null; $userRange = $null; $userRule = $null
$cleanup = [pscustomobject][ordered]@{exit_mode='NOT_STARTED';failure=$null;owned_pid=0}
$diagnostic = $null
$excelVersion = $null; $excelBuild = $null; $moveP95 = $null

try {
    $beforeBuild = Get-ExcelProcessBaseline
    if (@($beforeBuild.process_ids).Count -ne 0) { throw 'Focus conditional suite requires zero pre-existing Excel processes' }
    if (Test-Path -LiteralPath $artifact) { throw 'Existing Product.xlam evidence artifact is rejected' }
    & (Join-Path $SourceRoot 'build/Build-Xlam.ps1') -DataRoot $DataRoot -OutputPath $artifact `
        -ManifestPath $manifestPath -RibbonPath (Join-Path $SourceRoot 'src/ribbon/customUI14.xml')
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $artifact -PathType Leaf)) { throw 'Current Product.xlam build failed' }
    $artifactSha = Get-Sha256 $artifact

    $baseline = Get-ExcelProcessBaseline
    if (@($baseline.process_ids).Count -ne 0) { throw 'Product build left an Excel process behind' }
    $excel = New-Object -ComObject Excel.Application
    $binding = Get-ExactExcelProcessOwnership $excel $baseline 'FocusConditionalProduct Excel'
    if (-not [bool]$binding.owned) { throw ('Excel ownership rejected: ' + [string]$binding.failure) }
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excelVersion = [string]$excel.Version
    try { $excelBuild = [string]$excel.Build } catch { $excelBuild = $excelVersion }
    $books = $excel.Workbooks
    $addin = $books.Open($artifact, $false, $true)
    $macro = "'" + ([string]$addin.Name).Replace("'", "''") + "'!"
    $workbook = $books.Add()
    $sheet = $workbook.Worksheets.Item(1)
    $sheet.Name = 'FocusPrimary'
    $userRange = $sheet.Range('A1:Z100')
    $userRule = $userRange.FormatConditions.Add(2, $null, '=MOD(ROW()+COLUMN(),3)=0')
    $userRule.StopIfTrue = $false
    $userRule.Interior.Color = 65535
    $fixturePath = Join-Path $EvidenceRoot 'FocusConditionalFixture.xlsx'
    $workbook.SaveAs($fixturePath, 51)
    $userRuleBefore = Get-UserRuleFingerprint $sheet
    $workbook.Saved = $true

    [void]$sheet.Activate()
    [void]$sheet.Range('B2').Select()
    $enableResult = [string]$excel.Run($macro + 'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#5FC8D8',40,$false,$true)
    Start-Sleep -Milliseconds 250
    $backend = [string]$excel.Run($macro + 'NxFocusControllerBackend')
    Add-Case 'backend_is_cf' ($enableResult -ceq 'PASS' -and $backend -ceq 'cf') @{enable_result=$enableResult;backend=$backend;diagnostic=[string]$excel.Run($macro+'NxFocusControllerBackendDiagnostic')}

    $singleCount = Get-ManagedRuleCount $excel $macro $sheet
    $singleBounds = Get-CacheBounds $addin
    Add-Case 'single_cell_two_rules' ($singleCount -eq 2 -and $singleBounds.NX_FOCUS_TOP_V2 -eq 2 -and $singleBounds.NX_FOCUS_BOTTOM_V2 -eq 2 -and $singleBounds.NX_FOCUS_LEFT_V2 -eq 2 -and $singleBounds.NX_FOCUS_RIGHT_V2 -eq 2) @{managed_rules=$singleCount;bounds=$singleBounds}

    [void]$sheet.Range('A1:C3').Select()
    [void]$excel.Run($macro + 'NxFocusControllerRefreshActiveSelection', $true)
    $rangeBounds = Get-CacheBounds $addin
    $rangeRules = @(Get-RuleSnapshot $sheet)
    $rowBand = @($rangeRules | Where-Object {$_.formula -match 'NX_FOCUS_RULE_V002_ROW' -and $_.formula -match 'ROW\(\)>=' -and $_.formula -match 'ROW\(\)<='})
    $columnBand = @($rangeRules | Where-Object {$_.formula -match 'NX_FOCUS_RULE_V002_COLUMN' -and $_.formula -match 'COLUMN\(\)>=' -and $_.formula -match 'COLUMN\(\)<='})
    $rangePass = ($rangeRules.Count -eq 2 -and $rangeBounds.NX_FOCUS_TOP_V2 -eq 1 -and $rangeBounds.NX_FOCUS_BOTTOM_V2 -eq 3 -and $rangeBounds.NX_FOCUS_LEFT_V2 -eq 1 -and $rangeBounds.NX_FOCUS_RIGHT_V2 -eq 3 -and $rowBand.Count -eq 1 -and $columnBand.Count -eq 1 -and @($rangeRules | Where-Object {$_.formula -match 'NOT\('}).Count -eq 2)
    Add-Case 'range_a1_c3_outer_axes' $rangePass @{bounds=$rangeBounds;rules=$rangeRules}

    $horizontal = [string]$excel.Run($macro+'NxFocusTryCommitSettingsValues','horizontal','wide-stripe','#5FC8D8',40,$false,$true)
    $horizontalRules = @(Get-RuleSnapshot $sheet)
    $vertical = [string]$excel.Run($macro+'NxFocusTryCommitSettingsValues','vertical','wide-stripe','#5FC8D8',40,$false,$true)
    $verticalRules = @(Get-RuleSnapshot $sheet)
    $selected = [string]$excel.Run($macro+'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#5FC8D8',40,$true,$true)
    $selectedRules = @(Get-RuleSnapshot $sheet)
    $expectedColor = Get-ConditionalColor '#5FC8D8' 40
    $shapePass = ($horizontal -ceq 'PASS' -and $vertical -ceq 'PASS' -and $selected -ceq 'PASS' -and
        @($horizontalRules | Where-Object {$_.formula -match 'NX_FOCUS_RULE_V002_COLUMN' -and $_.formula -match 'FALSE'}).Count -eq 1 -and
        @($verticalRules | Where-Object {$_.formula -match 'NX_FOCUS_RULE_V002_ROW' -and $_.formula -match 'FALSE'}).Count -eq 1 -and
        @($selectedRules | Where-Object {$_.formula -match 'NOT\('}).Count -eq 0 -and
        @($selectedRules | Where-Object {$_.color -eq $expectedColor}).Count -eq 2)
    Add-Case 'shape_modes_and_selection_option' $shapePass @{horizontal=$horizontalRules;vertical=$verticalRules;selected=$selectedRules;expected_color=$expectedColor}
    [void]$excel.Run($macro+'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#5FC8D8',40,$false,$true)

    $workbook.Saved = $true
    $timings = New-Object 'Collections.Generic.List[double]'
    for ($index = 0; $index -lt ($MoveCount + $WarmupCount); $index++) {
        $target = $null
        try {
            $target = $sheet.Range($MoveAddresses[$index % $MoveAddresses.Count])
            $watch = [Diagnostics.Stopwatch]::StartNew()
            [void]$target.Select()
            $watch.Stop()
            if ($index -ge $WarmupCount) { $timings.Add($watch.Elapsed.TotalMilliseconds) }
        } finally { Release-ComObject $target }
    }
    $moveP95 = Get-Percentile95 $timings.ToArray()
    $moveRules = Get-ManagedRuleCount $excel $macro $sheet
    Add-Case 'five_hundred_moves_exact_two_rules' ($timings.Count -eq $MoveCount -and $moveRules -eq 2) @{move_count=$MoveCount;warmup_count=$WarmupCount;managed_rules=$moveRules}
    Add-Case 'move_p95_within_50ms' ($moveP95 -le 50.0) @{p95_ms=$moveP95;threshold_ms=50.0}
    Add-Case 'saved_state_preserved' ([bool]$workbook.Saved) @{saved=[bool]$workbook.Saved}

    $undoResult = Invoke-NativeUndoProbe $excel $sheet
    Add-Case 'native_undo_preserved' ([bool]$undoResult.passed) $undoResult
    $userRuleAfter = Get-UserRuleFingerprint $sheet
    Add-Case 'user_rules_unchanged' ($userRuleAfter -ceq $userRuleBefore) @{before=$userRuleBefore;after=$userRuleAfter}

    $source = $null; $destination = $null
    try {
        $source = $sheet.Range('AA1:AA3'); $destination = $sheet.Range('AB1:AB3')
        $source.Value2 = 7
        [void]$source.Copy($destination)
        [void]$destination.Select()
        Start-Sleep -Milliseconds 150
    } finally { Release-ComObject $destination; Release-ComObject $source }
    $pasteRuleCount = Get-ManagedRuleCount $excel $macro $sheet
    $pasteUserFingerprint = Get-UserRuleFingerprint $sheet
    Add-Case 'paste_has_no_duplicates' ($pasteRuleCount -eq 2 -and $pasteUserFingerprint -ceq $userRuleBefore) @{managed_rules=$pasteRuleCount;user_rule_sha256=$pasteUserFingerprint}

    $protectedSheet = $workbook.Worksheets.Add()
    $protectedSheet.Name = 'FocusProtected'
    $protectedSheet.Protect()
    [void]$protectedSheet.Activate()
    [void]$protectedSheet.Range('B2').Select()
    [void]$excel.Run($macro + 'NxFocusControllerRefreshActiveSelection', $true)
    $protectedCount = Get-ManagedRuleCount $excel $macro $protectedSheet
    $protectedDiagnostic = [string]$excel.Run($macro + 'NxFocusControllerBackendDiagnostic')
    Add-Case 'protected_sheet_safe_exclusion' ($protectedCount -eq 0 -and $protectedDiagnostic.StartsWith('CF_UNSUPPORTED_SELECTION')) @{managed_rules=$protectedCount;diagnostic=$protectedDiagnostic}
    $protectedSheet.Unprotect()

    [void]$excel.Run($macro + 'NxFocusDisable')
    Start-Sleep -Milliseconds 150
    $primaryRules = Get-ManagedRuleCount $excel $macro $sheet
    $protectedRules = Get-ManagedRuleCount $excel $macro $protectedSheet
    $sheetNames = $sheet.Names; $protectedNames = $protectedSheet.Names; $addinNames = $addin.Names
    try {
        $sheetNameCount = Get-ManagedNameCount $sheetNames
        $protectedNameCount = Get-ManagedNameCount $protectedNames
        $addinNameCount = Get-ManagedNameCount $addinNames
    } finally { Release-ComObject $addinNames; Release-ComObject $protectedNames; Release-ComObject $sheetNames }
    $disabled = -not [bool]$excel.Run($macro + 'NxFocusIsEnabled')
    $offBackend = [string]$excel.Run($macro + 'NxFocusControllerBackend')
    Add-Case 'disable_cleanup_zero_assets' ($disabled -and $offBackend -ceq 'off' -and $primaryRules -eq 0 -and $protectedRules -eq 0 -and $sheetNameCount -eq 0 -and $protectedNameCount -eq 0 -and $addinNameCount -eq 0) @{enabled=(-not $disabled);backend=$offBackend;primary_rules=$primaryRules;protected_rules=$protectedRules;sheet_names=$sheetNameCount;protected_names=$protectedNameCount;addin_names=$addinNameCount}
} catch {
    $diagnostic = $_.Exception.Message + ' | ' + $_.InvocationInfo.PositionMessage
} finally {
    if ($null -ne $protectedSheet) { try { $protectedSheet.Unprotect() } catch {} }
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch {} }
    if ($null -ne $addin) { try { $addin.Close($false) } catch {} }
    $cleanup = Invoke-ExcelCleanup $excel $binding @($userRule,$userRange,$protectedSheet,$sheet,$workbook,$addin,$books) 'FocusConditionalProduct Excel'
    if ($originalProfilePresent) { $env:LHEXCEL_PROFILE_ROOT = $originalProfile } else { Remove-Item Env:LHEXCEL_PROFILE_ROOT -ErrorAction SilentlyContinue }
}

foreach ($name in $required) {
    if (@($Cases | Where-Object {$_.name -ceq $name}).Count -eq 0) {
        $Cases.Add([pscustomobject][ordered]@{ordinal=$Cases.Count+1;name=$name;status='FAIL';details=@{not_reached=$true;diagnostic=$diagnostic}})
    }
}
$ordered = @($Cases | Sort-Object {[Array]::IndexOf($required,[string]$_.name)})
for ($index = 0; $index -lt $ordered.Count; $index++) { $ordered[$index].ordinal = $index + 1 }
$names = @($ordered | ForEach-Object {[string]$_.name})
$denominatorClosed = ($ordered.Count -eq $required.Count -and @(Compare-Object -ReferenceObject $required -DifferenceObject $names).Count -eq 0)
$passed = @($ordered | Where-Object {$_.status -ceq 'PASS'}).Count
$excelResidue = Get-ExcelResidueCount
$cleanupPass = ($cleanup.exit_mode -ceq 'NATURAL' -and [string]::IsNullOrWhiteSpace([string]$cleanup.failure) -and $excelResidue -eq 0)
$allPass = ($denominatorClosed -and $passed -eq $required.Count -and [string]::IsNullOrWhiteSpace($diagnostic) -and $cleanupPass)
$receipt = [ordered]@{
    schema_version = 1
    suite = 'FocusConditionalProduct'
    status = if ($allPass) { 'PASS' } else { 'DIAGNOSTIC' }
    run_id = $RunId
    product_manifest_sha256 = $manifestSha
    source_sha256 = $sourceSha
    artifact_sha256 = $artifactSha
    environment = @{excel_version=$excelVersion;excel_build=$excelBuild;office_bitness='x64';profile_root=$isolatedProfileRoot}
    cases = $ordered
    measurements = @{move_count=$MoveCount;warmup_count=$WarmupCount;move_p95_ms=$moveP95}
    cleanup = @{exit_mode=$cleanup.exit_mode;failure=$cleanup.failure;owned_pid=$cleanup.owned_pid;excel_residue=$excelResidue}
    diagnostic = $diagnostic
}
Write-AtomicJson $receiptPath $receipt
if ($allPass) {
    Write-Output 'PASS|FocusConditionalProduct|12/12'
    Write-Output $receiptPath
    exit 0
}
Write-Output ('DIAGNOSTIC|FocusConditionalProduct|' + $passed + '/12')
Write-Output $receiptPath
exit 24
