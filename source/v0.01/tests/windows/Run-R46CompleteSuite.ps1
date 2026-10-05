param(
    [ValidateSet('Static', 'Xlam', 'Host', 'ProductUi', 'Focus', 'Hangul', 'All')]
    [string]$Phase = 'All',
    [Parameter(Mandatory = $true)][string]$DataRoot,
    [Parameter(Mandatory = $true)][string]$EvidenceRoot,
    [switch]$RegisterHost
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }

$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$DataRoot = [IO.Path]::GetFullPath($DataRoot)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not (Test-Path -LiteralPath $DataRoot -PathType Container)) { throw 'DataRoot must be an existing directory' }
if (Test-Path -LiteralPath $EvidenceRoot -PathType Leaf) { throw 'EvidenceRoot must be a directory path' }

$RunId = [Guid]::NewGuid().ToString().ToLowerInvariant()
$RunRoot = Join-Path $EvidenceRoot ('r46-complete-' + $RunId)
$TempRoot = Join-Path ([IO.Path]::GetTempPath()) ('naeexcel-r46-complete-' + $RunId)
if (Test-Path -LiteralPath $RunRoot) { throw 'Generated r46 evidence root already exists' }
if (Test-Path -LiteralPath $TempRoot) { throw 'Generated r46 temporary root already exists' }

function Write-Json([string]$Path, [object]$Value) {
    $utf8 = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 30), $utf8)
}

function Get-ExistingOfficeProcessSnapshot {
    $names = @('EXCEL', 'HWP', 'HwpFrame', 'HwpViewer', 'HancomOffice')
    $seen = @{}
    $rows = New-Object 'Collections.Generic.List[object]'
    foreach ($process in @(Get-Process -Name $names -ErrorAction SilentlyContinue)) {
        if (-not $seen.ContainsKey([int]$process.Id)) {
            $seen[[int]$process.Id] = $true
            $rows.Add([ordered]@{
                pid = [int]$process.Id
                name = [string]$process.ProcessName
                started_utc = try { $process.StartTime.ToUniversalTime().ToString('o') } catch { $null }
            })
        }
    }
    return @($rows | Sort-Object pid)
}

function Invoke-ExternalPhase([string]$PhaseName, [string]$PhaseRoot, [scriptblock]$Action) {
    [void](New-Item -ItemType Directory -Path $PhaseRoot -Force)
    $stdout = Join-Path $PhaseRoot 'stdout.log'
    $stderr = Join-Path $PhaseRoot 'stderr.log'
    $started = [DateTime]::UtcNow
    $exitCode = 1
    $errorMessage = $null
    try {
        $global:LASTEXITCODE = 0
        & $Action 1> $stdout 2> $stderr
        $exitCode = [int]$global:LASTEXITCODE
        if ($exitCode -ne 0) { throw ('external command returned exit code ' + $exitCode) }
        return [ordered]@{
            phase = $PhaseName; status = 'PASS'; started_utc = $started.ToString('o')
            completed_utc = [DateTime]::UtcNow.ToString('o'); exit_code = $exitCode
            stdout_path = $stdout; stderr_path = $stderr; error = $null
        }
    } catch {
        $errorMessage = $_.Exception.Message
        if (-not (Test-Path -LiteralPath $stdout)) { [IO.File]::WriteAllText($stdout, '', (New-Object Text.UTF8Encoding($false))) }
        if (-not (Test-Path -LiteralPath $stderr)) { [IO.File]::WriteAllText($stderr, '', (New-Object Text.UTF8Encoding($false))) }
        Add-Content -LiteralPath $stderr -Value $errorMessage -Encoding UTF8
        return [ordered]@{
            phase = $PhaseName; status = 'FAIL'; started_utc = $started.ToString('o')
            completed_utc = [DateTime]::UtcNow.ToString('o'); exit_code = $exitCode
            stdout_path = $stdout; stderr_path = $stderr; error = $errorMessage
        }
    }
}

function Invoke-NativePhase([string]$PhaseName, [string]$PhaseRoot, [scriptblock]$Action) {
    return Invoke-ExternalPhase $PhaseName $PhaseRoot $Action
}

function New-UnavailablePhaseRecord([string]$PhaseName, [string]$PhaseRoot, [string]$Reason) {
    [void](New-Item -ItemType Directory -Path $PhaseRoot -Force)
    $stdout = Join-Path $PhaseRoot 'stdout.log'
    $stderr = Join-Path $PhaseRoot 'stderr.log'
    [IO.File]::WriteAllText($stdout, '', (New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText($stderr, $Reason + [Environment]::NewLine, (New-Object Text.UTF8Encoding($false)))
    return [ordered]@{
        phase = $PhaseName; status = 'UNAVAILABLE'; started_utc = [DateTime]::UtcNow.ToString('o')
        completed_utc = [DateTime]::UtcNow.ToString('o'); exit_code = 1
        stdout_path = $stdout; stderr_path = $stderr; error = $Reason
    }
}

function Add-RefusedNativeRecords([string[]]$PhaseNames, [string]$Reason) {
    foreach ($phaseName in $PhaseNames) {
        $phaseRoot = Join-Path $RunRoot ('phase-' + $phaseName.ToLowerInvariant())
        [void]$phaseRecords.Add((New-UnavailablePhaseRecord $phaseName $phaseRoot $Reason))
    }
}

$requestedPhases = if ($Phase -eq 'All') {
    @('Static', 'Xlam', 'Host', 'ProductUi', 'Focus', 'Hangul')
} else { @($Phase) }
$nativePhases = @('Xlam', 'Host', 'ProductUi', 'Focus', 'Hangul')
$phaseRecords = New-Object 'Collections.Generic.List[object]'
$preexistingProcesses = @()
$overallFailed = $false

try {
    [void](New-Item -ItemType Directory -Path $RunRoot -Force)
    [void](New-Item -ItemType Directory -Path $TempRoot -Force)
    $preexistingProcesses = @(Get-ExistingOfficeProcessSnapshot)
    $requestedNative = @($requestedPhases | Where-Object { $nativePhases -contains $_ })

    # refuses native phases while user Office processes exist; this suite never closes user processes.
    $nativeRefused = $requestedNative.Count -gt 0 -and $preexistingProcesses.Count -gt 0
    if ($nativeRefused) {
        Add-RefusedNativeRecords $requestedNative ('pre-existing Excel/Hancom process(es) detected: ' + (($preexistingProcesses | ForEach-Object { $_.name + ':' + $_.pid }) -join ', '))
        $overallFailed = $true
    }

    if ($requestedPhases -contains 'Static') {
        $staticRoot = Join-Path $RunRoot 'phase-static'
        $python = (Get-Command python -ErrorAction Stop).Source
        $generator = Join-Path $SourceRoot 'tools/generate_command_surface.py'
        $checker = Join-Path $SourceRoot 'tools/check_contracts.py'
        if (-not (Test-Path -LiteralPath $generator -PathType Leaf)) { throw 'Command-surface generator is missing' }
        if (-not (Test-Path -LiteralPath $checker -PathType Leaf)) { throw 'Contract checker is missing' }
        $phaseRecords.Add((Invoke-ExternalPhase 'Static' $staticRoot {
            & $python $generator --check
            if ($LASTEXITCODE -ne 0) { throw ('command-surface generator failed: ' + $LASTEXITCODE) }
            & $python $checker --root $SourceRoot
            if ($LASTEXITCODE -ne 0) { throw ('contract checker failed: ' + $LASTEXITCODE) }
            Push-Location $SourceRoot
            try { & $python -m unittest discover -s tests/python -p 'test_*.py' }
            finally { Pop-Location }
            if ($LASTEXITCODE -ne 0) { throw ('Python unittest discovery failed: ' + $LASTEXITCODE) }
        }))
    }

    if (-not $nativeRefused -and ($requestedPhases -contains 'Xlam')) {
        $xlamRoot = Join-Path $RunRoot 'phase-xlam'
        $buildXlam = Join-Path $SourceRoot 'build/Build-Xlam.ps1'
        $ribbonSuite = Join-Path $SourceRoot 'tests/windows/Run-ProductRibbonSuite.ps1'
        $powershell = (Get-Command powershell -ErrorAction Stop).Source
        $artifact = Join-Path $xlamRoot 'Product.xlam'
        $ribbonTempRoot = Join-Path $TempRoot 'product-ribbon'
        $ribbonEvidenceRoot = Join-Path $xlamRoot 'product-ribbon'
        $phaseRecords.Add((Invoke-NativePhase 'Xlam' $xlamRoot {
            & $powershell -NoProfile -ExecutionPolicy Bypass -File $buildXlam -DataRoot $DataRoot -OutputPath $artifact
            if ($LASTEXITCODE -ne 0) { throw ('Product XLAM build failed: ' + $LASTEXITCODE) }
            & $powershell -NoProfile -ExecutionPolicy Bypass -File $ribbonSuite -Suite ProductRibbon -Mode Green -DataRoot $DataRoot -EvidenceRoot $ribbonTempRoot -RunId $RunId
            $ribbonExit = $LASTEXITCODE
            if (Test-Path -LiteralPath $ribbonTempRoot -PathType Container) {
                Copy-Item -LiteralPath $ribbonTempRoot -Destination $ribbonEvidenceRoot -Recurse -Force
            }
            if ($ribbonExit -ne 0) { throw ('Product Ribbon suite failed: ' + $ribbonExit) }
        }))
    }

    if (-not $nativeRefused -and ($requestedPhases -contains 'Host')) {
        $hostRoot = Join-Path $RunRoot 'phase-host'
        $hostSuite = Join-Path $SourceRoot 'tests/windows/Run-NxHostSuite.ps1'
        $powershell = (Get-Command powershell -ErrorAction Stop).Source
        if (-not (Test-Path -LiteralPath $hostSuite -PathType Leaf)) {
            $phaseRecords.Add((New-UnavailablePhaseRecord 'Host' $hostRoot 'NxHost suite is missing'))
        } else {
            $phaseRecords.Add((Invoke-NativePhase 'Host' $hostRoot {
                & $powershell -NoProfile -ExecutionPolicy Bypass -File $hostSuite -EvidenceRoot $hostRoot -Register:$RegisterHost
                if ($LASTEXITCODE -ne 0) { throw ('NxHost suite failed: ' + $LASTEXITCODE) }
            }))
        }
    }

    if (-not $nativeRefused -and ($requestedPhases -contains 'ProductUi')) {
        $uiRoot = Join-Path $RunRoot 'phase-productui'
        $uiSuite = Join-Path $SourceRoot 'tests/windows/Run-ProductUiSuite.ps1'
        $powershell = (Get-Command powershell -ErrorAction Stop).Source
        $phaseRecords.Add((Invoke-NativePhase 'ProductUi' $uiRoot {
            & $powershell -NoProfile -ExecutionPolicy Bypass -File $uiSuite -DataRoot $DataRoot -EvidenceRoot $uiRoot -RunId $RunId
            if ($LASTEXITCODE -ne 0) { throw ('Product UI suite failed: ' + $LASTEXITCODE) }
        }))
    }

    if (-not $nativeRefused -and ($requestedPhases -contains 'Focus')) {
        $focusRoot = Join-Path $RunRoot 'phase-focus'
        $focusSuite = Join-Path $SourceRoot 'tests/windows/Run-FocusOverlayProductSuite.ps1'
        $powershell = (Get-Command powershell -ErrorAction Stop).Source
        $phaseRecords.Add((Invoke-NativePhase 'Focus' $focusRoot {
            & $powershell -NoProfile -ExecutionPolicy Bypass -File $focusSuite -DataRoot $DataRoot -EvidenceRoot $focusRoot -RunId $RunId
            if ($LASTEXITCODE -ne 0) { throw ('Focus suite failed: ' + $LASTEXITCODE) }
        }))
    }

    if (-not $nativeRefused -and ($requestedPhases -contains 'Hangul')) {
        $hangulRoot = Join-Path $RunRoot 'phase-hangul'
        $hangulCandidates = @(
            (Join-Path $SourceRoot 'tests/windows/Run-HangulTableSuite.ps1'),
            (Join-Path $SourceRoot 'tests/windows/Run-HwpTableSendSuite.ps1'),
            (Join-Path $SourceRoot 'tests/windows/Run-HancomTableSuite.ps1')
        )
        $hangulSuite = @($hangulCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1)
        if ($hangulSuite.Count -eq 0) {
            $phaseRecords.Add((New-UnavailablePhaseRecord 'Hangul' $hangulRoot 'requested phase has no executable suite'))
        } else {
            $powershell = (Get-Command powershell -ErrorAction Stop).Source
            $phaseRecords.Add((Invoke-NativePhase 'Hangul' $hangulRoot {
                & $powershell -NoProfile -ExecutionPolicy Bypass -File $hangulSuite[0] -DataRoot $DataRoot -EvidenceRoot $hangulRoot -RunId $RunId
                if ($LASTEXITCODE -ne 0) { throw ('Hangul suite failed: ' + $LASTEXITCODE) }
            }))
        }
    }

    if (@($phaseRecords | Where-Object { $_.status -ne 'PASS' }).Count -gt 0) { $overallFailed = $true }
} catch {
    $overallFailed = $true
    $fatalRoot = Join-Path $RunRoot 'phase-orchestrator'
    $phaseRecords.Add((New-UnavailablePhaseRecord 'Orchestrator' $fatalRoot $_.Exception.Message))
} finally {
    $receipt = [ordered]@{
        suite = 'R46Complete'; run_id = $RunId
        status = if ($overallFailed) { 'FAIL' } else { 'PASS' }
        requested_phase = $Phase; requested_phases = $requestedPhases
        register_host = [bool]$RegisterHost
        source_root = $SourceRoot; data_root = $DataRoot
        evidence_root = $RunRoot; temporary_root = $TempRoot
        preexisting_processes = $preexistingProcesses
        phase_records = @($phaseRecords.ToArray())
        completed_utc = [DateTime]::UtcNow.ToString('o')
    }
    if (Test-Path -LiteralPath $RunRoot) { Write-Json (Join-Path $RunRoot 'Run-R46CompleteSuite.json') $receipt }
    if (Test-Path -LiteralPath $TempRoot) { Remove-Item -LiteralPath $TempRoot -Recurse -Force }
}

if ($overallFailed) {
    Write-Output 'FAIL|R46Complete'
    exit 1
}

Write-Output 'PASS|R46Complete'
exit 0
