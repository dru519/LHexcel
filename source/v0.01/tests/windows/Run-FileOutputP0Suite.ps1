param(
    [Parameter(Mandatory=$true)][string]$ArtifactPath,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$SourceDigest,
    [Parameter(Mandatory=$true)][string]$SnapshotDigest,
    [string]$RunId = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }

$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$ArtifactPath = [IO.Path]::GetFullPath($ArtifactPath)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
if ($RunId -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$') { throw 'RunId must be a lowercase GUID' }
if ($SourceDigest -notmatch '^[0-9a-f]{64}$' -or $SnapshotDigest -notmatch '^[0-9a-f]{64}$') { throw 'Source and snapshot digests must be lowercase SHA-256 values' }
if (-not (Test-Path -LiteralPath $ArtifactPath -PathType Leaf) -or [IO.Path]::GetExtension($ArtifactPath) -ine '.xlam') { throw 'FileOutputP0 artifact must be an XLAM file' }
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing FileOutputP0 evidence root is rejected' }

$routeIds = @(
    'NX-FILE-MANNER-SAVE',
    'NX-FILE-SHEET-COPY-SAVE',
    'NX-FILE-RANGE-COPY-SAVE',
    'NX-FILE-RANGE-PNG',
    'NX-FILE-CHART-PNG',
    'NX-FILE-PDF-CURRENT-SHEET',
    'NX-FILE-PDF-EACH-SHEET',
    'NX-FILE-PDF-SELECTED-COMBINED',
    'NX-FILE-PDF-ALL-COMBINED',
    'NX-FILE-PDF-SETTINGS'
)

function Get-Sha256([string]$Path) {
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
}

function Write-AtomicJson([string]$Path, [object]$Value) {
    $directory = Split-Path -Parent $Path
    [void](New-Item -ItemType Directory -Path $directory -Force)
    $temporary = Join-Path $directory (([IO.Path]::GetFileName($Path)) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')
    $utf8 = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($temporary, ($Value | ConvertTo-Json -Depth 50), $utf8)
    [IO.File]::Move($temporary, $Path)
}

function ConvertTo-PowerShellLiteral([string]$Value) {
    return "'" + $Value.Replace("'", "''") + "'"
}

$suite = Join-Path $PSScriptRoot 'Run-ProductExhaustiveSuite.ps1'
$windowsPowerShell = Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
if (-not (Test-Path -LiteralPath $suite -PathType Leaf) -or -not (Test-Path -LiteralPath $windowsPowerShell -PathType Leaf)) { throw 'FileOutputP0 suite dependency is missing' }

$routeLiteral = ($routeIds | ForEach-Object { ConvertTo-PowerShellLiteral $_ }) -join ','
$childCommand = (
    '& ' + (ConvertTo-PowerShellLiteral $suite) +
    ' -SourceRoot ' + (ConvertTo-PowerShellLiteral $SourceRoot) +
    ' -ArtifactPath ' + (ConvertTo-PowerShellLiteral $ArtifactPath) +
    ' -EvidenceRoot ' + (ConvertTo-PowerShellLiteral $EvidenceRoot) +
    ' -RunId ' + (ConvertTo-PowerShellLiteral $RunId) +
    ' -SourceDigest ' + (ConvertTo-PowerShellLiteral $SourceDigest) +
    ' -SnapshotDigest ' + (ConvertTo-PowerShellLiteral $SnapshotDigest) +
    ' -RouteIds @(' + $routeLiteral + ')'
)
$encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($childCommand))
& $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encodedCommand
$suiteExitCode = $LASTEXITCODE
$targetedPath = Join-Path $EvidenceRoot 'ProductExhaustive.Targeted.json'
if (-not (Test-Path -LiteralPath $targetedPath -PathType Leaf)) { throw 'ProductExhaustive.Targeted.json was not produced' }
$receipt = Get-Content -LiteralPath $targetedPath -Raw -Encoding UTF8 | ConvertFrom-Json
$summaryMatches = (
    [string]$receipt.scope -ceq 'TARGETED' -and
    [int]$receipt.expected_routes -eq 10 -and
    [int]$receipt.executed_routes -eq 10 -and
    [int]$receipt.passed_routes -eq 10 -and
    [int]$receipt.held_routes -eq 0 -and
    [int]$receipt.failed_routes -eq 0 -and
    [string]$receipt.cleanup.status -ceq 'PASS'
)
$caseEvidenceMatches = @($receipt.cases | Where-Object {
    $expectedEvidenceCount = if ([string]$_.route_id -ceq 'NX-FILE-PDF-EACH-SHEET') { 3 } else { 1 }
    [string]$_.status -ceq 'PASS' -and @($_.evidence_outputs).Count -eq $expectedEvidenceCount
}).Count -eq 10
$status = if ($suiteExitCode -eq 0 -and $summaryMatches -and $caseEvidenceMatches) { 'PASS' } else { 'FAIL' }

$fileOutputReceipt = [ordered]@{
    schema_version=1
    suite='FileOutputP0'
    status=$status
    run_id=$RunId
    source_tree_sha256=$SourceDigest
    source_snapshot_sha256=$SnapshotDigest
    artifact_sha256=Get-Sha256 $ArtifactPath
    targeted_receipt_sha256=Get-Sha256 $targetedPath
    expected_routes=10
    executed_routes=[int]$receipt.executed_routes
    passed_routes=[int]$receipt.passed_routes
    held_routes=[int]$receipt.held_routes
    failed_routes=[int]$receipt.failed_routes
    cleanup_status=[string]$receipt.cleanup.status
    route_ids=$routeIds
}
Write-AtomicJson (Join-Path $EvidenceRoot 'FileOutputP0.json') $fileOutputReceipt

if ($status -cne 'PASS') { throw ('FileOutputP0 failed; ProductExhaustive exit=' + $suiteExitCode) }
Write-Host 'PASS|FileOutputP0|10/10'
