param(
    [Parameter(Mandatory = $true)][string]$EvidenceRoot,
    [string]$NxHostRoot = '',
    [string]$ArtifactPath = '',
    [string]$RunId = '',
    [string]$SourceDigest = '',
    [string]$SnapshotDigest = '',
    [ValidateSet('BuildOnly','ReleaseBound')][string]$EvidenceMode = 'BuildOnly',
    [switch]$Register
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitOperatingSystem) { throw 'NxHost validation requires 64-bit Windows' }

$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
$BuildScript = Join-Path $SourceRoot 'build/Build-NxHost.ps1'
$RegisterScript = Join-Path $SourceRoot 'build/Register-NxHost.ps1'
$RunRoot = Join-Path $EvidenceRoot 'nxhost'
if ($EvidenceMode -eq 'ReleaseBound' -and (
    [string]::IsNullOrWhiteSpace($ArtifactPath) -or
    [string]::IsNullOrWhiteSpace($RunId) -or
    [string]::IsNullOrWhiteSpace($SourceDigest) -or
    [string]::IsNullOrWhiteSpace($SnapshotDigest))) {
    throw 'ReleaseBound NxHost evidence requires ArtifactPath, RunId, SourceDigest, and SnapshotDigest'
}
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
if ($RunId -notmatch '^[0-9a-f-]{36}$') { throw 'RunId must be a lowercase GUID' }
if (-not [string]::IsNullOrWhiteSpace($SourceDigest) -and $SourceDigest -notmatch '^[0-9a-f]{64}$') { throw 'SourceDigest must be a lowercase SHA-256' }
if (-not [string]::IsNullOrWhiteSpace($SnapshotDigest) -and $SnapshotDigest -notmatch '^[0-9a-f]{64}$') { throw 'SnapshotDigest must be a lowercase SHA-256' }

function Test-PeMachine([string]$Path, [UInt16]$ExpectedMachine) {
    $stream = [IO.File]::OpenRead($Path)
    try {
        $reader = New-Object IO.BinaryReader($stream)
        $stream.Position = 0x3c
        $offset = $reader.ReadInt32()
        $stream.Position = $offset + 4
        return $reader.ReadUInt16() -eq $ExpectedMachine
    } finally {
        $stream.Dispose()
    }
}

function Get-Sha256([string]$Path) {
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Write-Utf8Json([string]$Path, [object]$Value) {
    $encoding = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 12), $encoding)
}

if (Test-Path -LiteralPath $RunRoot) { throw 'Existing NxHost evidence root is rejected' }
if (Test-Path -LiteralPath $EvidenceRoot -PathType Leaf) { throw 'EvidenceRoot must be a directory path' }
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { throw 'NxHost suite requires zero pre-existing Excel processes' }
if (-not (Test-Path -LiteralPath $BuildScript -PathType Leaf)) { throw 'NxHost build script is missing' }

[void](New-Item -ItemType Directory -Path $RunRoot -Force)
if ([string]::IsNullOrWhiteSpace($NxHostRoot)) {
    & $BuildScript -OutputRoot $RunRoot -Configuration Release
    if ($LASTEXITCODE -ne 0) { throw 'NxHost build failed' }
} else {
    $NxHostRoot = [IO.Path]::GetFullPath($NxHostRoot)
    $suppliedX86 = Join-Path $NxHostRoot 'NxHost32.dll'
    $suppliedX64 = Join-Path $NxHostRoot 'NxHost64.dll'
    if (-not (Test-Path -LiteralPath $suppliedX86 -PathType Leaf) -or -not (Test-Path -LiteralPath $suppliedX64 -PathType Leaf)) { throw 'NxHost supplied artifacts rejected' }
    [void](New-Item -ItemType Directory -Path (Join-Path $RunRoot 'x86') -Force)
    [void](New-Item -ItemType Directory -Path (Join-Path $RunRoot 'x64') -Force)
    Copy-Item -LiteralPath $suppliedX86 -Destination (Join-Path $RunRoot 'x86/NxHost32.dll')
    Copy-Item -LiteralPath $suppliedX64 -Destination (Join-Path $RunRoot 'x64/NxHost64.dll')
    if ((Get-Sha256 (Join-Path $RunRoot 'x86/NxHost32.dll')) -cne (Get-Sha256 $suppliedX86) -or (Get-Sha256 (Join-Path $RunRoot 'x64/NxHost64.dll')) -cne (Get-Sha256 $suppliedX64)) { throw 'NxHost supplied artifact copy mismatch' }
}

$x86Dll = Join-Path $RunRoot 'x86/NxHost32.dll'
$x64Dll = Join-Path $RunRoot 'x64/NxHost64.dll'
if (-not (Test-Path -LiteralPath $x86Dll -PathType Leaf)) { throw 'NxHost x86 DLL is missing' }
if (-not (Test-Path -LiteralPath $x64Dll -PathType Leaf)) { throw 'NxHost x64 DLL is missing' }
if (-not (Test-PeMachine $x86Dll 0x14c)) { throw 'NxHost32.dll is not x86' }
if (-not (Test-PeMachine $x64Dll 0x8664)) { throw 'NxHost64.dll is not x64' }
$artifactSha = $null
if (-not [string]::IsNullOrWhiteSpace($ArtifactPath)) {
    $ArtifactPath = [IO.Path]::GetFullPath($ArtifactPath)
    if (-not (Test-Path -LiteralPath $ArtifactPath -PathType Leaf) -or [IO.Path]::GetExtension($ArtifactPath) -ine '.xlam') { throw 'NxHost supplied Product artifact rejected' }
    $artifactSha = Get-Sha256 $ArtifactPath
}

$registrationMode = 'not-requested'
if ($Register) {
    if (-not (Test-Path -LiteralPath $RegisterScript -PathType Leaf)) { throw 'NxHost registration script is missing' }
    $registrationMode = 'requested'
    try {
        & $RegisterScript -Action Register -OptIn -X86Dll $x86Dll -X64Dll $x64Dll
        if ($LASTEXITCODE -ne 0) { throw 'NxHost registration failed' }
    } finally {
        & $RegisterScript -Action Unregister -X86Dll $x86Dll -X64Dll $x64Dll
        if ($LASTEXITCODE -ne 0) { throw 'NxHost unregistration failed' }
    }
}

$receipt = [ordered]@{
    schema_version = 2
    suite = 'NxHost'
    status = if ($EvidenceMode -eq 'ReleaseBound') { 'PASS' } else { 'BUILD_ONLY_PASS' }
    evidence_mode = $EvidenceMode
    run_id = $RunId
    artifact_sha256 = $artifactSha
    source_tree_sha256 = $SourceDigest
    source_snapshot_sha256 = $SnapshotDigest
    nxhost32_sha256 = Get-Sha256 $x86Dll
    nxhost64_sha256 = Get-Sha256 $x64Dll
    registration_mode = $registrationMode
    artifacts = [ordered]@{
        x86 = Get-Sha256 $x86Dll
        x64 = Get-Sha256 $x64Dll
    }
}
Write-Utf8Json (Join-Path $RunRoot 'NxHostSuite.json') $receipt

$statusLabel = if ($EvidenceMode -eq 'ReleaseBound') { 'PASS' } else { 'BUILD_ONLY_PASS' }
Write-Output ($statusLabel + '|NxHost|2/2')
