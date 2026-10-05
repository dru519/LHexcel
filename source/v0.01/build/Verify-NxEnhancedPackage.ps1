param(
    [string]$PackageRoot = (Split-Path -Parent $PSScriptRoot),
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$PackageRoot = [IO.Path]::GetFullPath($PackageRoot)

# Hashes alone cannot detect a package that Explorer hides. Inspect the exact
# package including hidden entries, without changing any source attributes.
$publicationItems = @((Get-Item -LiteralPath $PackageRoot -Force)) + @(Get-ChildItem -LiteralPath $PackageRoot -Recurse -Force)
foreach ($item in $publicationItems) {
    if ($item.Name -eq '.DS_Store') { continue }
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "reparse publication path: $($item.FullName)" }
    if ($item.Attributes -band [IO.FileAttributes]::Hidden) { throw "hidden publication path: $($item.FullName)" }
}

function Get-RelativePath([string]$Root, [string]$Path) {
    $prefix = $Root.TrimEnd('\') + '\'
    if (-not $Path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw "path escapes package root: $Path" }
    $Path.Substring($prefix.Length).Replace('\', '/')
}
function Get-Sha256([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Read-HashManifest([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "missing hash manifest: $Path" }
    $records = @{};$previous = ''
    foreach ($line in [IO.File]::ReadAllLines($Path, [Text.Encoding]::UTF8)) {
        if ($line -notmatch '^([0-9a-f]{64})  ([^\\]+(?:/[^\\]+)*)$') { throw "malformed hash line: $line" }
        $relative = [string]$Matches[2]
        if ([IO.Path]::IsPathRooted($relative) -or $relative.Split('/') -contains '..') { throw "unsafe hash path: $relative" }
        if ($records.ContainsKey($relative)) { throw "duplicate hash path: $relative" }
        if ($previous -and [string]::CompareOrdinal($previous, $relative) -ge 0) { throw 'hash manifest is not strictly sorted' }
        $records[$relative] = [string]$Matches[1];$previous = $relative
    }
    $records
}
function Assert-Manifest([string]$Root, [string]$ManifestName) {
    $manifestPath = Join-Path $Root $ManifestName
    $records = Read-HashManifest $manifestPath
    $actual = @(Get-ChildItem -LiteralPath $Root -Recurse -File -Force | Where-Object { $_.Name -ne '.DS_Store' -and $_.FullName -ne $manifestPath } | ForEach-Object { Get-RelativePath $Root $_.FullName } | Sort-Object)
    $declared = @($records.Keys | Sort-Object)
    if (@(Compare-Object -ReferenceObject $declared -DifferenceObject $actual).Count -ne 0) { throw "$ManifestName file inventory mismatch" }
    foreach ($relative in $declared) {
        $fullPath = Join-Path $Root $relative.Replace('/', '\')
        if ((Get-Sha256 $fullPath) -cne [string]$records[$relative]) { throw "SHA-256 mismatch: $relative" }
    }
}
function Get-PeMachine([string]$Path) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $reader = New-Object IO.BinaryReader($stream)
    try {
        if ($reader.ReadUInt16() -ne 0x5A4D) { throw "invalid MZ header: $Path" }
        $stream.Position = 0x3C;$peOffset = $reader.ReadInt32()
        if ($peOffset -lt 0x40 -or $peOffset -gt ($stream.Length - 6)) { throw "invalid PE offset: $Path" }
        $stream.Position = $peOffset
        if ($reader.ReadUInt32() -ne 0x00004550) { throw "invalid PE signature: $Path" }
        $reader.ReadUInt16()
    } finally { $reader.Dispose();$stream.Dispose() }
}

$payloadRoot = Join-Path $PackageRoot 'payload'
if (-not (Test-Path -LiteralPath $payloadRoot -PathType Container)) { throw 'payload directory is missing' }
$expectedPayload = @('NxHost32.dll','NxHost64.dll','NxCore32.dll','NxCore64.dll','Product.xlam','SHA256SUMS','docs/r105-build.md') | Sort-Object
$actualPayload = @(Get-ChildItem -LiteralPath $payloadRoot -Recurse -File -Force | Where-Object { $_.Name -ne '.DS_Store' } | ForEach-Object { Get-RelativePath $payloadRoot $_.FullName } | Sort-Object)
if (@(Compare-Object -ReferenceObject $expectedPayload -DifferenceObject $actualPayload).Count -ne 0) { throw 'payload does not match the enhanced-dll contract file set' }

Assert-Manifest $payloadRoot 'SHA256SUMS'
Assert-Manifest $PackageRoot 'PACKAGE_SHA256SUMS'
$x86Machine = Get-PeMachine (Join-Path $payloadRoot 'NxHost32.dll')
$x64Machine = Get-PeMachine (Join-Path $payloadRoot 'NxHost64.dll')
if ($x86Machine -ne 0x014C) { throw ('NxHost32.dll PE machine mismatch: 0x{0:X4}' -f $x86Machine) }
if ($x64Machine -ne 0x8664) { throw ('NxHost64.dll PE machine mismatch: 0x{0:X4}' -f $x64Machine) }
if ((Get-PeMachine (Join-Path $payloadRoot 'NxCore32.dll')) -ne 0x014C) { throw 'NxCore32.dll PE machine mismatch' }
if ((Get-PeMachine (Join-Path $payloadRoot 'NxCore64.dll')) -ne 0x8664) { throw 'NxCore64.dll PE machine mismatch' }
$forbidden = @(Get-ChildItem -LiteralPath $PackageRoot -Recurse -File -Force | Where-Object { $_.Extension.ToLowerInvariant() -in @('.exe','.pyd','.zip') })
if ($forbidden.Count -ne 0) { throw "forbidden executable/archive payload: $($forbidden.FullName -join ', ')" }
if (-not $Quiet) {
    Write-Output ('PACKAGE_ROOT|' + $PackageRoot)
    Write-Output 'PASS|PACKAGE_VISIBILITY'
    Write-Output ('PE_MACHINE|NxHost32.dll|0x{0:X4}' -f $x86Machine)
    Write-Output ('PE_MACHINE|NxHost64.dll|0x{0:X4}' -f $x64Machine)
    Write-Output 'PASS|PACKAGE_VERIFY'
}
