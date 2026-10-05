param(
    [string]$PackageRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$ExcelPath = '',
    [switch]$OptIn,
    [switch]$ProbeOnly,
    [switch]$InstallXlamToXLStart,
    [switch]$ReplaceExistingXlam,
    [switch]$ReplaceExistingRegistration
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Resolve-ExcelExecutable([string]$ExplicitPath) {
    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        $candidate = [IO.Path]::GetFullPath($ExplicitPath)
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { throw "Excel executable not found: $candidate" }
        return $candidate
    }
    $candidates = New-Object 'Collections.Generic.List[string]'
    foreach ($hiveName in @('CurrentUser','LocalMachine')) {
        foreach ($viewName in @('Registry64','Registry32')) {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::$hiveName,
                [Microsoft.Win32.RegistryView]::$viewName
            )
            $key = $null
            try {
                $key = $base.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', $false)
                if ($null -ne $key) {
                    $value = [string]$key.GetValue('')
                    if (-not [string]::IsNullOrWhiteSpace($value)) { $candidates.Add($value.Trim('"')) }
                }
            } finally {
                if ($null -ne $key) { $key.Dispose() }
                $base.Dispose()
            }
        }
    }
    foreach ($programRoot in @(
        [Environment]::GetFolderPath('ProgramFiles'),
        [Environment]::GetFolderPath('ProgramFilesX86')
    )) {
        if ([string]::IsNullOrWhiteSpace($programRoot)) { continue }
        foreach ($relative in @('Microsoft Office\root\Office16\EXCEL.EXE','Microsoft Office\Office16\EXCEL.EXE')) {
            $candidates.Add((Join-Path $programRoot $relative))
        }
    }
    foreach ($candidate in @($candidates | Select-Object -Unique)) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return [IO.Path]::GetFullPath($candidate) }
    }
    throw 'Excel executable could not be resolved without launching Excel'
}

function Get-PeMachine([string]$Path) {
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $reader = New-Object IO.BinaryReader($stream)
    try {
        if ($reader.ReadUInt16() -ne 0x5A4D) { throw "invalid MZ header: $Path" }
        $stream.Position = 0x3C
        $peOffset = $reader.ReadInt32()
        if ($peOffset -lt 0x40 -or $peOffset -gt ($stream.Length - 6)) { throw "invalid PE offset: $Path" }
        $stream.Position = $peOffset
        if ($reader.ReadUInt32() -ne 0x00004550) { throw "invalid PE signature: $Path" }
        return $reader.ReadUInt16()
    } finally {
        $reader.Dispose()
        $stream.Dispose()
    }
}

function Get-Sha256([string]$Path) {
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Write-Utf8Json([string]$Path, [object]$Value) {
    $encoding = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject $Value -Depth 20), $encoding)
}

$resolvedExcel = Resolve-ExcelExecutable $ExcelPath
$excelMachine = Get-PeMachine $resolvedExcel
$excelArchitecture = switch ($excelMachine) {
    0x014C { 'x86'; break }
    0x8664 { 'x64'; break }
    0xAA64 { throw 'NATIVE_ARM64_EXCEL_UNSUPPORTED: PE machine 0xAA64' }
    0xA641 { throw 'NATIVE_ARM64EC_EXCEL_UNSUPPORTED: PE machine 0xA641' }
    default { throw ('UNSUPPORTED_EXCEL_ARCHITECTURE: PE machine 0x{0:X4}' -f $excelMachine) }
}
Write-Output ('EXCEL_PATH|' + $resolvedExcel)
Write-Output ('EXCEL_ARCHITECTURE|' + $excelArchitecture + '|0x{0:X4}' -f $excelMachine)
# -ProbeOnly performs architecture discovery without package, registry, or XLSTART changes.
if ($ProbeOnly) { Write-Output 'PASS|PROBE'; return }
if (-not $OptIn) { throw 'Installation is opt-in: pass -OptIn explicitly' }

$PackageRoot = [IO.Path]::GetFullPath($PackageRoot)
$payloadRoot = Join-Path $PackageRoot 'payload'
$verifyScript = Join-Path $PSScriptRoot 'Verify-NxEnhancedPackage.ps1'
$registerScript = Join-Path $PSScriptRoot 'Register-NxHost.ps1'
$x86Dll = Join-Path $payloadRoot 'NxHost32.dll'
$x64Dll = Join-Path $payloadRoot 'NxHost64.dll'
$product = Join-Path $payloadRoot 'Product.xlam'
& $verifyScript -PackageRoot $PackageRoot -Quiet
if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'EXCEL_RUNNING: save documents and close every Excel process before installation'
}

$stateRoot = Join-Path $env:LOCALAPPDATA 'LHexcel\NxHost\state\r105'
[void](New-Item -ItemType Directory -Path $stateRoot -Force)
$installStatePath = Join-Path $stateRoot 'install-state.json'
if (Test-Path -LiteralPath $installStatePath -PathType Leaf) {
    $existingState = [IO.File]::ReadAllText($installStatePath) | ConvertFrom-Json
    if ([string]$existingState.status -eq 'ACTIVE') { throw 'ACTIVE_INSTALL_STATE_CONFLICT: uninstall the current r105 installation first' }
}

$xlstartRoot = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
$productName = -join @([char]0xB0B4, [char]0xC5D1, [char]0xC140)
$xlstartTarget = Join-Path $xlstartRoot ($productName + ' v0.01_r105.xlam')
$xlstartBackup = Join-Path $stateRoot 'xlstart-backup.xlam'
$hadPrevious = $false
if ($InstallXlamToXLStart) {
    [void](New-Item -ItemType Directory -Path $xlstartRoot -Force)
    $otherProducts = @(
        Get-ChildItem -LiteralPath $xlstartRoot -Filter '*.xlam' -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like ($productName + '*.xlam') -and $_.FullName -ine $xlstartTarget }
    )
    if ($otherProducts.Count -ne 0) {
        throw ('XLSTART_CONFLICT: another ' + $productName + ' add-in exists: ' + ($otherProducts.Name -join ', '))
    }
    if (Test-Path -LiteralPath $xlstartTarget -PathType Leaf) {
        if (-not $ReplaceExistingXlam) { throw 'XLSTART_CONFLICT: r105 target already exists; pass -ReplaceExistingXlam to back it up' }
        Copy-Item -LiteralPath $xlstartTarget -Destination $xlstartBackup -Force
        $hadPrevious = $true
    }
}

$registered = $false
try {
    $registrationArgs = @{
        Action = 'Register'; X86Dll = $x86Dll; X64Dll = $x64Dll; OptIn = $true
    }
    if ($ReplaceExistingRegistration) { $registrationArgs.ReplaceExisting = $true }
    & $registerScript @registrationArgs
    $registered = $true

    if ($InstallXlamToXLStart) {
        $temporaryTarget = Join-Path $xlstartRoot ('.nx-r105-' + [Guid]::NewGuid().ToString('N') + '.tmp')
        try {
            Copy-Item -LiteralPath $product -Destination $temporaryTarget
            if ((Get-Sha256 $temporaryTarget) -cne (Get-Sha256 $product)) { throw 'XLSTART staged Product.xlam hash mismatch' }
            Move-Item -LiteralPath $temporaryTarget -Destination $xlstartTarget -Force
        } finally {
            if (Test-Path -LiteralPath $temporaryTarget -PathType Leaf) { Remove-Item -LiteralPath $temporaryTarget -Force }
        }
    }

    $receipt = [pscustomobject][ordered]@{
        schema_version = 1
        release_id = 'r105'
        status = 'ACTIVE'
        excel_path = $resolvedExcel
        excel_architecture = $excelArchitecture
        excel_pe_machine = ('0x{0:X4}' -f $excelMachine)
        installed_product_sha256 = Get-Sha256 $product
        xlstart_installed = [bool]$InstallXlamToXLStart
        xlstart_target = if ($InstallXlamToXLStart) { $xlstartTarget } else { '' }
        xlstart_had_previous = $hadPrevious
        xlstart_backup = if ($hadPrevious) { $xlstartBackup } else { '' }
    }
    Write-Utf8Json $installStatePath $receipt
} catch {
    $failure = $_.Exception.Message
    if ($InstallXlamToXLStart -and (Test-Path -LiteralPath $xlstartTarget -PathType Leaf)) {
        Remove-Item -LiteralPath $xlstartTarget -Force
    }
    if ($hadPrevious -and (Test-Path -LiteralPath $xlstartBackup -PathType Leaf)) {
        Copy-Item -LiteralPath $xlstartBackup -Destination $xlstartTarget -Force
    }
    if ($registered) {
        try { & $registerScript -Action Unregister -X86Dll $x86Dll -X64Dll $x64Dll | Out-Null } catch {}
    }
    throw ('INSTALL_ROLLED_BACK: ' + $failure)
}

Write-Output ('PRODUCT_XLAM|' + $product)
Write-Output ('XLSTART_MODE|' + $(if ($InstallXlamToXLStart) { 'INSTALLED' } else { 'UNCHANGED' }))
Write-Output 'PASS|INSTALL'
