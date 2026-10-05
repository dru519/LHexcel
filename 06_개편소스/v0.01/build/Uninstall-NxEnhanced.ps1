param([string]$PackageRoot = (Split-Path -Parent $PSScriptRoot))

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$PackageRoot = [IO.Path]::GetFullPath($PackageRoot)
$payloadRoot = Join-Path $PackageRoot 'payload'
$registerScript = Join-Path $PSScriptRoot 'Register-NxHost.ps1'
$x86Dll = Join-Path $payloadRoot 'NxHost32.dll'
$x64Dll = Join-Path $payloadRoot 'NxHost64.dll'
$stateRoot = Join-Path $env:LOCALAPPDATA 'LHexcel\NxHost\state\r105'
$installStatePath = Join-Path $stateRoot 'install-state.json'

function Get-Sha256([string]$Path) {
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}
function Write-Utf8Json([string]$Path, [object]$Value) {
    $encoding = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, (ConvertTo-Json -InputObject $Value -Depth 20), $encoding)
}

if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'EXCEL_RUNNING: save documents and close every Excel process before uninstall'
}

$state = $null
if (-not (Test-Path -LiteralPath $installStatePath -PathType Leaf)) { throw 'NO_ACTIVE_INSTALL_RECEIPT' }
if (Test-Path -LiteralPath $installStatePath -PathType Leaf) {
    $state = [IO.File]::ReadAllText($installStatePath) | ConvertFrom-Json
    if ([string]$state.status -ne 'ACTIVE' -or [string]$state.release_id -ne 'r105') { throw 'NO_ACTIVE_INSTALL_RECEIPT' }
    if ([string]$state.status -eq 'ACTIVE' -and [bool]$state.xlstart_installed) {
        $xlstartTarget = [IO.Path]::GetFullPath([string]$state.xlstart_target)
        $productName = -join @([char]0xB0B4, [char]0xC5D1, [char]0xC140)
        $expectedTarget = [IO.Path]::GetFullPath((Join-Path $env:APPDATA ('Microsoft\Excel\XLSTART\' + $productName + ' v0.01_r105.xlam')))
        if (-not [string]::Equals($xlstartTarget, $expectedTarget, [StringComparison]::OrdinalIgnoreCase)) { throw 'XLSTART_RECEIPT_PATH_REJECTED' }
        if ([bool]$state.xlstart_had_previous) {
            $backup = [IO.Path]::GetFullPath([string]$state.xlstart_backup)
            $expectedBackup = [IO.Path]::GetFullPath((Join-Path $stateRoot 'xlstart-backup.xlam'))
            if (-not [string]::Equals($backup, $expectedBackup, [StringComparison]::OrdinalIgnoreCase) -or
                -not (Test-Path -LiteralPath $backup -PathType Leaf)) { throw 'XLSTART_BACKUP_REJECTED' }
        }
        if (-not (Test-Path -LiteralPath $xlstartTarget -PathType Leaf)) {
            throw 'XLSTART_CURRENT_FILE_CHANGED: installed r105 XLAM is missing'
        }
        if ((Get-Sha256 $xlstartTarget) -cne [string]$state.installed_product_sha256) {
            throw 'XLSTART_CURRENT_FILE_CHANGED: installed r105 XLAM no longer matches the receipt'
        }
    }
}

& $registerScript -Action Unregister -X86Dll $x86Dll -X64Dll $x64Dll

if ($null -ne $state -and [string]$state.status -eq 'ACTIVE') {
    if ([bool]$state.xlstart_installed) {
        $xlstartTarget = [IO.Path]::GetFullPath([string]$state.xlstart_target)
        if (Test-Path -LiteralPath $xlstartTarget -PathType Leaf) { Remove-Item -LiteralPath $xlstartTarget -Force }
        if ([bool]$state.xlstart_had_previous) {
            $backup = [IO.Path]::GetFullPath([string]$state.xlstart_backup)
            if (-not (Test-Path -LiteralPath $backup -PathType Leaf)) { throw 'XLSTART backup is missing after registry restore' }
            Copy-Item -LiteralPath $backup -Destination $xlstartTarget -Force
        }
    }
    $state.status = 'RESTORED'
    Write-Utf8Json $installStatePath $state
}

Write-Output 'PASS|UNINSTALL'
