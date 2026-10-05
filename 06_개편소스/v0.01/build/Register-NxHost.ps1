param(
    [ValidateSet('Register','Unregister')][string]$Action = 'Register',
    [Parameter(Mandatory=$true)][string]$X86Dll,
    [Parameter(Mandatory=$true)][string]$X64Dll,
    [switch]$OptIn,
    [switch]$ReplaceExisting
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$reg32 = Join-Path $env:windir 'Microsoft.NET/Framework/v4.0.30319/RegAsm.exe'
$reg64 = Join-Path $env:windir 'Microsoft.NET/Framework64/v4.0.30319/RegAsm.exe'
if (!(Test-Path $reg32) -or !(Test-Path $reg64)) { throw 'RegAsm x86/x64 not found' }
if ($Action -eq 'Register' -and -not $OptIn) { throw 'Registration is opt-in: pass -OptIn explicitly' }
if (!(Test-Path -LiteralPath $X86Dll)) { throw "missing DLL: $X86Dll" }
if (!(Test-Path -LiteralPath $X64Dll)) { throw "missing DLL: $X64Dll" }

$X86Dll = [IO.Path]::GetFullPath($X86Dll)
$X64Dll = [IO.Path]::GetFullPath($X64Dll)
$runtimeRoot = Join-Path $env:LOCALAPPDATA 'LHexcel\NxHost\r105'
$runtimeX86 = Join-Path $runtimeRoot 'x86\NxHost32.dll'
$runtimeX64 = Join-Path $runtimeRoot 'x64\NxHost64.dll'
$coreX86 = Join-Path (Split-Path -Parent $X86Dll) 'NxCore32.dll'
$coreX64 = Join-Path (Split-Path -Parent $X64Dll) 'NxCore64.dll'
$runtimeCoreX86 = Join-Path $runtimeRoot 'x86\NxCore32.dll'
$runtimeCoreX64 = Join-Path $runtimeRoot 'x64\NxCore64.dll'
$backup = Join-Path $env:LOCALAPPDATA 'LHexcel\NxHost\state\r105'
[void](New-Item -ItemType Directory -Path $backup -Force)
$statePath = Join-Path $backup 'registry-state.json'
$addinSubKey = 'Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect'
$classIds = @(
    '{4A4B4F02-76A6-4F22-BD4B-85E0308EC91D}',
    '{9A90294C-81C8-4AEE-A2D8-14B078A61CD2}',
    '{A13D5ED6-B0F2-43C0-9499-E53E6CE38DC0}',
    '{1434C649-18AA-4435-B0D2-2BD81B548A01}',
    '{C812F023-E912-49AA-98B7-75DC86501902}',
    '{C812F023-E912-49AA-98B7-75DC86501904}',
    '{C812F023-E912-49AA-98B7-75DC86501906}'
)
$progIds = @('LH.NxHost.Connect','LH.NxHost.BridgeService','LH.NxHost.NavigatorPane','LH.NxHost.HwpxExportService','LH.NxHost.PicturePreviewService','LH.NxHost.WorkbookCompareService','LH.NxHost.WorkbookCompareResultsService')
$affectedSubKeys = @(
    @($classIds | ForEach-Object { "Software\Classes\CLSID\$_" })
    @($progIds | ForEach-Object { "Software\Classes\$_" })
    $addinSubKey
)

function Test-NxHostRegistrationPresent {
    foreach ($viewName in @('Registry32','Registry64')) {
        $view = [Microsoft.Win32.RegistryView]::$viewName
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            foreach ($subKey in $affectedSubKeys) {
                $key = $null
                try {
                    $key = $base.OpenSubKey($subKey, $false)
                    if ($null -ne $key) { return $true }
                } finally {
                    if ($null -ne $key) { $key.Dispose() }
                }
            }
        } finally {
            $base.Dispose()
        }
    }
    return $false
}

function Convert-RegistryValue([Microsoft.Win32.RegistryKey]$key, [string]$name) {
    $kind = $key.GetValueKind($name)
    $value = $key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    if ($value -is [byte[]]) {
        $value = [Convert]::ToBase64String($value)
    } elseif ($value -is [string[]]) {
        $value = @($value)
    }
    [pscustomobject][ordered]@{ name = $name; kind = $kind.ToString(); value = $value }
}

function Get-RegistryTree([Microsoft.Win32.RegistryKey]$key) {
    $values = New-Object 'Collections.Generic.List[object]'
    foreach ($name in @($key.GetValueNames() | Sort-Object)) { $values.Add((Convert-RegistryValue $key $name)) }
    $children = New-Object 'Collections.Generic.List[object]'
    foreach ($name in @($key.GetSubKeyNames() | Sort-Object)) {
        $child = $null
        try {
            $child = $key.OpenSubKey($name, $false)
            if ($null -ne $child) { $children.Add([pscustomobject][ordered]@{ name = $name; tree = Get-RegistryTree $child }) }
        } finally {
            if ($null -ne $child) { $child.Dispose() }
        }
    }
    [pscustomobject][ordered]@{ values = $values.ToArray(); subkeys = $children.ToArray() }
}

function Get-RegistrySnapshot {
    $rows = New-Object 'Collections.Generic.List[object]'
    foreach ($viewName in @('Registry32','Registry64')) {
        $view = [Microsoft.Win32.RegistryView]::$viewName
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        try {
            foreach ($subKey in $affectedSubKeys) {
                $key = $null
                try {
                    $key = $base.OpenSubKey($subKey, $false)
                    $rows.Add([pscustomobject][ordered]@{
                        view = $viewName
                        sub_key = $subKey
                        present = ($null -ne $key)
                        tree = if ($null -ne $key) { Get-RegistryTree $key } else { $null }
                    })
                } finally {
                    if ($null -ne $key) { $key.Dispose() }
                }
            }
        } finally {
            $base.Dispose()
        }
    }
    $rows.ToArray()
}

function Convert-SnapshotJson([object]$snapshot) {
    ConvertTo-Json -InputObject @($snapshot) -Depth 100 -Compress
}

function Get-TextSha256([string]$value) {
    $algorithm = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($value)
        ([BitConverter]::ToString($algorithm.ComputeHash($bytes))).Replace('-', '').ToLowerInvariant()
    } finally {
        $algorithm.Dispose()
    }
}

function Write-Utf8Json([string]$path, [object]$value) {
    $encoding = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($path, (ConvertTo-Json -InputObject $value -Depth 100), $encoding)
}

function Invoke-HkcuRegAsm([string]$dll, [string]$regasm, [string]$tag, [int]$viewBits) {
    $regfile = Join-Path $backup ("$tag.reg")
    $previousPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $regasmOutput = @(& $regasm $dll /codebase /regfile:$regfile 2>&1)
    $regasmExitCode = $LASTEXITCODE
    $ErrorActionPreference = $previousPreference
    if ($regasmExitCode -ne 0) { throw ("RegAsm $tag export failed: " + ($regasmOutput -join ' ')) }
    if (@($regasmOutput | Where-Object { [string]$_ -match 'warning RA0000' }).Count -gt 0) {
        Write-Output ("WARN|REGASM_UNSIGNED_CODEBASE|$tag")
    }
    $raw = [IO.File]::ReadAllBytes($regfile)
    if ($raw.Length -ge 2 -and $raw[0] -eq 0xFF -and $raw[1] -eq 0xFE) {
        $text = [Text.Encoding]::Unicode.GetString($raw, 2, $raw.Length - 2)
    } elseif ($raw.Length -ge 2 -and $raw[1] -eq 0) {
        $text = [Text.Encoding]::Unicode.GetString($raw)
    } else {
        $text = [Text.Encoding]::Default.GetString($raw)
    }
    $classesRoot = ('HKEY_' + 'CLASSES_ROOT\')
    $text = $text -replace [regex]::Escape($classesRoot), 'HKEY_CURRENT_USER\Software\Classes\'
    [IO.File]::WriteAllText($regfile, $text, [Text.Encoding]::Unicode)
    $previousPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $importOutput = @(& reg.exe import $regfile ("/reg:$viewBits") 2>&1)
    $importExitCode = $LASTEXITCODE
    $ErrorActionPreference = $previousPreference
    if ($importExitCode -ne 0) { throw ("HKCU import failed: $tag; " + ($importOutput -join ' ')) }
}

function Set-OfficeAddinView([Microsoft.Win32.RegistryView]$view) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    $key = $null
    try {
        $key = $base.CreateSubKey($addinSubKey, $true)
        $key.SetValue('FriendlyName', '내엑셀 NxHost', [Microsoft.Win32.RegistryValueKind]::String)
        $key.SetValue('Description', '내엑셀 Enhanced DLL host', [Microsoft.Win32.RegistryValueKind]::String)
        $key.SetValue('LoadBehavior', 3, [Microsoft.Win32.RegistryValueKind]::DWord)
        # Excel started through COM automation loads only command-line-safe add-ins.
        $key.SetValue('CommandLineSafe', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
    } finally {
        if ($null -ne $key) { $key.Dispose() }
        $base.Dispose()
    }
}

function Remove-RegistryView([Microsoft.Win32.RegistryView]$view) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
    try {
        foreach ($subKey in $affectedSubKeys) { try { $base.DeleteSubKeyTree($subKey, $false) } catch {} }
    } finally { $base.Dispose() }
}

function Remove-RuntimeFile([string]$path) {
    for ($attempt = 0; $attempt -lt 50; $attempt++) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return }
        try { Remove-Item -LiteralPath $path -Force -ErrorAction Stop; return } catch { Start-Sleep -Milliseconds 100 }
    }
    if (Test-Path -LiteralPath $path -PathType Leaf) { throw "NxHost runtime file remained locked: $path" }
}

function Restore-NxHostRegistrationState {
    if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) { throw 'NxHost registry recovery state is missing' }
    $state = [IO.File]::ReadAllText($statePath) | ConvertFrom-Json
    $expectedJson = Convert-SnapshotJson @($state.snapshot)
    $expectedHash = Get-TextSha256 $expectedJson
    if ($expectedHash -cne [string]$state.snapshot_sha256) { throw 'NxHost registry recovery state hash mismatch' }

    Remove-RegistryView ([Microsoft.Win32.RegistryView]::Registry32)
    Remove-RegistryView ([Microsoft.Win32.RegistryView]::Registry64)
    foreach ($saved in @($state.exports)) {
        $savedPath = Join-Path $backup ([string]$saved.file)
        if (-not (Test-Path -LiteralPath $savedPath -PathType Leaf)) { throw "registry backup file missing: $savedPath" }
        $viewBits = if ([string]$saved.view -eq 'Registry32') { 32 } else { 64 }
        & reg.exe import $savedPath ("/reg:$viewBits") | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "registry restore failed: $($saved.view) $($saved.sub_key)" }
    }

    $restored = @(Get-RegistrySnapshot)
    $restoredJson = Convert-SnapshotJson $restored
    $restoredHash = Get-TextSha256 $restoredJson
    if ($restoredJson -cne $expectedJson -or $restoredHash -cne $expectedHash) { throw 'registry snapshot mismatch after restore' }
    $state.status = 'RESTORED'
    $state | Add-Member -NotePropertyName restored_snapshot_sha256 -NotePropertyValue $restoredHash -Force
    Write-Utf8Json $statePath $state

    foreach ($runtimeFile in @($runtimeX86,$runtimeX64,$runtimeCoreX86,$runtimeCoreX64)) { Remove-RuntimeFile $runtimeFile }
    foreach ($runtimeDirectory in @((Split-Path -Parent $runtimeX86),(Split-Path -Parent $runtimeX64),$runtimeRoot)) {
        if ((Test-Path -LiteralPath $runtimeDirectory -PathType Container) -and @(Get-ChildItem -LiteralPath $runtimeDirectory -Force).Count -eq 0) { Remove-Item -LiteralPath $runtimeDirectory -Force }
    }
    return $restoredHash
}

if ($Action -eq 'Register') {
    if (!(Test-Path -LiteralPath $coreX86 -PathType Leaf) -or !(Test-Path -LiteralPath $coreX64 -PathType Leaf)) { throw 'Verified core DLL pair is required' }
    if (Test-Path -LiteralPath $statePath -PathType Leaf) {
        $existingState = [IO.File]::ReadAllText($statePath) | ConvertFrom-Json
        if ([string]$existingState.status -eq 'ACTIVE') { throw 'an active NxHost registry recovery state already exists' }
    }
    if ((Test-NxHostRegistrationPresent) -and -not $ReplaceExisting) {
        throw 'NXHOST_REGISTRATION_CONFLICT: existing LH.NxHost registration detected; uninstall/recover it or pass -ReplaceExisting explicitly'
    }
    $snapshot = @(Get-RegistrySnapshot)
    $snapshotJson = Convert-SnapshotJson $snapshot
    $snapshotHash = Get-TextSha256 $snapshotJson
    $exports = New-Object 'Collections.Generic.List[object]'
    for ($index = 0; $index -lt $snapshot.Count; $index++) {
        $entry = $snapshot[$index]
        if (-not [bool]$entry.present) { continue }
        $viewBits = if ([string]$entry.view -eq 'Registry32') { 32 } else { 64 }
        $fileName = ('before-{0:d2}-{1}.reg' -f $index, $viewBits)
        $filePath = Join-Path $backup $fileName
        if (Test-Path -LiteralPath $filePath) { Remove-Item -LiteralPath $filePath -Force }
        & reg.exe export ("HKCU\" + [string]$entry.sub_key) $filePath /y ("/reg:$viewBits") | Out-Null
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $filePath -PathType Leaf)) { throw "registry backup failed: $($entry.view) $($entry.sub_key)" }
        $exports.Add([pscustomobject][ordered]@{ view = [string]$entry.view; sub_key = [string]$entry.sub_key; file = $fileName })
    }
    Write-Utf8Json $statePath ([pscustomobject][ordered]@{ schema_version = 1; status = 'ACTIVE'; snapshot_sha256 = $snapshotHash; snapshot = $snapshot; exports = $exports.ToArray() })

    try {
        [void](New-Item -ItemType Directory -Path (Split-Path -Parent $runtimeX86) -Force)
        [void](New-Item -ItemType Directory -Path (Split-Path -Parent $runtimeX64) -Force)
        Copy-Item -LiteralPath $X86Dll -Destination $runtimeX86 -Force
        Copy-Item -LiteralPath $X64Dll -Destination $runtimeX64 -Force
        Copy-Item -LiteralPath $coreX86 -Destination $runtimeCoreX86 -Force
        Copy-Item -LiteralPath $coreX64 -Destination $runtimeCoreX64 -Force
        Invoke-HkcuRegAsm $runtimeX86 $reg32 'x86' 32
        Invoke-HkcuRegAsm $runtimeX64 $reg64 'x64' 64
        Set-OfficeAddinView ([Microsoft.Win32.RegistryView]::Registry32)
        Set-OfficeAddinView ([Microsoft.Win32.RegistryView]::Registry64)
    } catch {
        $registrationFailure = $_.Exception.Message
        try {
            $automaticRestoreHash = Restore-NxHostRegistrationState
        } catch {
            throw ('NxHost registration failed and automatic registry restore failed: primary=' + $registrationFailure + '; restore=' + $_.Exception.Message)
        }
        throw ('NxHost registration failed; automatic registry restore completed: ' + $registrationFailure + '; restored_snapshot_sha256=' + $automaticRestoreHash)
    }
    Write-Output ('REGISTRY_SNAPSHOT_BEFORE|' + $snapshotHash)
} else {
    $restoredHash = Restore-NxHostRegistrationState
    Write-Output ('REGISTRY_SNAPSHOT_AFTER|' + $restoredHash)
}

Write-Output ('PASS|' + $Action)
