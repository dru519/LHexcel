param([Parameter(Mandatory=$true)][string]$SetupExe,[Parameter(Mandatory=$true)][string]$PreviousSetupExe,[Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [ValidatePattern('^v0\.01_r[0-9]+$')][string]$PreviousVersion='v0.01_r64',
    [ValidatePattern('^v0\.01_r[0-9]+$')][string]$ExpectedVersion='v0.01_r86')
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Fresh evidence required'}
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Existing Excel preserved'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$prior=$env:LHEXCEL_PROFILE_ROOT
$env:LHEXCEL_PROFILE_ROOT=Join-Path $env:LOCALAPPDATA ('LHExcel/verification/upgrade-profile-'+[Guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $env:LHEXCEL_PROFILE_ROOT)
$oldInstalled=$false
$newInstalled=$false
$result=[ordered]@{status='RUNNING';previous_version=$PreviousVersion;expected_version=$ExpectedVersion;profile=$env:LHEXCEL_PROFILE_ROOT;old_install='NOT_RUN';conflict='NOT_RUN';settings_preserved='NOT_RUN';new_install='NOT_RUN'}
function Snapshot {
    $rows=@(Get-ChildItem -LiteralPath (Join-Path $env:LHEXCEL_PROFILE_ROOT 'Settings') -File | Sort-Object Name | ForEach-Object {$_.Name+'|'+(Get-FileHash -LiteralPath $_.FullName).Hash})
    if($rows.Count -lt 3){throw 'Expected favorites, shortcuts and focus settings'}
    return ($rows -join "`n")
}
function RunSetup([string]$path,[string]$action) {
    $process=Start-Process -FilePath $path -ArgumentList @($action,'--accept-current-user-changes') -Wait -PassThru -WindowStyle Hidden
    return $process.ExitCode
}
try {
    & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Run-R63SetupSuite.ps1') -SetupExe $PreviousSetupExe -EvidenceRoot (Join-Path $EvidenceRoot 'previous') -ExpectedVersion $PreviousVersion -SeedPreferences -KeepInstalled
    # The receipt distinguishes ownership even if the child acceptance check fails.
    $oldReceipt=Join-Path $EvidenceRoot 'previous/installed-acceptance.json'
    if(Test-Path -LiteralPath $oldReceipt){$oldInstalled=(Get-Content -LiteralPath $oldReceipt -Raw|ConvertFrom-Json).install -eq 'PASS'}
    if($LASTEXITCODE -ne 0){throw 'Previous version installed fixture failed'}
    $result.old_install='PASS'
    $before=Snapshot
    $oldName=(-join @([char]0xB0B4,[char]0xC5D1,[char]0xC140))+' '+$PreviousVersion+'.xlam'
    $oldXlam=Join-Path $env:APPDATA ('Microsoft/Excel/XLSTART/'+$oldName)
    $oldHash=(Get-FileHash -LiteralPath $oldXlam).Hash
    if((RunSetup $SetupExe '--install') -eq 0){throw 'New installer overwrote an active prior installation'}
    if((Get-FileHash -LiteralPath $oldXlam).Hash -ne $oldHash -or (Snapshot) -cne $before){throw 'Conflict rejection changed previous files'}
    $result.conflict='PASS'
    if([int]($ExpectedVersion -replace '^v0\.01_r','') -ge 93) {
        if((RunSetup $SetupExe '--update') -ne 0){throw 'Automatic update failed'}
        $oldInstalled=$false;$newInstalled=$true
        $result.automatic_update='PASS'
        $newXlam=Join-Path $env:APPDATA ('Microsoft/Excel/XLSTART/'+(-join @([char]0xB0B4,[char]0xC5D1,[char]0xC140))+' '+$ExpectedVersion+'.xlam')
        $receiptPath=Join-Path $env:LOCALAPPDATA ('LHexcel/NxHost/state/'+$ExpectedVersion.Split('_')[1]+'/native-install.xml')
        [xml]$installedReceipt=Get-Content -LiteralPath $receiptPath -Raw
        $owned=@($installedReceipt.installation.file|Where-Object {$_.path -eq $newXlam})
        if($owned.Count -ne 1 -or (Get-FileHash -LiteralPath $newXlam).Hash -ne $owned[0].sha256){throw 'Updated fixture ownership mismatch'}
        $newHash=(Get-FileHash -LiteralPath $newXlam).Hash
        Remove-Item -LiteralPath $newXlam
        foreach($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
            $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
            $key=$base.OpenSubKey('Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect',$true)
            try {$key.SetValue('LoadBehavior',0,[Microsoft.Win32.RegistryValueKind]::DWord)} finally {$key.Dispose();$base.Dispose()}
        }
        if((RunSetup $SetupExe '--repair') -ne 0){throw 'Repair failed'}
        if((Get-FileHash -LiteralPath $newXlam).Hash -ne $newHash){throw 'Repaired XLAM differs'}
        $result.missing_file_disabled_registration_repair='PASS'
        & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Run-R63SetupSuite.ps1') -SetupExe $SetupExe -EvidenceRoot (Join-Path $EvidenceRoot 'current') -ExpectedVersion $ExpectedVersion -VerifyPreferences -AlreadyInstalled
        $newInstalled=$false
    } else {
        if((RunSetup $PreviousSetupExe '--uninstall') -ne 0){throw 'Previous version removal failed'}
        $oldInstalled=$false
        if((Snapshot) -cne $before){throw 'Previous version removal changed settings'}
        & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Run-R63SetupSuite.ps1') -SetupExe $SetupExe -EvidenceRoot (Join-Path $EvidenceRoot 'current') -ExpectedVersion $ExpectedVersion -VerifyPreferences
    }
    if($LASTEXITCODE -ne 0){throw 'Current version installed acceptance failed'}
    $result.new_install='PASS'
    if((Snapshot) -cne $before){throw 'Upgrade changed stored preferences'}
    $result.settings_preserved='PASS';$result.settings_hashes=$before
    $result.status='PASS'
} catch {
    $result.status='FAIL';$result.failure=$_.Exception.Message
} finally {
    if($newInstalled -and -not @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) {
        $result.new_cleanup=RunSetup $SetupExe '--uninstall'
    }
    if($oldInstalled -and -not @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) {
        $result.old_cleanup=RunSetup $PreviousSetupExe '--uninstall'
    }
    $env:LHEXCEL_PROFILE_ROOT=$prior
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'upgrade.json'),($result|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
}
$result|ConvertTo-Json -Depth 8
if($result.status -ne 'PASS'){exit 1}
