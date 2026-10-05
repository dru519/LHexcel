param([Parameter(Mandatory=$true)][string]$EvidenceRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Close user Excel before registered inventory'}
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Existing evidence rejected'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$runId=[Guid]::NewGuid().ToString()
$register=Join-Path $PSScriptRoot 'Register-DocumentNavigatorCtpProbe.ps1'
$build=Join-Path $PSScriptRoot 'build'
$statePath=Join-Path $EvidenceRoot 'registry-recovery/registry-state.json'
$failure=$null
$restored=$false
$expectedSnapshotHash=$null
$restoreCommandCompleted=$false;$restoreCommandFailure=$null;$recoveryReadFailure=$null
$runtimeDllPresent=$null;$runtimeFileCheckFailure=$null
$runtimeDll=Join-Path $env:LOCALAPPDATA ('LHexcel\CtpProbe\'+$runId+'\x64\DocumentNavigatorCtpProbe64.dll')
function Test-R60RecoveryReceipt($state,[string]$expectedRunId,[string]$expectedBitness,[string]$expectedSnapshotHash){
    if($null -eq $state -or $expectedSnapshotHash -cnotmatch '^[0-9a-f]{64}$'){return $false}
    foreach($name in @('run_id','office_bitness','status','snapshot_sha256','restored_snapshot_sha256')){
        if($null -eq $state.PSObject.Properties[$name] -or $state.$name -isnot [string]){return $false}
    }
    return ($state.run_id -ceq $expectedRunId -and $state.office_bitness -ceq $expectedBitness -and
        $state.status -ceq 'RESTORED' -and $state.snapshot_sha256 -ceq $expectedSnapshotHash -and
        $state.restored_snapshot_sha256 -ceq $expectedSnapshotHash)
}
$rows=New-Object 'Collections.Generic.List[object]'
try{
    & $register -Action Register -OfficeBitness x64 -RunId $runId -BuildRoot $build -EvidenceRoot $EvidenceRoot -OptIn
    $initialState=Get-Content -LiteralPath $statePath -Raw|ConvertFrom-Json
    foreach($name in @('run_id','office_bitness','status','snapshot_sha256')){
        if($null -eq $initialState -or $null -eq $initialState.PSObject.Properties[$name] -or $initialState.$name -isnot [string]){
            throw 'Initial registration receipt field type mismatch'
        }
    }
    if($initialState.run_id -cne $runId -or $initialState.office_bitness -cne 'x64' -or
       $initialState.status -cne 'ACTIVE' -or $initialState.snapshot_sha256 -cnotmatch '^[0-9a-f]{64}$'){
        throw 'Initial registration receipt identity mismatch'
    }
    $expectedSnapshotHash=[string]$initialState.snapshot_sha256
    foreach($viewName in @('Registry32','Registry64')){
        $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::$viewName)
        try{
            foreach($path in @('Software\Microsoft\Office\Excel\Addins\LH.MyExcel.DocNavCtpProbe.Connect',
                'Software\Classes\LH.MyExcel.DocNavCtpProbe.Connect\CLSID',
                'Software\Classes\CLSID\{7E08DD6E-3E56-426B-890C-02B1E9F2A752}\InprocServer32',
                'Software\Classes\LH.MyExcel.DocNavCtpProbe.Pane\CLSID',
                'Software\Classes\CLSID\{CBC48A05-8444-46E4-AEBE-375551497DD2}\InprocServer32',
                'Software\Microsoft\Office\Common\Security',
                'Software\Policies\Microsoft\Office\Common\Security')){
                $key=$base.OpenSubKey($path,$false)
                try{
                    $values=@{}
                    if($null -ne $key){foreach($name in $key.GetValueNames()){$values[$name]=$key.GetValue($name)}}
                    $rows.Add([pscustomobject]@{view=$viewName;path=$path;present=($null -ne $key);values=$values})
                }finally{if($null -ne $key){$key.Dispose()}}
            }
        }finally{$base.Dispose()}
    }
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'registered-keys.json'),(ConvertTo-Json -InputObject @($rows.ToArray()) -Depth 8))
    & powershell.exe -NoProfile -File (Join-Path $PSScriptRoot 'Run-R60PaneActivation.ps1') -EvidenceRoot $EvidenceRoot
    if($LASTEXITCODE -ne 0){throw 'Independent pane activation failed'}
    & (Join-Path $PSScriptRoot 'Run-R60ComInventory.ps1') -EvidenceRoot (Join-Path $EvidenceRoot 'inventory')
    & (Join-Path $PSScriptRoot 'Run-R60CtpCreation.ps1') -EvidenceRoot $EvidenceRoot
}catch{$failure=$_.Exception.Message}
finally{
    if(Test-Path -LiteralPath $statePath){
        try{
            $state=Get-Content -LiteralPath $statePath -Raw|ConvertFrom-Json
            if($state.status -eq 'ACTIVE'){
                & $register -Action Unregister -OfficeBitness x64 -RunId $runId -BuildRoot $build -EvidenceRoot $EvidenceRoot
                $restoreCommandCompleted=$true
            }
        }catch{$restoreCommandFailure=$_.Exception.Message}
        finally{
            # The helper writes RESTORED before temporary DLL cleanup. Always reread it.
            try{
                $state=Get-Content -LiteralPath $statePath -Raw|ConvertFrom-Json
                $restored=Test-R60RecoveryReceipt $state $runId 'x64' $expectedSnapshotHash
            }catch{$recoveryReadFailure=$_.Exception.Message}
        }
    }
    try{$runtimeDllPresent=Test-Path -LiteralPath $runtimeDll -PathType Leaf}catch{$runtimeFileCheckFailure=$_.Exception.Message}
    $receipt=[pscustomobject]@{run_id=$runId;failure=$failure;registry_restored=$restored;
        initial_snapshot_captured=($null -ne $expectedSnapshotHash);expected_snapshot_sha256=$expectedSnapshotHash;
        registry_evidence='recovery receipt hashes; not an independent live readback';
        restore_command_completed=$restoreCommandCompleted;restore_command_failure=$restoreCommandFailure;
        recovery_receipt_failure=$recoveryReadFailure;runtime_dll_present=$runtimeDllPresent;runtime_file_check_failure=$runtimeFileCheckFailure;
        remaining_excel=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id);evidence_class='registration and discovery only; not CTP acceptance'}
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'result.json'),($receipt|ConvertTo-Json -Depth 5))
}
if($null -ne $failure -or -not $restored -or $null -ne $restoreCommandFailure -or $null -ne $recoveryReadFailure -or
   $null -ne $runtimeFileCheckFailure -or $runtimeDllPresent -ne $false -or $receipt.remaining_excel.Count){throw ('Registered inventory failed: '+$failure)}
Write-Output 'PASS|REGISTERED_INVENTORY|discovery only'
