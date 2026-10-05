param([Parameter(Mandatory=$true)][string]$EvidenceRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Existing evidence rejected'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$rows=New-Object 'Collections.Generic.List[object]'
function Assert-Case([string]$name,[bool]$passed){
    $rows.Add([pscustomobject]@{name=$name;passed=$passed})
    if(-not $passed){throw ('FAIL|'+$name)}
}
$trees=@{}
foreach($definition in @(
    @('Run-R60CtpCreation.ps1','ConvertTo-CtpFailureDetail'),
    @('Run-R60RegisteredInventory.ps1','Test-R60RecoveryReceipt'))){
    $errors=$null
    $tree=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $definition[0]),[ref]$null,[ref]$errors)
    if($errors.Count){throw 'Runner parse failure'}
    $trees[$definition[0]]=$tree
    $functions=@($tree.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $definition[1]},$true))
    if($functions.Count -ne 1){throw 'Expected exactly one allowlisted helper'}
    # Never dot-source a runner: its top-level registration/Excel code is not executed.
    . ([ScriptBlock]::Create($functions[0].Extent.Text))
}
$id='aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';$hash='a'*64
function New-TestState {
    [pscustomobject]@{run_id=$id;office_bitness='x64';status='RESTORED';snapshot_sha256=$hash;restored_snapshot_sha256=$hash}
}
$inner=[Runtime.InteropServices.COMException]::new('inner COM failure',-2147221164)
try{throw [Exception]::new('primary CTP failure',$inner)}catch{$detail=ConvertTo-CtpFailureDetail $_ 'ensure_ctp'}
Assert-Case 'phase_and_hresult' ($detail.caller_phase -ceq 'ensure_ctp' -and @($detail.exceptions|Where-Object {$_.hresult -ceq '0x80040154' -and $_.message -ceq 'inner COM failure'}).Count -eq 1)
Assert-Case 'valid_recovery' (Test-R60RecoveryReceipt (New-TestState) $id 'x64' $hash)
foreach($variant in @('run_id','office_bitness','status','snapshot_sha256','restored_snapshot_sha256')){
    $state=New-TestState;$state.$variant='wrong'
    Assert-Case ('reject_'+$variant) (-not (Test-R60RecoveryReceipt $state $id 'x64' $hash))
}
$state=New-TestState;$state.PSObject.Properties.Remove('restored_snapshot_sha256')
Assert-Case 'missing_hash' (-not (Test-R60RecoveryReceipt $state $id 'x64' $hash))
$state=New-TestState;$state.snapshot_sha256='b'*64;$state.restored_snapshot_sha256='b'*64
Assert-Case 'different_initial_snapshot' (-not (Test-R60RecoveryReceipt $state $id 'x64' $hash))
Assert-Case 'initial_hash_unavailable' (-not (Test-R60RecoveryReceipt (New-TestState) $id 'x64' ''))
Assert-Case 'invalid_initial_hash' (-not (Test-R60RecoveryReceipt (New-TestState) $id 'x64' ('z'*64)))
foreach($field in @('run_id','office_bitness','status','snapshot_sha256','restored_snapshot_sha256')){
    foreach($kind in @('array','number','null')){
        $state=New-TestState
        switch($kind){
            'array' {$state.$field=@($state.$field,'wrong')}
            'number' {$state.$field=1}
            'null' {$state.$field=$null}
        }
        Assert-Case ('reject_type_'+$field+'_'+$kind) (-not (Test-R60RecoveryReceipt $state $id 'x64' $hash))
    }
}
$badJsonRejected=$false
try{$null='{'|ConvertFrom-Json}catch{$badJsonRejected=$true}
Assert-Case 'corrupt_json' $badJsonRejected

# Run only the creation finally block with non-COM objects and faulting doubles.
# No Excel process, registry or native lifecycle method is used.
function Release-CtpCom($value){throw 'injected release failure'}
function Stop-ExactProcessAfterGrace($process,$label,$grace,[switch]$Detailed){throw 'injected stop failure'}
function Start-Sleep {param($Milliseconds)}
function Get-Process {param($Name,$ErrorAction) @()}
$cleanupFailures=New-Object 'Collections.Generic.List[string]'
$widths=New-Object 'Collections.Generic.List[object]'
$failure='original failure';$failureDetail=$detail;$cleanup=$null
$automation=[pscustomobject]@{};$addin=[pscustomobject]@{};$addins=[pscustomobject]@{};$books=[pscustomobject]@{}
$book=[pscustomobject]@{};$book|Add-Member ScriptMethod Close {param($save) throw 'injected close failure'}
$excel=[pscustomobject]@{};$excel|Add-Member ScriptMethod Quit {throw 'injected quit failure'}
$process=[pscustomobject]@{};$process|Add-Member ScriptMethod Dispose {throw 'injected dispose failure'}
$binding=[pscustomobject]@{owned=$true;process=$process}
$creation=$trees['Run-R60CtpCreation.ps1']
$mainTry=@($creation.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst]})
if($mainTry.Count -ne 1){throw 'Expected one creation main try'}
. ([ScriptBlock]::Create(($mainTry[0].Finally.Statements.Extent.Text -join "`n")))
Assert-Case 'primary_survives_cleanup' ($failure -ceq 'original failure' -and $failureDetail.caller_phase -ceq 'ensure_ctp')
Assert-Case 'all_cleanup_attempted' ($cleanupFailures.Count -eq 10 -and @($cleanupFailures|Where-Object {$_ -like 'process_dispose:*'}).Count -eq 1)
Assert-Case 'receipt_written_after_cleanup_failure' (Test-Path -LiteralPath (Join-Path $EvidenceRoot 'ctp-creation.json'))

# Evaluate only each runner's final gate to prove cleanup failures never become PASS.
$gate=@($creation.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.IfStatementAst] -and $_.Extent.Text -like '*cleanupFailures.Count*'})
if($gate.Count -ne 1){throw 'Creation gate missing'}
$failure=$null;$cleanup=[pscustomobject]@{exit_mode='NATURAL'};$rejected=$false
try{. ([ScriptBlock]::Create($gate[0].Extent.Text))}catch{$rejected=$true}
Assert-Case 'cleanup_only_failure_rejected' $rejected
$restored=$true;$restoreCommandFailure='injected DLL cleanup failure';$recoveryReadFailure=$null
$runtimeFileCheckFailure=$null;$runtimeDllPresent=$true;$rejected=$false
$registryGate=@($trees['Run-R60RegisteredInventory.ps1'].EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.IfStatementAst] -and $_.Extent.Text -like '*runtimeDllPresent*'})
if($registryGate.Count -ne 1){throw 'Registry gate missing'}
try{. ([ScriptBlock]::Create($registryGate[0].Extent.Text))}catch{$rejected=$true}
Assert-Case 'restored_but_cleanup_failed_rejected' $rejected
[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'diagnostics-unit.json'),([pscustomobject]@{evidence_class='synthetic PowerShell; no Excel or registry execution';cases=@($rows.ToArray())}|ConvertTo-Json -Depth 6))
Write-Output ('PASS|R60_CTP_DIAGNOSTICS|'+$rows.Count)
