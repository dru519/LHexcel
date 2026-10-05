param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RunId = '',
    [string]$CaseNames = 'taxonomy,xlsx,xlsm,xlsb,cancel,read_only,events_disabled,frozen_protected,chart,multiwindow,focus_cf,already_clean,no_book,addin,grouped,cancel_grouped,split,cancel_frozen'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$requestedCases=@($CaseNames.Split(','))
$allowedCases=@('taxonomy','xlsx','xlsm','xlsb','cancel','read_only','events_disabled','frozen_protected','chart','multiwindow','focus_cf','already_clean','no_book','addin','grouped','cancel_grouped','frozen','protected','split','cancel_frozen','new_save','new_cancel','write_error')
$allowedCases += 'draw_popup'
if (-not $requestedCases.Count -or @($requestedCases | Where-Object {$_ -notin $allowedCases}).Count) { throw 'Unknown R59 case' }
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$ProductXlam = [IO.Path]::GetFullPath($ProductXlam)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing R59 evidence is rejected' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close user Excel before the R59 suite' }
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
function Release-Com([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch {}
    }
}
function Add-Probe([object]$Components,[string]$File,[int]$Type,[string]$Name) {
    $component=$null;$module=$null
    try {
        $code=[IO.File]::ReadAllText($File,[Text.Encoding]::UTF8)
        $code=$code.Substring($code.IndexOf('Option Explicit')).Replace("`r`n","`n").Replace("`n","`r`n")
        $component=$Components.Add($Type);$component.Name=$Name
        $module=$component.CodeModule
        [void]$module.AddFromString($code)
    } finally { Release-Com $module;Release-Com $component }
}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$copyPath=Join-Path $EvidenceRoot 'R59Probe.xlam'
Copy-Item -LiteralPath $ProductXlam -Destination $copyPath
$sourceHash=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash
$previousProfile=$env:LHEXCEL_PROFILE_ROOT
$env:LHEXCEL_PROFILE_ROOT=Join-Path $EvidenceRoot 'profile'
$excel=$null;$books=$null;$hostBook=$null;$addin=$null;$project=$null;$components=$null;$binding=$null;$cleanup=$null;$failure=$null
$cases=New-Object 'Collections.Generic.List[object]'
$started=[DateTime]::UtcNow.ToString('o')
try {
    $baseline=Get-ExcelProcessBaseline
    $startupFixture=Join-Path $EvidenceRoot 'startup-fixture.xlsx'
    Copy-Item -LiteralPath (Join-Path $SourceRoot 'tests/fixtures/r60-document-navigator/simple.xlsx') -Destination $startupFixture
    $interactive=Start-ExactInteractiveExcel $baseline 'R59 Manner Excel' -TimeoutMs 120000 -StartupWorkbook $startupFixture
    $excel=$interactive.excel;$binding=$interactive.binding;$interactive=$null
    if (-not $binding.owned) { throw 'Excel process ownership rejected' }
    # Pane geometry needs a realized Excel window, matching the user's UI.
    $excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.EnableEvents=$true
    $books=$excel.Workbooks
    $startupBook=$books.Item('startup-fixture.xlsx')
    try {
        if([IO.Path]::GetFullPath($startupBook.FullName) -ne [IO.Path]::GetFullPath($startupFixture)){throw 'Startup fixture identity mismatch'}
        $startupBook.Close($false)
    } finally {Release-Com $startupBook}
    # Keep a host window while an individual fixture is closed and reopened.
    $hostBook=$books.Add()
    $addin=$books.Open($copyPath,$false,$false)
    $project=$addin.VBProject;$components=$project.VBComponents
    Add-Probe $components (Join-Path $SourceRoot 'tests/vba/product/CNxR59CancelSave.cls') 2 'CNxR59CancelSave'
    Add-Probe $components (Join-Path $SourceRoot 'tests/vba/product/T_R59MannerSave.bas') 1 'T_R59MannerSave'
    if(@($requestedCases|Where-Object {$_ -in @('new_save','new_cancel','write_error')}).Count){
        Add-Probe $components (Join-Path $SourceRoot 'tests/vba/product/T_R59SaveExceptions.bas') 1 'T_R59SaveExceptions'
    }
    $dispatch="'"+$addin.Name+"'!NxR59RunCase"
    foreach($kind in $requestedCases) {
        Write-Output ('START|'+$kind)
        if($kind -in @('new_save','new_cancel','write_error')){
            $path=[string]$excel.Run(("'"+$addin.Name+"'!NxR59ExceptionSetup"),$kind,$EvidenceRoot)
            $faultComponent=$null;$faultModule=$null;$faultLine=0;$originalLine=''
            try{
                if($kind -eq 'write_error'){
                    # Fault injection only in the disposable, never-saved probe XLAM.
                    # This verifies the real error handler, not NTFS/disk-full behavior.
                    $faultComponent=$components.Item('NxMannerSave')
                    $faultModule=$faultComponent.CodeModule
                    $matches=@(1..$faultModule.CountOfLines|Where-Object {$faultModule.Lines($_,1).Trim() -eq 'sourceBook.Save'})
                    if($matches.Count -ne 1){throw 'Unique save injection point not found'}
                    $faultLine=$matches[0];$originalLine=$faultModule.Lines($faultLine,1)
                    $faultModule.ReplaceLine($faultLine,'        Err.Raise 70, "R59 test-only injection", "Simulated write failure"')
                }
                Write-Output ('NATIVE_SAVE_READY|'+$kind+'|'+$path)
                $result=[string]$excel.Run(("'"+$addin.Name+"'!NxR59ExceptionRun"),$kind,$path)
            }finally{
                if($faultLine -gt 0){$faultModule.ReplaceLine($faultLine,$originalLine)}
                Release-Com $faultModule;Release-Com $faultComponent
            }
            if($kind -eq 'new_cancel' -and (Test-Path -LiteralPath $path)){throw 'Cancel created unexpected file'}
            if($kind -eq 'new_save' -and -not (Test-Path -LiteralPath $path)){throw 'SaveAs output missing'}
        }else{
            $result=[string]$excel.Run($dispatch,$kind,$EvidenceRoot)
        }
        if ($kind -eq 'read_only' -and $result.StartsWith('READY|')) {
            $readOnlyBook=$null
            try {
                $readOnlyBook=$books.Open($result.Substring(6),0,$true)
                $result=[string]$excel.Run(("'"+$addin.Name+"'!NxR59CheckReadOnly"))
            } finally {
                if ($null -ne $readOnlyBook) { try {$readOnlyBook.Close($false)} catch {}; Release-Com $readOnlyBook }
            }
        }
        $cases.Add([pscustomobject]@{name=$kind;status=if($result.StartsWith('PASS|')){'PASS'}else{'FAIL'};result=$result})
        Write-Output $result
        if (-not $result.StartsWith('PASS|')) { throw ('R59 case failed: '+$result) }
    }
} catch { $failure=$_.Exception.Message }
finally {
    if($null -ne $addin -and @($requestedCases|Where-Object {$_ -in @('new_save','new_cancel','write_error')}).Count){
        try{[void]$excel.Run(("'" + $addin.Name + "'!NxR59ExceptionDispose"))}catch{}
    }
    if ($null -ne $addin) { try { $addin.Close($false) } catch {} }
    if ($null -ne $hostBook) { try { $hostBook.Close($false) } catch {} }
    Release-Com $hostBook
    Release-Com $components;Release-Com $project;Release-Com $addin;Release-Com $books
    if ($null -ne $excel) { try {$excel.Quit()}catch{};Release-Com $excel }
    $components=$null;$project=$null;$addin=$null;$books=$null;$excel=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($null -ne $binding -and $null -ne $binding.process) {
        $cleanup=Stop-ExactProcessAfterGrace $binding.process 'R59 Manner Excel' 15000 -Detailed
        $binding.process.Dispose()
    }
    $env:LHEXCEL_PROFILE_ROOT=$previousProfile
}
$settled=[Diagnostics.Stopwatch]::StartNew()
do {
    $remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    if ($remaining.Count -eq 0) { break }
    Start-Sleep -Milliseconds 100
} while ($settled.ElapsedMilliseconds -lt 2000)
$afterHash=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash
$passed=@($cases | Where-Object status -eq 'PASS').Count
$status=if($null -eq $failure -and $passed -eq $requestedCases.Count -and $remaining.Count -eq 0 -and $cleanup.exit_mode -eq 'NATURAL' -and $sourceHash -eq $afterHash){'PASS'}else{'FAIL'}
$receipt=[ordered]@{suite='R59Manner';status=$status;run_id=$RunId;started_utc=$started;completed_utc=[DateTime]::UtcNow.ToString('o');cases=@($cases.ToArray());failure=$failure;cleanup=$cleanup;remaining_excel=$remaining;source_sha256=$sourceHash;source_sha256_after=$afterHash;scope='owned native Excel; no install or COM registration'}
[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'R59Manner.json'),(($receipt | ConvertTo-Json -Depth 12)+"`n"),(New-Object Text.UTF8Encoding($false)))
if($status -ne 'PASS'){throw ('R59 Manner verification failed: '+$failure)}
Write-Output ('PASS|R59Manner|'+$passed+'/'+$requestedCases.Count)
