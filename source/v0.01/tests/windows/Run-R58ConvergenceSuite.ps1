param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RunId = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$ProductXlam = [IO.Path]::GetFullPath($ProductXlam)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
if (-not (Test-Path -LiteralPath $ProductXlam -PathType Leaf)) { throw 'Product.xlam not found' }
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing R58 convergence evidence root is rejected' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -ne 0) { throw 'R58 convergence requires zero pre-existing Excel processes' }
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

function Write-Json([string]$Path,[object]$Value) { [IO.File]::WriteAllText($Path,(($Value|ConvertTo-Json -Depth 20)+"`n"),(New-Object Text.UTF8Encoding($false))) }
function Write-R58LifecyclePhase([string]$Phase,[string]$Status,[string]$Detail='') {
    $script:LifecyclePhase=$Phase
    $script:LifecycleEvidence.Add([pscustomobject][ordered]@{phase=$Phase;status=$Status;detail=$Detail;observed_utc=[DateTime]::UtcNow.ToString('o')})
    Write-Json $lifecycleReceiptPath ([ordered]@{schema_version=1;suite='R58Convergence';run_id=$RunId;phase=$Phase;status=$Status;source_product_sha256=$sourceHash;owned_pid=if($null -ne $binding){$binding.pid}else{$null};events=@($script:LifecycleEvidence.ToArray())})
}
function Release-Com([object]$Value) { if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)}catch{}} }
function Release-ComObject([object]$Value) { Release-Com $Value }
function Get-FreshExcelProcessIds() {
    $processes=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
    try { return @($processes|ForEach-Object{[int]$_.Id}) }
    finally { foreach($process in $processes){try{$process.Dispose()}catch{}} }
}
function Wait-FreshExcelProcessExit([int]$TimeoutMs=2000) {
    $watch=[Diagnostics.Stopwatch]::StartNew();$ids=@()
    do {
        $ids=@(Get-FreshExcelProcessIds)
        if($ids.Count -eq 0){return [pscustomobject][ordered]@{status='CLEARED';elapsed_ms=$watch.ElapsedMilliseconds;remaining_pids=@()}}
        if($watch.ElapsedMilliseconds -lt $TimeoutMs){Start-Sleep -Milliseconds 100}
    } while($watch.ElapsedMilliseconds -lt $TimeoutMs)
    return [pscustomobject][ordered]@{status='REMAINING';elapsed_ms=$watch.ElapsedMilliseconds;remaining_pids=$ids}
}
function Get-PortableExecutableMachine([string]$Path) {
    $stream=$null;$reader=$null
    try {
        $stream=[IO.File]::OpenRead($Path);$reader=New-Object IO.BinaryReader($stream)
        if($reader.ReadUInt16() -ne 0x5A4D){return 'NOT_MZ'}
        $stream.Position=0x3C;$peOffset=$reader.ReadInt32()
        if($peOffset -lt 0){return 'INVALID_PE_OFFSET'}
        $stream.Position=$peOffset
        if($reader.ReadUInt32() -ne 0x00004550){return 'NOT_PE'}
        switch($reader.ReadUInt16()){
            0x014C {return 'x86'}
            0x8664 {return 'x64'}
            0xAA64 {return 'ARM64'}
            default {return 'UNKNOWN'}
        }
    } catch {return 'UNAVAILABLE: '+$_.Exception.Message}
    finally {if($null -ne $reader){$reader.Dispose()}elseif($null -ne $stream){$stream.Dispose()}}
}
function Add-Case([string]$Name,[scriptblock]$Action,[scriptblock]$Cleanup) {
    $script:CaseSequence++
    $progressPath=Join-Path $EvidenceRoot ('R58Convergence.case.{0:D2}.json' -f $script:CaseSequence)
    $script:CurrentCaseVbaResults=New-Object 'Collections.Generic.List[object]'
    $startedUtc=[DateTime]::UtcNow.ToString('o')
    Write-Json $progressPath ([ordered]@{schema_version=1;suite='R58Convergence';run_id=$RunId;sequence=$script:CaseSequence;name=$Name;status='RUNNING';phase='started';started_utc=$startedUtc;vba_results=@()})
    $failure=$null;$hostBefore=$null;$hostAfter=$null
    try {$hostBefore=Activate-R58ContextHost} catch {$failure='host before: '+$_.Exception.Message}
    if($null -eq $failure){try { & $Action } catch { $failure=$_.Exception.Message }}
    $rpcLost=Test-R58RpcLoss $failure
    if(-not $rpcLost){try { & $Cleanup } catch { if ($null -eq $failure){$failure='cleanup: '+$_.Exception.Message}else{$failure+=' | cleanup: '+$_.Exception.Message} }}
    if(-not $rpcLost){try {$hostAfter=Assert-R58ContextHostIntegrity ('after '+$Name)} catch { if ($null -eq $failure){$failure='host after: '+$_.Exception.Message}else{$failure+=' | host after: '+$_.Exception.Message} }}
    $case=[pscustomobject][ordered]@{name=$Name;status=if($null -eq $failure){'PASS'}else{'FAIL'};failure=$failure;progress_receipt=$progressPath;vba_results=@($script:CurrentCaseVbaResults.ToArray());host_before=$hostBefore;host_after=$hostAfter;rpc_lost=$rpcLost}
    $script:Cases.Add($case)
    Write-Json $progressPath ([ordered]@{schema_version=1;suite='R58Convergence';run_id=$RunId;sequence=$script:CaseSequence;name=$Name;status=$case.status;phase='completed';started_utc=$startedUtc;completed_utc=[DateTime]::UtcNow.ToString('o');failure=$failure;vba_results=$case.vba_results;host_before=$hostBefore;host_after=$hostAfter;rpc_lost=$rpcLost})
    if($rpcLost){throw ('R58 Excel RPC loss after case '+$Name+': '+$failure)}
}
function Get-R58ContextHostObservation() {
    $sheet=$null;$marker=$null;$value=$null;$formula=$null;$activeBook=$null
    try {
        if($null -eq $script:ContextHostBook){throw 'R58 context host workbook was not created'}
        $sheet=$script:ContextHostBook.Worksheets.Item(1)
        $marker=$sheet.Range('A1');$value=$sheet.Range('A2');$formula=$sheet.Range('A3');$activeBook=$script:ContextHostBook.Application.ActiveWorkbook
        return [ordered]@{status='OBSERVED';name=[string]$script:ContextHostBook.Name;marker=[string]$marker.Value2;value=[string]$value.Value2;formula=[string]$formula.Formula;saved=[bool]$script:ContextHostBook.Saved;active_book=if($null -eq $activeBook){$null}else{[string]$activeBook.Name}}
    } catch {return [ordered]@{status='UNAVAILABLE';failure=$_.Exception.Message}}
    # ActiveWorkbook can be the exact same RCW retained as ContextHostBook.
    # FinalRelease on this borrowed alias would invalidate the owning handle.
    finally {$activeBook=$null;Release-Com $formula;Release-Com $value;Release-Com $marker;Release-Com $sheet}
}
function Assert-R58ContextHostIntegrity([string]$Phase,[switch]$RequireActive) {
    $observed=Get-R58ContextHostObservation;$expected=$script:ContextHostExpected;$failure=$null
    if($null -eq $expected -or $observed.status -ne 'OBSERVED'){$failure='host observation unavailable: '+[string]$observed.failure}
    elseif($observed.name -ne $expected.name -or $observed.marker -ne $expected.marker -or $observed.value -ne $expected.value -or $observed.formula -ne $expected.formula -or $observed.saved -ne $expected.saved){$failure='host marker/value/formula/saved state changed'}
    elseif($RequireActive -and $observed.active_book -ne $expected.name){$failure='host workbook was not active before dispatch'}
    if($null -ne $failure){$script:ContextHostEvidence.status='FAIL';$script:ContextHostEvidence.failure=$failure;$script:ContextHostEvidence.failed_phase=$Phase;throw ('R58 context host rejected at '+$Phase+': '+$failure)}
    $script:ContextHostEvidence.status='PASS';$script:ContextHostEvidence.last_verified_phase=$Phase;$script:ContextHostEvidence.last_observation=$observed
    return $observed
}
function Activate-R58ContextHost() {
    $sheet=$null
    try {$script:ContextHostBook.Activate();$sheet=$script:ContextHostBook.Worksheets.Item(1);$sheet.Activate()}
    finally {Release-Com $sheet}
    return Assert-R58ContextHostIntegrity 'before dispatch' -RequireActive
}
function Get-ModuleBody([string]$Path) {
    $text=[IO.File]::ReadAllText($Path,[Text.Encoding]::UTF8).Replace("`r`n","`n").Replace("`r","`n")
    $lines=@($text -split "`n")
    $body=@($lines|Where-Object{$_ -notmatch '^\s*Attribute\s+'}) -join "`r`n"
    return $body
}
function Add-VbaModule([object]$Components,[string]$Path,[string]$Name) {
    $existing=$null;$component=$null;$module=$null;$exists=$false
    try {$existing=$Components.Item($Name);$exists=$null -ne $existing}catch{}
    finally {Release-Com $existing}
    if($exists){return}
    try {
        $component=$Components.Add(1)
        $component.Name=$Name
        $module=$component.CodeModule
        $code = (Get-ModuleBody $Path) -replace '\r?\n', "`r`n"
        [void]$module.AddFromString($code)
    } finally {Release-Com $module;Release-Com $component}
}
function Test-R58RpcLoss([string]$Message) {
    if([string]::IsNullOrWhiteSpace($Message)){return $false}
    return $Message -match '0x800706BE|0x800706BA|remote procedure call failed|RPC server is unavailable'
}
function Invoke-R58NativeCase([object]$Excel,[string]$Dispatcher,[string]$Kind,[string]$CaseName='') {
    $result=$null
    try {
        $result=[string]$Excel.Run($Dispatcher,$Kind,$CaseName)
    } catch {
        $message=$_.Exception.Message
        if(Test-R58RpcLoss $message){$script:RpcLossEvidence=[ordered]@{status='RPC_LOST';kind=$Kind;case_name=$CaseName;failure=$message;observed_utc=[DateTime]::UtcNow.ToString('o')}
        }
        throw
    }
    $entry=[ordered]@{kind=$Kind;case_name=$CaseName;result=$result;observed_utc=[DateTime]::UtcNow.ToString('o')}
    if($null -ne $script:CurrentCaseVbaResults){$script:CurrentCaseVbaResults.Add($entry)}
    if($result -ne 'PASS'){throw ('R58 native static dispatch failed: kind='+$Kind+'; case='+$CaseName+'; result='+$result)}
}
function Invoke-R58ComparePrepare([object]$Excel,[string]$Dispatcher) {
    $result=$null
    try {$result=[string]$Excel.Run($Dispatcher,'COMPARE_PREPARE','')} catch {
        $message=$_.Exception.Message
        if(Test-R58RpcLoss $message){$script:RpcLossEvidence=[ordered]@{status='RPC_LOST';kind='COMPARE_PREPARE';case_name=$null;failure=$message;observed_utc=[DateTime]::UtcNow.ToString('o')}}
        throw
    }
    if($result -notmatch '^REPORT\|.+$'){throw ('R58 compare producer returned invalid receipt: '+$result)}
    $reportPath=[IO.Path]::GetFullPath($result.Substring(7))
    if(-not (Test-Path -LiteralPath $reportPath -PathType Leaf)){throw ('R58 compare producer report missing: '+$reportPath)}
    if($null -ne $script:CurrentCaseVbaResults){$script:CurrentCaseVbaResults.Add([ordered]@{kind='COMPARE_PREPARE';case_name=$null;result=$result;observed_utc=[DateTime]::UtcNow.ToString('o')})}
    $script:CompareVerificationEvidence.producer_returned=$true;$script:CompareVerificationEvidence.report_path=$reportPath
    return $reportPath
}
function Invoke-R58CompareOutsideProducerVerification([object]$Excel,[string]$Dispatcher,[string]$ReportPath) {
    $report=$null;$failure=$null
    try {
        $report=$Excel.Workbooks.Open($ReportPath,0,$true)
        $script:CompareVerificationEvidence.opened_outside_producer=$true
        $script:CompareVerificationEvidence.opened_book_name=[string]$report.Name
        $script:CompareVerificationEvidence.opened_read_only=[bool]$report.ReadOnly
        Invoke-R58NativeCase $Excel $Dispatcher 'COMPARE_VERIFY_OPEN' $ReportPath
        $script:CompareVerificationEvidence.contents_verified=$true
    } catch {$failure=$_.Exception.Message}
    finally {if($null -ne $report){try{$report.Close($false)}catch{if($null -eq $failure){$failure='report close: '+$_.Exception.Message}};Release-Com $report}}
    try {
        Invoke-R58NativeCase $Excel $Dispatcher 'COMPARE_CLEANUP' $ReportPath
        if(Test-Path -LiteralPath $ReportPath -PathType Leaf){throw 'R58 compare producer report remained after cleanup'}
        $script:CompareVerificationEvidence.cleanup_completed=$true
    } catch {if($null -eq $failure){$failure='report cleanup: '+$_.Exception.Message}else{$failure+=' | report cleanup: '+$_.Exception.Message}}
    if($null -ne $failure){$script:CompareVerificationEvidence.status='FAIL';$script:CompareVerificationEvidence.failure=$failure;throw $failure}
    $script:CompareVerificationEvidence.status='PASS'
}
function Get-R58NativeDispatchModuleEvidence([object]$Components) {
    $component=$null;$module=$null
    try {
        $component=$Components.Item('R58NativeDispatch')
        $module=$component.CodeModule
        [void]$component.Activate()
        $lineCount=[int]$module.CountOfLines
        $source=if($lineCount -gt 0){[string]$module.Lines(1,$lineCount)}else{''}
        $dispatchLine=0
        $lines=@($source -split "`r?`n")
        for($index=0;$index -lt $lines.Count;$index++){
            if($lines[$index] -match '^\s*Public Function NxR58TestDispatch\b'){$dispatchLine=$index+1}
        }
        return [ordered]@{
            status='FOUND';component_name=[string]$component.Name;component_type=[int]$component.Type;activated=$true;line_count=$lineCount;
            proc_of_line_static_dispatch=if($dispatchLine -gt 0){[string]$module.ProcOfLine($dispatchLine,0)}else{$null};
            source=$source
        }
    } catch {
        return [ordered]@{status='MISSING_OR_UNREADABLE';failure=$_.Exception.Message}
    } finally {Release-Com $module;Release-Com $component}
}
function Invoke-R58NoArgumentEntryProbe([object]$Excel,[string]$EntryName) {
    $failure=$null;$result=$null
    try {$result=$Excel.Run($EntryName)}catch{$failure=$_.Exception.Message}
    return [ordered]@{entry=$EntryName;com_failure=$failure;result=if($null -eq $result){$null}else{[string]$result}}
}
function Invoke-R58StaticDispatchProbe([object]$Excel,[string]$EntryName) {
    $failure=$null;$result=$null
    try {$result=[string]$Excel.Run($EntryName,'PING','')}catch{$failure=$_.Exception.Message}
    return [ordered]@{entry=$EntryName;com_failure=$failure;result=$result}
}
function Get-R58MacroContext([object]$Excel,[object]$Book) {
    $vbe=$null;$activeProject=$null
    try {
        $vbe=$Excel.VBE;$activeProject=$vbe.ActiveVBProject
        return [ordered]@{
            automation_security=[int]$Excel.AutomationSecurity;enable_events=[bool]$Excel.EnableEvents;
            book_name=[string]$Book.Name;book_full_name=[string]$Book.FullName;book_is_addin=[bool]$Book.IsAddin;book_read_only=[bool]$Book.ReadOnly;
            vbe_active_project=if($null -eq $activeProject){$null}else{[string]$activeProject.FileName}
        }
    } catch {
        return [ordered]@{status='UNAVAILABLE';failure=$_.Exception.Message}
    # ActiveVBProject can share the retained project RCW; drop only this borrowed alias.
    } finally {$activeProject=$null;Release-Com $vbe}
}
function Set-InstrumentedConst([object]$Components,[string]$Name) {
    $component=$null;$module=$null
    try {
        $component=$Components.Item($Name);$module=$component.CodeModule
        $text=$module.Lines(1,$module.CountOfLines)
        if ($text -notmatch '#Const NX_R58_TEST_BUILD = False'){throw "instrumentation marker missing: $Name"}
        $text=$text.Replace('#Const NX_R58_TEST_BUILD = False','#Const NX_R58_TEST_BUILD = True') -replace '\r?\n', "`r`n"
        $module.DeleteLines(1,$module.CountOfLines);[void]$module.AddFromString($text)
    } finally {Release-Com $module;Release-Com $component}
}
function Invoke-R58VbaCompile([object]$Excel,[object]$Binding,[object]$Components,[string]$CopyPath,[string]$CompileModuleName='T_R58Data',[switch]$AllowAlreadyCompiled) {
    $vbe=$null;$bars=$null;$compileControl=$null;$compileTarget=$null;$activeProject=$null;$mainWindow=$null;$compileWatcher=$null
    $script:CompileEvidence=[ordered]@{status='RUNNING';control_id=578;active_project=$null;dialog=$null;failure=$null}
    try {
        $vbe=$Excel.VBE;$bars=$vbe.CommandBars;$compileControl=Find-VbeCompileControl $bars
        if($null -eq $compileControl -or [int]$compileControl.Id -ne 578){throw 'R58 VBA compile control 578 unavailable'}
        $compileTarget=$Components.Item($CompileModuleName);[void]$compileTarget.Activate()
        $activeProject=$vbe.ActiveVBProject;$activeProjectPath=[IO.Path]::GetFullPath([string]$activeProject.FileName)
        $script:CompileEvidence.active_project=$activeProjectPath
        if($activeProjectPath -ne [IO.Path]::GetFullPath($CopyPath)){throw 'R58 VBA compile active project identity mismatch'}
        if(-not [bool]$compileControl.Enabled){
            if($AllowAlreadyCompiled){
                $script:CompileEvidence.status='COMPILE_RETAINED'
                $script:CompileEvidence.state='NOT_REQUIRED'
                return
            }
            throw 'R58 VBA compile control was disabled before the required compile'
        }
        $mainWindow=$vbe.MainWindow
        $compileWatcher=Start-ExactOwnedCompileDialogWatcher -ExcelPid ([int]$Binding.pid) -ExcelHwnd ([Int64]$Binding.hwnd) -VbeHwnd ([Int64]$mainWindow.HWnd)
        $compileExecuteError=$null
        try{[void]$compileControl.Execute()}catch{$compileExecuteError=$_.Exception.Message}
        $watchCompleted=Wait-Job -Job $compileWatcher -Timeout 30
        if($null -eq $watchCompleted){Stop-Job -Job $compileWatcher -ErrorAction SilentlyContinue;throw 'R58 VBA compile dialog watcher timed out'}
        $watchResult=@(Receive-Job -Job $compileWatcher -Wait -AutoRemoveJob -ErrorAction Stop|Select-Object -First 1);$compileWatcher=$null
        if(-not [string]::IsNullOrWhiteSpace([string]$compileExecuteError)){throw ('R58 VBA compile 578 execution failed: '+$compileExecuteError)}
        if($watchResult.Count -ne 1 -or [string]$watchResult[0].status -ne 'NO_DIALOG'){
            $dialogText=if($watchResult.Count -eq 1){[string]$watchResult[0].dialog_text}else{'watcher returned no result'}
            $script:CompileEvidence.dialog=$dialogText
            $compilePane=$null;$compileModule=$null;$compileModuleName='UNAVAILABLE';[int]$compileLine=0;[int]$compileColumn=0;[int]$compileEndLine=0;[int]$compileEndColumn=0;$compileSource='UNAVAILABLE';$paneFailure=$null
            try{$compilePane=$vbe.ActiveCodePane;if($null -eq $compilePane){throw 'ActiveCodePane unavailable after compile dialog'};$compileModule=$compilePane.CodeModule;$compileModuleName=[string]$compileModule.Name;$compilePane.GetSelection([ref]$compileLine,[ref]$compileColumn,[ref]$compileEndLine,[ref]$compileEndColumn);if($compileLine -gt 0){$compileSource=([string]$compileModule.Lines($compileLine,1)).Replace("`r",' ').Replace("`n",' ')}}catch{$paneFailure=$_.Exception.Message}finally{Release-Com $compileModule;Release-Com $compilePane}
            $location='compile_module='+$compileModuleName+'; compile_line='+[string]$compileLine+'; compile_column='+[string]$compileColumn+'; compile_source='+$compileSource;if(-not [string]::IsNullOrWhiteSpace([string]$paneFailure)){$location+='; compile_pane_error='+$paneFailure}
            throw ('R58 VBA compile dialog detected: '+$dialogText+"`n"+$location)
        }
        if([bool]$compileControl.Enabled){throw 'R58 VBA compile did not reach saved disabled state'}
        $script:CompileEvidence.status='PASS'
        $script:CompileEvidence.state='EXECUTED'
    } catch {$script:CompileEvidence.status='FAIL';$script:CompileEvidence.failure=$_.Exception.Message;throw}
    # The caller owns the project; ActiveVBProject is only a borrowed alias here.
    finally {if($null -ne $compileWatcher){try{Stop-Job -Job $compileWatcher -ErrorAction SilentlyContinue;Remove-Job -Job $compileWatcher -Force -ErrorAction SilentlyContinue}catch{}};$activeProject=$null;Release-Com $mainWindow;Release-Com $compileTarget;Release-Com $compileControl;Release-Com $bars;Release-Com $vbe}
}

$Cases=New-Object 'Collections.Generic.List[object]';$CaseSequence=0;$CurrentCaseVbaResults=$null;$excel=$null;$binding=$null;$books=$null;$copy=$null;$copyPath=$null;$project=$null;$components=$null;$failure=$null;$cleanup=$null;$sourceHash=$null;$copyBeforeHash=$null;$sourceMarkerHashes=[ordered]@{};$CompileEvidence=[ordered]@{status='NOT_RUN';state=$null;control_id=578;active_project=$null;dialog=$null;failure=$null};$PostReopenCompileEvidence=[ordered]@{status='NOT_RUN';state=$null;control_id=578;active_project=$null;dialog=$null;failure=$null};$DispatchEvidence=[ordered]@{before_save=$null;after_reopen_before_compile=$null;after_reopen_after_compile=$null;context=$null;product_profile_root=$null;static_dispatch_ping=$null;test_entry_callable=$false;selected_entry_form=$null;compare_trace_path=$null};$HostEvidence=[ordered]@{excel_pid=$null;excel_executable=$null;excel_machine='NOT_RUN';powershell_is_64bit=[Environment]::Is64BitProcess;excel_version=$null;excel_operating_system=$null};$script:ContextHostBook=$null;$ContextHostSheet=$null;$script:ContextHostExpected=$null;$script:ContextHostEvidence=[ordered]@{status='NOT_CREATED';expected=$null;created=$null;last_verified_phase=$null;last_observation=$null;failure=$null;failed_phase=$null};$script:CompareVerificationEvidence=[ordered]@{status='NOT_RUN';producer_returned=$false;report_path=$null;opened_outside_producer=$false;opened_book_name=$null;opened_read_only=$null;contents_verified=$false;cleanup_completed=$false;failure=$null};$RpcLossEvidence=[ordered]@{status='NOT_DETECTED';kind=$null;case_name=$null;failure=$null;observed_utc=$null};$CrashEvidence=[ordered]@{status='NOT_DETECTED';failure=$null;observed_utc=$null};$executionScope='FULL_CONVERGENCE';$isolatedProfileRoot=[IO.Path]::GetFullPath((Join-Path $EvidenceRoot 'Profile'));$previousProfileRoot=[Environment]::GetEnvironmentVariable('LHEXCEL_PROFILE_ROOT','Process');$compareTracePath=Join-Path $EvidenceRoot 'R58Compare.trace.jsonl';$previousCompareTracePath=[Environment]::GetEnvironmentVariable('LHEXCEL_R58_COMPARE_TRACE','Process');$ownedProcessReceiptPath=$null;$ownedProcessReceipt=$null
$sourceMarkers=@('src/vba/features/file/rename/NxBatchRenameEngine.bas','src/vba/features/file/pdf/NxPdfController.bas')
$script:LifecycleEvidence=New-Object 'Collections.Generic.List[object]'
$script:LifecyclePhase='preflight'
$lifecycleReceiptPath=Join-Path $EvidenceRoot 'R58Convergence.lifecycle.json'
try {
    foreach($relative in $sourceMarkers){$sourceMarkerPath=Join-Path $SourceRoot $relative;if (([IO.File]::ReadAllText($sourceMarkerPath,[Text.Encoding]::UTF8)) -notmatch '#Const NX_R58_TEST_BUILD = False'){throw "shipped source test marker is not False: $relative"};$sourceMarkerHashes[$relative]=(Get-FileHash -LiteralPath $sourceMarkerPath -Algorithm SHA256).Hash.ToLowerInvariant()}
    [void](New-Item -ItemType Directory -Path $EvidenceRoot)
    [void](New-Item -ItemType Directory -Path $isolatedProfileRoot)
    [Environment]::SetEnvironmentVariable('LHEXCEL_PROFILE_ROOT',$isolatedProfileRoot,'Process')
    [Environment]::SetEnvironmentVariable('LHEXCEL_R58_COMPARE_TRACE',$compareTracePath,'Process')
    $DispatchEvidence.compare_trace_path=$compareTracePath
    $sourceHash=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash.ToLowerInvariant();$copyPath=Join-Path $EvidenceRoot 'Product.r58.instrumented.xlam';Copy-Item -LiteralPath $ProductXlam -Destination $copyPath -ErrorAction Stop
    $copyBeforeHash=(Get-FileHash -LiteralPath $copyPath -Algorithm SHA256).Hash.ToLowerInvariant()
    $baseline=Get-ExcelProcessBaseline;$excel=New-Object -ComObject Excel.Application;$binding=Get-ExactExcelProcessOwnership $excel $baseline 'R58 convergence Excel'
    if(-not $binding.owned){throw ('Excel ownership rejected: '+$binding.failure)}
    $excelExecutable=$null
    try{$excelExecutable=[string]$binding.process.MainModule.FileName}catch{}
    $HostEvidence=[ordered]@{excel_pid=[int]$binding.pid;excel_executable=$excelExecutable;excel_machine=if([string]::IsNullOrWhiteSpace($excelExecutable)){'UNAVAILABLE'}else{Get-PortableExecutableMachine $excelExecutable};powershell_is_64bit=[Environment]::Is64BitProcess;excel_version=[string]$excel.Version;excel_operating_system=[string]$excel.OperatingSystem}
    $ownedProcessReceiptPath=Join-Path $EvidenceRoot 'R58Convergence.owned-process.json'
    $ownedProcessReceipt=[ordered]@{schema_version=1;suite='R58Convergence';status='OWNED';phase='pre_compile';run_id=$RunId;owned_pid=[int]$binding.pid;owned_hwnd=[Int64]$binding.hwnd;source_product_sha256=$sourceHash;instrumented_copy=$copyPath;copy_before_sha256=$copyBeforeHash;written_utc=[DateTime]::UtcNow.ToString('o')}
    Write-Json $ownedProcessReceiptPath $ownedProcessReceipt
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$books=$excel.Workbooks;$copy=$books.Open($copyPath,$false,$false)
    if($copy.ReadOnly){throw 'Instrumented Product copy opened read-only'}
    try{$project=$copy.VBProject}catch{throw ('VBOM access blocked: '+$_.Exception.Message)};$components=$project.VBComponents
    Set-InstrumentedConst $components 'NxBatchRenameEngine';Set-InstrumentedConst $components 'NxPdfController'
    Add-VbaModule $components (Join-Path $SourceRoot 'tests/vba/file/T_R58Compare.bas') 'T_R58Compare'|Out-Null
    Add-VbaModule $components (Join-Path $SourceRoot 'tests/vba/file/T_R58Recovery.bas') 'T_R58Recovery'|Out-Null
    Add-VbaModule $components (Join-Path $SourceRoot 'tests/vba/file/R58NativeDispatch.bas') 'R58NativeDispatch'|Out-Null
    Add-VbaModule $components (Join-Path $SourceRoot 'tests/vba/data/T_R58Data.bas') 'T_R58Data'|Out-Null
    Add-VbaModule $components (Join-Path $SourceRoot 'tests/vba/draw/T_R58ModelMerge.bas') 'T_R58ModelMerge'|Out-Null
    $DispatchEvidence.before_save=Get-R58NativeDispatchModuleEvidence $components
    Write-R58LifecyclePhase 'compile' 'STARTED'
    Invoke-R58VbaCompile $excel $binding $components $copyPath
    Write-R58LifecyclePhase 'compile' 'COMPLETED'
    $initialCompileEvidence=$CompileEvidence
    Write-R58LifecyclePhase 'save' 'STARTED'
    [void](Assert-ExactExcelProcessOwnership $binding)
    $copy.Save()
    Write-R58LifecyclePhase 'save' 'COMPLETED'
    Write-R58LifecyclePhase 'release_children' 'STARTED'
    Release-Com $components;$components=$null
    Release-Com $project;$project=$null
    Write-R58LifecyclePhase 'release_children' 'COMPLETED'
    Write-R58LifecyclePhase 'close' 'STARTED'
    $copy.Close($false)
    Write-R58LifecyclePhase 'close' 'COMPLETED'
    Write-R58LifecyclePhase 'release_copy' 'STARTED'
    Release-Com $copy;$copy=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    Write-R58LifecyclePhase 'release_copy' 'COMPLETED'
    Write-R58LifecyclePhase 'reopen' 'STARTED'
    [void](Assert-ExactExcelProcessOwnership $binding)
    $copy=$books.Open($copyPath,$false,$false)
    if($copy.ReadOnly -or -not $copy.IsAddin -or -not [string]::Equals([IO.Path]::GetFullPath([string]$copy.FullName),[IO.Path]::GetFullPath($copyPath),[StringComparison]::OrdinalIgnoreCase)){throw 'R58 reopened Product copy identity rejected'}
    Write-R58LifecyclePhase 'reopen' 'COMPLETED'
    Write-R58LifecyclePhase 'bind_project' 'STARTED'
    try{$project=$copy.VBProject}catch{throw ('VBOM access blocked after R58 reopen: '+$_.Exception.Message)};$components=$project.VBComponents
    Write-R58LifecyclePhase 'bind_project' 'COMPLETED'
    $DispatchEvidence.after_reopen_before_compile=Get-R58NativeDispatchModuleEvidence $components
    try {
        Write-R58LifecyclePhase 'compile_after_reopen' 'STARTED'
        Invoke-R58VbaCompile $excel $binding $components $copyPath 'R58NativeDispatch' -AllowAlreadyCompiled
        $PostReopenCompileEvidence=$CompileEvidence
        Write-R58LifecyclePhase 'compile_after_reopen' 'COMPLETED'
    } catch {
        $PostReopenCompileEvidence=$CompileEvidence
        throw
    } finally {
        $CompileEvidence=$initialCompileEvidence
    }
    $DispatchEvidence.after_reopen_after_compile=Get-R58NativeDispatchModuleEvidence $components
    Write-R58LifecyclePhase 'dispatch_preflight' 'STARTED'
    $macro="'"+$copy.Name.Replace("'","''")+"'!"
    $DispatchEvidence.context=Get-R58MacroContext $excel $copy
    $DispatchEvidence.product_profile_root=Invoke-R58NoArgumentEntryProbe $excel ($macro+'NxLHexcelProfileRoot')
    if($DispatchEvidence.product_profile_root.com_failure -ne $null -or [IO.Path]::GetFullPath([string]$DispatchEvidence.product_profile_root.result) -cne $isolatedProfileRoot){throw ('R58 profile isolation rejected: expected='+ $isolatedProfileRoot+'; actual='+[string]$DispatchEvidence.product_profile_root.result+'; failure='+[string]$DispatchEvidence.product_profile_root.com_failure)}
    $dispatcher=$macro+'NxR58TestDispatch'
    $DispatchEvidence.static_dispatch_ping=Invoke-R58StaticDispatchProbe $excel $dispatcher
    if($DispatchEvidence.static_dispatch_ping.com_failure -ne $null -or $DispatchEvidence.static_dispatch_ping.result -cne 'PASS'){throw 'R58 native dispatch preflight found no callable static test entry'}
    $DispatchEvidence.test_entry_callable=$true;$DispatchEvidence.selected_entry_form='BOOK_BARE_STATIC'
    Write-R58LifecyclePhase 'dispatch_preflight' 'COMPLETED'
    Write-R58LifecyclePhase 'cases' 'STARTED'
    $script:ContextHostBook=$books.Add();$ContextHostSheet=$script:ContextHostBook.Worksheets.Item(1)
    $ContextHostSheet.Range('A1').Value2='R58_HOST_MARKER';$ContextHostSheet.Range('A2').Value2='R58_HOST_VALUE';$ContextHostSheet.Range('A3').Formula='=40+2'
    $script:ContextHostExpected=[ordered]@{name=[string]$script:ContextHostBook.Name;marker='R58_HOST_MARKER';value='R58_HOST_VALUE';formula='=40+2';saved=[bool]$script:ContextHostBook.Saved}
    $script:ContextHostEvidence=[ordered]@{status='PASS';expected=$script:ContextHostExpected;created=Get-R58ContextHostObservation;last_verified_phase=$null;last_observation=$null;failure=$null;failed_phase=$null}
    Add-Case 'compare_run_all' {
        $reportPath=Invoke-R58ComparePrepare $excel $dispatcher
        Invoke-R58CompareOutsideProducerVerification $excel $dispatcher $reportPath
        Invoke-R58NativeCase $excel $dispatcher 'COMPARE_REMAINING'
    } {}
    foreach($name in @('TestFileChainTemporaryMoveFailure','TestFileChainTargetMoveFailure','TestFileRollbackBlockedByReplacement','TestFileTargetReplacementIdentityFailsClosed','TestSheetChainRollbackUsesObjects','TestLockedPdfCleanupReportsResidual','TestRenameInjectionRejectsInvalidStage')){ $caseName=$name;Add-Case ('recovery_'+$caseName) { Invoke-R58NativeCase $excel $dispatcher 'RECOVERY' $caseName } { Invoke-R58NativeCase $excel $dispatcher 'RECOVERY_CLEANUP' } }
    Add-Case 'data_run_all' { Invoke-R58NativeCase $excel $dispatcher 'DATA' } {}
    Add-Case 'model_merge_run_all' { Invoke-R58NativeCase $excel $dispatcher 'MODEL_MERGE' } {}
    Write-R58LifecyclePhase 'cases' 'COMPLETED'
} catch {
    $failure=$_.Exception.Message
    if((Test-R58RpcLoss $failure) -and $RpcLossEvidence.status -ne 'RPC_LOST'){
        $RpcLossEvidence=[ordered]@{status='RPC_LOST';kind='LIFECYCLE';case_name=$script:LifecyclePhase;failure=$failure;observed_utc=[DateTime]::UtcNow.ToString('o')}
    }
    if(Test-Path -LiteralPath $EvidenceRoot -PathType Container){
        try{Write-R58LifecyclePhase $script:LifecyclePhase 'FAILED' $failure}catch{$failure+=' | lifecycle receipt: '+$_.Exception.Message}
    }
} finally {
    if($RpcLossEvidence.status -eq 'RPC_LOST' -and $CrashEvidence.status -eq 'NOT_DETECTED'){
        $CrashEvidence=[ordered]@{status='ABNORMAL_EXIT';failure=('Owned Excel RPC was lost: '+$RpcLossEvidence.failure);observed_utc=[DateTime]::UtcNow.ToString('o')}
    }
    if($null -ne $binding -and $null -ne $binding.process){
        try {
            if($binding.process.HasExited){$CrashEvidence=[ordered]@{status='ABNORMAL_EXIT';failure='Owned Excel process exited before harness cleanup';observed_utc=[DateTime]::UtcNow.ToString('o')}}
        } catch {}
    }
    Release-Com $components;$components=$null;Release-Com $project;$project=$null
    if ($null -ne $copy){try{$copy.Close($false)}catch{}};if ($null -ne $script:ContextHostBook){try{$script:ContextHostBook.Close($false)}catch{}};foreach($item in @($copy,$ContextHostSheet,$script:ContextHostBook,$books)){Release-Com $item}
    if ($null -ne $excel){if($CrashEvidence.status -eq 'NOT_DETECTED'){try{$excel.Quit()}catch{}};Release-Com $excel};[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($null -ne $binding -and $null -ne $binding.process){
        $cleanup=Stop-ExactProcessAfterGrace $binding.process 'R58 convergence Excel' 15000 -Detailed
        if($CrashEvidence.status -eq 'ABNORMAL_EXIT'){$cleanup.exit_mode='ABNORMAL_EXIT';$cleanup.failure=$CrashEvidence.failure}
        try{$binding.process.Dispose()}catch{}
    }
    [Environment]::SetEnvironmentVariable('LHEXCEL_PROFILE_ROOT',$previousProfileRoot,'Process')
    [Environment]::SetEnvironmentVariable('LHEXCEL_R58_COMPARE_TRACE',$previousCompareTracePath,'Process')
}
$remainingImmediate=@(Get-FreshExcelProcessIds);$remainingCheck=Wait-FreshExcelProcessExit 2000;$remaining=@($remainingCheck.remaining_pids);$passed=@($Cases|Where-Object status -eq 'PASS').Count;$cleanupNatural=($null -ne $cleanup -and $cleanup.exit_mode -eq 'NATURAL' -and [string]::IsNullOrWhiteSpace([string]$cleanup.failure));$compileAccepted=($CompileEvidence.status -eq 'PASS' -and ($PostReopenCompileEvidence.status -eq 'PASS' -or $PostReopenCompileEvidence.status -eq 'COMPILE_RETAINED'));$hostAccepted=($script:ContextHostEvidence.status -eq 'PASS' -and $null -eq $script:ContextHostEvidence.failure -and $script:ContextHostEvidence.last_verified_phase -like 'after *');$compareVerificationAccepted=($script:CompareVerificationEvidence.status -eq 'PASS' -and $script:CompareVerificationEvidence.producer_returned -and $script:CompareVerificationEvidence.opened_outside_producer -and $script:CompareVerificationEvidence.opened_read_only -and $script:CompareVerificationEvidence.contents_verified -and $script:CompareVerificationEvidence.cleanup_completed);$status=if ($null -eq $failure -and $Cases.Count -eq 10 -and $passed -eq 10 -and $remaining.Count -eq 0 -and $compileAccepted -and $cleanupNatural -and $hostAccepted -and $compareVerificationAccepted -and $RpcLossEvidence.status -eq 'NOT_DETECTED' -and $CrashEvidence.status -eq 'NOT_DETECTED'){'PASS'}else{'DIAGNOSTIC'}
$sourceHashAfter=if(Test-Path $ProductXlam){(Get-FileHash $ProductXlam -Algorithm SHA256).Hash.ToLowerInvariant()}else{$null};if($null -ne $sourceHash -and $sourceHashAfter -ne $sourceHash){$failure='Supplied Product.xlam changed during convergence suite';$status='DIAGNOSTIC'}
$receipt=[ordered]@{schema_version=1;suite='R58Convergence';status=$status;execution_scope=$executionScope;run_id=$RunId;evidence_root=$EvidenceRoot;isolated_profile_root=$isolatedProfileRoot;source_markers_sha256=$sourceMarkerHashes;source_product_sha256=$sourceHash;source_product_sha256_after=$sourceHashAfter;instrumented_copy=$copyPath;copy_before_sha256=$copyBeforeHash;copy_after_sha256=if($null -ne $copyPath -and (Test-Path $copyPath)){(Get-FileHash -LiteralPath $copyPath -Algorithm SHA256).Hash.ToLowerInvariant()}else{$null};owned_process_receipt=$ownedProcessReceiptPath;owned_process=$ownedProcessReceipt;compile=$CompileEvidence;compile_after_reopen=$PostReopenCompileEvidence;dispatch=$DispatchEvidence;compare_verification=$script:CompareVerificationEvidence;compare_verification_accepted=$compareVerificationAccepted;host_context=$script:ContextHostEvidence;host_preservation_accepted=$hostAccepted;rpc=$RpcLossEvidence;crash=$CrashEvidence;cases=@($Cases.ToArray());failure=$failure;cleanup=$cleanup;remaining_excel_immediate=$remainingImmediate;remaining_excel_recheck=$remainingCheck;remaining_excel=$remaining;completed_utc=[DateTime]::UtcNow.ToString('o')}
$receipt.lifecycle=@($script:LifecycleEvidence.ToArray())
$receipt.lifecycle_receipt=$lifecycleReceiptPath
Write-Json (Join-Path $EvidenceRoot 'R58Convergence.json') $receipt
if ($status -eq 'PASS'){Write-Output 'PASS|R58Convergence|10/10';exit 0};Write-Output ('DIAGNOSTIC|R58Convergence|'+$passed+'/'+$Cases.Count);throw ('R58 convergence failed; receipt='+ (Join-Path $EvidenceRoot 'R58Convergence.json'))
