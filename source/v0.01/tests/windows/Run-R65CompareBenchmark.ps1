param([Parameter(Mandatory=$true)][string]$ProductXlam,[Parameter(Mandatory=$true)][string]$EvidenceRoot)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value) } catch {}
    }
}
. (Join-Path $PSScriptRoot '../../build/Excel-ProcessLifecycle.ps1')
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Fresh benchmark evidence required' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close user Excel first' }
New-Item -ItemType Directory -Path $EvidenceRoot | Out-Null
$env:LHEXCEL_PROFILE_ROOT=$EvidenceRoot
$excel=$null; $binding=$null; $books=$null; $product=$null; $failure=$null
$report=[ordered]@{status='RUNNING';product_sha256=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash;cases=@()}
try {
    $baseline=Get-ExcelProcessBaseline
    $session=Start-ExactInteractiveExcel $baseline 'r65-compare-benchmark'
    $excel=$session.excel; $binding=$session.binding
    # This host coerces StatusBar=False into BSTR "FALSE" even without product code.
    # Keep the independent probe visible in receipts; benchmark the interactive host.
    $excel.Visible=$true; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    $books=$excel.Workbooks
    $copy=Join-Path $EvidenceRoot 'benchmark-product.xlam'
    Copy-Item -LiteralPath $ProductXlam -Destination $copy
    $product=$books.Open($copy,0,$false)
    $report.activation='interactive_x_rot'
    $observer=$product.VBProject.VBComponents.Add(2); $observer.Name='CNxR66CompareProgress'
    $observerCode=[IO.File]::ReadAllText((Join-Path $PSScriptRoot '../vba/product/CNxR66CompareProgress.cls'),[Text.Encoding]::UTF8)
    $observerCode=$observerCode.Substring($observerCode.IndexOf('Option Explicit'))
    $observer.CodeModule.AddFromString($observerCode.Replace("`r`n","`n").Replace("`n","`r`n"))
    Release-ComObject $observer; $observer=$null
    $component=$product.VBProject.VBComponents.Add(1); $component.Name='R65BenchmarkProbe'
    $probeCode=@'
Option Explicit
Public Function NxR66StatusProbe() As String
    Dim appObject As Object, prior As Variant
    prior = Application.StatusBar
    Application.StatusBar = "probe"
    Application.StatusBar = False
    NxR66StatusProbe = "early=" & CStr(VarType(Application.StatusBar)) & ":" & CStr(Application.StatusBar)
    Set appObject = Application
    appObject.StatusBar = False
    NxR66StatusProbe = NxR66StatusProbe & "|late=" & CStr(VarType(Application.StatusBar)) & ":" & CStr(Application.StatusBar)
    CallByName Application, "StatusBar", VbLet, False
    NxR66StatusProbe = NxR66StatusProbe & "|dispatch=" & CStr(VarType(Application.StatusBar)) & ":" & CStr(Application.StatusBar)
    If VarType(prior) = vbBoolean Then appObject.StatusBar = False Else appObject.StatusBar = prior
End Function
Public Function NxR65Measure(ByVal scenario As String, ByVal kind As String, ByVal repeat As String) As String
    Dim result As CNxResult, root As String, output As String, service As Object
    Dim progress As New CNxR66CompareProgress, status As Variant, cancelKey As XlEnableCancelKey, books As Long
    Dim failure As Long
    On Error GoTo Failed
    root = Environ$("LHEXCEL_PROFILE_ROOT")
    root = root & "\" & scenario
    output = root & "\report-" & kind & "-" & repeat & ".xlsx"
    Set service = NxHostCreateWorkbookCompare()
    If service Is Nothing Then Err.Raise vbObjectError + 966, , "Exact DLL compare service required"
    status = Application.StatusBar: cancelKey = Application.EnableCancelKey: books = Workbooks.Count
    If kind = "cancel" Then progress.CancelEnabled = True
    NxWorkbookCompareBindProgress progress
    On Error Resume Next
    Set result = NxWorkbookCompareReport(root & "\base.xlsx", root & "\other.xlsx", output, (kind = "True"))
    failure = Err.Number: Err.Clear
    On Error GoTo Failed
    If Application.StatusBar <> status Or Application.EnableCancelKey <> cancelKey Or Workbooks.Count <> books Then Err.Raise vbObjectError + 966, , "Excel state leak: status=" & CStr(status) & "(" & CStr(VarType(status)) & ")/" & CStr(Application.StatusBar) & "(" & CStr(VarType(Application.StatusBar)) & "); cancel=" & CStr(cancelKey) & "/" & CStr(Application.EnableCancelKey) & "; books=" & CStr(books) & "/" & CStr(Workbooks.Count)
    If kind = "cancel" Then
        If failure <> 18 Or Not progress.Triggered Or Len(Dir$(output)) <> 0 Then Err.Raise vbObjectError + 966, , "Cancel not clean"
        NxR65Measure = "PASS|CANCEL|state-restored"
        Exit Function
    End If
    If failure <> 0 Then Err.Raise failure, , "Comparison failed"
    If result Is Nothing Then Err.Raise vbObjectError + 966, , "No result"
    If result.Outcome <> NxSuccess Then Err.Raise vbObjectError + 966, , "Outcome not success"
    If kind = "True" Then
        If Not progress.VbaReport Or progress.NativeReport Then Err.Raise vbObjectError + 966, , "VBA route not observed"
        NxR65Measure = "PASS|VBA|state-restored"
    Else
        If Not progress.NativeReport Or progress.VbaReport Then Err.Raise vbObjectError + 966, , "DLL route not observed"
        NxR65Measure = "PASS|DLL|state-restored"
    End If
    Exit Function
Failed:
    NxR65Measure = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@
    $component.CodeModule.AddFromString($probeCode.Replace("`r`n","`n").Replace("`n","`r`n"))
    $component.Activate()
    $control=Find-VbeCompileControl $excel.VBE.CommandBars
    $watcher=Start-ExactOwnedCompileDialogWatcher -ExcelPid ([int]$binding.pid) -ExcelHwnd ([Int64]$binding.hwnd) -VbeHwnd ([Int64]$excel.VBE.MainWindow.HWnd)
    if ($control.Enabled) { $control.Execute() }
    if ($null -eq (Wait-Job $watcher -Timeout 30)) { throw 'Compile watcher timeout' }
    $observed=@(Receive-Job $watcher -Wait -AutoRemoveJob)
    if ($observed.Count -ne 1 -or $observed[0].status -ne 'NO_DIALOG' -or $control.Enabled) { throw 'Benchmark compile failed' }
    $product.Save()
    Release-ComObject $control; $control=$null
    Release-ComObject $component; $component=$null
    $product.Close($false); Release-ComObject $product; $product=$null
    $product=$books.Open($copy,0,$false)
    $macro="'"+$product.Name.Replace("'","''")+"'!"
    $report.profile_verified=([string]$excel.Run($macro+'NxLHexcelProfileRoot') -eq $EvidenceRoot)
    $report.status_roundtrip=[string]$excel.Run($macro+'NxR66StatusProbe')
    Write-Output ('STATUS_PROBE|'+$report.status_roundtrip)
    $sourceHashes=@{}
    foreach($scenario in @('sparse','dense','large')) {
    $fixtureRoot=Join-Path $EvidenceRoot $scenario
    New-Item -ItemType Directory -Path $fixtureRoot | Out-Null
    $height=if($scenario -eq 'large'){10000}elseif($scenario -eq 'dense'){100}else{500}
    foreach ($side in @('base','other')) {
        $book=$books.Add(-4167)
        $sheet=$book.Worksheets.Item(1)
        $sheet.Name='Data'
        $data=New-Object 'object[,]' $height,20
        for ($r=0;$r -lt $height;$r++) { for ($c=0;$c -lt 20;$c++) { $value=[double]($r*20+$c);if($side -eq 'other' -and $scenario -eq 'dense'){$value=-$value-1};$data[$r,$c]=$value } }
        $range=$sheet.Range(('A1:T'+$height)); $range.Value2=$data
        if ($side -eq 'other') { $sheet.Range('B2').Value2=[double]42; $sheet.Range('D4').Formula='=1+2';$literal=$sheet.Range('C6');$literal.NumberFormat='@';$literal.Value2='=literal-not-formula';Release-ComObject $literal }
        # Keep formatting identical so the format-aware VBA engine is an exact value/formula oracle.
        if ($side -eq 'base') { $sheet.Range('C6').NumberFormat='@' }
        $path=Join-Path $fixtureRoot ($side+'.xlsx')
        $book.SaveAs($path,51); $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($range)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
        $sourceHashes[$path]=(Get-FileHash -LiteralPath $path).Hash
    }
    }
    $context=$books.Add(-4167); $context.SaveAs((Join-Path $EvidenceRoot 'context.xlsx'),51)
    $report.profile_after_fixtures=([string]$excel.Run($macro+'NxLHexcelProfileRoot') -eq $EvidenceRoot)
    foreach($scenario in @('sparse','dense','large')) {
    $runs=@('False:0','False:1','False:2','False:3')
    if($scenario -ne 'large'){$runs+= 'True:1'}
    foreach ($run in $runs) {
        $formats,$repeat=$run.Split(':')
        $output=Join-Path (Join-Path $EvidenceRoot $scenario) ('report-'+$formats+'-'+$repeat+'.xlsx')
        $watch=[Diagnostics.Stopwatch]::StartNew()
        $result=$excel.Run($macro+'NxR65Measure',$scenario,$formats,$repeat)
        $watch.Stop()
        if (-not ([string]$result).StartsWith('PASS|')) { throw ([string]$result) }
        if (-not (Test-Path -LiteralPath $output)) { throw 'Report not created' }
        $report.cases+= [ordered]@{scenario=$scenario;formats=$formats;repeat=$repeat;warmup=($repeat -eq '0');engine=([string]$result).Split('|')[1];elapsed_ms=$watch.Elapsed.TotalMilliseconds;output_sha256=(Get-FileHash -LiteralPath $output).Hash}
        if ($null -ne $result -and [Runtime.InteropServices.Marshal]::IsComObject($result)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($result) }
    }
    }
    $report.cancel=[string]$excel.Run($macro+'NxR65Measure','sparse','cancel','1')
    if($report.cancel -ne 'PASS|CANCEL|state-restored'){throw $report.cancel}
    foreach($path in $sourceHashes.Keys) { if((Get-FileHash -LiteralPath $path).Hash -ne $sourceHashes[$path]) {throw 'Source changed'} }
    if(@(Get-ChildItem -LiteralPath $EvidenceRoot -Filter '*.tmp.xlsx' -Recurse).Count){throw 'Temporary report residue'}
    $report.sources_preserved=$true
    $report.status='PASS'
} catch { $failure=$_.Exception.Message; $report.status='FAIL'; $report.failure=$failure }
finally {
    if ($null -ne $binding -and $null -ne $excel) {
        [void](Assert-ExactExcelProcessOwnership $binding)
        if ($null -ne $books) {
            for ($i=$books.Count;$i -ge 1;$i--) {
                $book=$books.Item($i)
                $full=[IO.Path]::GetFullPath([string]$book.FullName)
                if ($full -ne [IO.Path]::GetFullPath($ProductXlam) -and -not $full.StartsWith([IO.Path]::GetFullPath($EvidenceRoot)+'\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected workbook preserved' }
                $book.Close($false)
                Release-ComObject $book
            }
        }
        $excel.Quit()
    }
    foreach ($com in @($context,$product,$books,$excel)) { if ($null -ne $com -and [Runtime.InteropServices.Marshal]::IsComObject($com)) { try {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)} catch {} } }
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null -ne $binding) {
        $report.cleanup=Stop-ExactProcessAfterGrace $binding.process 'compare benchmark' 15000 -Detailed
        $binding.process.Dispose()
        if($report.cleanup.exit_mode -ne 'NATURAL'){$report.status='FAIL';$failure='Excel did not exit naturally'}
    }
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'benchmark.json'),($report|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
}
if ($null -ne $failure) { throw $failure }
Write-Output ($report|ConvertTo-Json -Depth 8)
