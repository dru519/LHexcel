param([Parameter(Mandatory=$true)][string]$ProductXlam,[Parameter(Mandatory=$true)][string]$SourceRoot,[Parameter(Mandatory=$true)][string]$EvidenceRoot,[switch]$FullFrame,[switch]$RegistryOverride,[switch]$CompareResults)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot '../../build/Excel-ProcessLifecycle.ps1')
function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Fresh evidence required' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Existing Excel preserved' }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$env:LHEXCEL_PROFILE_ROOT=$EvidenceRoot
$excel=$null;$binding=$null;$books=$null;$product=$null;$fixture=$null;$component=$null;$control=$null;$failure=$null
$report=[ordered]@{status='RUNNING';product_sha256=(Get-FileHash -LiteralPath $ProductXlam).Hash;cases=@();evidence_class='native_excel_vba_with_fake_views'}
try {
    $baseline=Get-ExcelProcessBaseline
    $session=Start-ExactInteractiveExcel $baseline 'r66-frame-probe'
    $excel=$session.excel;$binding=$session.binding
    $excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    $books=$excel.Workbooks
    if($books.Count -ne 0){throw 'Unexpected startup workbook'}
    $copy=Join-Path $EvidenceRoot 'frame-product.xlam'
    Copy-Item -LiteralPath $ProductXlam -Destination $copy
    $product=$books.Open($copy,0,$false)
    if($RegistryOverride){
        $registrySource=Join-Path $SourceRoot 'src/vba/frame/NxRegistryFactory.bas'
        $registryText=[IO.File]::ReadAllText($registrySource,[Text.Encoding]::UTF8)
        $component=$product.VBProject.VBComponents.Item('NxRegistryFactory')
        $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
        $component.CodeModule.AddFromString($registryText.Substring($registryText.IndexOf('Option Explicit')).Replace("`r`n","`n").Replace("`n","`r`n"))
        $report.registry_override_sha256=(Get-FileHash -LiteralPath $registrySource).Hash
        Release-ComObject $component;$component=$null
    }
    $imports=@('tests/vba/CFakeFeatureCommand.cls','tests/vba/frame/CFakeExecutionPlanView.cls','tests/vba/frame/CFakeExecutionResultView.cls','tests/vba/frame/T_FrameUi.bas')
    if($FullFrame){$imports+='tests/vba/frame/T_Frame.bas'}
    if($CompareResults){$imports=@('tests/vba/product/T_R67CompareResults.bas');$report.evidence_class='native_excel_compare_results_session'}
    foreach($relative in $imports){
        $text=[IO.File]::ReadAllText((Join-Path $SourceRoot $relative),[Text.Encoding]::UTF8)
        $kind=if($relative.EndsWith('.cls')){2}else{1}
        $component=$product.VBProject.VBComponents.Add($kind)
        $component.Name=[IO.Path]::GetFileNameWithoutExtension($relative)
        $text=$text.Substring($text.IndexOf('Option Explicit'))
        $component.CodeModule.AddFromString($text.Replace("`r`n","`n").Replace("`n","`r`n"))
        Release-ComObject $component;$component=$null
    }
    # Import the actual assertion implementations, without the unrelated suite inventory.
    $harness=[IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/NxTestHarness.bas'),[Text.Encoding]::UTF8)
    $start=$harness.IndexOf('Public Sub AssertTrue(');$end=$harness.IndexOf('Private Sub AppendCase(')
    if($start -lt 0 -or $end -le $start){throw 'Assertion source boundary changed'}
    $component=$product.VBProject.VBComponents.Add(1);$component.Name='NxTestHarness'
    $component.CodeModule.AddFromString("Option Explicit`r`n"+$harness.Substring($start,$end-$start))
    Release-ComObject $component;$component=$null
    $component=$product.VBProject.VBComponents.Add(1);$component.Name='R66FrameProbe'
    $code=@'
Option Explicit
Public Function FrameNames() As String
    Dim item As Variant
    For Each item In T_FrameUi.TestNames()
        FrameNames = FrameNames & CStr(item) & vbLf
    Next item
End Function
Public Function FrameCase(ByVal name As String) As String
    On Error GoTo Failed
    T_FrameUi.RunCase name
    FrameCase = "PASS|" & name
    Exit Function
Failed:
    FrameCase = "FAIL|" & name & "|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@
    if($FullFrame){$code=$code.Replace('T_FrameUi.','T_Frame.')}
    if($CompareResults){
        $code=@'
Option Explicit
Public Function FrameNames() As String
    FrameNames = Replace(T_R67CompareResults.Names(), "|", vbLf)
End Function
Public Function FrameCase(ByVal name As String) As String
    FrameCase = T_R67CompareResults.RunCase(name)
End Function
'@
    }
    $component.CodeModule.AddFromString($code.Replace("`r`n","`n").Replace("`n","`r`n"))
    $component.Activate()
    $control=Find-VbeCompileControl $excel.VBE.CommandBars
    $watcher=Start-ExactOwnedCompileDialogWatcher -ExcelPid ([int]$binding.pid) -ExcelHwnd ([Int64]$binding.hwnd) -VbeHwnd ([Int64]$excel.VBE.MainWindow.HWnd)
    if($control.Enabled){$control.Execute()}
    if($null -eq (Wait-Job $watcher -Timeout 30)){throw 'Compile watcher timeout'}
    $observed=@(Receive-Job $watcher -Wait -AutoRemoveJob)
    if($observed.Count -ne 1 -or $observed[0].status -ne 'NO_DIALOG' -or $control.Enabled){throw 'Frame probe compile failed'}
    $report.compile='PASS'
    $product.Save()
    Release-ComObject $control;$control=$null
    Release-ComObject $component;$component=$null
    $product.Close($false);Release-ComObject $product;$product=$null
    $report.instrumented_sha256=(Get-FileHash -LiteralPath $copy).Hash
    $product=$books.Open($copy,0,$false)
    $macro="'"+$product.Name.Replace("'","''")+"'!"
    $fixture=$books.Add(-4167)
    $fixturePath=Join-Path $EvidenceRoot 'fixture.xlsx';$fixture.SaveAs($fixturePath,51)
    $names=@(([string]$excel.Run($macro+'FrameNames')).Split([char]10)|Where-Object {$_})
    $expectedCount=if($FullFrame){28}else{12}
    if($CompareResults){$expectedCount=15}
    if($names.Count -ne $expectedCount){throw 'Unexpected frame case inventory'}
    foreach($name in $names){
        $result=[string]$excel.Run($macro+'FrameCase',$name)
        $report.cases+= [ordered]@{name=$name;result=$result}
        Write-Output $result
        if($result -ne ('PASS|'+$name)){throw $result}
    }
    if((Get-FileHash -LiteralPath $ProductXlam).Hash -ne $report.product_sha256){throw 'Source Product changed'}
    $report.status='PASS'
} catch {$failure=$_.Exception.Message;$report.status='FAIL';$report.failure=$failure}
finally {
    if($null -ne $binding -and $null -ne $excel){
        [void](Assert-ExactExcelProcessOwnership $binding)
        if($null -ne $books){
            for($i=$books.Count;$i -ge 1;$i--){
                $book=$books.Item($i);$full=[IO.Path]::GetFullPath([string]$book.FullName)
                if(-not $full.StartsWith([IO.Path]::GetFullPath($EvidenceRoot)+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Unexpected workbook preserved'}
                $book.Close($false);Release-ComObject $book
            }
        }
        $excel.Quit()
    }
    foreach($com in @($control,$component,$fixture,$product,$books,$excel)){Release-ComObject $com}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null -ne $binding){
        $report.cleanup=Stop-ExactProcessAfterGrace $binding.process 'r66 frame probe' 15000 -Detailed
        $binding.process.Dispose()
        if($report.cleanup.exit_mode -ne 'NATURAL'){$report.status='FAIL';$failure='Excel did not exit naturally'}
    }
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'frame.json'),($report|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
}
if($null -ne $failure){throw $failure}
Write-Output ('PASS|Frame|'+$expectedCount)
