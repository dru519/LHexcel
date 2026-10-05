param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [switch]$MemoryCacheProbe,
    [switch]$NoCalculateProbe,
    [switch]$PauseForVisual,
    [switch]$RepaintProbe,
    [switch]$AcceptClipboardSuspension
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing evidence rejected' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close user Excel before native probe' }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$originalHash=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash
$copy=Join-Path $EvidenceRoot 'Product.xlam'
Copy-Item -LiteralPath $ProductXlam -Destination $copy
$oldProfile=$env:LHEXCEL_PROFILE_ROOT
$env:LHEXCEL_PROFILE_ROOT=Join-Path $EvidenceRoot 'profile'
function Release-Com([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch {}
    }
}
function Cache-Bounds([object]$addin) {
    $bounds=@{}
    foreach($axis in @('TOP','BOTTOM','LEFT','RIGHT')) {
        $name=$null
        try {
            $name=$addin.Names.Item('NX_FOCUS_'+$axis+'_V2')
            if ($MemoryCacheProbe) { $bounds[$axis]=[int]$addin.Application.Evaluate($name.RefersTo) }
            else { $bounds[$axis]=[int]([string]$name.RefersTo).TrimStart('=') }
        } catch { $bounds[$axis]=-1 } finally { Release-Com $name }
    }
    return $bounds
}
function Owned-Rules([object]$sheet) {
    $count=0
    foreach($rule in $sheet.Cells.FormatConditions) {
        try { if ($rule.Type -eq 2 -and [string]$rule.Formula1 -match 'NX_FOCUS_RULE_') { $count++ } }
        finally { Release-Com $rule }
    }
    return $count
}
$excel=$null;$books=$null;$hostBook=$null;$addin=$null;$book=$null;$sheet=$null
$binding=$null;$cleanup=$null;$failure=$null
$stage='initialize';$probeState=@{}
$cases=New-Object 'Collections.Generic.List[object]'
function Install-MemoryCacheProbe([object]$addin) {
    # Feasibility experiment on the disposable, unsaved add-in copy ONLY.
    $project=$addin.VBProject;$components=$project.VBComponents
    $component=$components.Add(1);$component.Name='NxFocusMemoryProbe'
    $module=$component.CodeModule
    $code=@'
Option Explicit
Public Function NxFocusProbeBound(ByVal axis As Long) As Long
    Dim target As Range
    Set target = Application.Selection
    Select Case axis
        Case 0: NxFocusProbeBound = target.Row
        Case 1: NxFocusProbeBound = target.Row + target.Rows.Count - 1
        Case 2: NxFocusProbeBound = target.Column
        Case 3: NxFocusProbeBound = target.Column + target.Columns.Count - 1
    End Select
End Function
'@
    $module.AddFromString($code)
    Release-Com $module;Release-Com $component
    $component=$components.Item('CNxFocusRuleEngine');$module=$component.CodeModule
    $start=$module.ProcStartLine('UpdateAddinCache',0);$count=$module.ProcCountLines('UpdateAddinCache',0)
    $module.DeleteLines($start,$count)
    $code=@'
Private Sub UpdateAddinCache(ByVal selection As CNxFocusSelection)
    Dim keys As Variant, i As Long, formula As String
    keys = Array("NX_FOCUS_TOP_V2", "NX_FOCUS_BOTTOM_V2", "NX_FOCUS_LEFT_V2", "NX_FOCUS_RIGHT_V2")
    For i = 0 To 3
        formula = "='" & ThisWorkbook.Name & "'!NxFocusProbeBound(" & CStr(i) & ")"
        If InStr(ThisWorkbook.Names(keys(i)).RefersTo, "NxFocusProbeBound") = 0 Then
            ThisWorkbook.Names(keys(i)).RefersTo = formula
        End If
    Next i
End Sub
'@
    $module.InsertLines($start,$code)
    if($NoCalculateProbe){
        $all=$module.Lines(1,$module.CountOfLines)
        $replacement="    ' Experiment: rely on CF repaint, not explicit Calculate"
        if($RepaintProbe){$replacement="    If Application.ScreenUpdating Then"+[Environment]::NewLine+"        Application.ScreenUpdating = False"+[Environment]::NewLine+"        Application.ScreenUpdating = True"+[Environment]::NewLine+"    End If"}
        $all=$all.Replace('    sheet.Calculate', $replacement)
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($all)
    }
    Release-Com $module;Release-Com $component
    $component=$components.Item('CNxFocusApplicationEvents');$module=$component.CodeModule
    $start=$module.ProcStartLine('ExcelApp_SheetSelectionChange',0);$count=$module.ProcCountLines('ExcelApp_SheetSelectionChange',0)
    $module.DeleteLines($start,$count)
    $module.InsertLines($start,@'
Private Sub ExcelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not mListening Or Target Is Nothing Then Exit Sub
    NxFocusControllerSelectionChanged Target, False
End Sub
'@)
    Release-Com $module;Release-Com $component;Release-Com $components;Release-Com $project
}
try {
    $startupFixture=Join-Path $EvidenceRoot 'startup-fixture.xlsx'
    Copy-Item -LiteralPath (Join-Path $SourceRoot 'tests/fixtures/r60-document-navigator/simple.xlsx') -Destination $startupFixture
    $interactive=Start-ExactInteractiveExcel (Get-ExcelProcessBaseline) 'Focus Copy Paste' -TimeoutMs 120000 -StartupWorkbook $startupFixture
    $excel=$interactive.excel;$binding=$interactive.binding;$interactive=$null
    $excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.EnableEvents=$true
    $startupBook=$excel.Workbooks.Item('startup-fixture.xlsx')
    try {
        if([IO.Path]::GetFullPath($startupBook.FullName) -ne [IO.Path]::GetFullPath($startupFixture)){throw 'Startup fixture identity mismatch'}
        $startupBook.Close($false)
    } finally {Release-Com $startupBook}
    $books=$excel.Workbooks;$hostBook=$books.Add()
    $addin=$books.Open($copy,$false,$true)
    if($MemoryCacheProbe){Install-MemoryCacheProbe $addin}
    $macro="'"+$addin.Name+"'!"
    foreach($size in @(@(1,1),@(2,2),@(2,3),@(3,2))) {
        $book=$books.Add();$sheet=$book.Worksheets.Item(1)
        $rows=$size[0];$columns=$size[1]
        $source=$sheet.Range('B2').Resize($rows,$columns)
        $target=$sheet.Range('K12').Resize($rows,$columns)
        $source.FormulaR1C1='=ROW()*100+COLUMN()'
        $source.NumberFormat='0.00';$source.Font.Bold=$true
        $source.Select()
        $enabled=[string]$excel.Run($macro+'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#5FC8D8',40,$false,$true)
        $backend=[string]$excel.Run($macro+'NxFocusControllerBackend')
        $before=Cache-Bounds $addin
        # Native Copy/Paste commands keep the real clipboard phase separate.
        $stage='copy'
        $excel.CommandBars.ExecuteMso('Copy')
        $copied=[int]$excel.CutCopyMode
        $stage='select-destination'
        $target.Select()
        $excel.ActiveWindow.ScrollRow=8;$excel.ActiveWindow.ScrollColumn=5
        Start-Sleep -Milliseconds 200
        if($PauseForVisual -and $rows -eq 2 -and $columns -eq 2){
            Write-Output ('VISUAL_READY|pid='+$binding.pid+'|copy_mode='+[int]$excel.CutCopyMode)
            Start-Sleep -Seconds 45
        }
        $afterMove=Cache-Bounds $addin
        $destinationRowColor=[long]$sheet.Range('A12').DisplayFormat.Interior.Color
        $oldRowColor=[long]$sheet.Range('A2').DisplayFormat.Interior.Color
        $copyAfterMove=[int]$excel.CutCopyMode
        $probeState=@{size=("$rows"+"x"+"$columns");backend=$backend;copy_before=$copied;copy_after=$copyAfterMove;bounds=$afterMove;destination_color=$destinationRowColor;old_color=$oldRowColor;paste_enabled=[bool]$excel.CommandBars.GetEnabledMso('Paste')}
        $stage='paste'
        $excel.CommandBars.ExecuteMso('Paste')
        Start-Sleep -Milliseconds 200
        $valuesOk=$true
        for($row=1;$row -le $rows;$row++){
            for($column=1;$column -le $columns;$column++){
                $cell=$target.Cells.Item($row,$column)
                try { if ([double]$cell.Value2 -ne (($row+11)*100+$column+10)) { $valuesOk=$false } }
                finally { Release-Com $cell }
            }
        }
        $formatOk=([string]$target.NumberFormat -eq '0.00' -and [bool]$target.Font.Bold)
        $undoAvailable=[bool]$excel.CommandBars.GetEnabledMso('Undo')
        $afterPaste=Cache-Bounds $addin
        $rulesAfterPaste=Owned-Rules $sheet
        $undoRestored=$false
        if($undoAvailable){
            $excel.CommandBars.ExecuteMso('Undo')
            $undoRestored=($null -eq $target.Value2 -or [double]$excel.WorksheetFunction.CountA($target) -eq 0)
        }
        $excel.CutCopyMode=$false
        $sheet.Range('R20').Select()
        Start-Sleep -Milliseconds 200
        $afterResume=Cache-Bounds $addin
        $excel.Run($macro+'NxFocusDisable') | Out-Null
        $residual=Owned-Rules $sheet
        $cases.Add([pscustomobject]@{
            size=("$rows"+"x"+"$columns"); enabled=$enabled; backend=$backend
            copy_mode=$copied; copy_mode_after_move=$copyAfterMove
            cache_before=$before;cache_after_move=$afterMove;cache_after_paste=$afterPaste;cache_after_resume=$afterResume
            values_ok=$valuesOk;format_ok=$formatOk;undo_available=$undoAvailable;undo_restored=$undoRestored
            owned_rules_after_paste=$rulesAfterPaste;owned_rules_after_off=$residual
            cf_followed_destination=($afterMove.TOP -eq 12 -and $afterMove.BOTTOM -eq 11+$rows -and $afterMove.LEFT -eq 11 -and $afterMove.RIGHT -eq 10+$columns)
            cf_suspended_for_copy=($afterMove.TOP -eq $before.TOP -and $afterMove.BOTTOM -eq $before.BOTTOM -and $afterMove.LEFT -eq $before.LEFT -and $afterMove.RIGHT -eq $before.RIGHT)
            destination_row_color=$destinationRowColor;old_row_color=$oldRowColor
            cf_display_followed_destination=($destinationRowColor -ne 16777215 -and $oldRowColor -eq 16777215)
        })
        Write-Output ('CASE|'+$rows+'x'+$columns+'|'+$backend+'|values='+$valuesOk+'|undo='+$undoRestored)
        Release-Com $source;Release-Com $target;$source=$null;$target=$null
        $book.Close($false);Release-Com $sheet;Release-Com $book;$sheet=$null;$book=$null
    }
} catch { $failure=$_.Exception.ToString()+"; at "+$_.ScriptStackTrace }
finally {
    if($null -ne $excel -and $null -ne $addin){try{$excel.Run($macro+'NxFocusDisable')|Out-Null}catch{}}
    foreach($item in @($book,$addin,$hostBook)){if($null -ne $item){try{$item.Close($false)}catch{}}}
    Release-Com $sheet;Release-Com $book;Release-Com $addin;Release-Com $hostBook;Release-Com $books
    if($null -ne $excel){try{$excel.Quit()}catch{};Release-Com $excel}
    $sheet=$null;$book=$null;$addin=$null;$hostBook=$null;$books=$null;$excel=$null;$item=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null -ne $binding){$cleanup=Stop-ExactProcessAfterGrace $binding.process 'Focus Copy Paste' 15000 -Detailed;$binding.process.Dispose()}
    $env:LHEXCEL_PROFILE_ROOT=$oldProfile
}
$settled=[Diagnostics.Stopwatch]::StartNew()
do {
    $remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    if($remaining.Count -eq 0){break}
    Start-Sleep -Milliseconds 100
} while($settled.ElapsedMilliseconds -lt 2000)
$receipt=[ordered]@{
    suite='FocusCopyPaste';evidence_class='native Excel command execution; CF cache diagnostic, not independent pixel geometry'
    source_sha256=$originalHash;source_unchanged=($originalHash -eq (Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash)
    cases=@($cases.ToArray());failure=$failure;cleanup=$cleanup
    remaining_excel=$remaining
    memory_cache_experiment=[bool]$MemoryCacheProbe
    no_calculate_experiment=[bool]$NoCalculateProbe
    repaint_experiment=[bool]$RepaintProbe
    accepted_clipboard_suspension=[bool]$AcceptClipboardSuspension
    last_stage=$stage;last_probe_state=$probeState
}
[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'FocusCopyPaste.json'),($receipt|ConvertTo-Json -Depth 12),(New-Object Text.UTF8Encoding($false)))
if($null -ne $failure){throw $failure}
if(-not $receipt.source_unchanged){throw 'Source product changed during copy/paste verification'}
if($cases.Count -ne 4 -or @($cases|Where-Object {-not $_.values_ok -or -not $_.format_ok -or -not $_.undo_restored -or $_.copy_mode_after_move -ne 1 -or $_.owned_rules_after_off -ne 0}).Count){throw 'Copy/paste safety regression'}
if($AcceptClipboardSuspension){
    if(@($cases|Where-Object {$_.backend -like 'cf*' -and -not $_.cf_suspended_for_copy}).Count){throw 'CF copy suspension contract failed'}
} else {
    if(@($cases|Where-Object {$_.backend -like 'cf*' -and -not $_.cf_followed_destination}).Count){throw 'CF focus did not follow destination during clipboard phase'}
    if(@($cases|Where-Object {$_.backend -like 'cf*' -and -not $_.cf_display_followed_destination}).Count){throw 'CF displayed row did not follow destination'}
}
if($cleanup.exit_mode -ne 'NATURAL' -or $receipt.remaining_excel.Count){throw 'Excel cleanup failed'}
Write-Output 'PASS|FocusCopyPasteSafety|4/4|physical-highlight-verification-separate'
