param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [ValidateSet('baseline','duplicate','visible','combined','all','product','paint-batch','name-batch')][string]$Variant='product',
    [ValidateRange(20,160)][int]$MoveCount=160,
    [ValidateRange(10,80)][int]$DuplicateCount=80,
    [switch]$PauseForVisual
)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$SourceRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Existing evidence preserved'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close user Excel before probe'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$beforeHash=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash
$copy=Join-Path $EvidenceRoot 'Product.xlam'
Copy-Item -LiteralPath $ProductXlam -Destination $copy
$previousProfile=$env:LHEXCEL_PROFILE_ROOT
$env:LHEXCEL_PROFILE_ROOT=Join-Path $EvidenceRoot 'profile'
function Release-Com([object]$value){
    if($null-ne$value-and[Runtime.InteropServices.Marshal]::IsComObject($value)){
        try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}catch{}
    }
}
function Stats([double[]]$values){
    $sorted=@($values|Sort-Object)
    return @{count=$sorted.Count;median_ms=$sorted[[int][math]::Floor($sorted.Count/2)];p95_ms=$sorted[[int][math]::Ceiling($sorted.Count*.95)-1]}
}
$excel=$null;$books=$null;$addin=$null;$book=$null;$sheet=$null;$window=$null;$binding=$null;$cleanup=$null
$failure='';$stage='start';$measurements=New-Object 'Collections.Generic.List[object]'
$checks=[ordered]@{};$sourceHash='';$excelVersion='';$excelBuild='';$offscreenRecalculated=$null
function Check([string]$name,[bool]$passed){$checks[$name]=$passed;if(-not$passed){throw ('Check failed: '+$name)}}
function CellColor([string]$address){$r=$sheet.Range($address);try{return [long]$r.DisplayFormat.Interior.Color}finally{Release-Com $r}}
function Select-Range([string]$address){$r=$sheet.Range($address);try{[void]$r.Select()}finally{Release-Com $r}}
try{
    $baseline=Get-ExcelProcessBaseline
    $startupFixture=Join-Path $EvidenceRoot 'startup-fixture.xlsx'
    Copy-Item -LiteralPath (Join-Path $SourceRoot 'tests/fixtures/r60-document-navigator/simple.xlsx') -Destination $startupFixture
    $interactive=Start-ExactInteractiveExcel $baseline 'focus speed Excel' -TimeoutMs 120000 -StartupWorkbook $startupFixture
    $excel=$interactive.excel;$binding=$interactive.binding;$interactive=$null
    if(-not$binding.owned){throw 'Excel ownership not established'}
    $excel.Visible=$true;$excel.DisplayAlerts=$false
    $startupBook=$excel.Workbooks.Item('startup-fixture.xlsx')
    try {
        if([IO.Path]::GetFullPath($startupBook.FullName) -ne [IO.Path]::GetFullPath($startupFixture)){throw 'Startup fixture identity mismatch'}
        $startupBook.Close($false)
    } finally {Release-Com $startupBook}
    $excelVersion=[string]$excel.Version;$excelBuild=[string]$excel.Build
    $books=$excel.Workbooks;$book=$books.Add();$addin=$books.Open($copy,$false,$true)
    $macro="'"+$addin.Name.Replace("'","''")+"'!"
    if($Variant -ne 'product'){
        $project=$addin.VBProject;$components=$project.VBComponents;$component=$components.Add(1);$component.Name='NxSpeedProbe';$module=$component.CodeModule
        $module.AddFromString(@'
Option Explicit
Public NxSpeedErrors As String
Public NxSpeedStage As String
Public Function NxSpeedToggleSafe() As String
    On Error GoTo Failed
    NxRibbonExecuteTag "nx1|feature|NX-DATA-FOCUS-CELL"
    NxSpeedToggleSafe = "PASS"
    Exit Function
Failed:
    NxSpeedToggleSafe = CStr(Err.Number) & "|" & Err.Description & "|" & NxSpeedErrors & "|" & NxSpeedStage
End Function
'@)
        Release-Com $module;Release-Com $component;Release-Com $components;Release-Com $project
    }
    if($Variant -eq 'name-batch'){
        # One coordinate-vector mutation per selection instead of four name
        # mutations. Only this disposable test copy is instrumented.
        $project=$addin.VBProject;$components=$project.VBComponents
        $component=$components.Item('CNxFocusRuleEngine');$module=$component.CodeModule
        $code=$module.Lines(1,$module.CountOfLines)
        $start=$code.IndexOf('Private Function EnsureAddinCacheNames()')
        $end=$code.IndexOf('End Function',$start)+'End Function'.Length
        $body=$code.Substring($start,$end-$start)
        if(([regex]::Matches($body,'RefersTo:="=1"')).Count -ne 4){throw 'Coordinate cache probe anchors missing'}
        $body=$body.Replace('    hasTop = HasAddinName', '    If Not HasAddinName("NX_FOCUS_BOUNDS_V3") Then ThisWorkbook.Names.Add Name:="NX_FOCUS_BOUNDS_V3", RefersTo:="={1,1,1,1}", Visible:=False'+[Environment]::NewLine+'    hasTop = HasAddinName')
        foreach($axis in @(@('TOP',1),@('BOTTOM',2),@('LEFT',3),@('RIGHT',4))){
            $body=$body.Replace(('Name:=NX_FOCUS_NAME_V002_'+$axis[0]+', RefersTo:="=1"'),('Name:=NX_FOCUS_NAME_V002_'+$axis[0]+', RefersTo:="=INDEX(NX_FOCUS_BOUNDS_V3,1,'+$axis[1]+')"'))
        }
        $code=$code.Substring(0,$start)+$body+$code.Substring($end)
        $start=$code.IndexOf('Private Sub UpdateAddinCache(')
        $end=$code.IndexOf('End Sub',$start)+'End Sub'.Length
        $body=@'
Private Sub UpdateAddinCache(ByVal selection As CNxFocusSelection)
    ThisWorkbook.Names("NX_FOCUS_BOUNDS_V3").RefersTo = "={" & CStr(selection.TopRow) & "," & CStr(selection.BottomRow) & "," & CStr(selection.LeftColumn) & "," & CStr(selection.RightColumn) & "}"
End Sub
'@
        $code=$code.Substring(0,$start)+$body+$code.Substring($end)
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($code)
        Release-Com $module;Release-Com $component
        $component=$components.Item('NxFocusRules');$module=$component.CodeModule
        $code=$module.Lines(1,$module.CountOfLines)
        $anchor='Right$(value, Len(NX_FOCUS_NAME_V002_RIGHT)) = NX_FOCUS_NAME_V002_RIGHT'
        $pattern=[regex]::Escape($anchor)
        if(-not [regex]::IsMatch($code,$pattern,[Text.RegularExpressions.RegexOptions]::IgnoreCase)){throw 'Coordinate cache cleanup anchor missing'}
        $code=[regex]::Replace($code,$pattern,[Text.RegularExpressions.MatchEvaluator]{param($match) $match.Value+' Or Right$(value, 18) = "NX_FOCUS_BOUNDS_V3"'},[Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($code)
        Release-Com $module;Release-Com $component;Release-Com $components;Release-Com $project
    }
    if($Variant -eq 'paint-batch'){
        # Experimental copy only: keep the product's own engine and suppress
        # intermediate name-change paints while retaining Worksheet.Calculate.
        $project=$addin.VBProject;$components=$project.VBComponents
        $component=$components.Item('CNxFocusRuleEngine');$module=$component.CodeModule
        $code=$module.Lines(1,$module.CountOfLines)
        $start=$code.IndexOf('Public Function ApplySelection(')
        $end=$code.IndexOf('End Function',$start)+'End Function'.Length
        $body=$code.Substring($start,$end-$start)
        foreach($anchor in @('    Dim sheet As Worksheet','    UpdateAddinCache selection','    sheet.Calculate','Failed:')){
            if(-not $body.Contains($anchor)){throw ('Paint probe anchor missing: '+$anchor)}
        }
        $body=$body.Replace('    Dim sheet As Worksheet','    Dim priorPaint As Boolean'+[Environment]::NewLine+'    Dim paintChanged As Boolean'+[Environment]::NewLine+'    Dim sheet As Worksheet')
        $body=$body.Replace('    UpdateAddinCache selection','    priorPaint = Application.ScreenUpdating'+[Environment]::NewLine+'    paintChanged = True'+[Environment]::NewLine+'    Application.ScreenUpdating = False'+[Environment]::NewLine+'    UpdateAddinCache selection')
        $body=$body.Replace('    sheet.Calculate','    sheet.Calculate'+[Environment]::NewLine+'    Application.ScreenUpdating = priorPaint'+[Environment]::NewLine+'    paintChanged = False')
        $body=$body.Replace('Failed:','Failed:'+[Environment]::NewLine+'    On Error Resume Next'+[Environment]::NewLine+'    If paintChanged Then Application.ScreenUpdating = priorPaint')
        $code=$code.Substring(0,$start)+$body+$code.Substring($end)
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($code)
        Release-Com $module;Release-Com $component;Release-Com $components;Release-Com $project
    }
    if($Variant -in @('duplicate','visible','combined','all')){
        $path=Join-Path $SourceRoot 'src/vba/features/data/focus/CNxFocusRuleEngine.cls'
        $sourceHash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $code=Get-Content -LiteralPath $path -Raw -Encoding UTF8
        $code=$code.Substring($code.IndexOf('Option Explicit'))
        foreach($line in @('        rulesValid = RulesMatchSettings(sheet, selection.Target, verifiedRange)','        RememberInspection sheet, verifiedRange','    Set repaint = VisibleRefreshRange(sheet, selection.Target)','    refreshKey = CStr(selection.TopRow)','    If mLastSheet Is sheet Then','    UpdateAddinCache selection','    For Each rule In sheet.Cells.FormatConditions','        If IsManagedFormula(formula) Then','            Set ManagedAppliesTo = rule.AppliesTo','    If Not Application.ActiveWindow Is Nothing Then','        If Application.ActiveSheet Is sheet Then','            For Each pane In Application.ActiveWindow.Panes','                Set visible = pane.VisibleRange','    If combined Is Nothing Then Set combined = fallback')){
            $code=$code.Replace($line,'    NxSpeedStage = "'+$line.Trim().Replace('"','""')+'"'+[Environment]::NewLine+$line)
        }
        $code=[regex]::Replace($code,'(?m)^Failed:', 'Failed:'+[Environment]::NewLine+'    NxSpeedErrors = NxSpeedErrors & "|" & CStr(Err.Number) & ":" & Err.Description')
        if($Variant -ne 'all'){$code=$code.Replace('If Not CanReuseInspection(sheet, selection.Target) Then','If True Then')}
        if($Variant -eq 'duplicate'){
            $anchor='    UpdateAddinCache selection'
            if(-not$code.Contains($anchor)){throw 'Calculation probe anchor missing'}
            $code=$code.Replace($anchor,$anchor+[Environment]::NewLine+'    sheet.Calculate')
        }
        if($Variant -eq 'visible'){
            $pattern='(?ms)^    If mLastSheet Is sheet Then\r?\n.*?^    End If\r?\n'
            if(-not[regex]::IsMatch($code,$pattern)){throw 'Duplicate probe block missing'}
            $code=[regex]::Replace($code,$pattern,'',1)
        }
        $project=$addin.VBProject;$components=$project.VBComponents;$component=$components.Item('CNxFocusRuleEngine');$module=$component.CodeModule
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($code)
        Release-Com $module;Release-Com $component
        $controller=Get-Content -LiteralPath (Join-Path $SourceRoot 'src/vba/features/data/focus/NxFocusController.bas') -Raw -Encoding UTF8
        $controller=$controller.Substring($controller.IndexOf('Option Explicit'))
        $component=$components.Item('NxFocusController');$module=$component.CodeModule
        $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($controller)
        Release-Com $module;Release-Com $component;Release-Com $components;Release-Com $project
    }
    $sheet=$book.Worksheets.Item(1);$sheet.Name='FocusSpeed';$book.Activate();$sheet.Activate()
    $window=$excel.ActiveWindow;$window.Zoom=100
    $excel.Calculation=-4135
    $fixture=$sheet.Range('J1:J3');$fixture.Value2=2
    $rule=$fixture.FormatConditions.AddColorScale(2);Release-Com $rule
    $rule=$fixture.FormatConditions.AddDatabar();Release-Com $rule
    $rule=$fixture.FormatConditions.AddIconSetCondition();Release-Com $rule;Release-Com $fixture
    $addresses=@('B2:C3','F6:G7','D4:E5','H8:I9','C12:D13','K15:L16')
    foreach($scenario in @('empty','90000_formulas','90000_automatic')){
        $stage=$scenario
        Write-Output ('STAGE|'+$scenario)
        if($scenario -eq '90000_formulas'){
            $grid=$sheet.Range('A1:AX1800');$grid.Formula='=ROW()+COLUMN()+NOW()*0';Release-Com $grid
        }
        if($scenario -eq '90000_automatic'){$excel.Calculation=-4105}
        $sheet.Calculate();Select-Range 'B2:C3'
        if($Variant -eq 'product'){[void]$excel.Run($macro+'NxRibbonExecuteTag','nx1|feature|NX-DATA-FOCUS-CELL')}
        else{$startResult=[string]$excel.Run($macro+'NxSpeedToggleSafe');if($startResult -ne 'PASS'){throw $startResult}}
        Check ($scenario+'_cf_backend') ([string]$excel.Run($macro+'NxFocusControllerBackend') -eq 'cf')
        $book.Saved=$true
        $moves=New-Object 'Collections.Generic.List[double]'
        for($i=0;$i-lt$MoveCount;$i++){
            $r=$sheet.Range($addresses[$i%$addresses.Count]);$watch=[Diagnostics.Stopwatch]::StartNew();[void]$r.Select();$watch.Stop();Release-Com $r
            if($i-ge10){$moves.Add($watch.Elapsed.TotalMilliseconds)}
        }
        Select-Range 'F6:G7'
        if($PauseForVisual -and $scenario -eq '90000_automatic'){
            Write-Output 'VISUAL_READY|F6:G7'
            Start-Sleep -Seconds 30
            Select-Range 'K15:L16'
            Write-Output 'VISUAL_READY|K15:L16'
            Start-Sleep -Seconds 30
            Select-Range 'F6:G7'
        }
        Check ($scenario+'_current_row') ((CellColor 'A6') -ne 16777215)
        Check ($scenario+'_current_column') ((CellColor 'F1') -ne 16777215)
        Check ($scenario+'_old_row_cleared') ((CellColor 'A2') -eq 16777215)
        $duplicates=New-Object 'Collections.Generic.List[double]'
        for($i=0;$i-lt$DuplicateCount;$i++){
            $watch=[Diagnostics.Stopwatch]::StartNew();[void]$excel.Run($macro+'NxFocusControllerRefreshActiveSelection',$true);$watch.Stop()
            if($i-ge5){$duplicates.Add($watch.Elapsed.TotalMilliseconds)}
        }
        Check ($scenario+'_saved_preserved') ([bool]$book.Saved)
        $measurements.Add(@{scenario=$scenario;moves=(Stats $moves.ToArray());duplicates=(Stats $duplicates.ToArray())})
        Write-Output ('MEASURED|'+$scenario+'|move_p95='+[string](Stats $moves.ToArray()).p95_ms)
        $measurements.ToArray()|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $EvidenceRoot 'measurements.partial.json') -Encoding UTF8
        [void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)
    }
    $stage='200_user_rules'
    Write-Output 'STAGE|200_user_rules'
    $fixtureCalculation=$excel.Calculation;$fixturePaint=$excel.ScreenUpdating
    try {
        $excel.Calculation=-4135;$excel.ScreenUpdating=$false
        for($i=1;$i-le200;$i++){
            $r=$sheet.Cells.Item($i,53)
            $rule=$r.FormatConditions.Add(2,$null,'=MOD(ROW()+COLUMN(),3)=0')
            Release-Com $rule;Release-Com $r
        }
    } finally {
        $excel.Calculation=$fixtureCalculation;$excel.ScreenUpdating=$fixturePaint
    }
    Write-Output 'READY|200_user_rules'
    $userRuleCountBefore=[int]$sheet.Cells.FormatConditions.Count
    Select-Range 'B2:C3'
    if($Variant -eq 'product'){[void]$excel.Run($macro+'NxRibbonExecuteTag','nx1|feature|NX-DATA-FOCUS-CELL')}
    else{$startResult=[string]$excel.Run($macro+'NxSpeedToggleSafe');if($startResult -ne 'PASS'){throw $startResult}}
    $book.Saved=$true
    $moves=New-Object 'Collections.Generic.List[double]'
    for($i=0;$i-lt$MoveCount;$i++){
        $r=$sheet.Range($addresses[$i%$addresses.Count]);$watch=[Diagnostics.Stopwatch]::StartNew();[void]$r.Select();$watch.Stop();Release-Com $r
        if($i-ge10){$moves.Add($watch.Elapsed.TotalMilliseconds)}
    }
    Check '200_user_rules_saved_preserved' ([bool]$book.Saved)
    [void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)
    $userRuleCountAfter=[int]$sheet.Cells.FormatConditions.Count
    Check '200_user_rules_preserved' ($userRuleCountAfter-eq$userRuleCountBefore)
    $measurements.Add(@{scenario='200_user_rules';moves=(Stats $moves.ToArray());user_rules_before=$userRuleCountBefore;user_rules_after=$userRuleCountAfter})
    Write-Output ('MEASURED|200_user_rules|move_p95='+[string](Stats $moves.ToArray()).p95_ms)
    $excel.Calculation=-4135
    $stage='calculation_scope'
    $sentinel=$sheet.Range('AZ5000');$sentinel.Formula='=NOW()';$sentinel.Calculate();$sentinelBefore=[double]$sentinel.Value2
    $scopeSheet=$book.Worksheets.Add();$scopeSheet.Name='CalculationScope'
    $otherSentinel=$scopeSheet.Range('A1');$otherSentinel.Formula='=NOW()';$otherSentinel.Calculate();$otherBefore=[double]$otherSentinel.Value2
    $sheet.Activate()
    Start-Sleep -Milliseconds 1100
    Select-Range 'B2:C3';[void]$excel.Run($macro+'NxFocusControllerSetEnabled',$true);Select-Range 'F6:G7'
    $offscreenRecalculated=([double]$sentinel.Value2-ne$sentinelBefore);Release-Com $sentinel
    if($Variant -eq 'product'){
        Check 'active_sheet_formula_recalculated_for_grid_refresh' $offscreenRecalculated
        Check 'other_sheet_formula_not_recalculated' ([double]$otherSentinel.Value2-eq$otherBefore)
        Check 'manual_calculation_mode_preserved' ([int]$excel.Calculation-eq-4135)
    }
    Release-Com $otherSentinel;Release-Com $scopeSheet
    $stage='settings'
    # Excel may reset ScreenUpdating when an external automation call returns.
    # Observe the nested caller's state within one VBA call instead.
    $project=$addin.VBProject;$components=$project.VBComponents
    $component=$components.Add(1);$component.Name='NxPaintStateProbe';$module=$component.CodeModule
    $module.AddFromString(@'
Option Explicit
Public Function NxPaintStatePreserved() As Boolean
    Dim prior As Boolean
    prior = Application.ScreenUpdating
    On Error GoTo Failed
    Application.ScreenUpdating = False
    ActiveSheet.Range("D4:E5").Select
    NxPaintStatePreserved = Not Application.ScreenUpdating
Done:
    Application.ScreenUpdating = prior
    Exit Function
Failed:
    NxPaintStatePreserved = False
    Resume Done
End Function
'@)
    Release-Com $module;Release-Com $component;Release-Com $components;Release-Com $project
    Check 'caller_disabled_paint_preserved' ([bool]$excel.Run($macro+'NxPaintStatePreserved'))
    Select-Range 'F6:G7'
    Check 'caller_enabled_paint_preserved' ([bool]$excel.ScreenUpdating)
    $oldColor=CellColor 'A6'
    $settings=[string]$excel.Run($macro+'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#5FC8D8',70,$false,$true)
    Check 'same_selection_settings_apply' ($settings-eq'PASS'-and(CellColor 'A6')-ne$oldColor)
    $stage='rule_repair'
    [void]$excel.Run($macro+'NxFocusDeleteManagedRules',$sheet)
    Start-Sleep -Milliseconds 300
    Select-Range 'B2:C3'
    Check 'externally_removed_rules_repaired' ([int]$excel.Run($macro+'NxFocusManagedRuleCount',$sheet)-eq2)
    Check 'repaired_row_visible' ((CellColor 'A2')-ne16777215)
    foreach($zoom in @(75,125,100)){
        $window.Zoom=$zoom;Select-Range 'F6:G7'
        [void]$excel.Run($macro+'NxFocusControllerRefreshActiveSelection',$false)
        Check ('zoom_'+$zoom) ((CellColor 'A6')-ne16777215)
    }
    $stage='freeze_scroll'
    $window.SplitRow=2;$window.SplitColumn=2;$window.FreezePanes=$true
    Select-Range 'F6:G7';[void]$excel.Run($macro+'NxFocusControllerRefreshActiveSelection',$false)
    Check 'frozen_column_painted' ((CellColor 'F1')-ne16777215)
    $window.ScrollRow=100;$window.ScrollColumn=6;Select-Range 'H105:I106'
    Check 'scrolled_row_painted' ((CellColor 'F105')-ne16777215)
    $window.FreezePanes=$false;$window.SplitRow=0;$window.SplitColumn=0;$window.ScrollRow=1;$window.ScrollColumn=1
    $window.SplitRow=8;$window.SplitColumn=4;Select-Range 'F6:G7'
    [void]$excel.Run($macro+'NxFocusControllerRefreshActiveSelection',$false)
    Check 'split_painted' ((CellColor 'A6')-ne16777215)
    $window.SplitRow=0;$window.SplitColumn=0
    $stage='sheet_switch'
    $other=$book.Worksheets.Add();$other.Name='Other';$r=$other.Range('B2:C3');$r.Select();Release-Com $r
    $sheet.Activate();Select-Range 'F6:G7'
    Check 'sheet_return_painted' ((CellColor 'A6')-ne16777215);Release-Com $other
    $stage='save'
    $book.SaveAs((Join-Path $EvidenceRoot 'FocusFixture.xlsx'),51)
    Check 'save_restores_focus' ((CellColor 'A6')-ne16777215)
    $stage='undo'
    $r=$sheet.Range('Z205');$r.Value2='FOCUS_UNDO';$r.Select()
    $excel.CommandBars.ExecuteMso('ClearContents');Start-Sleep -Milliseconds 100
    Check 'undo_created' ([bool]$excel.CommandBars.GetEnabledMso('Undo'))
    Select-Range 'Z206';Start-Sleep -Milliseconds 150
    Check 'undo_retained' ([bool]$excel.CommandBars.GetEnabledMso('Undo'))
    $excel.Undo();Check 'undo_value_restored' ([string]$r.Value2-eq'FOCUS_UNDO');Release-Com $r
    $stage='disable'
    [void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)
    Check 'zero_focus_rules' ([int]$excel.Run($macro+'NxFocusManagedRuleCount',$sheet)-eq0)
    $r=$sheet.Range('J1:J3');Check 'user_three_rules_preserved' ([int]$r.FormatConditions.Count-eq3);Release-Com $r
    $stage='disabled_events_start'
    Select-Range 'A1'
    $excel.EnableEvents=$false
    [void]$excel.Run($macro+'NxFocusControllerSetEnabled',$true)
    Select-Range 'F6:G7'
    Check 'disabled_events_start_tracks_selection' ((CellColor 'J6')-ne16777215 -and (CellColor 'A8')-eq16777215)
    Check 'disabled_events_start_recovers_events' ([bool]$excel.EnableEvents)
    $excel.EnableEvents=$false
    Select-Range 'A1'
    [void]$excel.Run($macro+'NxFocusControllerSetEnabled',$true)
    Select-Range 'D4:E5'
    Check 'already_enabled_recovers_tracking' ([bool]$excel.EnableEvents -and (CellColor 'J4')-ne16777215 -and (CellColor 'J6')-eq16777215)
    [void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)
    Check 'recovered_start_zero_focus_rules' ([int]$excel.Run($macro+'NxFocusManagedRuleCount',$sheet)-eq0)
    $stage='complete'
}catch{$failure=$_.Exception.Message}finally{
    if($null-ne$excel){try{[void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)}catch{}}
    if($null-ne$addin){try{$addin.Close($false)}catch{}}
    if($null-ne$book){try{$book.Close($false)}catch{}}
    Release-Com $window;Release-Com $sheet;Release-Com $addin;Release-Com $book;Release-Com $books
    if($null-ne$excel){try{$excel.Quit()}catch{};Release-Com $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null-ne$binding){$cleanup=Stop-ExactProcessAfterGrace $binding.process 'focus speed Excel' 15000 -Detailed;$binding.process.Dispose()}
    $env:LHEXCEL_PROFILE_ROOT=$previousProfile
}
$afterHash=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash
$remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
for($i=0;$i-lt20-and$remaining.Count-gt0;$i++){Start-Sleep -Milliseconds 100;$remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue)}
$passed=([string]::IsNullOrEmpty($failure) -and $cleanup.exit_mode -eq 'NATURAL' -and $beforeHash -eq $afterHash -and $remaining.Count -eq 0)
$receipt=@{status=$(if($passed){'PASS'}else{'FAIL'});variant=$Variant;stage=$stage;failure=$failure;checks=$checks;measurements=$measurements.ToArray();product_sha256=$beforeHash;product_sha256_after=$afterHash;remaining_pids=@($remaining|ForEach-Object{$_.Id});candidate_source_sha256=$sourceHash;excel_version=$excelVersion;excel_build=$excelBuild;offscreen_recalculated=$offscreenRecalculated;cleanup=$cleanup}
$receipt|ConvertTo-Json -Depth 12|Set-Content -LiteralPath (Join-Path $EvidenceRoot 'focus-speed.json') -Encoding UTF8
Write-Output ($receipt|ConvertTo-Json -Depth 7)
if(-not$passed){exit 1}
