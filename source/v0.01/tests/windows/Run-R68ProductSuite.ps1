param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$SourceRoot,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$Suites = 'performance',
    [string]$CaseName = '',
    [switch]$ObserveBaseline,
    [switch]$OverridePerformance
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$ProductXlam=[IO.Path]::GetFullPath($ProductXlam).Replace('/', '\')
$SourceRoot=[IO.Path]::GetFullPath($SourceRoot).Replace('/', '\')
$EvidenceRoot=[IO.Path]::GetFullPath($EvidenceRoot).Replace('/', '\')
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
function Release-ComObject([object]$Value) {
    if($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)}catch{}
    }
}
function Wait-ProbeExcelReady([object]$Application) {
    $watch=[Diagnostics.Stopwatch]::StartNew()
    do {
        try {if([bool]$Application.Ready){return}} catch {}
        Start-Sleep -Milliseconds 100
    } while($watch.ElapsedMilliseconds -lt 15000)
    throw 'Owned test Excel did not become ready'
}
$map=@{
    'product-about'='tests/vba/product/T_ProductAbout.bas'
    'unit-symbols'='tests/vba/product/T_R102Units.bas'
    'picture-titles'='tests/vba/product/T_R103PictureTitles.bas'
    'consolidate-list'='tests/vba/product/T_ConsolidateList.bas'
    'compare-consolidate-links'='tests/vba/product/T_CompareConsolidateLinks.bas'
    'r100-focus-preview'='tests/vba/product/T_R100FocusPreview.bas'
    'r97-preview'='tests/vba/product/T_R97Preview.bas'
    'r96-ribbon'='tests/vba/product/T_R96Ribbon.bas'
    'r94-compare-output'='tests/vba/product/T_R94CompareOutput.bas'
    'r94-formulatools'='tests/vba/product/T_R94FormulaTools.bas'
    'r94-functions'='tests/vba/product/T_R94Functions.bas'
    'r94-template'='tests/vba/product/T_R94Template.bas'
    'r93-output'='tests/vba/product/T_R93Outputs.bas'
    'r93-compare'='tests/vba/product/T_R93Compare.bas'
    'r86-policy'='tests/vba/product/T_R76Policy.bas'
    'r86-data'='tests/vba/product/T_R86Data.bas'
    'r85-data'='tests/vba/product/T_R85Data.bas'
    'r80-table'='tests/vba/product/T_R80Table.bas'
    'r81-table'='tests/vba/product/T_R81Table.bas'
    'r82-file'='tests/vba/product/T_R82File.bas'
    'r83-files'='tests/vba/product/T_R83Files.bas'
    'r84-ai'='tests/vba/product/T_R84Ai.bas'
    'r79-table'='tests/vba/product/T_R79Table.bas'
    'r78-files'='tests/vba/product/T_R78Files.bas'
    'r77-normalization'='tests/vba/product/T_R77Normalization.bas'
    'r76-policy'='tests/vba/product/T_R76Policy.bas'
    'r76-format-backup'='tests/vba/product/T_R76FormatBackup.bas'
    'r75-guard'='tests/vba/product/T_R75Guard.bas'
    'r75-journal'='tests/vba/product/T_R75Journal.bas'
    'r74-folder'='tests/vba/product/T_R74Folder.bas'
    'r74-features'='tests/vba/product/T_R74Features.bas'
    'range-properties'='tests/vba/product/T_R72RangeProperties.bas'
    'stage-profile'='tests/vba/product/T_R70LargeRange.bas'
    'fast-snapshot'='tests/vba/product/T_R71FastSnapshot.bas'
    'serialization-probe'='tests/vba/product/T_R71Serialization.bas'
    'alignment'='tests/vba/product/T_R70Alignment.bas'
    'large-range'='tests/vba/product/T_R70LargeRange.bas'
    performance='tests/vba/product/T_R68PerformanceExport.bas'
    table='tests/vba/draw/T_R68Table.bas'
    shortcuts='tests/vba/product/T_R68Shortcuts.bas'
    quickformat='tests/vba/product/T_R68QuickFormat.bas'
    'privacy-flow'='tests/vba/product/T_R102PrivacyFlow.bas'
    visual='tests/vba/product/T_R69Visual.bas'
    benchmark='tests/vba/product/T_R69Benchmark.bas'
    'runner-benchmark'='tests/vba/product/T_R69RunnerBenchmark.bas'
    'runner-safety'='tests/vba/product/T_R69RunnerSafety.bas'
    'repaint-benchmark'='tests/vba/product/T_R69RepaintBenchmark.bas'
    'snapshot-probe'='tests/vba/product/T_R69SnapshotProbe.bas'
    'compare-internal'='tests/vba/product/T_R69CompareMigration.bas'
    'compare-enhanced'='tests/vba/product/T_R69CompareMigration.bas'
    userflows='tests/vba/product/T_R69UserFlows.bas'
    'popup-layout'='tests/vba/product/T_R69PopupLayout.bas'
    'output-flows'='tests/vba/product/T_R69OutputFlows.bas'
}
$selected=@($Suites.Split(','))
if(@($selected|Where-Object {-not $map.ContainsKey($_)}).Count){throw 'Unknown R68 suite'}
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Fresh evidence required'}
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Existing Excel preserved'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$env:LHEXCEL_PROFILE_ROOT=$EvidenceRoot
$excel=$null;$binding=$null;$books=$null;$product=$null;$fixture=$null;$component=$null;$control=$null;$failure=$null
$report=[ordered]@{status='RUNNING';product_sha256=(Get-FileHash -LiteralPath $ProductXlam).Hash;source_hashes=@{};cases=@();images=@();evidence_class='native_excel_instrumented_product';baseline=[bool]$ObserveBaseline}
function Add-Probe([string]$Name,[string]$Text,[int]$Kind=1){
    $component=$null
    try {
        $component=$product.VBProject.VBComponents.Add($Kind)
        $component.Name=$Name
        $start=$Text.IndexOf('Option Explicit')
        if($start -lt 0){throw 'Probe requires Option Explicit'}
        $component.CodeModule.AddFromString($Text.Substring($start).Replace("`r`n","`n").Replace("`n","`r`n"))
    } finally {Release-ComObject $component}
}
try {
    $startupFixture=Join-Path $EvidenceRoot 'startup.xlsx'
    Copy-Item -LiteralPath (Join-Path $SourceRoot 'tests/fixtures/r60-document-navigator/simple.xlsx') -Destination $startupFixture
    $session=Start-ExactInteractiveExcel (Get-ExcelProcessBaseline) 'r68-product-probe' -ShowWindow -StartupWorkbook $startupFixture
    $excel=$session.excel;$binding=$session.binding
    Wait-ProbeExcelReady $excel
    $excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
    $books=$excel.Workbooks
    if($books.Count -eq 1){
        $startup=$books.Item(1)
        try {if([string]$startup.FullName -eq $startupFixture){$startup.Close($false)}}finally{Release-ComObject $startup}
    }
    # /x can open a pristine default workbook on this Office installation.
    # Only the freshly launched, owned process and an unmodified empty sheet qualify.
    if($books.Count -eq 1){
        $startup=$books.Item(1);$startupSheet=$null;$startupRange=$null
        try {
            if($startup.Path -eq '' -and $startup.Saved -and $startup.Sheets.Count -eq 1 -and $startup.Worksheets.Count -eq 1 -and $startup.Names.Count -eq 0){
                $startupSheet=$startup.Worksheets.Item(1);$startupRange=$startupSheet.UsedRange
                if($startupRange.Address() -eq '$A$1' -and $null -eq $startupRange.Value2 -and -not $startupRange.HasFormula -and $startupSheet.Shapes.Count -eq 0 -and $startupSheet.Comments.Count -eq 0){
                    $startup.Close($false)
                    $report.startup_empty_workbook_closed=$true
                }
            }
        }finally{Release-ComObject $startupRange;Release-ComObject $startupSheet;Release-ComObject $startup}
    }
    if($books.Count -ne 0){throw 'Unexpected startup workbook'}
    $copy=Join-Path $EvidenceRoot 'r68-probe-product.xlam'
    if(($selected -contains 'compare-internal') -or ($selected -contains 'compare-enhanced')){
        $profile=if($selected -contains 'compare-enhanced'){'enhanced-dll'}else{'internal-xlam'}
        & python -X utf8 (Join-Path $SourceRoot 'tools/distribution_source_profile.py') --source $ProductXlam --destination $copy --profile $profile
        if($LASTEXITCODE -ne 0){throw 'Distribution source profile transformation failed'}
    }else{Copy-Item -LiteralPath $ProductXlam -Destination $copy}
    $product=$books.Open($copy,0,$false)
    Wait-ProbeExcelReady $excel
    if($selected -contains 'r96-ribbon'){
        Add-Probe 'T_R96Control' ([IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/product/T_R96Control.cls'),[Text.Encoding]::UTF8)) 2
    }
    if($selected -contains 'r94-template'){
        $env:LHEXCEL_TEMPLATE_TEST_ROOT=Join-Path $env:TEMP ('nx94-ui-'+[Guid]::NewGuid().ToString('N').Substring(0,12))
        $component=$product.VBProject.VBComponents.Item('NxTemplatePaths');$module=$component.CodeModule
        try{
            $text=$module.Lines(1,$module.CountOfLines)
            $pathSource=[IO.File]::ReadAllText((Join-Path $SourceRoot 'src/vba/features/template/NxTemplatePaths.bas'),[Text.Encoding]::UTF8)
            $sites=[regex]::Matches($pathSource,'(?m)^    NxTemplateStoreRoot = ThisWorkbook.Path & [^\r\n]+')
            if($sites.Count -ne 1){throw 'Template root source site changed'}
            $site=$sites[0].Value
            $sitePattern=[regex]::Escape($site)
            if(-not [regex]::IsMatch($text,$sitePattern,[Text.RegularExpressions.RegexOptions]::IgnoreCase)){throw 'Template root probe site changed'}
            $replacement='    NxTemplateStoreRoot = "'+$env:LHEXCEL_TEMPLATE_TEST_ROOT+'"'
            $text=[regex]::Replace($text,$sitePattern,$replacement,[Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text)
        }finally{Release-ComObject $module;Release-ComObject $component}
        $component=$product.VBProject.VBComponents.Item('FNxTemplateRegister');$module=$component.CodeModule
        try{
            $text=$module.Lines(1,$module.CountOfLines)
            $text=$text.Replace('    MsgBox detail, vbExclamation, Me.Caption','    Err.Raise 5, , detail')
            $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text)
            $module.AddFromString("Public Sub R94ApplyDirect()`r`n    chkConfirmRisk.Value = True`r`n    cmdApply_Click`r`nEnd Sub")
        }finally{Release-ComObject $module;Release-ComObject $component}
    }
    if($selected -contains 'r83-files'){
        Add-Probe 'T_R83Picker' ([IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/product/T_R83Picker.cls'))) 2
        foreach($name in @('NxCopySave','NxFileController','NxHostBridge')){
            $component=$product.VBProject.VBComponents.Item($name);$module=$component.CodeModule
            try {
                $text=$module.Lines(1,$module.CountOfLines)
                if($name -eq 'NxCopySave'){
                    $sites=@('    sourceRange.Copy Destination:=outputBook.Worksheets(1).Range("A1")',
                        '    NxCopyRangeDimensions sourceRange, outputBook.Worksheets(1).Range("A1"), False',
                        '    outputBook.SaveAs Filename:=tempPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False',
                        '    outputBook.Close SaveChanges:=False')
                    for($i=0;$i -lt $sites.Count;$i++){
                        $match=[regex]::Match($text,[regex]::Escape($sites[$i]),'IgnoreCase')
                        if(-not $match.Success){[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'copy-probe-source.txt'),$text);throw "Copy timing site changed $i"}
                        $replacement=$match.Value+ "`r`n    T_R83Files.Mark ""stage-$i"""
                        if($i -eq 0){$replacement="    T_R83Files.Mark ""copy-start""`r`n"+$replacement}
                        $text=$text.Replace($match.Value,$replacement)
                    }
                    $text=$text.Replace('    NxFileCommitTemporaryOutput tempPath, outputPath',"    T_R83Files.AfterSave`r`n    T_R83Files.BeforeCommit tempPath`r`n    NxFileCommitTemporaryOutput tempPath, outputPath`r`n    T_R83Files.Mark ""commit""")
                    $text=$text.Replace('    NxFileDeleteCreatedFile tempPath, True',"    T_R83Files.ReleaseLock`r`n    NxFileDeleteCreatedFile tempPath, True")
                }elseif($name -eq 'NxFileController'){
                    $text=$text.Replace('        Set picker = Application.FileDialog(4)','        Set picker = New T_R83Picker')
                }else{
                    $text=$text.Replace('Public Function NxHostCreateWorkbookCompare() As Object',"Public Function NxHostCreateWorkbookCompare() As Object`r`n    Exit Function")
                }
                $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text)
            }finally{Release-ComObject $module;Release-ComObject $component}
        }
    }
    if($selected -contains 'r86-data'){
        foreach($moduleName in @('NxDataSpecialController','CNxDataSpecialCommand')){
            $component=$product.VBProject.VBComponents.Item($moduleName);$module=$component.CodeModule
            try{
                $text=$module.Lines(1,$module.CountOfLines)
                if($moduleName -eq 'NxDataSpecialController'){
                    $text=$text.Replace('    commandObject.ConfigureOptions featureId, source, output, options', "    T_R86Data.Mark ""validated""`r`n    commandObject.ConfigureOptions featureId, source, output, options`r`n    T_R86Data.Mark ""configured""")
                }else{
                    $text=$text.Replace('        resultValues = BuildTransformValues()', "        T_R86Data.Mark ""execution_enter""`r`n        resultValues = BuildTransformValues()`r`n        T_R86Data.Mark ""transformed""")
                    $text=$text.Replace('        sourceChanged = True', "        T_R86Data.Mark ""written""`r`n        sourceChanged = True")
                }
                if(-not $text.Contains('T_R86Data.Mark')){throw 'r86 timing hook missing'}
                $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text)
            }finally{Release-ComObject $module;Release-ComObject $component}
        }
    }
    if($selected -contains 'r84-ai'){
        $component=$product.VBProject.VBComponents.Item('FNxAi');$module=$component.CodeModule
        try{$module.AddFromString("Public Sub R84Execute()`r`n    cmdExecute_Click`r`nEnd Sub")}
        finally{Release-ComObject $module;Release-ComObject $component}
        $component=$product.VBProject.VBComponents.Item('NxAiUi');$module=$component.CodeModule
        try{
            $text=$module.Lines(1,$module.CountOfLines)
            $text=$text.Replace('Public Function NxAiShowExecutionPlan(ByVal session As CNxAiInputSession) As NxFrameState',"Public Function NxAiShowExecutionPlan(ByVal session As CNxAiInputSession) As NxFrameState`r`n    T_R84Ai.CopiedPrompt = session.PreviewPrompt()`r`n    NxAiShowExecutionPlan = NxFrameSuccess`r`n    Exit Function")
            $text=$text.Replace('Public Function NxAiOpenTargetAfterCopy(ByVal targetName As String, ByVal copyState As NxFrameState) As Boolean',"Public Function NxAiOpenTargetAfterCopy(ByVal targetName As String, ByVal copyState As NxFrameState) As Boolean`r`n    Exit Function")
            if(-not $text.Contains('T_R84Ai.CopiedPrompt = session.PreviewPrompt()')){throw 'r84 AI execution stub not installed'}
            if(-not $text.Contains("As Boolean`r`n    Exit Function")){throw 'r84 AI browser stub not installed'}
            $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text)
        }finally{Release-ComObject $module;Release-ComObject $component}
    }
    if($selected -contains 'r78-files'){
        $component=$product.VBProject.VBComponents.Item('NxWorkbookCompare');$module=$component.CodeModule
        try {
            $text=$module.Lines(1,$module.CountOfLines)
            $wrapper=[IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/product/R78CompareWrapper.txt'),[Text.Encoding]::UTF8)
            if($text.Contains('ByRef values As Variant, ByRef formulas As Variant')){
                $wrapper=$wrapper.Replace('leftRows = CompareCapture(leftRange, formats): rightRows = CompareCapture(rightRange, formats)',
                    'leftRows = CompareCapture(leftRange, formats, CompareGridValues(leftRange, False), CompareGridValues(leftRange, True)): rightRows = CompareCapture(rightRange, formats, CompareGridValues(rightRange, False), CompareGridValues(rightRange, True))')
            }
            if(-not $text.Contains('ByRef leftRows As Variant, ByRef rightRows As Variant')){
                $wrapper=$wrapper.Replace('leftRows = CompareCapture(leftRange, formats): rightRows = CompareCapture(rightRange, formats)','')
                $wrapper=$wrapper.Replace('NxWorkbookCompareCellDifference(leftRows, rightRows, i)','NxWorkbookCompareCellDifference(leftRange.Cells(row, column), rightRange.Cells(row, column), formats)')
            }
            $module.AddFromString($wrapper)
        }finally{Release-ComObject $module;Release-ComObject $component}
        $component=$product.VBProject.VBComponents.Item('NxHostBridge');$module=$component.CodeModule
        try{
            $text=$module.Lines(1,$module.CountOfLines)
            $signature='Public Function NxHostCreateWorkbookCompare() As Object'
            if(-not $text.Contains($signature)){throw 'Compare factory changed'}
            $module.DeleteLines(1,$module.CountOfLines)
            $module.AddFromString($text.Replace($signature,$signature+"`r`n    Exit Function"))
        }finally{Release-ComObject $module;Release-ComObject $component}
    }
    if($selected -contains 'r75-guard'){
        $component=$product.VBProject.VBComponents.Item('NxHostIntegrity');$module=$component.CodeModule
        try {$module.AddFromString("Public Sub NxHostProbeFile(ByVal path As String, ByVal expected As String)`r`n    NxHostRequireFileHash path, expected`r`nEnd Sub`r`nPublic Function NxHostProbeUrl(ByVal url As String) As String`r`n    NxHostProbeUrl = NxHostLocalCodeBase(url)`r`nEnd Function")}
        finally {Release-ComObject $module;Release-ComObject $component}
    }
    if($selected -contains 'r75-journal'){
        $component=$product.VBProject.VBComponents.Item('NxDrawController');$module=$component.CodeModule
        try {
            $text=$module.Lines(1,$module.CountOfLines)
            $needle='        NxDrawApplyTableBorders target, header, tableStyle, displayMode'
            if(-not $text.Contains($needle)){throw 'Drawing fault injection target changed'}
            $text=$text.Replace($needle,$needle+"`r`n"+'        If T_R75Journal.FailDraw Then T_R75Journal.DrawFaultObserved = True: Err.Raise 5, "R76Probe", "Injected drawing failure"')
            $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text)
        }finally{Release-ComObject $module;Release-ComObject $component}
    }
    if(($selected -contains 'r76-policy') -or ($selected -contains 'r86-policy')){
        if($selected.Count -ne 1){throw 'Policy probe requires an isolated suite'}
        $policyScript=if($selected -contains 'r86-policy'){'tests/windows/make_r86_policy_fixtures.py'}else{'tests/windows/make_r76_policy_fixtures.py'}
        & python -X utf8 (Join-Path $SourceRoot $policyScript) --out (Join-Path $EvidenceRoot 'policy-fixtures')
        if($LASTEXITCODE -ne 0){throw 'Policy fixtures failed'}
        Add-Probe 'T_R76Control' "Option Explicit`r`nPublic Tag As String" 2
        $gateComponent=$product.VBProject.VBComponents.Item('NxDistributionGate')
        $gateModule=$gateComponent.CodeModule
        try {$gateModule.AddFromString("Public Function NxDistributionProbeTimer() As Boolean`r`n    NxDistributionProbeTimer = mRefreshScheduled`r`nEnd Function`r`nPublic Function NxDistributionProbeDetail() As String`r`nEnd Function")}
        finally {Release-ComObject $gateModule;Release-ComObject $gateComponent}
    }
    if($selected -contains 'r76-format-backup'){
        if(-not (@($product.VBProject.VBComponents | ForEach-Object {$_.Name}) -contains 'CNxRangeFormatBackup')){
            Add-Probe 'CNxRangeFormatBackup' ([IO.File]::ReadAllText((Join-Path $SourceRoot 'src/vba/core/CNxRangeFormatBackup.cls'),[Text.Encoding]::UTF8)) 2
        }
    }
    if($selected -contains 'r74-features'){
        foreach($name in @('FNxNumberFormat','FNxViewPresets')){
            $component=$product.VBProject.VBComponents.Item($name);$module=$component.CodeModule
            try {$module.AddFromString("Public Sub NxProbeExecute()`r`n    cmdExecute_Click`r`nEnd Sub")}
            finally {Release-ComObject $module;Release-ComObject $component}
        }
    }
    if($selected -contains 'r94-formulatools'){
        $component=$product.VBProject.VBComponents.Item('FNxFormulaTools');$module=$component.CodeModule
        try {$module.AddFromString("Public Sub NxProbeExecute()`r`n    cmdExecute_Click`r`nEnd Sub")}
        finally {Release-ComObject $module;Release-ComObject $component}
    }
    if($selected -contains 'r94-compare-output'){
        $component=$product.VBProject.VBComponents.Item('FNxWorkbookCompare');$module=$component.CodeModule
        try {$module.AddFromString("Public Sub NxProbeExecute()`r`n    cmdExecute_Click`r`nEnd Sub`r`nPublic Sub NxProbeLoadSheets()`r`n    cmdLoadSheets_Click`r`nEnd Sub")}
        finally {Release-ComObject $module;Release-ComObject $component}
    }
    if($selected -contains 'r94-functions'){
        $component=$product.VBProject.VBComponents.Item('FNxFunctionWrap');$module=$component.CodeModule
        try {$module.AddFromString("Public Sub NxProbeExecute()`r`n    cmdExecute_Click`r`nEnd Sub")}
        finally {Release-ComObject $module;Release-ComObject $component}
        $component=$product.VBProject.VBComponents.Item('NxFunctionWrap');$module=$component.CodeModule
        try {
            $text=$module.Lines(1,$module.CountOfLines)
            $needle='        WriteFormula mCells(i), CStr(mAfter(i))'
            if(-not $text.Contains($needle)){throw 'Function fault target changed'}
            $text=$text.Replace($needle,$needle+"`r`n"+'        If T_R94Functions.FailAfterFirst Then T_R94Functions.FailAfterFirst = False: Err.Raise 5, "FunctionProbe", "Injected write failure"')
            $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text)
        }finally{Release-ComObject $module;Release-ComObject $component}
    }
    if($selected -contains 'privacy-flow'){
        # Cancellation injection is restricted to the disposable instrumented product.
        $component=$product.VBProject.VBComponents.Item('NxPrivacyScan');$module=$component.CodeModule
        try {
            $text=$module.Lines(1,$module.CountOfLines)
            $scanNeedle='    checkedCells = 0'
            $reportNeedle='    Set report = reportBook.Worksheets(1)'
            if(-not $text.Contains($scanNeedle) -or -not $text.Contains($reportNeedle)){throw 'Privacy checkpoint target changed'}
            $text=$text.Replace($scanNeedle, $scanNeedle+"`r`n"+'    T_R102PrivacyFlow.Checkpoint "scan"')
            $text=$text.Replace($reportNeedle, $reportNeedle+"`r`n"+'    T_R102PrivacyFlow.Checkpoint "report"')
            $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text)
        }finally{Release-ComObject $module;Release-ComObject $component}
    }
    if($selected -contains 'quickformat'){
        $component=$product.VBProject.VBComponents.Item('NxQuickFormatUndo');$module=$component.CodeModule
        try {
            $text=$module.Lines(1,$module.CountOfLines)
            $text=[regex]::Replace($text,'(?m)^\s*MsgBox .* & detail,.*$','    Err.Raise 5, "UndoProbe", detail')
            $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text)
        }finally{Release-ComObject $module;Release-ComObject $component}
    }
    if($selected -contains 'output-flows'){
        # Expose the Click handler only. Production approvals are not replaced.
        foreach($name in @('FNxAgeCalculator','FNxDataSpecial','FNxDataNormalize','FNxFileOutput','FNxFileConsolidate','FNxWorkbookCompare','CNxDataOutput')){
            $component=$product.VBProject.VBComponents.Item($name);$module=$component.CodeModule
            try {
                $text=$module.Lines(1,$module.CountOfLines).Replace("`r`n","`n")
                if($name -eq 'CNxDataOutput'){
                    $text=$text.Replace('    Set CreateOutput = mOutput','    If T_R69OutputFlows.FailAfterCreate Then Err.Raise 5, "R69Probe", "Injected output failure"'+"`n"+'    Set CreateOutput = mOutput')
                }else{
                    $text=[regex]::Replace($text,'(?im)^\s*MsgBox detail,.*$','    Err.Raise 5, "PopupProbe", detail')
                    $text+="`nPublic Sub NxProbeExecute()`n    cmdExecute_Click`nEnd Sub`n"
                }
                $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text.Replace("`n","`r`n"))
            }finally{Release-ComObject $module;Release-ComObject $component;$module=$null;$component=$null}
        }
        $report.evidence_class='native_excel_production_approval_path_with_handler_access_and_explicit_fault'
    }
    if($selected -contains 'userflows'){
        # Disposable probe only: invoke actual popup handlers without clicking
        # Preview. Automated approvals still use the typed execution engine.
        Add-Probe 'T_R62DataSpecial' ([IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/data/T_R62DataSpecial.bas'),[Text.Encoding]::UTF8))
        foreach($name in @('FNxAgeCalculator','FNxDataSpecial','FNxDataNormalize','FNxNumberFormat','FNxResize','FNxFileOutput','FNxFileConsolidate','NxFileController','NxDataSpecialController')){
            $component=$product.VBProject.VBComponents.Item($name);$module=$component.CodeModule
            try{
                $text=$module.Lines(1,$module.CountOfLines).Replace("`r`n","`n")
                switch($name){
                    'NxDataSpecialController' {$text=$text.Replace('form.Show vbModal','form.Show vbModeless')}
                    'FNxDataSpecial' {$text=$text.Replace('NxDataSpecialRunOptions(mFeatureId, source, mOptions)','NxDataSpecialRunOptions(mFeatureId, source, mOptions, True)')}
                    'FNxAgeCalculator' {$text=$text.Replace('NxDataSpecialRunOptions(mFeatureId, source, mOptions)','NxDataSpecialRunOptions(mFeatureId, source, mOptions, True)')}
                    'FNxDataNormalize' {$text=$text.Replace('NxDataRunNormalization(source,','NxDataRunNormalizationForTest(source,')}
                    'NxFileController' {
                        $from='Set result = NxRunPlannedFile(ticket, command, definition)'
                        if(-not $text.Contains($from)){throw 'File approval test boundary changed'}
                        $text=$text.Replace($from,'Set result = router.RunApproved(ticket, router.TakePlannedApproval(ticket, True))')
                    }
                }
                if($name.StartsWith('FNx')){
                    $text=[regex]::Replace($text,'(?im)^\s*MsgBox detail,.*$','    Err.Raise 5, "PopupProbe", detail')
                    $text=[regex]::Replace($text,'(?im)^\s*MsgBox "[^"\r\n]+", vbInformation[^\r\n]*$','    Debug.Print "Popup operation complete"')
                    $text+="`nPublic Sub NxProbeExecute()`n    cmdExecute_Click`nEnd Sub`n"
                    if($name -eq 'FNxDataNormalize'){$text+="`nPublic Sub NxProbeDefaults()`n    cmdResetDefaults_Click`nEnd Sub`n"}
                }
                $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text.Replace("`n","`r`n"))
            }finally{Release-ComObject $module;Release-ComObject $component;$module=$null;$component=$null}
        }
        $report.evidence_class='native_excel_popup_handlers_with_test_only_typed_approvals'
    }
    if(($selected -contains 'compare-internal') -or ($selected -contains 'compare-enhanced')){
        if($selected.Count -ne 1){throw 'Compare profile probes must be isolated'}
        $profile=if($selected -contains 'compare-enhanced'){'enhanced-dll'}else{'internal-xlam'}
        $prepared=Join-Path $EvidenceRoot 'compare-profile.bas'
        & python -X utf8 (Join-Path $SourceRoot 'tools/distribution_source_profile.py') --source (Join-Path $SourceRoot 'src/vba/features/file/compare/NxWorkbookCompare.bas') --destination $prepared --profile $profile
        if($LASTEXITCODE -ne 0){throw 'Compare profile preparation failed'}
        $component=$product.VBProject.VBComponents.Item('NxWorkbookCompare')
        $text=[IO.File]::ReadAllText($prepared,[Text.Encoding]::UTF8)
        $text=$text.Substring($text.IndexOf('Option Explicit'))
        $text+=[IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/product/R69CompareWrapper.txt'),[Text.Encoding]::UTF8)
        $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
        $component.CodeModule.AddFromString($text.Replace("`r`n","`n").Replace("`n","`r`n"))
        Release-ComObject $component;$component=$null
        $component=$product.VBProject.VBComponents.Item('NxHostBridge')
        $text=$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)
        $signature='Public Function NxHostCreateWorkbookCompare() As Object'
        if(-not $text.Contains($signature)){throw 'Compare factory signature changed'}
        $text=$text.Replace($signature,$signature+"`r`n    If T_R69CompareMigration.DisableHost Then Exit Function")
        $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
        $component.CodeModule.AddFromString($text)
        Release-ComObject $component;$component=$null
        $report.evidence_class='native_excel_profile_transformation_with_test_only_host_unavailable_switch'
    }
    if($selected -contains 'stage-profile'){
        $component=$product.VBProject.VBComponents.Item('NxCommandRunner');$module=$component.CodeModule
        try {
            $text=$module.Lines(1,$module.CountOfLines).Replace("`r`n","`n")
            $sites=@('    Set context = NxContextFactory.CaptureCurrent()', '    Set plan = command.BuildPlan(context)', '    journal.Capture context, approvedPlan, mSelectionCapability', '    Set result = dispatcher.ExecuteApproved(command, context, approvedPlan)')
            for($i=0;$i -lt $sites.Count;$i++){
                if(-not $text.Contains($sites[$i])){throw 'Stage profile site missing'}
                $text=$text.Replace($sites[$i],"    NxStageProfile ""start-$i""`n"+$sites[$i]+"`n    NxStageProfile ""end-$i""")
            }
            $text+="`nPrivate Sub NxStageProfile(ByVal stage As String)`n    Dim handle As Integer`n    handle = FreeFile`n    Open Environ`$(""LHEXCEL_PROFILE_ROOT"") & ""\stages.tsv"" For Append As #handle`n    Print #handle, stage & vbTab & CStr(Timer)`n    Close #handle`nEnd Sub`n"
            $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text.Replace("`n","`r`n"))
        }finally{Release-ComObject $module;Release-ComObject $component;$module=$null;$component=$null}
        $component=$product.VBProject.VBComponents.Item('CNxCellRollbackJournal');$module=$component.CodeModule
        try {
            # Profile the candidate source in the disposable probe, not the published XLAM.
            $journalPath=Join-Path $SourceRoot 'src/vba/core/CNxCellRollbackJournal.cls'
            $text=[IO.File]::ReadAllText($journalPath,[Text.Encoding]::UTF8)
            $text=$text.Substring($text.IndexOf('Option Explicit')).Replace("`r`n","`n")
            $report.source_hashes['src/vba/core/CNxCellRollbackJournal.cls']=(Get-FileHash -LiteralPath $journalPath).Hash
            $text=$text.Replace('Option Explicit',"Option Explicit`nPrivate mProbeStart As Double`nPrivate mProbeAssert As Double`nPrivate mProbeSnapshot As Double`nPrivate mProbeShare As Double")
            $sites=@('                    AssertSupportedCell cell, target, effect.MutationMask','                    SnapshotUniqueCell cell, effect.MutationMask, captureLimit','    record.Add ShareExactProperties(properties)')
            $totals=@('mProbeAssert','mProbeSnapshot','mProbeShare')
            for($i=0;$i -lt $sites.Count;$i++){
                $match=[regex]::Match($text,'(?im)^'+[regex]::Escape($sites[$i])+'$')
                if(-not $match.Success){
                    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'journal-profile-source.txt'),$text)
                    throw "Journal profile site missing: $i"
                }
                $sites[$i]=$match.Value
                $text=$text.Replace($sites[$i],"    mProbeStart = Timer`n"+$sites[$i]+"`n    "+$totals[$i]+" = "+$totals[$i]+" + Timer - mProbeStart")
            }
            # Snapshot nests ShareExactProperties, so use a local timer for the outer measurement.
            $text=$text.Replace('    Dim captureLimit As Long',"    Dim captureLimit As Long`n    Dim snapshotStart As Double")
            $text=$text.Replace("    mProbeStart = Timer`n"+$sites[1],"    snapshotStart = Timer`n"+$sites[1]).Replace('mProbeSnapshot + Timer - mProbeStart','mProbeSnapshot + Timer - snapshotStart')
            $text=$text.Replace('    mCaptured = True',"    NxJournalProfile`n    mCaptured = True")
            $text+="`nPrivate Sub NxJournalProfile()`n    Dim handle As Integer`n    handle = FreeFile`n    Open Environ`$(""LHEXCEL_PROFILE_ROOT"") & ""\journal.tsv"" For Append As #handle`n    Print #handle, CStr(mProbeAssert) & vbTab & CStr(mProbeSnapshot) & vbTab & CStr(mProbeShare)`n    Close #handle`nEnd Sub`n"
            $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text.Replace("`n","`r`n"))
        }finally{Release-ComObject $module;Release-ComObject $component;$module=$null;$component=$null}
        $report.evidence_class='native_excel_test_only_stage_timing'
    }
    if($selected -contains 'snapshot-probe'){
        if($selected -notcontains 'stage-profile'){
            $journalPath=Join-Path $SourceRoot 'src/vba/core/CNxCellRollbackJournal.cls'
            $text=[IO.File]::ReadAllText($journalPath,[Text.Encoding]::UTF8)
            $component=$product.VBProject.VBComponents.Item('CNxCellRollbackJournal');$module=$component.CodeModule
            try {
                $module.DeleteLines(1,$module.CountOfLines)
                $module.AddFromString($text.Substring($text.IndexOf('Option Explicit')).Replace("`r`n","`n").Replace("`n","`r`n"))
                $report.source_hashes['src/vba/core/CNxCellRollbackJournal.cls']=(Get-FileHash -LiteralPath $journalPath).Hash
            } finally {Release-ComObject $module;Release-ComObject $component;$module=$null;$component=$null}
        }
        # Fault/benchmark instrumentation in disposable probe only.
        foreach($name in @('CNxDrawingFeatureCommand','NxDrawController')){
            $component=$product.VBProject.VBComponents.Item($name);$module=$component.CodeModule
            try{
                $text=$module.Lines(1,$module.CountOfLines).Replace("`r`n","`n")
                if($name -eq 'CNxDrawingFeatureCommand'){
                    $pattern='(?im)^    If mRequest.FeatureId (?:<> NX_FEATURE_DRAW_FIT_PICTURE|= NX_FEATURE_DRAW_CLEAR_INNER) Then snapshot.Capture mRequest.Target$'
                    if([regex]::Matches($text,$pattern).Count -ne 1){throw 'Snapshot capture site changed'}
                    $text=[regex]::Replace($text,$pattern,'    If T_R69SnapshotProbe.UseDuplicateBackup Or mRequest.FeatureId = NX_FEATURE_DRAW_CLEAR_INNER Then snapshot.Capture mRequest.Target')
                }else{
                    $pattern='(?im)^    If autoFitColumns Then target.Columns.AutoFit$'
                    if([regex]::Matches($text,$pattern).Count -ne 1){throw 'Table fault site changed'}
                    $text=[regex]::Replace($text,$pattern,"    If autoFitColumns Then target.Columns.AutoFit`n    T_R69SnapshotProbe.MaybeFail")
                }
                $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString($text.Replace("`n","`r`n"))
            }finally{Release-ComObject $module;Release-ComObject $component;$module=$null;$component=$null}
        }
        $report.evidence_class='native_excel_test_only_snapshot_switch_and_fault'
    }
    if($selected -contains 'runner-benchmark'){
        # A/B switch exists only in the disposable probe, never the built Product.
        $component=$product.VBProject.VBComponents.Item('NxCommandRunner')
        $module=$component.CodeModule
        try {
            $text=$module.Lines(1,$module.CountOfLines).Replace("`r`n","`n")
            $optimized="        sourceChanged = result.SourceChanged`n        If Not sourceChanged Then sourceChanged = (context.SourceStateMatchesCurrent() = False)"
            $pattern='(?i)'+[regex]::Escape($optimized)
            if([regex]::Matches($text,$pattern).Count -ne 1){throw 'Runner benchmark target must occur exactly once'}
            $text=$text.Replace('Private mActiveDispatcher As CNxSideEffectDispatcher',"Private mBenchmarkLegacy As Boolean`nPrivate mActiveDispatcher As CNxSideEffectDispatcher")
            $text=[regex]::Replace($text,$pattern,"        If mBenchmarkLegacy Then`n            sourceChanged = result.SourceChanged Or (context.SourceStateMatchesCurrent() = False)`n        Else`n"+$optimized+"`n        End If")
            $text+="`nPublic Sub NxSetLegacyBenchmark(ByVal enabled As Boolean)`n    mBenchmarkLegacy = enabled`nEnd Sub`n"
            $module.DeleteLines(1,$module.CountOfLines)
            $module.AddFromString($text.Replace("`n","`r`n"))
        }finally{Release-ComObject $module;Release-ComObject $component;$module=$null;$component=$null}
        $report.evidence_class='native_excel_paired_test_only_runner_switch'
    }
    if($OverridePerformance){
        foreach($relative in @('src/vba/core/CNxExecutionContext.cls','src/vba/core/NxContextFactory.bas','src/vba/features/file/NxPngExport.bas')){
            $path=Join-Path $SourceRoot $relative
            $text=[IO.File]::ReadAllText($path,[Text.Encoding]::UTF8)
            $component=$product.VBProject.VBComponents.Item([IO.Path]::GetFileNameWithoutExtension($path))
            $component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)
            $component.CodeModule.AddFromString($text.Substring($text.IndexOf('Option Explicit')).Replace("`r`n","`n").Replace("`n","`r`n"))
            $report.source_hashes[$relative]=(Get-FileHash -LiteralPath $path).Hash
            Release-ComObject $component;$component=$null
        }
        $report.evidence_class='native_excel_instrumented_source_override'
    }
    if($selected -contains 'runner-safety'){
        Add-Probe 'CFakeFeatureCommand' ([IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/CFakeFeatureCommand.cls'),[Text.Encoding]::UTF8)) 2
        Add-Probe 'T_Core' ([IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/T_Core.bas'),[Text.Encoding]::UTF8))
    }
    if(($selected -contains 'r93-compare') -or ($selected -contains 'r94-template')){
        Add-Probe 'C_R93CompareProgress' ([IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/product/C_R93CompareProgress.cls'),[Text.Encoding]::UTF8)) 2
    }
    foreach($suite in $selected){
        $path=Join-Path $SourceRoot $map[$suite]
        $report.source_hashes[$map[$suite]]=(Get-FileHash -LiteralPath $path).Hash
        Add-Probe ([IO.Path]::GetFileNameWithoutExtension($path)) ([IO.File]::ReadAllText($path,[Text.Encoding]::UTF8))
    }
    $harness=[IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/NxTestHarness.bas'),[Text.Encoding]::UTF8)
    $start=$harness.IndexOf('Public Sub AssertTrue(');$end=$harness.IndexOf('Private Sub AppendCase(')
    if($start -lt 0 -or $end -le $start){throw 'Assertion source boundary changed'}
    Add-Probe 'NxTestHarness' ("Option Explicit`r`n"+$harness.Substring($start,$end-$start))
    $component=$product.VBProject.VBComponents.Item([IO.Path]::GetFileNameWithoutExtension($map[$selected[0]]))
    $component.Activate()
    $control=Find-VbeCompileControl $excel.VBE.CommandBars
    $watcher=Start-ExactOwnedCompileDialogWatcher -ExcelPid ([int]$binding.pid) -ExcelHwnd ([Int64]$binding.hwnd) -VbeHwnd ([Int64]$excel.VBE.MainWindow.HWnd)
    if($control.Enabled){$control.Execute()}
    if($null -eq (Wait-Job $watcher -Timeout 30)){throw 'Compile watcher timeout'}
    $observed=@(Receive-Job $watcher -Wait -AutoRemoveJob)
    $report.compile_observation=$observed
    try {$report.compile_selection=[string]$excel.VBE.ActiveCodePane.CodeModule.Name} catch {}
    if($observed.Count -and $observed[0].status -ne 'NO_DIALOG'){
        $pane=$null;$module=$null
        try{
            [int]$line=0;[int]$column=0;[int]$lastLine=0;[int]$lastColumn=0
            $pane=$excel.VBE.ActiveCodePane;$module=$pane.CodeModule
            $pane.GetSelection([ref]$line,[ref]$column,[ref]$lastLine,[ref]$lastColumn)
            $report.compile_line=$line;$report.compile_source=[string]$module.Lines($line,1)
        }finally{Release-ComObject $module;Release-ComObject $pane}
    }
    if($observed.Count -ne 1 -or $observed[0].status -ne 'NO_DIALOG' -or $control.Enabled){throw 'R68 probe compile failed'}
    $report.compile='PASS';$product.Save()
    Release-ComObject $control;$control=$null
    Release-ComObject $component;$component=$null
    $product.Close($false);Release-ComObject $product;$product=$null
    $report.instrumented_sha256=(Get-FileHash -LiteralPath $copy).Hash
    $product=$books.Open($copy,0,$false)
    $macro="'"+$product.Name.Replace("'","''")+"'!"
    $fixture=$books.Add(-4167)
    $fixture.SaveAs((Join-Path $EvidenceRoot 'fixture.xlsx'),51)
    $excel.VBE.MainWindow.Visible=$false
    foreach($suite in $selected){
        $module=[IO.Path]::GetFileNameWithoutExtension($map[$suite])
        if($suite -in @('large-range','stage-profile')){[void]$excel.Run($macro+$module+'.Configure',$EvidenceRoot)}
        if($suite -in @('benchmark','r97-preview')){[void]$excel.Run($macro+$module+'.Configure',$EvidenceRoot)}
        if($suite -eq 'runner-benchmark'){[void]$excel.Run($macro+$module+'.Configure',$EvidenceRoot)}
        if($suite -eq 'repaint-benchmark'){[void]$excel.Run($macro+$module+'.Configure',$EvidenceRoot)}
        if($suite -eq 'snapshot-probe'){[void]$excel.Run($macro+$module+'.Configure',$EvidenceRoot)}
        if($suite -in @('compare-internal','compare-enhanced')){[void]$excel.Run($macro+$module+'.Configure',($suite -eq 'compare-enhanced'))}
        if($suite -eq 'performance'){[void]$excel.Run($macro+$module+'.Configure',$EvidenceRoot, $(if($ObserveBaseline){'baseline'}else{'candidate'}))}
        $names=@(([string]$excel.Run($macro+$module+'.Names')).Split('|')|Where-Object {$_})
        if(-not $names.Count){throw 'Empty case inventory'}
        if($CaseName){
            if($selected.Count -ne 1 -or $suite -notin @('compare-consolidate-links','consolidate-list','r94-template','table','output-flows','popup-layout') -or $CaseName -notin $names){throw 'Unknown focused case'}
            $names=@($CaseName)
        }
        foreach($name in $names){
            if($suite -eq 'r76-policy' -or $suite -eq 'r86-policy'){
                $gateText=[IO.File]::ReadAllText((Join-Path $EvidenceRoot ("policy-fixtures/"+$name+".bas")),[Text.Encoding]::UTF8)
                $gateText=$gateText.Substring($gateText.IndexOf('Option Explicit')).Replace("`r`n","`n")
                $gateText=$gateText.Replace('Private mVerified As Boolean',"Private mProbeDetail As String`r`nPrivate mVerified As Boolean")
                $gateText=$gateText.Replace("Invalid:", "Invalid:`r`n    mProbeDetail = mProbeDetail & ""|stateerr:"" & CStr(Err.Number) & "":"" & Err.Description")
                $gateText=$gateText.Replace("Clean:`n", "Clean:`n    mProbeDetail = mProbeDetail & ""|cryptoerr:"" & CStr(Err.Number) & "":"" & Err.Description & ""|dll:"" & CStr(Err.LastDllError)`n")
                $gateText=$gateText.Replace('    mVerified = NxDistributionVerify()', '    mProbeDetail = mProbeDetail & "|verify": mVerified = NxDistributionVerify()')
                $gateText+="`r`nPublic Function NxDistributionProbeDetail() As String`r`n    NxDistributionProbeDetail = mProbeDetail`r`nEnd Function"
                $gateText+="`r`nPublic Function NxDistributionProbeTimer() As Boolean`r`n    NxDistributionProbeTimer = mRefreshScheduled`r`nEnd Function"
                $gateComponent=$product.VBProject.VBComponents.Item('NxDistributionGate')
                $gateModule=$gateComponent.CodeModule
                try {$gateModule.DeleteLines(1,$gateModule.CountOfLines);$gateModule.AddFromString($gateText)}
                finally {Release-ComObject $gateModule;Release-ComObject $gateComponent}
                $payloadMatch=[regex]::Match($gateText,'NX_POLICY_PAYLOAD As String = "([^"]+)"')
                $policyFields=$payloadMatch.Groups[1].Value.Split('|')
                $expiry=[datetime]::ParseExact($policyFields[3],'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
                [void]$excel.Run($macro+'T_R76Policy.SetupMetadata',$expiry.ToString('yyyy-MM-dd'),$expiry.AddDays(-1).ToString('yyyy-MM-dd'),$policyFields[2],[bool]($name -eq 'tampered_metadata'))
            }
            $caseWatch=[Diagnostics.Stopwatch]::StartNew()
            $result=[string]$excel.Run($macro+$module+'.RunCase',$name)
            $caseWatch.Stop()
            $report.cases+=[ordered]@{suite=$suite;name=$name;result=$result;fixture_inclusive_elapsed_ms=$caseWatch.Elapsed.TotalMilliseconds}
            Write-Output $result
            if($suite -eq 'r94-template' -and $name -eq 'ui_layout'){
                [void]$excel.Run($macro+$module+'.OpenPanel')
            }
            if($suite -eq 'popup-layout' -or ($suite -eq 'r94-template' -and $name -eq 'ui_layout')){
                if(-not ('NxPopupCapture' -as [type])){. (Join-Path $SourceRoot 'tests/windows/Capture-OwnedPopup.ps1')}
                $imageRoot=Join-Path $EvidenceRoot 'ui';[void](New-Item -ItemType Directory -Force -Path $imageRoot)
                $caption=[string]$excel.Run($macro+$module+'.Caption')
                Start-Sleep -Milliseconds 150
                [NxPopupCapture]::Save($caption,[int]$binding.pid,(Join-Path $imageRoot ($name+'.png')))
                [void]$excel.Run($macro+$module+'.ClosePanel')
            }
            if($suite -eq 'visual'){
                $excel.EnableEvents=$true
                $until=(Get-Date).AddMinutes(15)
                Write-Output ('VISUAL_READY|'+$EvidenceRoot)
                while(-not (Test-Path -LiteralPath (Join-Path $EvidenceRoot 'visual.done')) -and (Get-Date) -lt $until){Start-Sleep -Seconds 1}
                if(-not (Test-Path -LiteralPath (Join-Path $EvidenceRoot 'visual.done'))){throw 'Visual verification timed out'}
                $excel.EnableEvents=$false
            }
        }
    }
    if($selected -contains 'quickformat'){
        [void]$excel.Run($macro+'T_R68QuickFormat.PrepareNativeUndo')
        $undo=$excel.CommandBars.FindControl(1,128)
        if($null -eq $undo -or -not $undo.Enabled){throw 'Native Undo control disabled'}
        try {$undo.Execute()} finally {Release-ComObject $undo}
        $result=[string]$excel.Run($macro+'T_R68QuickFormat.VerifyNativeUndo')
        $report.cases+=[ordered]@{suite='quickformat';name='native_undo';result=$result}
        Write-Output $result
    }
    if($selected -contains 'r94-compare-output'){
        if(-not ('NxPopupCapture' -as [type])){. (Join-Path $SourceRoot 'tests/windows/Capture-OwnedPopup.ps1')}
        $imageRoot=Join-Path $EvidenceRoot 'ui';[void](New-Item -ItemType Directory -Force -Path $imageRoot)
        [void]$excel.Run($macro+'T_R94CompareOutput.OpenPanel')
        $caption=[string]$excel.Run($macro+'T_R94CompareOutput.Caption')
        [NxPopupCapture]::Save($caption,[int]$binding.pid,(Join-Path $imageRoot 'file-compare.png'))
        [void]$excel.Run($macro+'T_R94CompareOutput.ClosePanel')
    }
    if($selected -contains 'r94-functions'){
        if(-not ('NxPopupCapture' -as [type])){. (Join-Path $SourceRoot 'tests/windows/Capture-OwnedPopup.ps1')}
        $imageRoot=Join-Path $EvidenceRoot 'ui';[void](New-Item -ItemType Directory -Force -Path $imageRoot)
        foreach($errorMode in @($false,$true)){
            [void]$excel.Run($macro+'T_R94Functions.OpenPanel',$errorMode)
            $caption=[string]$excel.Run($macro+'T_R94Functions.Caption')
            [NxPopupCapture]::Save($caption,[int]$binding.pid,(Join-Path $imageRoot $(if($errorMode){'iferror.png'}else{'round.png'})))
            [void]$excel.Run($macro+'T_R94Functions.ClosePanel')
        }
    }
    if($selected -contains 'r94-formulatools'){
        [void]$excel.Run($macro+'T_R94FormulaTools.PrepareNativeUndo')
        $undo=$excel.CommandBars.FindControl(1,128)
        if($null -eq $undo -or -not $undo.Enabled){throw 'Native notes Undo control disabled'}
        try {$undo.Execute()} finally {Release-ComObject $undo}
        $result=[string]$excel.Run($macro+'T_R94FormulaTools.VerifyNativeUndo')
        $report.cases+=[ordered]@{suite='r94-formulatools';name='native_notes_undo';result=$result}
        Write-Output $result
        if(-not ('NxPopupCapture' -as [type])){. (Join-Path $SourceRoot 'tests/windows/Capture-OwnedPopup.ps1')}
        $imageRoot=Join-Path $EvidenceRoot 'ui';[void](New-Item -ItemType Directory -Force -Path $imageRoot)
        foreach($references in @($false,$true)){
            [void]$excel.Run($macro+'T_R94FormulaTools.OpenPanel',$references)
            $caption=[string]$excel.Run($macro+'T_R94FormulaTools.Caption')
            [NxPopupCapture]::Save($caption,[int]$binding.pid,(Join-Path $imageRoot $(if($references){'references.png'}else{'notes.png'})))
            [void]$excel.Run($macro+'T_R94FormulaTools.ClosePanel')
        }
    }
    if($selected -contains 'r74-folder'){
        $shell=New-Object -ComObject Shell.Application
        $ownedFolders=@()
        try {
            $fixture.Activate()
            $beforeSaved=[bool]$fixture.Saved
            for($attempt=1;$attempt -le 2;$attempt++){
                $before=@($shell.Windows()|ForEach-Object {[long]$_.HWND})
                [void]$excel.Run($macro+'NxProductOpenSavedFolder')
                $deadline=(Get-Date).AddSeconds(15)
                $found=$null
                do {
                    foreach($candidate in @($shell.Windows())){
                        try {
                            if([long]$candidate.HWND -notin $before -and
                               [IO.Path]::GetFullPath([string]$candidate.Document.Folder.Self.Path) -eq [IO.Path]::GetFullPath($EvidenceRoot)){
                                $found=$candidate;break
                            }
                        }catch{}
                    }
                    if($null -eq $found){Start-Sleep -Milliseconds 100}
                }while($null -eq $found -and (Get-Date) -lt $deadline)
                if($null -eq $found){throw 'Saved folder did not open a new Explorer window at the expected path'}
                $ownedFolders+=,$found
                $report.cases+=[ordered]@{suite='r74-folder';name=('new_explorer_'+$attempt);result=('PASS|new_explorer_'+$attempt);hwnd=[long]$found.HWND;path=[string]$found.Document.Folder.Self.Path}
            }
            if([bool]$fixture.Saved -ne $beforeSaved){throw 'Opening folder changed workbook Saved state'}
        }finally{
            foreach($ownedFolder in $ownedFolders){try{$ownedFolder.Quit()}finally{Release-ComObject $ownedFolder}}
            Release-ComObject $shell
        }
    }
    Add-Type -AssemblyName System.Drawing
    foreach($path in @(Get-ChildItem -LiteralPath $EvidenceRoot -Filter '*.png' -File)){
        $bitmap=New-Object Drawing.Bitmap($path.FullName)
        try {
            $yellow=0;$dark=0;$samples=0
            for($y=0;$y -lt $bitmap.Height;$y+=2){for($x=0;$x -lt $bitmap.Width;$x+=2){
                $pixel=$bitmap.GetPixel($x,$y);$samples++
                if($pixel.R -gt 180 -and $pixel.G -gt 140 -and $pixel.B -lt 185){$yellow++}
                if($pixel.R -lt 100 -and $pixel.G -lt 100 -and $pixel.B -lt 100){$dark++}
            }}
            $report.images+=[ordered]@{file=$path.Name;width=$bitmap.Width;height=$bitmap.Height;yellow=$yellow;dark=$dark;samples=$samples;content_pass=($yellow -gt $samples*0.2 -and $dark -gt 10);sha256=(Get-FileHash -LiteralPath $path.FullName).Hash}
        } finally {$bitmap.Dispose()}
    }
    $failed=@($report.cases|Where-Object {
        if($_.suite -in @('serialization-probe','r76-format-backup')){ -not $_.result.StartsWith(('PASS|'+$_.name+'|'),[StringComparison]::Ordinal) }
        else { $_.result -cne ('PASS|'+$_.name) }
    }).Count
    $badImages=@($report.images|Where-Object {-not $_.content_pass}).Count
    if($selected -contains 'performance' -and @($report.images|Where-Object {$_.file -like 'export_*' -and $_.file -notlike 'export_bench_*'}).Count -ne 4){$badImages++}
    if($selected -contains 'benchmark' -and @($report.images|Where-Object {$_.file -like 'export_bench_*'}).Count -ne 9){$badImages++}
    if($selected -contains 'benchmark' -and @($report.images|Where-Object {$_.file -like 'png_*'}).Count -ne 9){$badImages++}
    if($failed -or $badImages){
        if($ObserveBaseline){$report.status='OBSERVED_DEFECT'}else{throw "R68 failed cases=$failed images=$badImages"}
    }else{$report.status='PASS'}
    if((Get-FileHash -LiteralPath $ProductXlam).Hash -ne $report.product_sha256){throw 'Source Product changed'}
}catch{$failure=$_.Exception.Message;$report.status='FAIL';$report.failure=$failure;$report.failure_location=$_.InvocationInfo.PositionMessage}
finally {
    $safeCleanup=$true
    try {
    if($null -ne $binding -and $null -ne $excel){
        [void](Assert-ExactExcelProcessOwnership $binding)
        if($null -ne $books){for($i=$books.Count;$i -ge 1;$i--){
            $book=$books.Item($i);$full=[IO.Path]::GetFullPath([string]$book.FullName)
            if(-not $full.StartsWith([IO.Path]::GetFullPath($EvidenceRoot)+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Unexpected workbook preserved'}
            $book.Close($false);Release-ComObject $book
        }}
        $excel.Quit()
    }
    } catch {
        $safeCleanup=$false
        $report.status='FAIL';$report.cleanup_failure=$_.Exception.Message
        $report.cleanup_location=$_.InvocationInfo.PositionMessage
        if($null -eq $failure){$failure=$_.Exception.Message}
    }
    foreach($com in @($control,$component,$fixture,$product,$books,$excel)){Release-ComObject $com}
    # Chained VBE/CodeModule property calls create temporary RCWs. A second
    # finalization pass releases their dependent wrappers before observing Quit.
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null -ne $binding -and $safeCleanup){
        $report.cleanup=Stop-ExactProcessAfterGrace $binding.process 'r68 product probe' 30000 -Detailed
        $binding.process.Dispose()
        if($report.cleanup.exit_mode -ne 'NATURAL'){$report.status='FAIL';$failure='Excel did not exit naturally'}
    }elseif($null -ne $binding){
        # Unknown workbook/COM state: retain the process, but still write evidence.
        $report.cleanup=[ordered]@{exit_mode='PRESERVED';pid=$binding.pid}
        $binding.process.Dispose()
    }
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'r68.json'),($report|ConvertTo-Json -Depth 10),(New-Object Text.UTF8Encoding($false)))
}
if($null -ne $failure){throw $failure}
Write-Output ($report.status+'|R68|'+$report.cases.Count)
