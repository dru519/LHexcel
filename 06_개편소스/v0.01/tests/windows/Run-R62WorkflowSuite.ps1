param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RunId = '',
    [string]$CaseNames = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$ProductXlam = [IO.Path]::GetFullPath($ProductXlam)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
$allowedCases = @(
    'special.TestR62AgeInputAndReference',
    'special.TestR62MoneyOptionsAndPrecision',
    'special.TestR62MaskOptionsAndAuditIsolation',
    'special.TestR62SpecialOutputModes',
    'special.TestR62SpecialPreviewAndCollision',
    'special.TestR62NormalizeOutputVisibility',
    'analysis.R62UniqueRecordCountsAndRatios',
    'analysis.R62UniqueColumnsAndBlankSentinel',
    'analysis.R62CompareCrossMembershipAndDense',
    'analysis.R62CompareRowsAndHidden',
    'analysis.R62AnalysisRejectsStaleValuesAndFormulas',
    'analysis.R62AnalysisLiteralNewWorkbookAndSourcePreservation',
    'template.TestR62TemplateWorkbookAndCopyPreserveOriginal',
    'template.TestR62TemplateSheetRegistrationUsesWholeSheet',
    'template.TestR62TemplateMultisheetFailureRollsBack',
    'template.TestR62TemplateCatalogIsLiteralAndCollisionSafe',
    'template.TestR62TemplateFileRegistrationClosesOwnedSource',
    'template.TestR62TemplateManagerControlsAreActionable',
    'rename.rules',
    'rename.files',
    'rename.sheets',
    'rename.recovery',
    'ai.tasks',
    'ai.compatibility',
    'ai.masked_preview',
    'ai.reselect',
    'ai.reject_mismatch',
    'ai.incomplete',
    'ai.form',
    'shortcut.draft_cancel',
    'shortcut.draft_remove',
    'shortcut.draft_clear',
    'shortcut.conflict',
    'shortcut.stale_apply',
    'shortcut.reserved_apply',
    'shortcut.apply_roundtrip',
    'shortcut.save_failure_rollback',
    'picture.range_default',
    'picture.shape_default',
    'picture.multi_shape_default',
    'picture.mode_cancel',
    'picture.fit_preview',
    'picture.fit_drift',
    'picture.insert_preview',
    'picture.engine_rollback',
    'picture.picker_state'
)
function Resolve-R62Cases([string]$Value,[string[]]$Allowed) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Allowed }
    $selected = @($Value.Split(',') | ForEach-Object { $_.Trim() })
    if (-not $selected.Count -or @($selected | Where-Object { $Allowed -cnotcontains $_ }).Count -or
        @($selected | Select-Object -Unique).Count -ne $selected.Count) { throw 'Unknown, empty, or duplicate R62 workflow case' }
    return $selected
}
$requestedCases = @(Resolve-R62Cases $CaseNames $allowedCases)
if (-not (Test-Path -LiteralPath $ProductXlam -PathType Leaf)) { throw 'Product XLAM unavailable' }
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing R62 workflow evidence is rejected' }
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'R62 workflow requires zero pre-existing Excel processes' }
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

function Write-Json([string]$Path,[object]$Value) {
    [IO.File]::WriteAllText($Path,(($Value | ConvertTo-Json -Depth 16)+"`n"),(New-Object Text.UTF8Encoding($false)))
}
function Release-Com([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}
function Release-ComObject([object]$Value) { Release-Com $Value }
function Get-Sha256([string]$Path) { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Get-SourceSnapshot {
    $files = @(foreach ($folder in @('src','contracts','tests/vba')) {
        Get-ChildItem -LiteralPath (Join-Path $SourceRoot $folder) -File -Recurse
    })
    $files += Get-Item -LiteralPath (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1'),$PSCommandPath
    $hashes = @{}
    foreach ($file in $files) {
        $relative = $file.FullName.Substring($SourceRoot.Length + 1).Replace('\','/')
        $hashes[$relative] = Get-Sha256 $file.FullName
    }
    [string[]]$keys = @($hashes.Keys)
    [Array]::Sort($keys,[StringComparer]::Ordinal)
    $ordered = [ordered]@{}
    $lines = foreach ($key in $keys) { $ordered[$key]=$hashes[$key]; $hashes[$key]+'  '+$key }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $digest=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($lines -join "`n")+"`n")))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
    return [ordered]@{sha256=$digest;files=$ordered}
}
function Add-ProbeModule([object]$Components,[string]$Name,[string]$Code) {
    $existing=$null; $component=$null; $module=$null
    try { $existing=$Components.Item($Name) } catch {}
    if ($null -ne $existing) { Release-Com $existing; throw ('Unexpected existing probe module: '+$Name) }
    try {
        $body=[regex]::Replace($Code,'(?m)^Attribute [^\r\n]*\r?\n','').Replace("`r`n","`n").Replace("`r","`n")
        if ($body -notmatch '(?m)^Option Explicit$') { throw ('Probe Option Explicit missing: '+$Name) }
        $component=$Components.Add(1); $component.Name=$Name; $module=$component.CodeModule
        [void]$module.AddFromString($body.Replace("`n","`r`n"))
    } finally { Release-Com $module; Release-Com $component }
}
function Set-RenameProbeConst([object]$Components) {
    $component=$null; $module=$null
    try {
        $component=$Components.Item('NxBatchRenameEngine'); $module=$component.CodeModule
        $code=[string]$module.Lines(1,$module.CountOfLines)
        if ([regex]::Matches($code,'(?m)^#Const NX_R58_TEST_BUILD = False\r?$').Count -ne 1) { throw 'Rename test marker absent or ambiguous in disposable copy' }
        $code=$code.Replace('#Const NX_R58_TEST_BUILD = False','#Const NX_R58_TEST_BUILD = True')
        [void]$module.DeleteLines(1,$module.CountOfLines); [void]$module.AddFromString($code)
    } finally { Release-Com $module; Release-Com $component }
}
function Enable-PicturePickerStateProbe([object]$Components) {
    # Expose the existing implementation only inside the hash-bound disposable XLAM.
    # The three-line signatures must be unique; no production body is replaced.
    $component=$null; $module=$null
    try {
        $component=$Components.Item('NxPictureInsertController'); $module=$component.CodeModule
        $code=([string]$module.Lines(1,$module.CountOfLines)).Replace("`r`n","`n").Replace("`r","`n")
        $options=[Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant -bor [Text.RegularExpressions.RegexOptions]::Multiline
        foreach ($name in @('RestorePictureInsertPickerState','PictureInsertPickerStateMatches')) {
            $signature="Private Function $name(ByVal originalWorkbook As Workbook, ByVal originalSheet As Object, ByVal originalSelection As Object, _`n    ByVal originalScreenUpdating As Boolean, ByVal originalEnableEvents As Boolean, ByVal originalDisplayAlerts As Boolean, _`n    ByVal originalCalculation As XlCalculation) As Boolean"
            $matches=[regex]::Matches($code,'^'+[regex]::Escape($signature)+'$',$options)
            $declarations=[regex]::Matches($code,'^(?:Private|Public) Function '+[regex]::Escape($name)+'\b',$options)
            if ($matches.Count -ne 1 -or $declarations.Count -ne 1) { throw ('Picture picker private signature absent or ambiguous in disposable copy: '+$name) }
            $match=$matches[0]
            $code=$code.Remove($match.Index,$match.Length).Insert($match.Index,'Public'+$match.Value.Substring('Private'.Length))
        }
        [void]$module.DeleteLines(1,$module.CountOfLines)
        [void]$module.AddFromString($code.Replace("`n","`r`n"))
    } finally { Release-Com $module; Release-Com $component }
}
function Ensure-AssertHarness([object]$Components) {
    # Full NxTestHarness.RunSuite statically depends on unrelated T_Core tests.
    # Project only exact AssertTrue required by R62 analysis and template fixtures.
    $source=[IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/NxTestHarness.bas'),[Text.Encoding]::UTF8).Replace("`r`n","`n")
    $matches=[regex]::Matches($source,'(?ms)^Public Sub AssertTrue\(ByVal condition As Boolean, ByVal message As String\)\n.*?^End Sub$')
    if ($matches.Count -ne 1) { throw 'Canonical AssertTrue template absent or ambiguous' }
    $assertion=$matches[0].Value
    if ($assertion -cne "Public Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)`n    If Not condition Then Err.Raise vbObjectError + 711, `"NxTestHarness`", message`nEnd Sub") { throw 'Canonical AssertTrue template changed; review required' }
    $sha=[Security.Cryptography.SHA256]::Create()
    try { $assertionHash=([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($assertion)))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
    $script:assertionEvidence=[ordered]@{source='tests/vba/NxTestHarness.bas';source_sha256=(Get-Sha256 (Join-Path $SourceRoot 'tests/vba/NxTestHarness.bas'));assert_true_sha256=$assertionHash;mode='EXACT_ASSERT_TRUE_PROJECTION'}
    $existing=$null; $module=$null
    try { $existing=$Components.Item('NxTestHarness') } catch {}
    if ($null -ne $existing) {
        try {
            $module=$existing.CodeModule
            $body=([string]$module.Lines(1,$module.CountOfLines)).Replace("`r`n","`n")
            $existingMatches=[regex]::Matches($body,'(?ms)^Public Sub AssertTrue\(ByVal condition As Boolean, ByVal message As String\)\n.*?^End Sub$')
            if ($existingMatches.Count -ne 1 -or $existingMatches[0].Value -cne $assertion) { throw 'Existing disposable AssertTrue differs from canonical source' }
        } finally { Release-Com $module; Release-Com $existing }
    } else { Add-ProbeModule $Components 'NxTestHarness' ("Option Explicit`n`n"+$assertion+"`n") }
}
function Invoke-PictureRollbackProbe([object]$Excel,[object]$Components,[string]$Macro,[string]$TempRoot) {
    # Mutate only the unsaved disposable test copy, only for this failure case.
    $component=$null; $module=$null; $original=$null; $changed=$false
    try {
        $component=$Components.Item('NxPictureLayout'); $module=$component.CodeModule
        $original=[string]$module.Lines(1,$module.CountOfLines)
        $needle='pictureShape.LockAspectRatio = msoTrue: pictureShape.Width = originalWidth * scaleFactor'
        # VBE canonicalizes property spelling (.Width -> .width) on import.
        # Keep an exact statement match, ignoring only VBA-insignificant case.
        $locations=[regex]::Matches($original,[regex]::Escape($needle),([Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::CultureInvariant))
        if ($locations.Count -ne 1) { throw 'Picture fault location must be unique' }
        $updated=$original.Insert($locations[0].Index+$locations[0].Length,': T_R62PictureModes.NxR62PictureFitFaultPoint pictureShape')
        $changed=$true
        $module.DeleteLines(1,$module.CountOfLines); $module.AddFromString($updated)
        return [string]$Excel.Run($Macro+'NxR62WorkflowDispatch','picture','engine_rollback',$TempRoot)
    } finally {
        try {
            if ($changed) { $module.DeleteLines(1,$module.CountOfLines); $module.AddFromString($original) }
        } finally { Release-Com $module; Release-Com $component }
    }
}
function Invoke-ProbeCompile([object]$Excel,[object]$Binding,[object]$Components,[string]$CopyPath,[switch]$AllowAlreadyCompiled) {
    $vbe=$null; $bars=$null; $control=$null; $target=$null; $active=$null; $window=$null; $watcher=$null
    $script:compile=[ordered]@{status='RUNNING';control_id=578;active_project=$null;dialog=$null;failure=$null}
    try {
        [void](Assert-ExactExcelProcessOwnership $Binding)
        $vbe=$Excel.VBE; $bars=$vbe.CommandBars; $control=Find-VbeCompileControl $bars
        if ($null -eq $control -or [int]$control.Id -ne 578) { throw 'R62 VBA compile control 578 unavailable' }
        $target=$Components.Item('R62WorkflowDispatch'); [void]$target.Activate()
        $active=$vbe.ActiveVBProject; $activePath=[IO.Path]::GetFullPath([string]$active.FileName)
        $script:compile.active_project=$activePath
        if (-not [string]::Equals($activePath,$CopyPath,[StringComparison]::OrdinalIgnoreCase)) { throw 'R62 compile active project identity mismatch' }
        if (-not [bool]$control.Enabled) {
            if ($AllowAlreadyCompiled) { $script:compile.status='COMPILE_RETAINED'; return }
            throw 'R62 compile was disabled before required instrumentation compile'
        }
        $window=$vbe.MainWindow
        $watcher=Start-ExactOwnedCompileDialogWatcher -ExcelPid ([int]$Binding.pid) -ExcelHwnd ([Int64]$Binding.hwnd) -VbeHwnd ([Int64]$window.HWnd)
        $executeFailure=$null
        try { [void]$control.Execute() } catch { $executeFailure=$_.Exception.Message }
        if ($null -eq (Wait-Job -Job $watcher -Timeout 30)) { throw 'R62 compile watcher timed out' }
        $observed=@(Receive-Job -Job $watcher -Wait -AutoRemoveJob -ErrorAction Stop); $watcher=$null
        if ($null -ne $executeFailure) { throw ('R62 compile execution failed: '+$executeFailure) }
        if ($observed.Count -ne 1 -or $observed[0].status -cne 'NO_DIALOG') {
            $script:compile.dialog=@($observed)
            $compilePane=$null; $compileModule=$null; $compileLine=0; $compileColumn=0; $compileEndLine=0; $compileEndColumn=0
            $compileLocation='unavailable'
            try {
                $compilePane=$vbe.ActiveCodePane; $compileModule=$compilePane.CodeModule
                $compilePane.GetSelection([ref]$compileLine,[ref]$compileColumn,[ref]$compileEndLine,[ref]$compileEndColumn)
                $compileLocation=([string]$compileModule.Name)+':'+$compileLine+':'+$compileColumn+' '+([string]$compileModule.Lines($compileLine,1))
            } catch { $compileLocation='capture failed: '+$_.Exception.Message }
            finally { Release-Com $compileModule; Release-Com $compilePane }
            throw ('R62 compile dialog detected or watcher result unavailable; '+$compileLocation)
        }
        if ([bool]$control.Enabled) { throw 'R62 compile did not reach disabled state' }
        $script:compile.status='PASS'
    } catch { $script:compile.status='FAIL'; $script:compile.failure=$_.Exception.Message; throw }
    finally {
        if ($null -ne $watcher) { Stop-Job -Job $watcher -ErrorAction SilentlyContinue; Remove-Job -Job $watcher -Force -ErrorAction SilentlyContinue }
        # ActiveVBProject is a borrowed alias of the retained project.
        $active=$null
        Release-Com $window; Release-Com $target; Release-Com $control; Release-Com $bars; Release-Com $vbe
    }
}
function Assert-OnlyProbeWorkbook([object]$Books,[string]$CopyPath) {
    for ($index=1; $index -le $Books.Count; $index++) {
        $item=$null
        try {
            $item=$Books.Item($index)
            if (-not [string]::Equals([IO.Path]::GetFullPath([string]$item.FullName),$CopyPath,[StringComparison]::OrdinalIgnoreCase)) {
                $script:preserveExcel=$true
                throw 'Unexpected workbook remains; Excel preserved for user review'
            }
        } finally { $item=$null } # Borrowed alias may be the retained instrumented workbook.
    }
}
function Get-RemainingExcel {
    $processes=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
    try { return @($processes | Where-Object { -not $_.HasExited } | ForEach-Object { [int]$_.Id }) }
    finally { foreach ($item in $processes) { try { $item.Dispose() } catch {} } }
}

# Only seven current R62 fixture modules plus this dispatcher and exact AssertTrue
# projection are imported into the disposable XLAM. Only that copy enables the
# rename test constant, exposes the two unchanged picture picker state functions,
# and temporarily injects the picture rollback failure point.
$dispatchCode=@'
Option Explicit
Public Function NxR62WorkflowDispatch(ByVal family As String, ByVal kind As String, ByVal fixtureRoot As String) As String
    Dim detail As String
    On Error GoTo Failed
    Select Case family
        Case "special": T_R62DataSpecial.RunCase kind
        Case "analysis": T_R62DataAnalysis.RunCase kind
        Case "template": T_R62TemplateManager.RunCase kind
        Case "rename"
            Select Case kind
                Case "rules": detail = T_R62RenameWorkbench.T_R62RenameRules()
                Case "files": detail = T_R62RenameWorkbench.T_R62RenameFiles(fixtureRoot)
                Case "sheets": detail = T_R62RenameWorkbench.T_R62RenameSheets()
                Case "recovery": detail = T_R62RenameWorkbench.T_R62RenameRecovery(fixtureRoot)
                Case Else: Err.Raise vbObjectError + 962, , "Unknown R62 rename case"
            End Select
        Case "ai": detail = T_R62AiWorkflow.NxR62AiRunCase(kind)
        Case "shortcut": detail = T_R62ShortcutWorkflow.NxR62ShortcutRunCase(kind)
        Case "picture": detail = T_R62PictureModes.NxR62PictureModesCase(kind)
        Case Else: Err.Raise vbObjectError + 962, , "Unknown R62 workflow family"
    End Select
    If family = "rename" Or family = "ai" Or family = "shortcut" Or family = "picture" Then
        If Len(detail) = 0 Or Left$(detail, 5) <> "PASS|" Then Err.Raise vbObjectError + 962, , "Fixture returned a non-PASS result: " & detail
    End If
    NxR62WorkflowDispatch = "PASS|" & kind
    If Len(detail) > 0 Then NxR62WorkflowDispatch = NxR62WorkflowDispatch & "|" & detail
    Exit Function
Failed:
    NxR62WorkflowDispatch = "FAIL|" & kind & "|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function NxR62WorkflowCatalog() As String
    Dim name As Variant, result As String
    For Each name In T_R62DataSpecial.TestNames()
        result = result & "special." & CStr(name) & vbLf
    Next name
    For Each name In T_R62DataAnalysis.TestNames()
        result = result & "analysis." & CStr(name) & vbLf
    Next name
    For Each name In T_R62TemplateManager.TestNames()
        result = result & "template." & CStr(name) & vbLf
    Next name
    NxR62WorkflowCatalog = result
End Function

Public Function NxR62WorkflowState() As String
    NxR62WorkflowState = "forms=" & CStr(VBA.UserForms.Count) & ";wheel=" & CStr(NxFormWheelIsAttached()) & ";bindings=" & CStr(NxShortcutsBindingCount())
End Function

Public Function NxR62WorkflowTempRoot() As String
    NxR62WorkflowTempRoot = Environ$("TEMP")
End Function
'@

$started=[DateTime]::UtcNow.ToString('o')
$isolatedProfileRoot=Join-Path $EvidenceRoot 'profile'
# Native Excel SaveAs must not inherit a deeply nested evidence path.
# Keep a unique owned directory; preserve residual files for investigation.
$isolatedTempRoot=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) ('LHExcel\ValidationTemp\r62-'+[Guid]::NewGuid().ToString('N'))
$previousProfileRoot=[Environment]::GetEnvironmentVariable('LHEXCEL_PROFILE_ROOT','Process')
$previousTempRoot=[Environment]::GetEnvironmentVariable('TEMP','Process')
$previousTmpRoot=[Environment]::GetEnvironmentVariable('TMP','Process')
$copyPath=Join-Path $EvidenceRoot 'Product.r62.instrumented.xlam'
$sourceHash=$null; $sourceHashAfter=$null; $copyBeforeHash=$null; $copyAfterHash=$null; $instrumentedHash=$null
$sourceSnapshot=$null; $sourceSnapshotAfter=$null; $sourceUnchanged=$false
$excel=$null; $binding=$null; $books=$null; $copy=$null; $project=$null; $components=$null
$failure=$null; $preserveExcel=$false; $abnormalExit=$false; $owner=$null; $profileObserved=$null; $tempObserved=$null
$catalogObserved=$null
$assertionEvidence=[ordered]@{source='tests/vba/NxTestHarness.bas';source_sha256=$null;assert_true_sha256=$null;mode='NOT_RUN'}
$compile=[ordered]@{status='NOT_RUN';control_id=578;active_project=$null;dialog=$null;failure=$null}
$compileAfterReopen=[ordered]@{status='NOT_RUN';control_id=578;active_project=$null;dialog=$null;failure=$null}
$cleanup=[ordered]@{exit_mode='NOT_STARTED';failure=$null}
$cleanupErrors=New-Object 'Collections.Generic.List[string]'
$cases=New-Object 'Collections.Generic.List[object]'
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
try {
    $sourceHash=Get-Sha256 $ProductXlam
    $sourceSnapshot=Get-SourceSnapshot
    $renameSource=Join-Path $SourceRoot 'src/vba/features/file/rename/NxBatchRenameEngine.bas'
    if ([regex]::Matches([IO.File]::ReadAllText($renameSource,[Text.Encoding]::UTF8),'(?m)^#Const NX_R58_TEST_BUILD = False\r?$').Count -ne 1) { throw 'Shipped rename test marker must remain False' }
    [void](New-Item -ItemType Directory -Path $isolatedProfileRoot)
    [void](New-Item -ItemType Directory -Path $isolatedTempRoot)
    [Environment]::SetEnvironmentVariable('LHEXCEL_PROFILE_ROOT',$isolatedProfileRoot,'Process')
    [Environment]::SetEnvironmentVariable('TEMP',$isolatedTempRoot,'Process')
    [Environment]::SetEnvironmentVariable('TMP',$isolatedTempRoot,'Process')
    Copy-Item -LiteralPath $ProductXlam -Destination $copyPath -ErrorAction Stop
    $copyBeforeHash=Get-Sha256 $copyPath
    if ($copyBeforeHash -cne $sourceHash) { throw 'R62 disposable copy SHA256 mismatch' }
    $baseline=Get-ExcelProcessBaseline
    if (@($baseline.process_ids).Count) { throw 'Excel appeared during preflight; ownership rejected' }
    $excel=New-Object -ComObject Excel.Application
    $binding=Get-ExactExcelProcessOwnership $excel $baseline 'R62 workflow Excel'
    if (-not $binding.owned) { $preserveExcel=$true; throw ('R62 Excel ownership rejected: '+$binding.failure) }
    $owner=[ordered]@{pid=[int]$binding.pid;hwnd=[Int64]$binding.hwnd;started_utc=$binding.started_utc;executable=$binding.executable;excel_version=[string]$excel.Version}
    Write-Json (Join-Path $EvidenceRoot 'R62Workflow.owned-process.json') ([ordered]@{schema_version=1;suite='R62Workflow';status='OWNED';run_id=$RunId;owner=$owner;source_product_sha256=$sourceHash;copy_before_sha256=$copyBeforeHash;instrumented_copy=$copyPath})
    $books=$excel.Workbooks
    if ($books.Count -ne 0) { $preserveExcel=$true; throw 'Unexpected startup workbook; Excel preserved for user review' }
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
    $copy=$books.Open($copyPath,$false,$false)
    if ($copy.ReadOnly -or -not $copy.IsAddin -or -not [string]::Equals([IO.Path]::GetFullPath([string]$copy.FullName),$copyPath,[StringComparison]::OrdinalIgnoreCase)) { throw 'Disposable add-in identity or writable state rejected' }
    $project=$copy.VBProject; $components=$project.VBComponents
    Set-RenameProbeConst $components
    Enable-PicturePickerStateProbe $components
    Ensure-AssertHarness $components
    foreach ($probe in @(
        @('tests/vba/data/T_R62DataSpecial.bas','T_R62DataSpecial'),
        @('tests/vba/data/T_R62DataAnalysis.bas','T_R62DataAnalysis'),
        @('tests/vba/file/T_R62RenameWorkbench.bas','T_R62RenameWorkbench'),
        @('tests/vba/template/T_R62TemplateManager.bas','T_R62TemplateManager'),
        @('tests/vba/product/T_R62AiWorkflow.bas','T_R62AiWorkflow'),
        @('tests/vba/product/T_R62ShortcutWorkflow.bas','T_R62ShortcutWorkflow'),
        @('tests/vba/draw/T_R62PictureModes.bas','T_R62PictureModes')
    )) {
        $probeCode=[IO.File]::ReadAllText((Join-Path $SourceRoot $probe[0]),[Text.Encoding]::UTF8)
        if ($probe[1] -ceq 'T_R62RenameWorkbench') {
            if ([regex]::Matches($probeCode,'(?m)^#Const NX_R58_TEST_BUILD = False\r?$').Count -ne 1) { throw 'Rename fixture test marker absent or ambiguous; production source must remain False' }
            $probeCode=$probeCode.Replace('#Const NX_R58_TEST_BUILD = False','#Const NX_R58_TEST_BUILD = True')
        }
        Add-ProbeModule $components $probe[1] $probeCode
    }
    Add-ProbeModule $components 'R62WorkflowDispatch' $dispatchCode
    Invoke-ProbeCompile $excel $binding $components $copyPath
    $initialCompileEvidence=$compile
    $copy.Save()
    [void](Assert-ExactExcelProcessOwnership $binding)
    Assert-OnlyProbeWorkbook $books $copyPath
    # Excel retains a sharing lock on the saved add-in while it is open.
    # Hash only after closing the exact owned copy and releasing its references.
    Release-Com $components; $components=$null
    Release-Com $project; $project=$null
    $copy.Close($false)
    Release-Com $copy; $copy=$null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $copyAfterHash=Get-Sha256 $copyPath
    $instrumentedHash=$copyAfterHash
    [void](Assert-ExactExcelProcessOwnership $binding)
    $copy=$books.Open($copyPath,$false,$false)
    if ($copy.ReadOnly -or -not $copy.IsAddin -or -not [string]::Equals([IO.Path]::GetFullPath([string]$copy.FullName),$copyPath,[StringComparison]::OrdinalIgnoreCase)) { throw 'Reopened disposable add-in identity or writable state rejected' }
    $project=$copy.VBProject; $components=$project.VBComponents
    try { Invoke-ProbeCompile $excel $binding $components $copyPath -AllowAlreadyCompiled }
    finally { $compileAfterReopen=$compile; $compile=$initialCompileEvidence }
    $macro="'"+$copy.Name.Replace("'","''")+"'!"
    $profileObserved=[string]$excel.Run($macro+'NxLHexcelProfileRoot')
    if (-not [string]::Equals([IO.Path]::GetFullPath($profileObserved),$isolatedProfileRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'R62 isolated profile root verification failed' }
    $tempObserved=[string]$excel.Run($macro+'NxR62WorkflowTempRoot')
    if (-not [string]::Equals([IO.Path]::GetFullPath($tempObserved),$isolatedTempRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'R62 isolated TEMP verification failed' }
    $catalogObserved=[string]$excel.Run($macro+'NxR62WorkflowCatalog')
    $catalogNames=@($catalogObserved.Split([char]10) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $catalogExpected=@($allowedCases | Where-Object { $_ -match '^(special|analysis|template)\.' })
    if (($catalogNames -join "`n") -cne ($catalogExpected -join "`n")) { throw 'Native TestNames differs from closed R62 runner catalog' }
    $excel.EnableEvents=$true; $excel.Visible=$true
    Assert-OnlyProbeWorkbook $books $copyPath
    foreach ($name in $requestedCases) {
        $parts=$name.Split('.',2); $family=$parts[0]; $kind=$parts[1]
        $casePath=Join-Path $EvidenceRoot ('R62Workflow.case.'+$name+'.json')
        $case=[ordered]@{schema_version=1;suite='R62Workflow';run_id=$RunId;name=$name;status='RUNNING';phase='started';started_utc=[DateTime]::UtcNow.ToString('o');completed_utc=$null;evidence_class='native_vba_disposable_fixture';source_product_sha256=$sourceHash;instrumented_copy_sha256=$copyAfterHash;source_tree_sha256=$sourceSnapshot.sha256;result=$null;state=$null;failure=$null;receipt_path=$casePath}
        Write-Json $casePath $case
        Write-Output ('START|'+$name)
        try {
            [void](Assert-ExactExcelProcessOwnership $binding)
            Assert-OnlyProbeWorkbook $books $copyPath
            if ($name -ceq 'picture.engine_rollback') {
                $case.result=Invoke-PictureRollbackProbe $excel $components $macro $isolatedTempRoot
            } else {
                $case.result=[string]$excel.Run($macro+'NxR62WorkflowDispatch',$family,$kind,$isolatedTempRoot)
            }
            if ($case.result -cnotmatch ('^PASS\|'+[regex]::Escape($kind)+'(?:\||$)')) { throw ('R62 fixture failed: '+$case.result) }
            Assert-OnlyProbeWorkbook $books $copyPath
            $case.state=[string]$excel.Run($macro+'NxR62WorkflowState')
            if ($case.state -cne 'forms=0;wheel=False;bindings=0') { throw ('R62 fixture lifecycle leaked state: '+$case.state) }
            # Rename fixtures intentionally retain uniquely owned files as evidence.
            # Workbook/form leaks still fail; do not delete evidence or infer pass from residual names.
            [void](Assert-ExactExcelProcessOwnership $binding)
            $case.status='PASS'
        } catch { $case.status='FAIL'; $case.failure=$_.Exception.Message }
        $case.phase='completed'; $case.completed_utc=[DateTime]::UtcNow.ToString('o')
        $cases.Add([pscustomobject]$case); Write-Json $casePath $case
        Write-Output ($case.status+'|'+$name+'|'+$case.result)
        if ($case.status -cne 'PASS') { throw ('R62 case failed: '+$name+'; '+$case.failure) }
    }
} catch { $failure=$_.Exception.Message }
finally {
    if ($null -ne $binding -and $binding.owned) {
        try {
            if ($binding.process.HasExited) { $abnormalExit=$true; $cleanupErrors.Add('Owned Excel exited before cleanup') }
            else { [void](Assert-ExactExcelProcessOwnership $binding); if ($null -ne $books) { Assert-OnlyProbeWorkbook $books $copyPath } }
        } catch { $preserveExcel=$true; $cleanupErrors.Add($_.Exception.Message) }
    }
    if ($null -ne $copy -and -not $preserveExcel -and -not $abnormalExit) {
        try { $copy.Close($false) } catch { $cleanupErrors.Add('Probe close: '+$_.Exception.Message) }
    }
    Release-Com $components; Release-Com $project; Release-Com $copy; Release-Com $books
    if ($null -ne $excel) {
        if ($null -ne $binding -and $binding.owned -and -not $preserveExcel -and -not $abnormalExit) {
            try { $excel.Quit() } catch { $cleanupErrors.Add('Excel Quit: '+$_.Exception.Message) }
        } elseif ($null -ne $binding -and $binding.owned -and $preserveExcel) {
            try { $excel.DisplayAlerts=$true; $excel.Visible=$true } catch {}
        }
        Release-Com $excel
    }
    $components=$null; $project=$null; $copy=$null; $books=$null; $excel=$null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($null -ne $binding -and $binding.owned -and $null -ne $binding.process) {
        try {
            if ($abnormalExit) { $cleanup=[ordered]@{exit_mode='ABNORMAL_EXIT';failure='Excel exited before cleanup'} }
            elseif ($preserveExcel) { $cleanup=[ordered]@{exit_mode='PRESERVED_UNEXPECTED_STATE';failure='User review required; no Quit or kill'} }
            else { $cleanup=Stop-ExactProcessAfterGrace $binding.process 'R62 workflow Excel' 15000 -Detailed }
        } catch { $cleanup=[ordered]@{exit_mode='CLEANUP_FAILED';failure=$_.Exception.Message} }
        finally { $binding.process.Dispose() }
    }
    [Environment]::SetEnvironmentVariable('LHEXCEL_PROFILE_ROOT',$previousProfileRoot,'Process')
    [Environment]::SetEnvironmentVariable('TEMP',$previousTempRoot,'Process')
    [Environment]::SetEnvironmentVariable('TMP',$previousTmpRoot,'Process')
}
$remainingImmediate=@(Get-RemainingExcel)
$watch=[Diagnostics.Stopwatch]::StartNew()
$remaining=@($remainingImmediate)
while ($remaining.Count -gt 0 -and $watch.ElapsedMilliseconds -lt 2000) { Start-Sleep -Milliseconds 100; $remaining=@(Get-RemainingExcel) }
try {
    $sourceHashAfter=Get-Sha256 $ProductXlam
    $sourceSnapshotAfter=Get-SourceSnapshot
    $sourceUnchanged=($null -ne $sourceSnapshot -and $sourceHash -ceq $sourceHashAfter -and $sourceSnapshot.sha256 -ceq $sourceSnapshotAfter.sha256)
    if (-not $sourceUnchanged) { throw 'Supplied product or source tree changed during R62 workflow' }
    if ($null -ne $instrumentedHash) {
        $copyAfterHash=Get-Sha256 $copyPath
        if ($copyAfterHash -cne $instrumentedHash) { throw 'Instrumented copy changed after the bound case snapshot' }
    }
} catch { if ($null -eq $failure) { $failure=$_.Exception.Message }; $sourceUnchanged=$false }
$retainedFixtureFiles=@()
try {
    if (Test-Path -LiteralPath $isolatedTempRoot -PathType Container) {
        $retainedFixtureFiles=@(Get-ChildItem -LiteralPath $isolatedTempRoot -File -Recurse | ForEach-Object {
            [ordered]@{path=$_.FullName;length=$_.Length;sha256=(Get-Sha256 $_.FullName)}
        })
    }
} catch { if ($null -eq $failure) { $failure='Fixture evidence inventory failed: '+$_.Exception.Message } }
$passed=@($cases | Where-Object status -eq 'PASS').Count
$cleanupNatural=($cleanup.exit_mode -ceq 'NATURAL' -and [string]::IsNullOrWhiteSpace([string]$cleanup.failure))
$selectedPassed=($null -eq $failure -and $cleanupErrors.Count -eq 0 -and $cases.Count -eq $requestedCases.Count -and $passed -eq $requestedCases.Count -and
    $remaining.Count -eq 0 -and $cleanupNatural -and -not $abnormalExit -and $compile.status -ceq 'PASS' -and
    $compileAfterReopen.status -cin @('PASS','COMPILE_RETAINED') -and $sourceUnchanged)
$status=if (-not $selectedPassed) { 'FAIL' } elseif ($requestedCases.Count -eq $allowedCases.Count) { 'PASS' } else { 'PARTIAL' }
$executedNames=@($cases | ForEach-Object { $_.name })
$notRunCases=@($allowedCases | Where-Object { $executedNames -cnotcontains $_ } | ForEach-Object { [ordered]@{name=$_;status='NOT_RUN'} })
$receipt=[ordered]@{
    schema_version=1;suite='R62Workflow';status=$status;run_id=$RunId;started_utc=$started;completed_utc=[DateTime]::UtcNow.ToString('o')
    execution_scope=if($requestedCases.Count -eq $allowedCases.Count){'FULL_WORKFLOW'}else{'SELECTED_CASES'};expected_case_count=$allowedCases.Count;requested_cases=$requestedCases;not_run_cases=$notRunCases;cases=@($cases.ToArray())
    failure=$failure;owner=$owner;isolated_profile_root=$isolatedProfileRoot;profile_root_observed=$profileObserved;isolated_temp_root=$isolatedTempRoot;temp_root_observed=$tempObserved;compile=$compile;compile_after_reopen=$compileAfterReopen
    cleanup=$cleanup;cleanup_errors=@($cleanupErrors.ToArray());abnormal_exit=$abnormalExit;remaining_excel_immediate=$remainingImmediate;remaining_excel=$remaining
    source_product_sha256=$sourceHash;source_product_sha256_after=$sourceHashAfter;copy_before_sha256=$copyBeforeHash;copy_after_sha256=$copyAfterHash;instrumented_copy=$copyPath
    source_tree_sha256=if($null -ne $sourceSnapshot){$sourceSnapshot.sha256}else{$null};source_tree_sha256_after=if($null -ne $sourceSnapshotAfter){$sourceSnapshotAfter.sha256}else{$null}
    source_files_sha256=if($null -ne $sourceSnapshot){$sourceSnapshot.files}else{$null};source_unchanged=$sourceUnchanged;assertion_fixture=$assertionEvidence;native_test_catalog=$catalogObserved
    retained_fixture_files=$retainedFixtureFiles;test_instrumentation=@('NxBatchRenameEngine.NX_R58_TEST_BUILD=True','T_R62RenameWorkbench.NX_R58_TEST_BUILD=True','disposable copy only: exact private signatures of RestorePictureInsertPickerState and PictureInsertPickerStateMatches made Public; bodies unchanged','picture.engine_rollback only: unsaved test-copy NxPictureLayout post-width failure injection, restored in finally')
    not_run=@('compare.navigation and compare.failures: external real-report-open orchestration requires separate runner; NOT_RUN','physical mouse/keyboard input and accessibility','ProductUi, R58Convergence and document navigator regression suites','install, COM registration, signing and release approval')
    scope='owned Windows Excel, disposable instrumented XLAM and isolated LHEXCEL_PROFILE_ROOT; native VBA fixture assertions, not physical UI acceptance'
}
Write-Json (Join-Path $EvidenceRoot 'R62Workflow.json') $receipt
if (-not $selectedPassed) { throw ('R62 workflow failed; receipt='+ (Join-Path $EvidenceRoot 'R62Workflow.json')) }
Write-Output ($status+'|R62Workflow|'+$passed+'/'+$allowedCases.Count)
