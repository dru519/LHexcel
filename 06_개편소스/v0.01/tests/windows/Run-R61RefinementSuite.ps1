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
    'sheet.validation','sheet.cycles','sheet.workbook_drift','sheet.sheet_drift','sheet.rollback','sheet.editor_cancel','sheet.editor_lifecycle',
    'style.palette','style.axes','style.drift','style.conflict_cancel',
    'shortcut.categories','shortcut.category_search','shortcut.empty_results','shortcut.reselection','shortcut.description','shortcut.binding_display','shortcut.lifecycle',
    'file.folder_hierarchy','file.row_consolidation','file.chart_jpg'
)
function Resolve-R61Cases([string]$Value,[string[]]$Allowed) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Allowed }
    $selected = @($Value.Split(',') | ForEach-Object { $_.Trim() })
    if (-not $selected.Count -or @($selected | Where-Object { $Allowed -cnotcontains $_ }).Count -or
        @($selected | Select-Object -Unique).Count -ne $selected.Count) { throw 'Unknown, empty, or duplicate R61 refinement case' }
    return $selected
}
$requestedCases = @(Resolve-R61Cases $CaseNames $allowedCases)
if (-not (Test-Path -LiteralPath $ProductXlam -PathType Leaf)) { throw 'Product XLAM unavailable' }
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing R61 refinement evidence is rejected' }
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'R61 refinement requires zero pre-existing Excel processes' }
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
function Ensure-AssertHarness([object]$Components) {
    # Full NxTestHarness.RunSuite statically depends on unrelated T_Core tests.
    # Project only the exact canonical assertion used by the three T_File cases.
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
function Invoke-ProbeCompile([object]$Excel,[object]$Binding,[object]$Components,[string]$CopyPath,[switch]$AllowAlreadyCompiled) {
    $vbe=$null; $bars=$null; $control=$null; $target=$null; $active=$null; $window=$null; $watcher=$null
    $script:compile=[ordered]@{status='RUNNING';control_id=578;active_project=$null;dialog=$null;failure=$null}
    try {
        [void](Assert-ExactExcelProcessOwnership $Binding)
        $vbe=$Excel.VBE; $bars=$vbe.CommandBars; $control=Find-VbeCompileControl $bars
        if ($null -eq $control -or [int]$control.Id -ne 578) { throw 'R61 VBA compile control 578 unavailable' }
        $target=$Components.Item('R61RefinementDispatch'); [void]$target.Activate()
        $active=$vbe.ActiveVBProject; $activePath=[IO.Path]::GetFullPath([string]$active.FileName)
        $script:compile.active_project=$activePath
        if (-not [string]::Equals($activePath,$CopyPath,[StringComparison]::OrdinalIgnoreCase)) { throw 'R61 compile active project identity mismatch' }
        if (-not [bool]$control.Enabled) {
            if ($AllowAlreadyCompiled) { $script:compile.status='COMPILE_RETAINED'; return }
            throw 'R61 compile was disabled before required instrumentation compile'
        }
        $window=$vbe.MainWindow
        $watcher=Start-ExactOwnedCompileDialogWatcher -ExcelPid ([int]$Binding.pid) -ExcelHwnd ([Int64]$Binding.hwnd) -VbeHwnd ([Int64]$window.HWnd)
        $executeFailure=$null
        try { [void]$control.Execute() } catch { $executeFailure=$_.Exception.Message }
        if ($null -eq (Wait-Job -Job $watcher -Timeout 30)) { throw 'R61 compile watcher timed out' }
        $observed=@(Receive-Job -Job $watcher -Wait -AutoRemoveJob -ErrorAction Stop); $watcher=$null
        if ($null -ne $executeFailure) { throw ('R61 compile execution failed: '+$executeFailure) }
        if ($observed.Count -ne 1 -or $observed[0].status -cne 'NO_DIALOG') {
            $script:compile.dialog=@($observed)
            throw 'R61 compile dialog detected or watcher result unavailable'
        }
        if ([bool]$control.Enabled) { throw 'R61 compile did not reach disabled state' }
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

# This static dispatcher is imported only into the copied XLAM. Production APIs
# create/remove the isolated binding; no source or real-user settings are edited.
$dispatchCode=@'
Option Explicit
Public Function NxR61RefinementDispatch(ByVal family As String, ByVal kind As String) As String
    Dim result As String, assigned As Boolean, routeKey As String, cleanupFailure As String
    On Error GoTo Failed
    Select Case family
        Case "sheet": result = T_R61SheetNameEditor.NxR61SheetNameEditorRunCase(kind)
        Case "style": result = T_R61RoleStylePreview.NxR61RoleStylePreviewCase(kind)
        Case "file"
            Select Case kind
                Case "folder_hierarchy": T_File.RunCase "TestFolderHierarchyFromRange"
                Case "row_consolidation": T_File.RunCase "TestRowsConsolidationPreservesOrderAndLiteralText"
                Case "chart_jpg": T_File.RunCase "TestChartJpgHasJpegSignature"
                Case Else: Err.Raise vbObjectError + 761, , "Unknown R61 file case"
            End Select
            result = "PASS|" & kind
        Case "shortcut"
            If kind = "binding_display" Then
                If NxShortcutsBindingCount() <> 0 Then Err.Raise vbObjectError + 761, , "Isolated profile binding inventory was not empty"
                routeKey = "feature:NX-DATA-UNIQUE-COUNT"
                NxShortcutsAssign routeKey, "Ctrl+Alt+Shift", "F12"
                assigned = True
            End If
            result = T_R61ShortcutCatalog.NxR61ShortcutCatalogRunCase(kind, routeKey)
        Case Else: Err.Raise vbObjectError + 761, , "Unknown R61 family"
    End Select
    GoTo CleanUp
Failed:
    result = "FAIL|" & kind & "|" & CStr(Err.Number) & "|" & Err.Description
CleanUp:
    On Error Resume Next
    If assigned Then
        Err.Clear
        NxShortcutsRemoveRoute routeKey
        If Err.Number <> 0 Then cleanupFailure = Err.Description
    End If
    On Error GoTo 0
    If Len(cleanupFailure) > 0 Then result = "FAIL|" & kind & "|binding cleanup|" & cleanupFailure
    NxR61RefinementDispatch = result
End Function

Public Function NxR61RefinementState() As String
    NxR61RefinementState = "forms=" & CStr(VBA.UserForms.Count) & ";wheel=" & CStr(NxFormWheelIsAttached()) & ";bindings=" & CStr(NxShortcutsBindingCount())
End Function

Public Function NxR61RefinementTempRoot() As String
    NxR61RefinementTempRoot = Environ$("TEMP")
End Function
'@

$started=[DateTime]::UtcNow.ToString('o')
$isolatedProfileRoot=Join-Path $EvidenceRoot 'profile'
# Native Excel SaveAs must not inherit a deeply nested evidence path.
# Keep a unique owned directory; preserve residual files for investigation.
$isolatedTempRoot=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) ('LHExcel\ValidationTemp\r61-'+[Guid]::NewGuid().ToString('N'))
$previousProfileRoot=[Environment]::GetEnvironmentVariable('LHEXCEL_PROFILE_ROOT','Process')
$previousTempRoot=[Environment]::GetEnvironmentVariable('TEMP','Process')
$previousTmpRoot=[Environment]::GetEnvironmentVariable('TMP','Process')
$copyPath=Join-Path $EvidenceRoot 'Product.r61.instrumented.xlam'
$sourceHash=$null; $sourceHashAfter=$null; $copyBeforeHash=$null; $copyAfterHash=$null; $instrumentedHash=$null
$sourceSnapshot=$null; $sourceSnapshotAfter=$null; $sourceUnchanged=$false
$excel=$null; $binding=$null; $books=$null; $copy=$null; $project=$null; $components=$null
$failure=$null; $preserveExcel=$false; $abnormalExit=$false; $owner=$null; $profileObserved=$null; $tempObserved=$null
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
    if ($copyBeforeHash -cne $sourceHash) { throw 'R61 disposable copy SHA256 mismatch' }
    $baseline=Get-ExcelProcessBaseline
    if (@($baseline.process_ids).Count) { throw 'Excel appeared during preflight; ownership rejected' }
    $excel=New-Object -ComObject Excel.Application
    $binding=Get-ExactExcelProcessOwnership $excel $baseline 'R61 refinement Excel'
    if (-not $binding.owned) { $preserveExcel=$true; throw ('R61 Excel ownership rejected: '+$binding.failure) }
    $owner=[ordered]@{pid=[int]$binding.pid;hwnd=[Int64]$binding.hwnd;started_utc=$binding.started_utc;executable=$binding.executable;excel_version=[string]$excel.Version}
    Write-Json (Join-Path $EvidenceRoot 'R61Refinement.owned-process.json') ([ordered]@{schema_version=1;suite='R61Refinement';status='OWNED';run_id=$RunId;owner=$owner;source_product_sha256=$sourceHash;copy_before_sha256=$copyBeforeHash;instrumented_copy=$copyPath})
    $books=$excel.Workbooks
    if ($books.Count -ne 0) { $preserveExcel=$true; throw 'Unexpected startup workbook; Excel preserved for user review' }
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
    $copy=$books.Open($copyPath,$false,$false)
    if ($copy.ReadOnly -or -not $copy.IsAddin -or -not [string]::Equals([IO.Path]::GetFullPath([string]$copy.FullName),$copyPath,[StringComparison]::OrdinalIgnoreCase)) { throw 'Disposable add-in identity or writable state rejected' }
    $project=$copy.VBProject; $components=$project.VBComponents
    Set-RenameProbeConst $components
    Ensure-AssertHarness $components
    foreach ($probe in @(
        @('tests/vba/file/T_R61SheetNameEditor.bas','T_R61SheetNameEditor'),
        @('tests/vba/draw/T_R61RoleStylePreview.bas','T_R61RoleStylePreview'),
        @('tests/vba/product/T_R61ShortcutCatalog.bas','T_R61ShortcutCatalog'),
        @('tests/vba/file/T_File.bas','T_File')
    )) { Add-ProbeModule $components $probe[1] ([IO.File]::ReadAllText((Join-Path $SourceRoot $probe[0]),[Text.Encoding]::UTF8)) }
    Add-ProbeModule $components 'R61RefinementDispatch' $dispatchCode
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
    if (-not [string]::Equals([IO.Path]::GetFullPath($profileObserved),$isolatedProfileRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'R61 isolated profile root verification failed' }
    $tempObserved=[string]$excel.Run($macro+'NxR61RefinementTempRoot')
    if (-not [string]::Equals([IO.Path]::GetFullPath($tempObserved),$isolatedTempRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'R61 isolated TEMP verification failed' }
    $excel.EnableEvents=$true; $excel.Visible=$true
    Assert-OnlyProbeWorkbook $books $copyPath
    foreach ($name in $requestedCases) {
        $parts=$name.Split('.',2); $family=$parts[0]; $kind=$parts[1]
        $casePath=Join-Path $EvidenceRoot ('R61Refinement.case.'+$name+'.json')
        $case=[ordered]@{schema_version=1;suite='R61Refinement';run_id=$RunId;name=$name;status='RUNNING';phase='started';started_utc=[DateTime]::UtcNow.ToString('o');completed_utc=$null;evidence_class='native_vba_disposable_fixture';source_product_sha256=$sourceHash;instrumented_copy_sha256=$copyAfterHash;source_tree_sha256=$sourceSnapshot.sha256;result=$null;state=$null;failure=$null;receipt_path=$casePath}
        Write-Json $casePath $case
        Write-Output ('START|'+$name)
        try {
            [void](Assert-ExactExcelProcessOwnership $binding)
            Assert-OnlyProbeWorkbook $books $copyPath
            $case.result=[string]$excel.Run($macro+'NxR61RefinementDispatch',$family,$kind)
            if ($case.result -cnotmatch ('^PASS\|'+[regex]::Escape($kind)+'(?:\||$)')) { throw ('R61 fixture failed: '+$case.result) }
            Assert-OnlyProbeWorkbook $books $copyPath
            $case.state=[string]$excel.Run($macro+'NxR61RefinementState')
            if ($case.state -cne 'forms=0;wheel=False;bindings=0') { throw ('R61 fixture lifecycle leaked state: '+$case.state) }
            if (@(Get-ChildItem -LiteralPath $isolatedTempRoot -File -Recurse | Where-Object { $_.Name -like 'LHexcel-*' }).Count) { throw 'R61 fixture files remain; residual artifacts preserved for review' }
            [void](Assert-ExactExcelProcessOwnership $binding)
            $case.status='PASS'
        } catch { $case.status='FAIL'; $case.failure=$_.Exception.Message }
        $case.phase='completed'; $case.completed_utc=[DateTime]::UtcNow.ToString('o')
        $cases.Add([pscustomobject]$case); Write-Json $casePath $case
        Write-Output ($case.status+'|'+$name+'|'+$case.result)
        if ($case.status -cne 'PASS') { throw ('R61 case failed: '+$name+'; '+$case.failure) }
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
            else { $cleanup=Stop-ExactProcessAfterGrace $binding.process 'R61 refinement Excel' 15000 -Detailed }
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
    if (-not $sourceUnchanged) { throw 'Supplied product or source tree changed during R61 refinement' }
    if ($null -ne $instrumentedHash) {
        $copyAfterHash=Get-Sha256 $copyPath
        if ($copyAfterHash -cne $instrumentedHash) { throw 'Instrumented copy changed after the bound case snapshot' }
    }
} catch { if ($null -eq $failure) { $failure=$_.Exception.Message }; $sourceUnchanged=$false }
$passed=@($cases | Where-Object status -eq 'PASS').Count
$cleanupNatural=($cleanup.exit_mode -ceq 'NATURAL' -and [string]::IsNullOrWhiteSpace([string]$cleanup.failure))
$selectedPassed=($null -eq $failure -and $cleanupErrors.Count -eq 0 -and $cases.Count -eq $requestedCases.Count -and $passed -eq $requestedCases.Count -and
    $remaining.Count -eq 0 -and $cleanupNatural -and -not $abnormalExit -and $compile.status -ceq 'PASS' -and
    $compileAfterReopen.status -cin @('PASS','COMPILE_RETAINED') -and $sourceUnchanged)
$status=if (-not $selectedPassed) { 'FAIL' } elseif ($requestedCases.Count -eq $allowedCases.Count) { 'PASS' } else { 'PARTIAL' }
$executedNames=@($cases | ForEach-Object { $_.name })
$notRunCases=@($allowedCases | Where-Object { $executedNames -cnotcontains $_ } | ForEach-Object { [ordered]@{name=$_;status='NOT_RUN'} })
$receipt=[ordered]@{
    schema_version=1;suite='R61Refinement';status=$status;run_id=$RunId;started_utc=$started;completed_utc=[DateTime]::UtcNow.ToString('o')
    execution_scope=if($requestedCases.Count -eq $allowedCases.Count){'FULL_REFINEMENT'}else{'SELECTED_CASES'};expected_case_count=$allowedCases.Count;requested_cases=$requestedCases;not_run_cases=$notRunCases;cases=@($cases.ToArray())
    failure=$failure;owner=$owner;isolated_profile_root=$isolatedProfileRoot;profile_root_observed=$profileObserved;isolated_temp_root=$isolatedTempRoot;temp_root_observed=$tempObserved;compile=$compile;compile_after_reopen=$compileAfterReopen
    cleanup=$cleanup;cleanup_errors=@($cleanupErrors.ToArray());abnormal_exit=$abnormalExit;remaining_excel_immediate=$remainingImmediate;remaining_excel=$remaining
    source_product_sha256=$sourceHash;source_product_sha256_after=$sourceHashAfter;copy_before_sha256=$copyBeforeHash;copy_after_sha256=$copyAfterHash;instrumented_copy=$copyPath
    source_tree_sha256=if($null -ne $sourceSnapshot){$sourceSnapshot.sha256}else{$null};source_tree_sha256_after=if($null -ne $sourceSnapshotAfter){$sourceSnapshotAfter.sha256}else{$null}
    source_files_sha256=if($null -ne $sourceSnapshot){$sourceSnapshot.files}else{$null};source_unchanged=$sourceUnchanged;assertion_fixture=$assertionEvidence
    not_run=@('physical mouse/keyboard input and accessibility','ProductUi, R58Convergence and document navigator regression suites','install, COM registration, signing and release approval')
    scope='owned Windows Excel, disposable instrumented XLAM and isolated LHEXCEL_PROFILE_ROOT; native VBA fixture assertions, not physical UI acceptance'
}
Write-Json (Join-Path $EvidenceRoot 'R61Refinement.json') $receipt
if (-not $selectedPassed) { throw ('R61 refinement failed; receipt='+ (Join-Path $EvidenceRoot 'R61Refinement.json')) }
Write-Output ($status+'|R61Refinement|'+$passed+'/'+$allowedCases.Count)
