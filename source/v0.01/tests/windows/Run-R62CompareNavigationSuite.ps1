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
if (-not (Test-Path -LiteralPath $ProductXlam -PathType Leaf)) { throw 'Product XLAM unavailable' }
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing R62 compare evidence is rejected' }
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'R62 compare requires zero pre-existing Excel processes' }
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

# Lifecycle, compile watcher, COM ownership and snapshot helpers match Run-R62WorkflowSuite.
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
function Invoke-ProbeCompile([object]$Excel,[object]$Binding,[object]$Components,[string]$CopyPath,[switch]$AllowAlreadyCompiled) {
    $vbe=$null; $bars=$null; $control=$null; $target=$null; $active=$null; $window=$null; $watcher=$null
    $script:compile=[ordered]@{status='RUNNING';control_id=578;active_project=$null;dialog=$null;failure=$null}
    try {
        [void](Assert-ExactExcelProcessOwnership $Binding)
        $vbe=$Excel.VBE; $bars=$vbe.CommandBars; $control=Find-VbeCompileControl $bars
        if ($null -eq $control -or [int]$control.Id -ne 578) { throw 'R62 VBA compile control 578 unavailable' }
        $target=$Components.Item('R62CompareDispatch'); [void]$target.Activate()
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
            throw 'R62 compile dialog detected or watcher result unavailable'
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


function Assert-OnlyCompareWorkbooks([object]$Books,[string[]]$AllowedPaths) {
    for ($index=1; $index -le $Books.Count; $index++) {
        $item=$null
        try {
            $item=$Books.Item($index)
            $itemPath=[IO.Path]::GetFullPath([string]$item.FullName)
            if (-not @($AllowedPaths | Where-Object { [string]::Equals($_,$itemPath,[StringComparison]::OrdinalIgnoreCase) }).Count) {
                $script:preserveExcel=$true
                throw 'Unexpected workbook remains; Excel preserved for user review'
            }
        } finally { $item=$null } # Borrowed aliases must not invalidate retained copy/report RCWs.
    }
}
function Assert-OwnedFixturePath([string]$Path,[string]$Kind) {
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Empty compare fixture path' }
    $full=[IO.Path]::GetFullPath($Path)
    if (-not [string]::Equals([IO.Path]::GetDirectoryName($full),$isolatedTempRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'Compare fixture escaped isolated TEMP' }
    if ([IO.Path]::GetFileName($full) -cnotmatch ('^LHexcel-r62-compare-'+[regex]::Escape($Kind)+'-[0-9a-fA-F]{32}\.xlsx$')) { throw 'Compare fixture identity rejected' }
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw 'Owned compare fixture is missing' }
    if ((Get-Item -LiteralPath $full -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Compare fixture reparse point rejected' }
    return $full
}
function Open-OwnedReport([string]$Path) {
    [void](Assert-ExactExcelProcessOwnership $binding)
    $priorSecurity=[int]$excel.AutomationSecurity
    $opened=$null
    try {
        $excel.AutomationSecurity=3
        $opened=$books.Open($Path,0,$true)
        if (-not $opened.ReadOnly -or [int]$opened.FileFormat -ne 51 -or
            -not [string]::Equals([IO.Path]::GetFullPath([string]$opened.FullName),$Path,[StringComparison]::OrdinalIgnoreCase)) {
            throw 'Externally reopened report identity, xlsx format or read-only state rejected'
        }
        return $opened
    } catch {
        if ($null -ne $opened) { try { $opened.Close($false) } finally { Release-Com $opened } }
        throw
    } finally { $excel.AutomationSecurity=$priorSecurity }
}
function Read-ReportSourcePaths([object]$Report) {
    $sheets=$null; $sheet=$null; $base=$null; $compared=$null
    try {
        # This BOM-free runner must remain invariant under Windows PowerShell 5.1 ANSI decoding.
        $summaryName=-join ([char[]]@(0xBE44,0xAD50,0xC694,0xC57D))
        $sheets=$Report.Worksheets; $sheet=$sheets.Item($summaryName)
        $base=$sheet.Range('B2'); $compared=$sheet.Range('B3')
        if ($base.HasFormula -or $compared.HasFormula) { throw 'Source metadata must be literal' }
        return [ordered]@{base=(Assert-OwnedFixturePath ([string]$base.Value2) 'base');compare=(Assert-OwnedFixturePath ([string]$compared.Value2) 'compare')}
    } finally { Release-Com $compared; Release-Com $base; Release-Com $sheet; Release-Com $sheets }
}
function Invoke-CompareCase([string]$Name,[scriptblock]$Action) {
    $casePath=Join-Path $EvidenceRoot ('R62CompareNavigation.case.'+$Name+'.json')
    $entry=[ordered]@{schema_version=1;suite='R62CompareNavigation';run_id=$RunId;name=$Name;status='RUNNING';started_utc=[DateTime]::UtcNow.ToString('o');completed_utc=$null;evidence_class='native_vba_disposable_fixture';source_product_sha256=$sourceHash;instrumented_copy_sha256=$instrumentedHash;source_tree_sha256=$sourceSnapshot.sha256;result=$null;failure=$null;failure_location=$null;receipt_path=$casePath}
    Write-Json $casePath $entry
    Write-Output ('START|'+$Name)
    try {
        [void](Assert-ExactExcelProcessOwnership $binding)
        $entry.result=& $Action
        [void](Assert-ExactExcelProcessOwnership $binding)
        $entry.status='PASS'
    } catch {
        $entry.status='FAIL'; $entry.failure=$_.Exception.Message
        $entry.failure_location=[ordered]@{script_name=$_.InvocationInfo.ScriptName;script_line_number=$_.InvocationInfo.ScriptLineNumber;position_message=$_.InvocationInfo.PositionMessage;script_stack_trace=$_.ScriptStackTrace;fully_qualified_error_id=$_.FullyQualifiedErrorId}
    }
    $entry.completed_utc=[DateTime]::UtcNow.ToString('o')
    $cases.Add([pscustomobject]$entry); Write-Json $casePath $entry
    Write-Output ($entry.status+'|'+$Name)
    if ($entry.status -cne 'PASS') { throw ('R62 compare case failed: '+$Name+'; '+$entry.failure) }
}
function Invoke-CompareFixture([string]$Kind) {
    $result=[string]$excel.Run($macro+'NxR62CompareDispatch',$Kind,$script:reportPath)
    if ($result -cne ('PASS|'+$Kind)) { throw ('R62 compare fixture failed: '+$result) }
    return $result
}

# Test-copy checkpoints locate a native crash without changing the published source.
# Unsafe external hyperlinks are NEVER followed in this suite.
function Add-CompareTrace([object]$Components,[string]$ModuleName,[string]$Needle,[string]$Stage,[bool]$Before) {
    $component=$null; $codeModule=$null
    try {
        $component=$Components.Item($ModuleName); $codeModule=$component.CodeModule
        $text=[string]$codeModule.Lines(1,$codeModule.CountOfLines)
        # VBE canonicalizes identifier casing (e.g. Filename/fileName) on import.
        $matches=[regex]::Matches($text,[regex]::Escape($Needle),[Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($matches.Count -ne 1) { throw ('Ambiguous compare trace location: '+$Stage) }
        $Needle=$matches[0].Value
        $call='    NxR62CompareTrace "'+$Stage+'"'
        $replacement=if($Before){$call+"`r`n"+$Needle}else{$Needle+"`r`n"+$call}
        $updated=$text.Replace($Needle,$replacement)
        $codeModule.DeleteLines(1,$codeModule.CountOfLines); $codeModule.AddFromString($updated)
    } finally { Release-Com $codeModule; Release-Com $component }
}
$dispatchCode=@'
Option Explicit
Public Sub NxR62CompareTrace(ByVal stage As String)
    Dim handle As Integer
    handle = FreeFile
    Open Environ$("TEMP") & "\r62-compare-stages.log" For Append As #handle
    Print #handle, stage
    Close #handle
End Sub
Public Function NxR62CompareDispatch(ByVal kind As String, ByVal reportPath As String) As String
    Dim issue As String
    On Error GoTo Failed
    Select Case kind
        Case "prepare"
            NxR62CompareDispatch = "REPORT|" & T_R62CompareNavigation.PrepareReport()
            Exit Function
        Case "navigation": T_R62CompareNavigation.VerifyOpenedReport reportPath
        Case "failures": T_R62CompareNavigation.RunFailures reportPath
        Case "unsafe_links": NxR62CompareUnsafeLinks reportPath
        Case "cleanup": T_R62CompareNavigation.CleanupPreparedReport reportPath
        Case Else: Err.Raise vbObjectError + 962, , "Unknown R62 compare case"
    End Select
    NxR62CompareDispatch = "PASS|" & kind
    Exit Function
Failed:
    NxR62CompareDispatch = "FAIL|" & kind & "|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Sub NxR62CompareUnsafeLinks(ByVal reportPath As String)
    Dim book As Workbook, candidate As Workbook, sheet As Worksheet, summary As Worksheet, target As Hyperlink
    Dim originalAddress As String, originalSubAddress As String, originalPath As String, issue As String, failure As String
    Dim priorEvents As Boolean, priorSecurity As MsoAutomationSecurity, beforeCount As Long, value As Variant
    Dim activeBook As Workbook, activeSheet As Object, priorCell As String
    On Error GoTo Failed
    priorEvents = Application.EnableEvents: priorSecurity = Application.AutomationSecurity
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, reportPath, vbTextCompare) = 0 Then Set book = candidate: Exit For
    Next candidate
    If book Is Nothing Then Err.Raise vbObjectError + 962, , "Prepared report is not open"
    Set sheet = book.Worksheets(ChrW$(&HC140) & ChrW$(&HCC28) & ChrW$(&HC774)): Set summary = book.Worksheets(ChrW$(&HBE44) & ChrW$(&HAD50) & ChrW$(&HC694) & ChrW$(&HC57D))
    Set target = sheet.Range("H2").Hyperlinks(1)
    originalAddress = target.Address: originalSubAddress = target.SubAddress: originalPath = CStr(summary.Range("B2").Value2)
    If Len(originalAddress) <> 0 Then Err.Raise vbObjectError + 962, , "Original link must be a self anchor"
    Application.EnableEvents = False
    book.Activate: sheet.Activate: sheet.Range("H2").Select
    Set activeBook = Application.ActiveWorkbook: Set activeSheet = Application.ActiveSheet
    priorCell = Application.ActiveCell.Address: beforeCount = Application.Workbooks.Count
    For Each value In Array("https://example.invalid/source.xlsx", "file:///C:/not-a-source.xlsx", "cmd.exe")
        target.Address = CStr(value)
        If NxWorkbookCompareNavigate(sheet, target, issue) Then Err.Raise vbObjectError + 962, , "External link injection accepted"
        If Len(issue) = 0 Then Err.Raise vbObjectError + 962, , "External link rejection lacked an explanation"
        target.Address = originalAddress: target.SubAddress = originalSubAddress
    Next value
    target.SubAddress = "'" & summary.Name & "'!A1"
    If NxWorkbookCompareNavigate(sheet, target, issue) Then Err.Raise vbObjectError + 962, , "Wrong report anchor accepted"
    target.SubAddress = originalSubAddress
    For Each value In Array("https://example.invalid/source.xlsx", "..\relative.xlsx", "\\?\C:\source.xlsx", "C:\source.xlsx:payload.xlsm")
        summary.Range("B2").Value2 = CStr(value)
        If NxWorkbookCompareNavigate(sheet, target, issue) Then Err.Raise vbObjectError + 962, , "Unsafe source field accepted"
        If Len(issue) = 0 Then Err.Raise vbObjectError + 962, , "Unsafe source rejection lacked an explanation"
    Next value
    If Application.Workbooks.Count <> beforeCount Then Err.Raise vbObjectError + 962, , "Rejected navigation changed open books"
    If Not Application.ActiveWorkbook Is activeBook Then Err.Raise vbObjectError + 962, , "Rejected navigation changed active workbook"
    If Not Application.ActiveSheet Is activeSheet Then Err.Raise vbObjectError + 962, , "Rejected navigation changed active sheet"
    If Application.ActiveCell.Address <> priorCell Then Err.Raise vbObjectError + 962, , "Rejected navigation changed active cell"
    If Application.EnableEvents Or Application.AutomationSecurity <> priorSecurity Then Err.Raise vbObjectError + 962, , "Rejected navigation changed Excel security/events"
CleanUp:
    On Error Resume Next
    If Not target Is Nothing Then target.Address = originalAddress: target.SubAddress = originalSubAddress
    If Not summary Is Nothing And Len(originalPath) > 0 Then summary.Range("B2").Value2 = originalPath
    Application.AutomationSecurity = priorSecurity: Application.EnableEvents = priorEvents
    On Error GoTo 0
    If Len(failure) > 0 Then Err.Raise vbObjectError + 962, , failure
    Exit Sub
Failed:
    failure = Err.Description
    Resume CleanUp
End Sub

Public Function NxR62CompareTempRoot() As String
    NxR62CompareTempRoot = Environ$("TEMP")
End Function
'@

$started=[DateTime]::UtcNow.ToString('o')
$isolatedProfileRoot=Join-Path $EvidenceRoot 'profile'
$isolatedTempRoot=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) ('LHExcel\ValidationTemp\r62cmp-'+[Guid]::NewGuid().ToString('N'))
$previousProfileRoot=[Environment]::GetEnvironmentVariable('LHEXCEL_PROFILE_ROOT','Process')
$previousTempRoot=[Environment]::GetEnvironmentVariable('TEMP','Process')
$previousTmpRoot=[Environment]::GetEnvironmentVariable('TMP','Process')
$copyPath=Join-Path $EvidenceRoot 'Product.r62.compare.instrumented.xlam'
$sourceHash=$null; $sourceHashAfter=$null; $copyBeforeHash=$null; $copyAfterHash=$null; $instrumentedHash=$null
$sourceSnapshot=$null; $sourceSnapshotAfter=$null; $sourceUnchanged=$false
$excel=$null; $binding=$null; $books=$null; $copy=$null; $project=$null; $components=$null; $report=$null
$failure=$null; $preserveExcel=$false; $abnormalExit=$false; $owner=$null; $profileObserved=$null; $tempObserved=$null
$reportPath=''; $fixturePaths=$null; $fixtureHashesBefore=$null; $fixtureHashesAfter=$null
$reportRoundTrip=[ordered]@{producer_closed=$false;first_open_outside_producer=$false;first_open_read_only=$false;closed_then_reopened=$false;reopened_read_only=$false;path=$null;before_sha256=$null;after_sha256=$null;disk_unchanged=$false;fixture_cleanup=$false}
$compile=[ordered]@{status='NOT_RUN';control_id=578;active_project=$null;dialog=$null;failure=$null}
$compileAfterReopen=[ordered]@{status='NOT_RUN';control_id=578;active_project=$null;dialog=$null;failure=$null}
$cleanup=[ordered]@{exit_mode='NOT_STARTED';failure=$null}
$cleanupErrors=New-Object 'Collections.Generic.List[string]'
$cases=New-Object 'Collections.Generic.List[object]'
$expectedCases=@('compare.prepare_reopen','compare.navigation','compare.failures','compare.unsafe_links','compare.cleanup')
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
try {
    $sourceHash=Get-Sha256 $ProductXlam
    $sourceSnapshot=Get-SourceSnapshot
    [void](New-Item -ItemType Directory -Path $isolatedProfileRoot)
    [void](New-Item -ItemType Directory -Path $isolatedTempRoot)
    [Environment]::SetEnvironmentVariable('LHEXCEL_PROFILE_ROOT',$isolatedProfileRoot,'Process')
    [Environment]::SetEnvironmentVariable('TEMP',$isolatedTempRoot,'Process')
    [Environment]::SetEnvironmentVariable('TMP',$isolatedTempRoot,'Process')
    Copy-Item -LiteralPath $ProductXlam -Destination $copyPath -ErrorAction Stop
    $copyBeforeHash=Get-Sha256 $copyPath
    if ($copyBeforeHash -cne $sourceHash) { throw 'R62 compare disposable copy SHA256 mismatch' }
    $baseline=Get-ExcelProcessBaseline
    if (@($baseline.process_ids).Count) { throw 'Excel appeared during preflight; ownership rejected' }
    # Exercise the same interactive /x activation used by a real Excel user.
    # Automation-created Excel is a distinct diagnostic environment.
    $session=Start-ExactInteractiveExcel $baseline 'R62 compare Excel'
    $excel=$session.excel; $binding=$session.binding
    if (-not $binding.owned) { $preserveExcel=$true; throw ('R62 compare Excel ownership rejected: '+$binding.failure) }
    $owner=[ordered]@{pid=[int]$binding.pid;hwnd=[Int64]$binding.hwnd;started_utc=$binding.started_utc;executable=$binding.executable;excel_version=[string]$excel.Version;activation='interactive_x_rot';user_control=[bool]$excel.UserControl}
    Write-Json (Join-Path $EvidenceRoot 'R62CompareNavigation.owned-process.json') ([ordered]@{schema_version=1;suite='R62CompareNavigation';status='OWNED';run_id=$RunId;owner=$owner;source_product_sha256=$sourceHash;copy_before_sha256=$copyBeforeHash;instrumented_copy=$copyPath})
    $books=$excel.Workbooks
    if ($books.Count -ne 0) { $preserveExcel=$true; throw 'Unexpected startup workbook; Excel preserved for user review' }
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false
    $copy=$books.Open($copyPath,$false,$false)
    if ($copy.ReadOnly -or -not $copy.IsAddin -or -not [string]::Equals([IO.Path]::GetFullPath([string]$copy.FullName),$copyPath,[StringComparison]::OrdinalIgnoreCase)) { throw 'Disposable add-in identity or writable state rejected' }
    $project=$copy.VBProject; $components=$project.VBComponents
    $probeCode=[IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/file/T_R62CompareNavigation.bas'),[Text.Encoding]::UTF8)
    Add-ProbeModule $components 'T_R62CompareNavigation' $probeCode
    Add-ProbeModule $components 'R62CompareDispatch' $dispatchCode
    Add-CompareTrace $components 'NxSafeWorkbook' '    snapshotPath = NxSafeWorkbookPrepareSnapshot(inputPath, snapshotFolder, issue, strictTemplate)' 'safe.package.begin' $true
    Add-CompareTrace $components 'NxSafeWorkbook' '    snapshotPath = NxSafeWorkbookPrepareSnapshot(inputPath, snapshotFolder, issue, strictTemplate)' 'safe.package.end' $false
    Add-CompareTrace $components 'NxSafeWorkbook' '    Set openedBook = Application.Workbooks.Open(Filename:=snapshotPath, UpdateLinks:=0, ReadOnly:=True, _' 'safe.open.begin' $true
    Add-CompareTrace $components 'NxSafeWorkbook' '        IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False, Editable:=strictTemplate)' 'safe.open.end' $false
    Add-CompareTrace $components 'NxWorkbookCompare' '    Set baseBook = NxSafeWorkbookOpen(basePath, baseOwned)' 'compare.base.begin' $true
    Add-CompareTrace $components 'NxWorkbookCompare' '    Set baseBook = NxSafeWorkbookOpen(basePath, baseOwned)' 'compare.base.end' $false
    Add-CompareTrace $components 'NxWorkbookCompare' '    Set compareBook = NxSafeWorkbookOpen(comparePath, compareOwned)' 'compare.other.begin' $true
    Add-CompareTrace $components 'NxWorkbookCompare' '    Set compareBook = NxSafeWorkbookOpen(comparePath, compareOwned)' 'compare.other.end' $false
    Invoke-ProbeCompile $excel $binding $components $copyPath
    $initialCompileEvidence=$compile
    $copy.Save()
    [void](Assert-ExactExcelProcessOwnership $binding)
    Assert-OnlyProbeWorkbook $books $copyPath
    Release-Com $components; $components=$null
    Release-Com $project; $project=$null
    $copy.Close($false); Release-Com $copy; $copy=$null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $instrumentedHash=Get-Sha256 $copyPath
    [void](Assert-ExactExcelProcessOwnership $binding)
    $copy=$books.Open($copyPath,$false,$false)
    if ($copy.ReadOnly -or -not $copy.IsAddin -or -not [string]::Equals([IO.Path]::GetFullPath([string]$copy.FullName),$copyPath,[StringComparison]::OrdinalIgnoreCase)) { throw 'Reopened disposable add-in identity rejected' }
    $project=$copy.VBProject; $components=$project.VBComponents
    try { Invoke-ProbeCompile $excel $binding $components $copyPath -AllowAlreadyCompiled }
    finally { $compileAfterReopen=$compile; $compile=$initialCompileEvidence }
    $macro="'"+$copy.Name.Replace("'","''")+"'!"
    $profileObserved=[string]$excel.Run($macro+'NxLHexcelProfileRoot')
    if (-not [string]::Equals([IO.Path]::GetFullPath($profileObserved),$isolatedProfileRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'R62 compare isolated profile root verification failed' }
    $tempObserved=[string]$excel.Run($macro+'NxR62CompareTempRoot')
    if (-not [string]::Equals([IO.Path]::GetFullPath($tempObserved),$isolatedTempRoot,[StringComparison]::OrdinalIgnoreCase)) { throw 'R62 compare isolated TEMP verification failed' }
    $excel.EnableEvents=$true; $excel.Visible=$true
    Assert-OnlyProbeWorkbook $books $copyPath

    Invoke-CompareCase 'compare.prepare_reopen' {
        $result=[string]$excel.Run($macro+'NxR62CompareDispatch','prepare','')
        if ($result -cnotmatch '^REPORT\|.+$') { throw ('Compare preparation failed: '+$result) }
        $script:reportPath=Assert-OwnedFixturePath $result.Substring(7) 'report'
        $reportRoundTrip.path=$script:reportPath
        Assert-OnlyProbeWorkbook $books $copyPath
        $reportRoundTrip.producer_closed=$true
        $script:report=Open-OwnedReport $script:reportPath
        $reportRoundTrip.first_open_outside_producer=$true; $reportRoundTrip.first_open_read_only=[bool]$script:report.ReadOnly
        $script:fixturePaths=Read-ReportSourcePaths $script:report
        Assert-OnlyCompareWorkbooks $books @($copyPath,$script:reportPath)
        $script:report.Close($false); Release-Com $script:report; $script:report=$null
        Assert-OnlyProbeWorkbook $books $copyPath
        $script:fixtureHashesBefore=[ordered]@{base=(Get-Sha256 $fixturePaths.base);compare=(Get-Sha256 $fixturePaths.compare)}
        $reportRoundTrip.before_sha256=Get-Sha256 $script:reportPath
        $script:report=Open-OwnedReport $script:reportPath
        $reportRoundTrip.closed_then_reopened=$true; $reportRoundTrip.reopened_read_only=[bool]$script:report.ReadOnly
        Assert-OnlyCompareWorkbooks $books @($copyPath,$script:reportPath)
        return 'PASS|prepare_reopen'
    }
    foreach ($kind in @('navigation','failures','unsafe_links')) {
        Invoke-CompareCase ('compare.'+$kind) {
            Assert-OnlyCompareWorkbooks $books @($copyPath,$script:reportPath)
            $priorEvents=[bool]$excel.EnableEvents; $priorSecurity=[int]$excel.AutomationSecurity
            $result=Invoke-CompareFixture $kind
            if ([bool]$excel.EnableEvents -ne $priorEvents -or [int]$excel.AutomationSecurity -ne $priorSecurity) { throw 'Compare fixture leaked events or macro security state' }
            Assert-OnlyCompareWorkbooks $books @($copyPath,$script:reportPath)
            return $result
        }
    }
    Invoke-CompareCase 'compare.cleanup' {
        $script:report.Close($false); Release-Com $script:report; $script:report=$null
        Assert-OnlyProbeWorkbook $books $copyPath
        $script:fixtureHashesAfter=[ordered]@{base=(Get-Sha256 $fixturePaths.base);compare=(Get-Sha256 $fixturePaths.compare)}
        $reportRoundTrip.after_sha256=Get-Sha256 $script:reportPath
        $reportRoundTrip.disk_unchanged=($reportRoundTrip.before_sha256 -ceq $reportRoundTrip.after_sha256 -and
            $fixtureHashesBefore.base -ceq $fixtureHashesAfter.base -and $fixtureHashesBefore.compare -ceq $fixtureHashesAfter.compare)
        if (-not $reportRoundTrip.disk_unchanged) { throw 'Comparison navigation changed source/report disk contents; artifacts retained' }
        foreach ($pair in @(@($script:reportPath,'report'),@($fixturePaths.base,'base'),@($fixturePaths.compare,'compare'))) { [void](Assert-OwnedFixturePath $pair[0] $pair[1]) }
        $result=Invoke-CompareFixture 'cleanup'
        foreach ($path in @($script:reportPath,$fixturePaths.base,$fixturePaths.compare)) {
            if (Test-Path -LiteralPath $path) { throw 'Owned compare fixture remained after cleanup' }
        }
        $reportRoundTrip.fixture_cleanup=$true
        Assert-OnlyProbeWorkbook $books $copyPath
        return $result
    }
} catch { $failure=$_.Exception.Message }
finally {
    # Failures retain fixture files. Close only exact known report/source books under
    # the unique owned TEMP; unexpected books preserve Excel without Quit or kill.
    if ($null -ne $binding -and $binding.owned) {
        try {
            if ($binding.process.HasExited) { $abnormalExit=$true; $cleanupErrors.Add('Owned Excel exited before cleanup') }
            else {
                [void](Assert-ExactExcelProcessOwnership $binding)
                $knownPaths=@($copyPath)
                if (-not [string]::IsNullOrWhiteSpace($reportPath)) { $knownPaths+=@($reportPath) }
                if ($null -ne $fixturePaths) { $knownPaths+=@($fixturePaths.base,$fixturePaths.compare) }
                if ($null -ne $books) { Assert-OnlyCompareWorkbooks $books $knownPaths }
                if (-not $preserveExcel -and $null -ne $books) {
                    for ($index=$books.Count; $index -ge 1; $index--) {
                        $item=$null
                        try {
                            $item=$books.Item($index)
                            if (-not [string]::Equals([IO.Path]::GetFullPath([string]$item.FullName),$copyPath,[StringComparison]::OrdinalIgnoreCase)) { $item.Close($false) }
                        } finally { $item=$null }
                    }
                }
            }
        } catch { $preserveExcel=$true; $cleanupErrors.Add($_.Exception.Message) }
    }
    Release-Com $report; $report=$null
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
            else { $cleanup=Stop-ExactProcessAfterGrace $binding.process 'R62 compare Excel' 15000 -Detailed }
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
    if (-not $sourceUnchanged) { throw 'Supplied product or source tree changed during R62 compare suite' }
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
$suitePassed=($null -eq $failure -and $cleanupErrors.Count -eq 0 -and $cases.Count -eq $expectedCases.Count -and $passed -eq $expectedCases.Count -and
    $remaining.Count -eq 0 -and $cleanupNatural -and -not $abnormalExit -and $compile.status -ceq 'PASS' -and
    $compileAfterReopen.status -cin @('PASS','COMPILE_RETAINED') -and $sourceUnchanged -and $reportRoundTrip.disk_unchanged -and $reportRoundTrip.fixture_cleanup)
$status=if ($suitePassed) { 'PASS' } else { 'FAIL' }
$executedNames=@($cases | ForEach-Object { $_.name })
$receipt=[ordered]@{
    schema_version=1;suite='R62CompareNavigation';status=$status;run_id=$RunId;started_utc=$started;completed_utc=[DateTime]::UtcNow.ToString('o')
    expected_case_count=$expectedCases.Count;cases=@($cases.ToArray());not_run_cases=@($expectedCases | Where-Object { $executedNames -cnotcontains $_ } | ForEach-Object { [ordered]@{name=$_;status='NOT_RUN'} })
    failure=$failure;owner=$owner;isolated_profile_root=$isolatedProfileRoot;profile_root_observed=$profileObserved;isolated_temp_root=$isolatedTempRoot;temp_root_observed=$tempObserved
    compile=$compile;compile_after_reopen=$compileAfterReopen;report_round_trip=$reportRoundTrip;fixture_source_paths=$fixturePaths;fixture_source_sha256_before=$fixtureHashesBefore;fixture_source_sha256_after=$fixtureHashesAfter
    cleanup=$cleanup;cleanup_errors=@($cleanupErrors.ToArray());abnormal_exit=$abnormalExit;remaining_excel_immediate=$remainingImmediate;remaining_excel=$remaining
    source_product_sha256=$sourceHash;source_product_sha256_after=$sourceHashAfter;copy_before_sha256=$copyBeforeHash;copy_after_sha256=$copyAfterHash;instrumented_copy=$copyPath
    source_tree_sha256=if($null -ne $sourceSnapshot){$sourceSnapshot.sha256}else{$null};source_tree_sha256_after=if($null -ne $sourceSnapshotAfter){$sourceSnapshotAfter.sha256}else{$null}
    source_files_sha256=if($null -ne $sourceSnapshot){$sourceSnapshot.files}else{$null};source_unchanged=$sourceUnchanged;retained_fixture_files=$retainedFixtureFiles
    test_instrumentation=@('T_R62CompareNavigation','R62CompareDispatch','test-copy-only NxSafeWorkbook/NxWorkbookCompare fixed-stage tracing');production_modules_rewritten=$true
    not_run=@('physical mouse/keyboard input and accessibility','UNC server connectivity and SMB execution: parser fixture only','hostile XLM/VBA execution fixtures','other feature suites, installation, COM registration, signing and release approval')
    scope='owned Windows Excel, disposable instrumented XLAM, external COM report open-close-reopen, native VBA navigation assertions; unsafe external targets never followed'
}
Write-Json (Join-Path $EvidenceRoot 'R62CompareNavigation.json') $receipt
if (-not $suitePassed) { throw ('R62 compare navigation failed; receipt='+ (Join-Path $EvidenceRoot 'R62CompareNavigation.json')) }
Write-Output ('PASS|R62CompareNavigation|'+$passed+'/'+$expectedCases.Count)
