param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RunId = '',
    [string]$CaseNames = 'no_documents,snapshot,form_filter_click,hidden_protected_chart,rename_closed_identity,multiwindow,event_redirect,event_reenter,event_close,wheel_owner,lifecycle',
    [switch]$VisualHold,
    [ValidateSet('off','cf')][string]$VisualFocusMode = 'off'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$requestedCases = @($CaseNames.Split(','))
$allowedCases = @('no_documents','snapshot','form_filter_click','hidden_protected_chart','rename_closed_identity','multiwindow','event_redirect','event_reenter','event_close','wheel_owner','lifecycle')
if (-not $requestedCases.Count -or @($requestedCases | Where-Object {$_ -notin $allowedCases}).Count -or @($requestedCases | Select-Object -Unique).Count -ne $requestedCases.Count) { throw 'Unknown or duplicate R60 navigator case' }
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$ProductXlam = [IO.Path]::GetFullPath($ProductXlam)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing R60 navigator evidence is rejected' }
if (-not (Test-Path -LiteralPath $ProductXlam -PathType Leaf)) { throw 'Product XLAM unavailable' }
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
function Release-Com([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch {}
    }
}
function Add-Probe([object]$Components,[string]$File,[string]$Name,[int]$Type=1) {
    $component=$null; $module=$null
    try {
        $code=[IO.File]::ReadAllText($File,[Text.Encoding]::UTF8)
        $start=$code.IndexOf('Option Explicit')
        if ($start -lt 0) { throw ('Probe Option Explicit unavailable: '+$Name) }
        $code=$code.Substring($start).Replace("`r`n","`n").Replace("`n","`r`n")
        $code=[regex]::Replace($code,'(?m)^Attribute [^\r\n]*\r?\n','')
        $component=$Components.Add($Type); $component.Name=$Name
        $module=$component.CodeModule; [void]$module.AddFromString($code)
    } finally { Release-Com $module; Release-Com $component }
}
function Add-Wrapper([object]$Components,[string]$Name,[string]$Code) {
    $component=$null; $module=$null
    try {
        $component=$Components.Item($Name); $module=$component.CodeModule
        if ($module.Lines(1,$module.CountOfLines) -match 'NxR60Test|NxR60ProbeNavigator') { throw ('Unexpected existing R60 wrapper: '+$Name) }
        [void]$module.AddFromString($Code)
    } finally { Release-Com $module; Release-Com $component }
}
function Add-WheelDiagnostics([object]$Components) {
    $component=$null; $module=$null
    try {
        # Counters are confined to the disposable probe, never the product.
        $component=$Components.Item('NxFormWheel'); $module=$component.CodeModule
        $code=([string]$module.Lines(1,$module.CountOfLines)).Replace("`r`n","`n")
        if ($code -match 'NxR60Wheel|mNxR60Wheel') { throw 'Unexpected existing R60 wheel diagnostics' }
        $anchors=@('Private mHost As Object', '    If mHandling Then GoTo ForwardMessage', '    If GetAncestor(targetWindow, GA_ROOT) <> mWindow Then GoTo ForwardMessage', '    NxFormWheelProc = 1')
        foreach ($anchor in $anchors) {
            if ([regex]::Matches($code,[regex]::Escape($anchor)).Count -ne 1) { throw ('R60 wheel diagnostic anchor absent or ambiguous: '+$anchor) }
        }
        $declarations=@'
#If VBA7 Then
Private Declare PtrSafe Function NxR60GetDpiForWindow Lib "user32" Alias "GetDpiForWindow" (ByVal hwnd As LongPtr) As Long
#Else
Private Declare Function NxR60GetDpiForWindow Lib "user32" Alias "GetDpiForWindow" (ByVal hwnd As Long) As Long
#End If
Private mNxR60WheelMessageCount As Long
Private mNxR60WheelHandledCount As Long
Private mNxR60WheelLastDelta As Long
Private mNxR60WheelTarget As String
'@
        $diagnostic=@'
Public Function NxR60WheelInfo() As String
    Dim hostCaption As String, dpi As Long, dpiError As Long
    On Error Resume Next
    If Not mHost Is Nothing Then hostCaption = CStr(mHost.Caption)
    If mWindow <> 0 Then
        Err.Clear
        dpi = NxR60GetDpiForWindow(mWindow)
        dpiError = Err.Number
    End If
    On Error GoTo 0
    NxR60WheelInfo = "wm_mousewheel_count=" & CStr(mNxR60WheelMessageCount) & _
        ";handled_count=" & CStr(mNxR60WheelHandledCount) & ";last_delta=" & CStr(mNxR60WheelLastDelta) & _
        ";attached=" & CStr(mAttached) & ";hooked_hwnd=" & CStr(mWindow) & ";hook_handle=" & CStr(mHook) & _
        ";target_hwnd=" & mNxR60WheelTarget & ";host_caption=" & hostCaption & _
        ";dpi=" & CStr(dpi) & ";dpi_error=" & CStr(dpiError)
End Function
'@
        $code=$code.Replace($anchors[0],$declarations.Replace("`r`n","`n")+"`n"+$anchors[0])
        $code=$code.Replace($anchors[1],"    mNxR60WheelMessageCount = mNxR60WheelMessageCount + 1`n"+$anchors[1])
        $code=$code.Replace($anchors[2],"    mNxR60WheelLastDelta = wheelDelta`n    mNxR60WheelTarget = CStr(targetWindow)`n"+$anchors[2])
        $code=$code.Replace($anchors[3],"    mNxR60WheelHandledCount = mNxR60WheelHandledCount + 1`n"+$anchors[3])
        $code=$code.TrimEnd()+"`n`n"+$diagnostic.Replace("`r`n","`n")+"`n"
        [void]$module.DeleteLines(1,$module.CountOfLines)
        [void]$module.AddFromString($code.Replace("`n","`r`n"))
    } finally { Release-Com $module; Release-Com $component }
}
function Get-PhysicalObservation([object]$Excel,[object]$Books,[string]$Label) {
    # Property reads only. No Run, Select, Evaluate, clipboard or Undo invocation.
    $book=$null;$sheets=$null;$sheet=$null;$cell=$null;$font=$null;$selection=$null;$active=$null;$bars=$null
    $cells=[ordered]@{}
    try {
        $book=$Books.Item('visual_Second.xlsx');$sheets=$book.Worksheets;$sheet=$sheets.Item('Start')
        foreach($address in @('B2','C2','D2','B3','C3','D3','B4','C4','D4','H2','K12','L12','M12','K13','L13','M13','K14','L14','M14')) {
            $cell=$sheet.Range($address);$font=$cell.Font
            $cells[$address]=[ordered]@{value=$cell.Value2;formula=$cell.Formula;number_format=$cell.NumberFormat;bold=$font.Bold}
            Release-Com $font;Release-Com $cell;$font=$null;$cell=$null
        }
        $active=$Excel.ActiveWorkbook;$selection=$Excel.Selection;$bars=$Excel.CommandBars
        $activeName=$null;if($null -ne $active){$activeName=$active.Name}
        $selectionAddress=$null;$selectionKind='none';$selectionError=$null
        if($null -ne $selection) {
            # Charts/shapes have no Range.Address; an unavailable address is not a suite failure.
            $selectionKind='non_range_or_unavailable'
            try {$selectionAddress=$selection.Address();$selectionKind='range'}
            catch {$selectionError=$_.Exception.Message}
        }
        return [ordered]@{label=$Label;evidence_class='native_com_property_reads_only';utc=[DateTime]::UtcNow.ToString('o');active_book=$activeName;selection=$selectionAddress;selection_kind=$selectionKind;selection_error=$selectionError;cut_copy_mode=$Excel.CutCopyMode;undo_enabled=$bars.GetEnabledMso('Undo');paste_enabled=$bars.GetEnabledMso('Paste');saved=$book.Saved;cells=$cells}
    } finally {Release-Com $bars;Release-Com $active;Release-Com $selection;Release-Com $font;Release-Com $cell;Release-Com $sheet;Release-Com $sheets;Release-Com $book}
}
function Test-BlankStartupBook([object]$Book) {
    $sheets=$null; $allSheets=$null; $names=$null; $sheet=$null; $cells=$null; $shapes=$null
    try {
        if ($Book.IsAddin -or -not $Book.Saved -or $Book.Path -ne '' -or $Book.HasVBProject) { return $false }
        $names=$Book.Names
        if ($names.Count -ne 0) { return $false }
        $sheets=$Book.Worksheets
        $allSheets=$Book.Sheets
        if ($sheets.Count -ne $allSheets.Count) { return $false }
        for ($index=1; $index -le $sheets.Count; $index++) {
            $sheet=$sheets.Item($index); $cells=$sheet.UsedRange; $shapes=$sheet.Shapes
            if ($sheet.Visible -ne -1 -or $sheet.ProtectContents -or $shapes.Count -ne 0 -or $cells.Count -ne 1 -or $null -ne $cells.Value2 -or $cells.HasFormula) { return $false }
            Release-Com $shapes; Release-Com $cells; Release-Com $sheet
            $shapes=$null; $cells=$null; $sheet=$null
        }
        return $true
    } finally { Release-Com $shapes; Release-Com $cells; Release-Com $sheet; Release-Com $allSheets; Release-Com $sheets; Release-Com $names }
}
function Get-UserDocumentCount([object]$Books) {
    $count=0
    for($index=1;$index -le $Books.Count;$index++) {
        $item=$null
        try { $item=$Books.Item($index);if(-not $item.IsAddin){$count++} }
        finally { Release-Com $item }
    }
    return $count
}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$fixtureRoot=Join-Path $EvidenceRoot 'fixtures'
[void](New-Item -ItemType Directory -Path $fixtureRoot)
$fixtureHashes=@{ 'rich.xlsx'='727C2C99502E30A22FA45C12374E17F613334FD865CD7EA6AE801A1990D94E6F'; 'simple.xlsx'='597C2F1E08480DA51CA2F4C417A4FAB32FC723A43EE96A30ADED5590A25D72D0' }
foreach($fixtureName in $fixtureHashes.Keys) {
    $fixtureSource=Join-Path $SourceRoot ('tests/fixtures/r60-document-navigator/'+$fixtureName)
    if((Get-FileHash -LiteralPath $fixtureSource -Algorithm SHA256).Hash -ne $fixtureHashes[$fixtureName]) { throw ('Fixture hash mismatch: '+$fixtureName) }
    Copy-Item -LiteralPath $fixtureSource -Destination (Join-Path $fixtureRoot $fixtureName)
}
$copyPath=Join-Path $EvidenceRoot 'R60NavigatorProbe.xlam'
$sourceHash=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash
$previousProfile=$env:LHEXCEL_PROFILE_ROOT
$env:LHEXCEL_PROFILE_ROOT=Join-Path $EvidenceRoot 'profile'
$excel=$null; $books=$null; $addin=$null; $project=$null; $components=$null; $binding=$null
$cleanup=[pscustomobject]@{exit_mode='NOT_STARTED';failure=$null}
$failure=$null; $preserveExcel=$false; $ownerPid=$null; $visualStatus='NOT_RUN'; $visualDetail=$null; $visualDiagnostics=$null
$probeReady=$false;$initialCollectionCount=$null;$initialUserDocumentCount=$null
$cases=New-Object 'Collections.Generic.List[object]'
$cleanupErrors=New-Object 'Collections.Generic.List[string]'
$started=[DateTime]::UtcNow.ToString('o')
try {
    if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close user Excel before the R60 navigator suite' }
    Copy-Item -LiteralPath $ProductXlam -Destination $copyPath
    if ((Get-FileHash -LiteralPath $copyPath -Algorithm SHA256).Hash -ne $sourceHash) { throw 'Probe copy hash mismatch' }
    $baseline=Get-ExcelProcessBaseline
    if (@($baseline.process_ids).Count) { throw 'Excel appeared during preflight; ownership rejected' }
    $interactive=Start-ExactInteractiveExcel $baseline 'R60 Document Navigator Excel' -TimeoutMs 120000
    $excel=$interactive.excel; $binding=$interactive.binding; $interactive=$null
    if (-not $binding.owned) { throw 'Excel process ownership rejected' }
    $ownerPid=[int]$binding.pid
    $books=$excel.Workbooks
    # Only a pristine, unsaved-path startup workbook in our exact new process may close.
    # Unknown startup/user documents cause preservation of the Excel process, not Quit/kill.
    for ($index=$books.Count; $index -ge 1; $index--) {
        $startup=$null
        try {
            $startup=$books.Item($index)
            if (-not (Test-BlankStartupBook $startup)) { $preserveExcel=$true; throw 'Unexpected startup workbook; Excel is preserved for user review' }
            $startup.Close($false)
        } finally { Release-Com $startup }
    }
    $excel.Visible=$true; $excel.DisplayAlerts=$false; $excel.EnableEvents=$true
    $addin=$books.Open($copyPath,$false,$false)
    $initialCollectionCount=[int]$books.Count
    $initialUserDocumentCount=Get-UserDocumentCount $books
    if (-not $addin.IsAddin -or $initialUserDocumentCount -ne 0) { throw ('Zero-document setup rejected; collection='+$initialCollectionCount+'; user_documents='+$initialUserDocumentCount+'; is_addin='+$addin.IsAddin) }
    $project=$addin.VBProject; $components=$project.VBComponents
    Add-Wrapper $components 'NxDocumentNavigator' @'
Public Function NxR60ProbeNavigator() As FNxDocumentNavigator
    Set NxR60ProbeNavigator = mNavigator
End Function
Public Function NxR60ProbeVisualDetail() As String
    NxR60ProbeVisualDetail = NxR60WheelInfo()
    If mNavigator Is Nothing Then
        NxR60ProbeVisualDetail = NxR60ProbeVisualDetail & "|navigator=NOT_OPEN"
    Else
        NxR60ProbeVisualDetail = NxR60ProbeVisualDetail & "|" & mNavigator.NxR60TestWheelState()
    End If
End Function
'@
    Add-Wrapper $components 'FNxDocumentNavigator' @'
Public Sub NxR60TestBookClick(ByVal row As Long)
    lstBooks.ListIndex = row
    lstBooks_Click
End Sub
Public Sub NxR60TestSheetClick(ByVal row As Long)
    lstSheets.ListIndex = row
    lstSheets_Click
End Sub
Public Sub NxR60TestMove()
    cmdMove_Click
End Sub
Public Sub NxR60TestClose()
    cmdClose_Click
End Sub
Public Function NxR60TestSelectionInfo() As String
    Dim selected As CNxDocumentNavigatorItem
    Set selected = SelectedItem()
    NxR60TestSelectionInfo = mActiveList
    If Not selected Is Nothing Then
        NxR60TestSelectionInfo = NxR60TestSelectionInfo & "|" & selected.DisplayName & "|" & TypeName(selected.SheetRef)
    End If
End Function
Public Function NxR60TestLifecycleInfo() As String
    NxR60TestLifecycleInfo = "closing=" & CStr(mClosing) & ";books=" & CStr(lstBooks.ListCount) & ";filter=" & txtBooks.Value
End Function
Public Function NxR60TestWheelState() As String
    Dim activeControlName As String
    On Error Resume Next
    activeControlName = Me.ActiveControl.Name
    On Error GoTo 0
    NxR60TestWheelState = "navigator=OPEN;visible=" & CStr(Me.Visible) & ";closing=" & CStr(mClosing) & _
        ";active_control=" & activeControlName & ";active_list=" & mActiveList & ";wheel_list=" & mWheelList & _
        ";books_count=" & CStr(lstBooks.ListCount) & ";books_selected=" & CStr(lstBooks.ListIndex) & ";books_top=" & CStr(lstBooks.TopIndex) & _
        ";sheets_count=" & CStr(lstSheets.ListCount) & ";sheets_selected=" & CStr(lstSheets.ListIndex) & ";sheets_top=" & CStr(lstSheets.TopIndex) & _
        ";selection=" & NxR60TestSelectionInfo()
End Function
'@
    Add-WheelDiagnostics $components
    Add-Probe $components (Join-Path $SourceRoot 'tests/vba/product/CNxR60NavigatorEvents.cls') 'CNxR60NavigatorEvents' 2
    Add-Probe $components (Join-Path $SourceRoot 'tests/vba/product/T_R60DocumentNavigator.bas') 'T_R60DocumentNavigator'
    $probeReady=$true
    $dispatch="'"+$addin.Name+"'!NxR60DocumentNavigatorCase"
    foreach ($kind in $requestedCases) {
        [void](Assert-ExactExcelProcessOwnership $binding)
        Write-Output ('START|'+$kind)
        $result=[string]$excel.Run($dispatch,$kind,$EvidenceRoot)
        $cases.Add([pscustomobject]@{name=$kind;status=if($result.StartsWith('PASS|')){'PASS'}else{'FAIL'};evidence_class='native_vba_fixture';result=$result})
        Write-Output $result
        if (-not $result.StartsWith('PASS|')) { $failure='R60 navigator case failed: '+$result }
        if ((Get-UserDocumentCount $books) -ne 0) { throw ('Fixture cleanup left unexpected workbooks after '+$kind) }
    }
    if ($VisualHold -and $null -eq $failure) {
        $visualDetail=[string]$excel.Run(("'"+$addin.Name+"'!NxR60DocumentNavigatorVisual"),$EvidenceRoot,$VisualFocusMode)
        if (-not $visualDetail.StartsWith('READY|')) { throw $visualDetail }
        $visualStatus='WAITING'
        $visualDiagnostics=[ordered]@{evidence_class='native_vba_probe_diagnostics';before=$null;after=$null;capture_errors=@()}
        $visualDispatch="'"+$addin.Name+"'!NxR60ProbeVisualDetail"
        try { $visualDiagnostics.before=[string]$excel.Run($visualDispatch) }
        catch { $visualDiagnostics.capture_errors+=('Before hold: '+$_.Exception.Message) }
        Write-Output ('UI_READY|'+$ownerPid+'|'+$visualDetail)
        $continuePath=Join-Path $EvidenceRoot 'continue.txt'
        $observationRequest=Join-Path $EvidenceRoot 'observe.txt';$lastObservation=''
        $watch=[Diagnostics.Stopwatch]::StartNew()
        while (-not (Test-Path -LiteralPath $continuePath -PathType Leaf) -and $watch.Elapsed.TotalSeconds -lt 600) {
            if(Test-Path -LiteralPath $observationRequest -PathType Leaf) {
                $label=[IO.File]::ReadAllText($observationRequest).Trim()
                if($label -match '^[a-z0-9_-]{1,48}$' -and $label -ne $lastObservation) {
                    $observationPath=Join-Path $EvidenceRoot ('physical_observation_'+$label+'.json')
                    if(Test-Path -LiteralPath $observationPath) {
                        Write-Output ('UI_OBSERVATION_REJECTED|'+$label+'|already_exists')
                    } else {
                        try {
                            try {$observation=Get-PhysicalObservation $excel $books $label}
                            catch {$observation=[ordered]@{label=$label;status='ERROR';evidence_class='native_com_read_only_observation_error';utc=[DateTime]::UtcNow.ToString('o');error=$_.Exception.Message}}
                            [IO.File]::WriteAllText($observationPath,($observation | ConvertTo-Json -Depth 7),[Text.UTF8Encoding]::new($false))
                            Write-Output ('UI_OBSERVED|'+$label)
                        } catch {Write-Output ('UI_OBSERVATION_ERROR|'+$label+'|'+$_.Exception.Message)}
                    }
                    $lastObservation=$label
                }
            }
            Start-Sleep -Milliseconds 250
        }
        try { $visualDiagnostics.after=[string]$excel.Run($visualDispatch) }
        catch { $visualDiagnostics.capture_errors+=('After hold: '+$_.Exception.Message) }
        Write-Output ('UI_DIAGNOSTICS|'+($visualDiagnostics | ConvertTo-Json -Compress -Depth 4))
        if (-not (Test-Path -LiteralPath $continuePath -PathType Leaf)) { $visualStatus='TIMEOUT'; throw 'Visual hold timed out after 600 seconds' }
        # A continuation is not a UI PASS receipt: the operator records observations separately.
        $visualStatus='RELEASED_WITHOUT_UI_VERDICT'
        Write-Output ('UI_RELEASED|'+$ownerPid)
    }
} catch { $failure=$_.Exception.Message }
finally {
    if ($null -ne $binding -and $binding.owned -and $null -ne $excel) {
        try { [void](Assert-ExactExcelProcessOwnership $binding) } catch { $preserveExcel=$true; $cleanupErrors.Add($_.Exception.Message) }
    }
    if ($probeReady -and -not $preserveExcel -and $null -ne $addin -and $null -ne $excel) {
        try { [void]$excel.Run(("'"+$addin.Name+"'!NxR60DocumentNavigatorDispose")) } catch { $cleanupErrors.Add('Probe dispose: '+$_.Exception.Message) }
    }
    if ($null -ne $addin -and -not $preserveExcel) { try { $addin.Close($false) } catch { $cleanupErrors.Add('Probe close: '+$_.Exception.Message) } }
    if ($null -ne $books -and -not $preserveExcel) {
        # A workbook that is not one of our saved fixture files is never silently discarded.
        for ($index=$books.Count; $index -ge 1; $index--) {
            $remainingBook=$null
            try {
                $remainingBook=$books.Item($index)
                $remainingPath=[IO.Path]::GetFullPath([string]$remainingBook.FullName)
                if ($remainingPath.StartsWith($EvidenceRoot+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {
                    $remainingBook.Close($false)
                } else { $preserveExcel=$true; $cleanupErrors.Add('Unexpected workbook remains; Excel preserved for user review') }
            } catch { $preserveExcel=$true; $cleanupErrors.Add('Workbook cleanup: '+$_.Exception.Message) }
            finally { Release-Com $remainingBook }
        }
    }
    Release-Com $components; Release-Com $project; Release-Com $addin; Release-Com $books
    if ($null -ne $excel) {
        if (-not $preserveExcel) { try { $excel.Quit() } catch { $cleanupErrors.Add('Excel Quit: '+$_.Exception.Message) } }
        else { try { $excel.DisplayAlerts=$true; $excel.Visible=$true } catch {} }
        Release-Com $excel
    }
    $components=$null; $project=$null; $addin=$null; $books=$null; $excel=$null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($null -ne $binding -and $null -ne $binding.process) {
        if ($preserveExcel) { $cleanup=[pscustomobject]@{exit_mode='PRESERVED_UNEXPECTED_WORKBOOK';failure='User review required; no Quit or kill'} }
        else {
            try { $cleanup=Stop-ExactProcessAfterGrace $binding.process 'R60 Document Navigator Excel' 15000 -Detailed }
            catch { $cleanup=[pscustomobject]@{exit_mode='CLEANUP_FAILED';failure=$_.Exception.Message} }
        }
        $binding.process.Dispose()
    }
    $env:LHEXCEL_PROFILE_ROOT=$previousProfile
}
$remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { -not $_.HasExited } | ForEach-Object Id)
$afterHash=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash
$passed=@($cases | Where-Object status -eq 'PASS').Count
$status=if($null -eq $failure -and $cleanupErrors.Count -eq 0 -and $passed -eq $requestedCases.Count -and $remaining.Count -eq 0 -and $cleanup.exit_mode -eq 'NATURAL' -and $sourceHash -eq $afterHash){'PASS'}else{'FAIL'}
$receipt=[ordered]@{
    suite='R60DocumentNavigator';status=$status;run_id=$RunId;started_utc=$started;completed_utc=[DateTime]::UtcNow.ToString('o')
    requested_cases=$requestedCases;cases=@($cases.ToArray());failure=$failure;owner_pid=$ownerPid
    initial_collection_count=$initialCollectionCount;initial_user_document_count=$initialUserDocumentCount
    fixture_hashes=$fixtureHashes
    cleanup=$cleanup;cleanup_errors=@($cleanupErrors.ToArray());remaining_excel=$remaining
    source_sha256=$sourceHash;source_sha256_after=$afterHash;source_unchanged=($sourceHash -eq $afterHash)
    visual_hold=[ordered]@{requested=[bool]$VisualHold;status=$visualStatus;detail=$visualDetail;ui_verdict='NOT_RUN_BY_SUITE'}
    visual_detail=$visualDiagnostics
    not_run=@('physical mouse click, double-click, Enter, Esc and X close','physical mouse wheel delivery and accessibility/screen-reader behavior','actual clipboard copy sizes and Excel Undo retention','full ProductUi/R58/copy-regression suites; untouched by this targeted suite','install, COM registration, signing and release acceptance')
    scope='owned Windows Excel; native VBA model and private control handlers in disposable XLAM copy; no production source edits, registry writes, clipboard simulation, or UIA injection'
}
[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'R60DocumentNavigator.json'),(($receipt | ConvertTo-Json -Depth 12)+"`n"),(New-Object Text.UTF8Encoding($false)))
if ($status -ne 'PASS') { throw ('R60 document navigator verification failed: '+$failure+'; receipt='+ (Join-Path $EvidenceRoot 'R60DocumentNavigator.json')) }
Write-Output ('PASS|R60DocumentNavigator|'+$passed+'/'+$requestedCases.Count)
