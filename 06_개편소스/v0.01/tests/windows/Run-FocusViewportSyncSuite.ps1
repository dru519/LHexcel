param(
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$ArtifactPath,
    [Parameter(Mandatory=$true)][string]$NxHostRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
$ArtifactPath = [IO.Path]::GetFullPath($ArtifactPath)
$NxHostRoot = [IO.Path]::GetFullPath($NxHostRoot)
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing viewport evidence root is rejected' }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public sealed class NxViewportWindow { public long Handle; public string Title; public int Left; public int Top; public int Right; public int Bottom; }
public static class NxViewportApi {
  public delegate bool EnumProc(IntPtr hwnd, IntPtr value);
  [StructLayout(LayoutKind.Sequential)] struct Rect { public int Left, Top, Right, Bottom; }
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, IntPtr value);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder value, int maximum);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
  [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
  [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int width, int height, uint flags);
  public static NxViewportWindow[] Find(int wantedPid, string prefix) {
    var rows=new List<NxViewportWindow>();
    EnumWindows((hwnd,value)=>{uint pid;GetWindowThreadProcessId(hwnd,out pid);if(pid!=(uint)wantedPid||!IsWindowVisible(hwnd))return true;var title=new StringBuilder(512);GetWindowText(hwnd,title,title.Capacity);string text=title.ToString();if(!text.StartsWith(prefix,StringComparison.Ordinal))return true;Rect rect;if(!GetWindowRect(hwnd,out rect))return true;rows.Add(new NxViewportWindow{Handle=hwnd.ToInt64(),Title=text,Left=rect.Left,Top=rect.Top,Right=rect.Right,Bottom=rect.Bottom});return true;},IntPtr.Zero);
    return rows.ToArray();
  }
  public static NxViewportWindow[] FindAll(int wantedPid, string prefix) {
    var rows=new List<NxViewportWindow>();
    EnumWindows((hwnd,value)=>{uint pid;GetWindowThreadProcessId(hwnd,out pid);if(pid!=(uint)wantedPid)return true;var title=new StringBuilder(512);GetWindowText(hwnd,title,title.Capacity);string text=title.ToString();if(!text.StartsWith(prefix,StringComparison.Ordinal))return true;Rect rect;if(!GetWindowRect(hwnd,out rect))return true;rows.Add(new NxViewportWindow{Handle=hwnd.ToInt64(),Title=text,Left=rect.Left,Top=rect.Top,Right=rect.Right,Bottom=rect.Bottom});return true;},IntPtr.Zero);
    return rows.ToArray();
  }
  public static NxViewportWindow Bounds(long handle) { Rect rect;if(!GetWindowRect(new IntPtr(handle),out rect))throw new InvalidOperationException("GetWindowRect failed");return new NxViewportWindow{Handle=handle,Left=rect.Left,Top=rect.Top,Right=rect.Right,Bottom=rect.Bottom}; }
  public static bool Place(long handle,int left,int top,int width,int height) { const uint flags=0x0004|0x0010;return SetWindowPos(new IntPtr(handle),IntPtr.Zero,left,top,width,height,flags); }
  public static bool Move(long handle,int left,int top) { const uint flags=0x0001|0x0004|0x0010;return SetWindowPos(new IntPtr(handle),IntPtr.Zero,left,top,0,0,flags); }
}
'@

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) }
}
function Get-Sha256([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Marker([object[]]$Rows,[string]$Prefix) { $row=@($Rows|ForEach-Object{[string]$_}|Where-Object{$_.StartsWith($Prefix,[StringComparison]::Ordinal)}|Select-Object -Last 1);if($row.Count-eq0){return ''};$row[0].Substring($Prefix.Length) }
function Write-Json([string]$Path,[object]$Value) { $encoding=New-Object Text.UTF8Encoding($false);[IO.File]::WriteAllText($Path,($Value|ConvertTo-Json -Depth 20),$encoding) }
function Overlay-Map([int]$ProcessId,[string]$Prefix) {
    $map=@{}
    foreach($window in [NxViewportApi]::Find($ProcessId,$Prefix)){
        $token=$window.Title.Substring($Prefix.Length).Split('_')[0]
        if($window.Title.StartsWith($Prefix+'ROW_LEFT')){$token='ROW_LEFT'}
        elseif($window.Title.StartsWith($Prefix+'ROW_RIGHT')){$token='ROW_RIGHT'}
        elseif($window.Title.StartsWith($Prefix+'COLUMN_TOP')){$token='COLUMN_TOP'}
        elseif($window.Title.StartsWith($Prefix+'COLUMN_BOTTOM')){$token='COLUMN_BOTTOM'}
        elseif($window.Title.StartsWith($Prefix+'SELECT_TOP')){$token='SELECT_TOP'}
        elseif($window.Title.StartsWith($Prefix+'SELECT_RIGHT')){$token='SELECT_RIGHT'}
        elseif($window.Title.StartsWith($Prefix+'SELECT_BOTTOM')){$token='SELECT_BOTTOM'}
        elseif($window.Title.StartsWith($Prefix+'SELECT_LEFT')){$token='SELECT_LEFT'}
        $map[$token]=[pscustomobject]@{handle=[int64]$window.Handle;left=[int]$window.Left;top=[int]$window.Top;right=[int]$window.Right;bottom=[int]$window.Bottom}
    }
    $map
}
function Try-Macro([object]$Excel,[string]$Macro,[object]$DefaultValue) {
    try { return $Excel.Run($Macro) } catch { return $DefaultValue }
}
function Try-StateSignature([object]$Excel,[string]$Macro) {
    try { return [string]$Excel.Run($Macro+'NxFocusGeometryStateSignature',$Excel) } catch { return '' }
}
function Get-FocusDiagnostic([object]$Excel,[int]$ProcessId,[string]$Macro,[string]$Prefix) {
    $activeWindow=$null;$activePane=$null;$activeSheet=$null;$selection=$null;$visibleRange=$null
    try {
        $activeWindow=$Excel.ActiveWindow;$activePane=$activeWindow.ActivePane;$activeSheet=$Excel.ActiveSheet;$selection=$Excel.Selection;$visibleRange=$activePane.VisibleRange
        return [pscustomobject][ordered]@{
            controller_backend = [string](Try-Macro $Excel ($Macro+'NxFocusControllerBackend') '')
            controller_diagnostic = [string](Try-Macro $Excel ($Macro+'NxFocusControllerBackendDiagnostic') '')
            geometry_diagnostic = [string](Try-Macro $Excel ($Macro+'NxFocusGeometryLastDiagnostic') '')
            overlay_last_error = [string](Try-Macro $Excel ($Macro+'NxFocusOverlayLastError') '')
            vba_window_count = [int](Try-Macro $Excel ($Macro+'NxFocusOverlayWindowCount') -1)
            vba_visible_count = [int](Try-Macro $Excel ($Macro+'NxFocusOverlayVisibleWindowCount') -1)
            vba_timer_count = [int](Try-Macro $Excel ($Macro+'NxFocusOverlayTimerCount') -1)
            state_signature = Try-StateSignature $Excel $Macro
            win32_total = @([NxViewportApi]::FindAll($ProcessId,$Prefix)).Count
            win32_visible = @([NxViewportApi]::Find($ProcessId,$Prefix)).Count
            excel_bounds = [NxViewportApi]::Bounds([int64]$Excel.Hwnd)
            active_window_hwnd = [int64]$activeWindow.Hwnd
            application_window_state = [int]$Excel.WindowState
            active_window_state = [int]$activeWindow.WindowState
            active_window_width = [double]$activeWindow.Width
            active_window_height = [double]$activeWindow.Height
            active_sheet = [string]$activeSheet.Name
            selection = [string]$selection.Address($false,$false,1,$true)
            visible_range = [string]$visibleRange.Address($false,$false,1,$true)
        }
    } finally {
        Release-ComObject $visibleRange;Release-ComObject $selection;Release-ComObject $activeSheet;Release-ComObject $activePane;Release-ComObject $activeWindow
    }
}
function Test-PhysicalMove([object]$Excel,[int]$ProcessId,[string]$Prefix,[int]$DeltaX,[int]$DeltaY) {
    $activeWindow=$null;$activePane=$null;$visibleRange=$null;$selection=$null
    try {
        $before=Overlay-Map $ProcessId $Prefix
        if($before.Count-lt4){return [pscustomobject]@{passed=$false;reason='initial_window_count';count=$before.Count}}
        $activeWindow=$Excel.ActiveWindow;$activePane=$activeWindow.ActivePane;$visibleRange=$activePane.VisibleRange;$selection=$Excel.Selection
        $visibleBefore=[string]$visibleRange.Address($false,$false,1,$true);$selectionBefore=[string]$selection.Address($false,$false,1,$true)
        Release-ComObject $visibleRange;$visibleRange=$null;Release-ComObject $selection;$selection=$null
        $excelBounds=[NxViewportApi]::Bounds([int64]$Excel.Hwnd)
        if(-not[NxViewportApi]::Move([int64]$Excel.Hwnd,$excelBounds.Left+$DeltaX,$excelBounds.Top+$DeltaY)){throw 'SetWindowPos failed'}
        Start-Sleep -Milliseconds 260
        $after=Overlay-Map $ProcessId $Prefix
        $visibleRange=$activePane.VisibleRange;$selection=$Excel.Selection
        $visibleAfter=[string]$visibleRange.Address($false,$false,1,$true);$selectionAfter=[string]$selection.Address($false,$false,1,$true)
        $rows=New-Object 'Collections.Generic.List[object]';$maximum=0;$sameHandles=$true
        foreach($token in @($before.Keys|Sort-Object)){
            if(-not$after.ContainsKey($token)){$maximum=999;$sameHandles=$false;continue}
            $dx=[Math]::Abs(([int]$after[$token].left-[int]$before[$token].left)-$DeltaX)
            $dy=[Math]::Abs(([int]$after[$token].top-[int]$before[$token].top)-$DeltaY)
            $error=[Math]::Max($dx,$dy);if($error-gt$maximum){$maximum=$error}
            if([int64]$after[$token].handle-ne[int64]$before[$token].handle){$sameHandles=$false}
            $rows.Add([pscustomobject]@{token=$token;delta_x=[int]$after[$token].left-[int]$before[$token].left;delta_y=[int]$after[$token].top-[int]$before[$token].top;error=$error;same_handle=([int64]$after[$token].handle-eq[int64]$before[$token].handle)})
        }
        [void][NxViewportApi]::Move([int64]$Excel.Hwnd,$excelBounds.Left,$excelBounds.Top)
        Start-Sleep -Milliseconds 260
        return [pscustomobject][ordered]@{
            passed=($after.Count-eq$before.Count-and$maximum-le2-and$sameHandles-and$visibleBefore-ceq$visibleAfter-and$selectionBefore-ceq$selectionAfter)
            visible_window_count=$before.Count
            max_error=$maximum
            same_handles=$sameHandles
            visible_range_unchanged=($visibleBefore-ceq$visibleAfter)
            selection_unchanged=($selectionBefore-ceq$selectionAfter)
            rows=$rows.ToArray()
        }
    } finally {
        Release-ComObject $selection;Release-ComObject $visibleRange;Release-ComObject $activePane;Release-ComObject $activeWindow
    }
}
function Close-OwnedExcel([object]$Excel,[object]$Binding,[object[]]$Objects) {
    foreach($item in $Objects){try{Release-ComObject $item}catch{}}
    if($null-ne$Excel){try{$Excel.DisplayAlerts=$false;$Excel.Quit()}catch{};Release-ComObject $Excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null-eq$Binding){return $null}
    $stop=Stop-ExactProcessAfterGrace $Binding.process 'Focus viewport Excel' 15000 -Detailed
    try{$Binding.process.Dispose()}catch{}
    $stop
}

if(-not(Test-Path -LiteralPath $ArtifactPath -PathType Leaf)){throw 'Product artifact missing'}
$x86=Join-Path $NxHostRoot 'NxHost32.dll';$x64=Join-Path $NxHostRoot 'NxHost64.dll'
if(-not(Test-Path -LiteralPath $x86 -PathType Leaf)-or-not(Test-Path -LiteralPath $x64 -PathType Leaf)){throw 'NxHost artifacts missing'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Focus viewport suite requires zero pre-existing Excel processes'}

$cases=New-Object 'Collections.Generic.List[object]';$registerOutput=@();$unregisterOutput=@();$beforeHash='';$afterHash='';$diagnostic='';$cleanup=$null;$registeredAddins=@()
$registered=$false;$interactive=$null;$excel=$null;$binding=$null;$books=$null;$addin=$null;$book=$null;$sheets=$null;$sheet=$null;$window=$null;$selectionRange=$null;$comAddins=$null;$hostAddin=$null
try{
    $registerOutput=@(& (Join-Path $SourceRoot 'build/Register-NxHost.ps1') -Action Register -X86Dll $x86 -X64Dll $x64 -OptIn)
    if(-not($registerOutput|Where-Object{[string]$_-ceq'PASS|Register'})){throw 'registration failed'}
    $beforeHash=Marker $registerOutput 'REGISTRY_SNAPSHOT_BEFORE|';$registered=$true
    $baseline=Get-ExcelProcessBaseline;$interactive=Start-ExactInteractiveExcel $baseline 'Focus viewport Excel';$excel=$interactive.excel;$binding=$interactive.binding;$interactive=$null
    if(-not[bool]$binding.owned){throw 'Excel ownership rejected'}
    $excel.Visible=$true;$excel.DisplayAlerts=$false
    $comAddins=$excel.COMAddIns
    for($addinIndex=1;$addinIndex-le[int]$comAddins.Count;$addinIndex++){
        $candidate=$null
        try{$candidate=$comAddins.Item($addinIndex);$registeredAddins+=([string]$candidate.ProgId)}finally{Release-ComObject $candidate}
    }
    $hostAddin=$comAddins.Item('LH.NxHost.Connect')
    if(-not[bool]$hostAddin.Connect){$hostAddin.Connect=$true;Start-Sleep -Milliseconds 400}
    if(-not[bool]$hostAddin.Connect){throw ('NxHost COM add-in did not connect; registered=' + ($registeredAddins -join ','))}
    $books=$excel.Workbooks;$addin=$books.Open($ArtifactPath,$false,$true);$book=$books.Add();$sheets=$book.Worksheets;$sheet=$sheets.Item(1);[void]$book.Activate();[void]$sheet.Activate();$excel.WindowState=-4137;Start-Sleep -Milliseconds 500;$window=$excel.ActiveWindow;$window.Zoom=100;$window.ScrollRow=1;$window.ScrollColumn=1;$selectionRange=$sheet.Range('B2:E6');[void]$selectionRange.Select()
    $macro="'Product.xlam'!";[void]$excel.Run($macro+'NxFocusEnable',255);Start-Sleep -Milliseconds 600;[void]$excel.Run($macro+'NxFocusControllerRefreshActiveSelection');Start-Sleep -Milliseconds 200
    $excel.WindowState=-4143;$excel.Top=40;$excel.Left=40;$excel.Width=760;$excel.Height=520;Start-Sleep -Milliseconds 450;[void]$excel.Run($macro+'NxFocusControllerRefreshActiveSelection');Start-Sleep -Milliseconds 200
    $backend=[string]$excel.Run($macro+'NxFocusControllerBackend')
    $dllDiagnostic=Get-FocusDiagnostic $excel ([int]$binding.pid) $macro 'NX_FOCUS_DLL_'
    $dll=Test-PhysicalMove $excel ([int]$binding.pid) 'NX_FOCUS_DLL_' 13 9
    $cases.Add([pscustomobject]@{name='dll_physical_viewport_sync';status=if($backend-ceq'dll'-and[bool]$dll.passed){'PASS'}else{'FAIL'};details=[pscustomobject]@{backend=$backend;diagnostic=$dllDiagnostic;measurement=$dll}})

    [void]$excel.Run($macro+'NxFocusSuspendForOperation');$hostAddin.Connect=$false;Start-Sleep -Milliseconds 180;[void]$excel.Run($macro+'NxFocusResumeAfterOperation');Start-Sleep -Milliseconds 400
    $fallbackBackend=[string]$excel.Run($macro+'NxFocusControllerBackend')
    $fallbackDiagnostic=Get-FocusDiagnostic $excel ([int]$binding.pid) $macro 'NX_FOCUS_OVERLAY_'
    $fallbackWindows=@([NxViewportApi]::Find([int]$binding.pid,'NX_FOCUS_DLL_')).Count+@([NxViewportApi]::Find([int]$binding.pid,'NX_FOCUS_OVERLAY_')).Count
    $cases.Add([pscustomobject]@{name='conditional_fallback_after_dll_disconnect';status=if($fallbackBackend-ceq'cf-fallback'-and$fallbackWindows-eq0){'PASS'}else{'FAIL'};details=[pscustomobject]@{backend=$fallbackBackend;diagnostic=$fallbackDiagnostic;overlay_windows=$fallbackWindows}})

    [void]$excel.Run($macro+'NxFocusDisable');Start-Sleep -Milliseconds 250
    $residualDll=@([NxViewportApi]::Find([int]$binding.pid,'NX_FOCUS_DLL_')).Count
    $residualVba=@([NxViewportApi]::Find([int]$binding.pid,'NX_FOCUS_OVERLAY_')).Count
    $timer=0 # The conditional-format fallback owns no overlay timer.
    $cases.Add([pscustomobject]@{name='residual_windows_zero';status=if($residualDll-eq0-and$residualVba-eq0){'PASS'}else{'FAIL'};details=[pscustomobject]@{dll=$residualDll;vba=$residualVba;timer=$timer;timer_source='conditional-backend-no-overlay-timer'}})
} catch {$diagnostic=$_.Exception.Message} finally {
    if($null-ne$book){try{$book.Close($false)}catch{}}
    if($null-ne$addin){try{$addin.Close($false)}catch{}}
    if($null-ne$binding){$cleanup=Close-OwnedExcel $excel $binding @($selectionRange,$window,$hostAddin,$comAddins,$sheet,$sheets,$book,$addin,$books);$excel=$null;$binding=$null}
    if($registered){
        try{$unregisterOutput=@(& (Join-Path $SourceRoot 'build/Register-NxHost.ps1') -Action Unregister -X86Dll $x86 -X64Dll $x64);$afterHash=Marker $unregisterOutput 'REGISTRY_SNAPSHOT_AFTER|'}catch{if([string]::IsNullOrEmpty($diagnostic)){$diagnostic=$_.Exception.Message}else{$diagnostic+=' | '+$_.Exception.Message}}
    }
}
$registryPass=(-not[string]::IsNullOrEmpty($beforeHash)-and$beforeHash-ceq$afterHash)
$cases.Add([pscustomobject]@{name='registry_snapshot_restored';status=if($registryPass){'PASS'}else{'FAIL'};details=[pscustomobject]@{before=$beforeHash;after=$afterHash}})
$passed=@($cases|Where-Object status -eq PASS).Count;$allPass=($cases.Count-eq4-and$passed-eq4-and[string]::IsNullOrEmpty($diagnostic)-and$null-ne$cleanup-and[string]$cleanup.exit_mode-ceq'NATURAL')
$receipt=[pscustomobject][ordered]@{schema_version=1;suite='FocusViewportSync';status=if($allPass){'PASS'}else{'DIAGNOSTIC'};artifact_sha256=Get-Sha256 $ArtifactPath;nxhost32_sha256=Get-Sha256 $x86;nxhost64_sha256=Get-Sha256 $x64;registered_com_addins=@($registeredAddins);cases=$cases.ToArray();cleanup=$cleanup;diagnostic=$diagnostic}
$path=Join-Path $EvidenceRoot 'FocusViewportSync.json';Write-Json $path $receipt
if($allPass){Write-Output ('PASS|FocusViewportSync|'+$passed+'/'+$cases.Count);Write-Output $path;exit 0}
Write-Output ('DIAGNOSTIC|FocusViewportSync|'+$passed+'/'+$cases.Count);Write-Output $path;exit 24
