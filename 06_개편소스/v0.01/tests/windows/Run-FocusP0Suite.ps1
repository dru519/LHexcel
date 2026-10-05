param(
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RunId = '',
    [string]$ArtifactPath = '',
    [string]$NxHostRoot = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$DataRoot = [IO.Path]::GetFullPath($DataRoot)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
if ($RunId -notmatch '^[0-9a-f-]{36}$') { throw 'RunId must be a lowercase GUID' }
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing FocusP0 evidence root is rejected' }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$originalProfileOverride = [string]$env:LHEXCEL_PROFILE_ROOT
$isolatedProfileRoot = Join-Path $EvidenceRoot 'profile/LHexcel'
[void](New-Item -ItemType Directory -Path $isolatedProfileRoot -Force)
$env:LHEXCEL_PROFILE_ROOT = $isolatedProfileRoot
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes,System.Windows.Forms

$PixelTolerance = 2
$MoveCount = 500
$DllMoveCount = 500
$WarmupCount = 20
$IncludeDll = $true
$script:Cases = New-Object 'Collections.Generic.List[object]'
$script:VbaParity = $null
$script:DllParity = $null

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}
function Get-Sha256([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Write-Json([string]$Path,[object]$Value) { $utf8=New-Object Text.UTF8Encoding($false);[IO.File]::WriteAllText($Path,($Value|ConvertTo-Json -Depth 40),$utf8) }
function Add-Case([string]$Name,[bool]$Passed,[object]$Details) { $script:Cases.Add([pscustomobject][ordered]@{ordinal=$script:Cases.Count+1;name=$Name;status=if($Passed){'PASS'}else{'FAIL'};details=$Details}) }
function Get-P95([double[]]$Values) { $sorted=@($Values|Sort-Object);if($sorted.Count-eq0){return 0.0};$index=[Math]::Max(0,[Math]::Ceiling($sorted.Count*.95)-1);[Math]::Round([double]$sorted[$index],3) }
function Get-MarkerValue([object[]]$Lines,[string]$Prefix) { $match=@($Lines|ForEach-Object{[string]$_}|Where-Object{$_.StartsWith($Prefix,[StringComparison]::Ordinal)}|Select-Object -Last 1);if($match.Count-eq0){return ''};$match[0].Substring($Prefix.Length) }

function Initialize-NativeApi {
    if ('NxFocusP0Api' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
public sealed class NxFocusP0Window { public long Handle; public string Title; }
public sealed class NxFocusP0Rect { public int Left; public int Top; public int Right; public int Bottom; }
public static class NxFocusP0Api {
  public delegate bool EnumProc(IntPtr hwnd, IntPtr parameter);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, IntPtr parameter);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr hwnd, StringBuilder value, int maximum);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool GetLayeredWindowAttributes(IntPtr hwnd, out uint colorKey, out byte alpha, out uint flags);
  [DllImport("user32.dll")] public static extern int GetGuiResources(IntPtr process, int flag);
  [DllImport("user32.dll")] public static extern int GetDpiForWindow(IntPtr hwnd);
  [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out NativeRect rect);
  [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr insertAfter, int x, int y, int width, int height, uint flags);
  struct NativeRect { public int Left, Top, Right, Bottom; }
  public static int CountWindows(int wantedPid, string prefix) {
    int count=0; EnumWindows((hwnd,p)=>{uint pid;GetWindowThreadProcessId(hwnd,out pid);if(pid!=(uint)wantedPid)return true;var text=new StringBuilder(512);GetWindowText(hwnd,text,text.Capacity);if(text.ToString().StartsWith(prefix,StringComparison.Ordinal))count++;return true;},IntPtr.Zero);return count;
  }
  public static NxFocusP0Window[] FindVisibleWindows(int wantedPid, string prefix) {
    var rows=new List<NxFocusP0Window>(); EnumWindows((hwnd,p)=>{uint pid;GetWindowThreadProcessId(hwnd,out pid);if(pid!=(uint)wantedPid||!IsWindowVisible(hwnd))return true;var text=new StringBuilder(512);GetWindowText(hwnd,text,text.Capacity);string title=text.ToString();if(title.StartsWith(prefix,StringComparison.Ordinal))rows.Add(new NxFocusP0Window{Handle=hwnd.ToInt64(),Title=title});return true;},IntPtr.Zero);return rows.ToArray();
  }
  public static NxFocusP0Rect WindowRect(long handle) {
    NativeRect rect; if(!GetWindowRect(new IntPtr(handle),out rect))throw new InvalidOperationException("GetWindowRect failed");
    return new NxFocusP0Rect{Left=rect.Left,Top=rect.Top,Right=rect.Right,Bottom=rect.Bottom};
  }
  public static bool MoveWindow(long handle, int left, int top) {
    const uint SWP_NOSIZE=0x0001, SWP_NOZORDER=0x0004, SWP_NOACTIVATE=0x0010;
    return SetWindowPos(new IntPtr(handle),IntPtr.Zero,left,top,0,0,SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
  }
}
'@
}

function Get-UiaCellBounds([object]$Excel,[int]$ProcessId,[string]$Address) {
    $root=[Windows.Automation.AutomationElement]::FromHandle([IntPtr][Int64]$Excel.Hwnd)
    $condition=[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,$Address)
    $matches=@($root.FindAll([Windows.Automation.TreeScope]::Descendants,$condition)|Where-Object{
        [int]$_.Current.ProcessId -eq $ProcessId -and [string]$_.Current.ClassName -ceq 'XLSpreadsheetCell' -and $_.Current.BoundingRectangle.Width -gt 0 -and $_.Current.BoundingRectangle.Height -gt 0
    })
    if($matches.Count -lt 1){throw ('UIA XLSpreadsheetCell unavailable: '+$Address)}
    $bounds=$matches[0].Current.BoundingRectangle
    [pscustomobject][ordered]@{left=[int]$bounds.Left;top=[int]$bounds.Top;width=[int]$bounds.Width;height=[int]$bounds.Height;right=[int]$bounds.Right;bottom=[int]$bounds.Bottom}
}

$script:SegmentTokens=@('ROW_LEFT','ROW_RIGHT','COLUMN_TOP','COLUMN_BOTTOM','SELECT_TOP','SELECT_RIGHT','SELECT_BOTTOM','SELECT_LEFT')
function Get-OverlayMap([int]$ProcessId,[string]$Prefix) {
    $map=@{}
    foreach($window in @([NxFocusP0Api]::FindVisibleWindows($ProcessId,$Prefix))){
        $name=[string]$window.Title
        $handle=[IntPtr][Int64]$window.Handle
        $element=[Windows.Automation.AutomationElement]::FromHandle($handle)
        if($null-eq$element){continue}
        foreach($token in $script:SegmentTokens){
            if(-not $name.StartsWith($Prefix+$token,[StringComparison]::Ordinal)){continue}
            $bounds=$element.Current.BoundingRectangle
            if($bounds.Width -le 0 -or $bounds.Height -le 0){continue}
            $colorKey=[uint32]0;$alpha=[byte]0;$flags=[uint32]0
            [void][NxFocusP0Api]::GetLayeredWindowAttributes($handle,[ref]$colorKey,[ref]$alpha,[ref]$flags)
            $map[$token]=[pscustomobject][ordered]@{left=[int]$bounds.Left;top=[int]$bounds.Top;width=[int]$bounds.Width;height=[int]$bounds.Height;right=[int]$bounds.Right;bottom=[int]$bounds.Bottom;alpha=[int]$alpha;handle=[int64]$handle}
        }
    }
    $map
}

function Get-SelectionBounds([object]$Excel,[int]$ProcessId,[string]$First,[string]$Last) {
    $a=Get-UiaCellBounds $Excel $ProcessId $First
    $b=Get-UiaCellBounds $Excel $ProcessId $Last
    [pscustomobject][ordered]@{left=[Math]::Min($a.left,$b.left);top=[Math]::Min($a.top,$b.top);right=[Math]::Max($a.right,$b.right);bottom=[Math]::Max($a.bottom,$b.bottom);width=[Math]::Max($a.right,$b.right)-[Math]::Min($a.left,$b.left);height=[Math]::Max($a.bottom,$b.bottom)-[Math]::Min($a.top,$b.top)}
}

function Test-CrossEdges([hashtable]$Map,[object]$Cell,[int]$Tolerance) {
    $required=@('ROW_LEFT','ROW_RIGHT','COLUMN_TOP','COLUMN_BOTTOM')
    $missing=@($required|Where-Object{-not $Map.ContainsKey($_)})
    if($missing.Count -gt 0){return [pscustomobject]@{passed=$false;max_error=999;missing=$missing;interior_overlap=$true}}
    $errors=@(
        [Math]::Abs($Map.ROW_LEFT.right-$Cell.left),[Math]::Abs($Map.ROW_LEFT.top-$Cell.top),[Math]::Abs($Map.ROW_LEFT.height-$Cell.height),
        [Math]::Abs($Map.ROW_RIGHT.left-$Cell.right),[Math]::Abs($Map.ROW_RIGHT.top-$Cell.top),[Math]::Abs($Map.ROW_RIGHT.height-$Cell.height),
        [Math]::Abs($Map.COLUMN_TOP.bottom-$Cell.top),[Math]::Abs($Map.COLUMN_TOP.left-$Cell.left),[Math]::Abs($Map.COLUMN_TOP.width-$Cell.width),
        [Math]::Abs($Map.COLUMN_BOTTOM.top-$Cell.bottom),[Math]::Abs($Map.COLUMN_BOTTOM.left-$Cell.left),[Math]::Abs($Map.COLUMN_BOTTOM.width-$Cell.width)
    )
    $interior=$false
    foreach($token in $required){$r=$Map[$token];$x=[Math]::Max(0,[Math]::Min($r.right,$Cell.right)-[Math]::Max($r.left,$Cell.left));$y=[Math]::Max(0,[Math]::Min($r.bottom,$Cell.bottom)-[Math]::Max($r.top,$Cell.top));if($x*$y-gt0){$interior=$true}}
    $max=[int](($errors|Measure-Object -Maximum).Maximum)
    [pscustomobject][ordered]@{passed=($max-le$Tolerance-and-not$interior);max_error=$max;missing=$missing;interior_overlap=$interior;cell=$Cell;segments=$Map}
}

function Test-OutlineEdges([hashtable]$Map,[object]$Cell,[int]$Tolerance,[int]$Dpi) {
    $required=@('SELECT_TOP','SELECT_RIGHT','SELECT_BOTTOM','SELECT_LEFT')
    $missing=@($required|Where-Object{-not $Map.ContainsKey($_)})
    if($missing.Count -gt 0){return [pscustomobject]@{passed=$false;max_error=999;missing=$missing;selection_interior_overlap=$true;expected_thickness=0}}
    $expected=[Math]::Max(1,[int][Math]::Floor((2.0*$Dpi/96.0)+0.5))
    $errors=@(
        [Math]::Abs($Map.SELECT_TOP.bottom-$Cell.top),[Math]::Abs($Map.SELECT_TOP.left-$Cell.left),[Math]::Abs($Map.SELECT_TOP.right-$Cell.right),[Math]::Abs($Map.SELECT_TOP.height-$expected),
        [Math]::Abs($Map.SELECT_RIGHT.left-$Cell.right),[Math]::Abs($Map.SELECT_RIGHT.top-$Cell.top),[Math]::Abs($Map.SELECT_RIGHT.bottom-$Cell.bottom),[Math]::Abs($Map.SELECT_RIGHT.width-$expected),
        [Math]::Abs($Map.SELECT_BOTTOM.top-$Cell.bottom),[Math]::Abs($Map.SELECT_BOTTOM.left-$Cell.left),[Math]::Abs($Map.SELECT_BOTTOM.right-$Cell.right),[Math]::Abs($Map.SELECT_BOTTOM.height-$expected),
        [Math]::Abs($Map.SELECT_LEFT.right-$Cell.left),[Math]::Abs($Map.SELECT_LEFT.top-$Cell.top),[Math]::Abs($Map.SELECT_LEFT.bottom-$Cell.bottom),[Math]::Abs($Map.SELECT_LEFT.width-$expected)
    )
    $selectionInteriorOverlap=$false
    foreach($token in $required){$r=$Map[$token];$x=[Math]::Max(0,[Math]::Min($r.right,$Cell.right)-[Math]::Max($r.left,$Cell.left));$y=[Math]::Max(0,[Math]::Min($r.bottom,$Cell.bottom)-[Math]::Max($r.top,$Cell.top));if($x*$y-gt0){$selectionInteriorOverlap=$true}}
    $max=[int](($errors|Measure-Object -Maximum).Maximum)
    [pscustomobject][ordered]@{passed=($max-le$Tolerance-and-not$selectionInteriorOverlap);max_error=$max;missing=$missing;selection_interior_overlap=$selectionInteriorOverlap;expected_thickness=$expected;segments=$Map}
}

function Test-FocusEdges([hashtable]$Map,[object]$Cell,[int]$Tolerance,[int]$Dpi) {
    $cross=Test-CrossEdges $Map $Cell $Tolerance
    $outline=Test-OutlineEdges $Map $Cell $Tolerance $Dpi
    [pscustomobject][ordered]@{passed=([bool]$cross.passed-and[bool]$outline.passed);max_error=[Math]::Max([int]$cross.max_error,[int]$outline.max_error);cross=$cross;outline=$outline;alpha=@($Map.Values|Select-Object -ExpandProperty alpha -Unique);cell=$Cell}
}

function Wait-Refresh([object]$Excel,[string]$Macro,[string]$Prefix='NX_FOCUS_OVERLAY_') {
    Start-Sleep -Milliseconds 120
    if($Prefix -ceq 'NX_FOCUS_DLL_') {[void]$Excel.Run($Macro+'NxFocusControllerRefreshActiveSelection')} else {[void]$Excel.Run($Macro+'NxFocusOverlayRefresh')}
    Start-Sleep -Milliseconds 60
}
function Invoke-EdgeCase([object]$Excel,[object]$Sheet,[int]$ProcessId,[string]$Macro,[string]$Name,[string]$RangeAddress,[string]$First,[string]$Last,[string]$Prefix='NX_FOCUS_OVERLAY_',[bool]$Record=$true) {
    $range=$Sheet.Range($RangeAddress)
    try{[void]$range.Select();Wait-Refresh $Excel $Macro $Prefix;$cell=Get-SelectionBounds $Excel $ProcessId $First $Last;$map=Get-OverlayMap $ProcessId $Prefix;$dpi=[NxFocusP0Api]::GetDpiForWindow([IntPtr][Int64]$Excel.Hwnd);$result=Test-FocusEdges $map $cell $PixelTolerance $dpi;if($Record){Add-Case $Name ([bool]$result.passed) $result};return $result}finally{Release-ComObject $range}
}

function Invoke-ViewportDragCase([object]$Excel,[object]$Sheet,[int]$ProcessId,[string]$Macro,[string]$Prefix) {
    # Moving the physical XLMAIN viewport without changing VisibleRange reproduces
    # the same cell-granularity gap seen during smooth scrollbar/thumb dragging.
    $previousState=[int]$Excel.WindowState
    $range=$Sheet.Range('B2')
    $samples=New-Object 'Collections.Generic.List[object]'
    $original=$null
    try{
        $Excel.WindowState=-4143
        $Excel.Left=120;$Excel.Top=90;$Excel.Width=940;$Excel.Height=680
        [void]$range.Select();Wait-Refresh $Excel $Macro $Prefix
        $original=[NxFocusP0Api]::WindowRect([int64]$Excel.Hwnd)
        $left=[int]$original.Left;$top=[int]$original.Top
        foreach($delta in @(@(11,7),@(9,13),@(-6,8))){
            $left += [int]$delta[0];$top += [int]$delta[1]
            if(-not[NxFocusP0Api]::MoveWindow([int64]$Excel.Hwnd,$left,$top)){throw 'physical viewport drag failed'}
            Start-Sleep -Milliseconds 180
            $cell=Get-SelectionBounds $Excel $ProcessId 'B2' 'B2'
            $map=Get-OverlayMap $ProcessId $Prefix
            $dpi=[NxFocusP0Api]::GetDpiForWindow([IntPtr][Int64]$Excel.Hwnd)
            $sample=Test-FocusEdges $map $cell $PixelTolerance $dpi
            $samples.Add([pscustomobject]@{left=$left;top=$top;passed=[bool]$sample.passed;max_error=[int]$sample.max_error})
        }
        $maximum=[int](($samples|Measure-Object -Property max_error -Maximum).Maximum)
        [pscustomobject][ordered]@{passed=(@($samples|Where-Object{-not $_.passed}).Count-eq0-and$maximum-le$PixelTolerance);max_error=$maximum;samples=$samples.ToArray()}
    } finally {
        if($null-ne$original){[void][NxFocusP0Api]::MoveWindow([int64]$Excel.Hwnd,[int]$original.Left,[int]$original.Top)}
        $Excel.WindowState=$previousState
        Wait-Refresh $Excel $Macro $Prefix
        Release-ComObject $range
    }
}

function Get-UserFingerprint([object]$Workbook,[object]$Sheet,[object]$Range) {
    [pscustomobject][ordered]@{saved=[bool]$Workbook.Saved;formats=[int]$Range.FormatConditions.Count;shapes=[int]$Sheet.Shapes.Count;names=[int]$Workbook.Names.Count}
}

function Close-OwnedExcel([object]$Excel,[object]$Binding,[object[]]$Objects,[string]$Label) {
    foreach($item in $Objects){try{Release-ComObject $item}catch{}}
    if($null-ne$Excel){try{$Excel.DisplayAlerts=$false;$Excel.Quit()}catch{};Release-ComObject $Excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null-eq$Binding){return [pscustomobject]@{exit_mode='NOT_STARTED';failure='';pid=0}}
    $stop=Stop-ExactProcessAfterGrace $Binding.process $Label 15000 -Detailed
    try{$Binding.process.Dispose()}catch{}
    [pscustomobject]@{exit_mode=[string]$stop.exit_mode;failure=[string]$stop.failure;pid=[int]$Binding.pid}
}

[void](Initialize-NativeApi)
$artifact=Join-Path $EvidenceRoot 'Product.xlam'
$manifest=Join-Path $SourceRoot 'build/manifests/Product.json'
if([string]::IsNullOrWhiteSpace($ArtifactPath)){
    & (Join-Path $SourceRoot 'build/Build-Xlam.ps1') -DataRoot $DataRoot -OutputPath $artifact -ManifestPath $manifest -RibbonPath (Join-Path $SourceRoot 'src/ribbon/customUI14.xml')
    if($LASTEXITCODE-ne0-or-not(Test-Path -LiteralPath $artifact)){throw 'FocusP0 Product build failed'}
}else{
    $ArtifactPath=[IO.Path]::GetFullPath($ArtifactPath)
    if(-not(Test-Path -LiteralPath $ArtifactPath -PathType Leaf)-or[IO.Path]::GetExtension($ArtifactPath)-ine'.xlam'){throw 'FocusP0 supplied Product artifact rejected'}
    Copy-Item -LiteralPath $ArtifactPath -Destination $artifact
    if((Get-Sha256 $artifact)-cne(Get-Sha256 $ArtifactPath)){throw 'FocusP0 supplied Product artifact copy mismatch'}
}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'FocusP0 requires zero pre-existing Excel processes'}

$excel=$null;$binding=$null;$books=$null;$addin=$null;$workbook=$null;$sheet=$null;$secondWorkbook=$null;$secondSheet=$null;$userRange=$null;$userRule=$null;$userShape=$null;$userName=$null
$vbaCleanup=$null;$diagnostic=$null;$refreshP95=0.0;$gdiDelta=0;$userDelta=0;$settingsPath=''
try{
    $baseline=Get-ExcelProcessBaseline
    $excel=New-Object -ComObject Excel.Application
    $binding=Get-ExactExcelProcessOwnership $excel $baseline 'FocusP0 VBA Excel'
    if(-not[bool]$binding.owned){throw 'FocusP0 VBA Excel ownership rejected'}
    $excel.Visible=$true;$excel.DisplayAlerts=$false;$books=$excel.Workbooks;$addin=$books.Open($artifact,$false,$true);$workbook=$books.Add();$sheet=$workbook.Worksheets.Item(1);$sheet.Name='FocusP0'
    $excel.WindowState=-4137;$excel.ActiveWindow.Zoom=100;$excel.ActiveWindow.ScrollRow=1;$excel.ActiveWindow.ScrollColumn=1
    $userRange=$sheet.Range('A1:Z100');$userRule=$userRange.FormatConditions.Add(2,$null,'=MOD(ROW()+COLUMN(),3)=0');$userShape=$sheet.Shapes.AddShape(1,10,10,30,20);$userName=$workbook.Names.Add('P0Keep','=FocusP0!$A$1')
    $fixture=Join-Path $EvidenceRoot 'FocusP0Fixture.xlsx';$workbook.SaveAs($fixture,51);$workbook.Saved=$true;$before=Get-UserFingerprint $workbook $sheet $userRange
    $macro="'Product.xlam'!";[void]$sheet.Range('B2').Select();$initialSettingsResult=[string]$excel.Run($macro+'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#FF0000',60,$true,$true);if($initialSettingsResult-ne'PASS'){throw ('Initial focus settings failed: '+$initialSettingsResult)};Wait-Refresh $excel $macro

    $script:VbaMatrix=[ordered]@{}
    $script:VbaMatrix['zoom_100']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'zoom_100_uia_edges' 'B2' 'B2' 'B2'
    $excel.ActiveWindow.Zoom=75;$script:VbaMatrix['zoom_75']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'zoom_75_uia_edges' 'B2' 'B2' 'B2'
    $excel.ActiveWindow.Zoom=125;$script:VbaMatrix['zoom_125']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'zoom_125_uia_edges' 'B2' 'B2' 'B2'
    $excel.ActiveWindow.Zoom=100;$script:VbaMatrix['single_cell']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'single_cell_uia_edges' 'B2' 'B2' 'B2'
    $script:VbaMatrix['multi_cell']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'multi_cell_uia_edges' 'B2:D4' 'B2' 'D4'
    $merge=$sheet.Range('B2:C3');[void]$merge.Merge();try{$script:VbaMatrix['merged_cell']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'merged_cell_uia_edges' 'B2' 'B2' 'B2'}finally{[void]$merge.UnMerge();Release-ComObject $merge}

    $excel.ActiveWindow.ScrollColumn=5;$excel.ActiveWindow.ScrollRow=20;$script:VbaMatrix['scroll']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'scroll_uia_edges' 'H25' 'H25' 'H25'
    $vbaViewportDrag=Invoke-ViewportDragCase $excel $sheet ([int]$binding.pid) $macro 'NX_FOCUS_OVERLAY_';Add-Case 'smooth_scroll_drag_vba_uia_edges' ([bool]$vbaViewportDrag.passed) $vbaViewportDrag
    $excel.ActiveWindow.ScrollColumn=1;$excel.ActiveWindow.ScrollRow=1;[void]$sheet.Range('C3').Select();$excel.ActiveWindow.FreezePanes=$true;$script:VbaMatrix['freeze_panes']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'freeze_panes_uia_edges' 'D4' 'D4' 'D4'
    $excel.ActiveWindow.FreezePanes=$false;$excel.ActiveWindow.SplitColumn=2;$excel.ActiveWindow.SplitRow=2;$script:VbaMatrix['split_window']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'split_window_uia_edges' 'D4' 'D4' 'D4';$excel.ActiveWindow.SplitColumn=0;$excel.ActiveWindow.SplitRow=0

    $sheet.Rows.Item(10).Hidden=$true;[void]$sheet.Range('B10').Select();Wait-Refresh $excel $macro;$hiddenRow=([int]$excel.Run($macro+'NxFocusOverlayVisibleWindowCount')-eq0);$sheet.Rows.Item(10).Hidden=$false
    $sheet.Columns.Item(2).Hidden=$true;[void]$sheet.Range('B2').Select();Wait-Refresh $excel $macro;$hiddenColumn=([int]$excel.Run($macro+'NxFocusOverlayVisibleWindowCount')-eq0);$sheet.Columns.Item(2).Hidden=$false
    $filterRange=$sheet.Range('A1:B5');$filterValues=New-Object 'object[,]' 5,2;$filterValues[0,0]='K';$filterValues[0,1]='V';$filterValues[1,0]='keep';$filterValues[1,1]=1;$filterValues[2,0]='drop';$filterValues[2,1]=2;$filterValues[3,0]='keep';$filterValues[3,1]=3;$filterValues[4,0]='keep';$filterValues[4,1]=4;$filterRange.Value2=$filterValues;[void]$filterRange.AutoFilter(1,'keep');Start-Sleep -Milliseconds 200;$filterResult=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'filter_internal_uia' 'B2' 'B2' 'B2' 'NX_FOCUS_OVERLAY_' $false;$sheet.AutoFilterMode=$false;Release-ComObject $filterRange
    $sheet.Protect();$protectedResult=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'protected_internal_uia' 'B2' 'B2' 'B2' 'NX_FOCUS_OVERLAY_' $false;$sheet.Unprotect()
    $script:VbaMatrix['hidden_filtered_protected']=[pscustomobject]@{passed=($hiddenRow-and$hiddenColumn-and$filterResult.passed-and$protectedResult.passed);hidden_row=$hiddenRow;hidden_column=$hiddenColumn;filter=$filterResult;protected=$protectedResult};Add-Case 'hidden_filtered_protected' ([bool]$script:VbaMatrix['hidden_filtered_protected'].passed) $script:VbaMatrix['hidden_filtered_protected']

    $secondWorkbook=$books.Add();$secondSheet=$secondWorkbook.Worksheets.Item(1);[void]$secondWorkbook.Activate();[void]$secondSheet.Activate();$secondResult=Invoke-EdgeCase $excel $secondSheet ([int]$binding.pid) $macro 'second_window_internal_uia' 'C3' 'C3' 'C3' 'NX_FOCUS_OVERLAY_' $false;[void]$workbook.Activate();[void]$sheet.Activate();$firstResult=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'first_window_internal_uia' 'B2' 'B2' 'B2' 'NX_FOCUS_OVERLAY_' $false;$script:VbaMatrix['multiple_workbooks']=[pscustomobject]@{passed=($secondResult.passed-and$firstResult.passed);second=$secondResult;first=$firstResult};Add-Case 'multiple_workbooks_windows' ([bool]$script:VbaMatrix['multiple_workbooks'].passed) $script:VbaMatrix['multiple_workbooks']

    $settingsResult=[string]$excel.Run($macro+'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#FF0000',60,$true,$true);$settingsPath=[string]$excel.Run($macro+'NxFocusSettingsPath');$normalHash=Get-Sha256 $settingsPath;[void]$excel.Run($macro+'NxFocusDisable');[void]$excel.Run($macro+'NxFocusEnable',255);Wait-Refresh $excel $macro;$normalMap=Get-OverlayMap ([int]$binding.pid) 'NX_FOCUS_OVERLAY_';$normalRestart=($settingsResult-eq'PASS'-and$normalMap.Count-eq8-and@($normalMap.Values|Where-Object{$_.alpha-ne153}).Count-eq0);Add-Case 'settings_normal_restart' $normalRestart @{result=$settingsResult;sha256=$normalHash;alpha=@($normalMap.Values|Select-Object -ExpandProperty alpha -Unique)}
    $backupPath=$settingsPath+'.bak';Add-Case 'settings_existing_backup' (Test-Path -LiteralPath $backupPath -PathType Leaf) @{backup=$backupPath;sha256=if(Test-Path $backupPath){Get-Sha256 $backupPath}else{''}}
    $appearanceBefore=[string]$excel.Run($macro+'NxFocusOverlayAppearanceSignature');$fileBefore=Get-Sha256 $settingsPath
    $lockPath=$settingsPath+'.lck';$lock=[IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None);try{$lockResult=[string]$excel.Run($macro+'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#00B050',70,$true,$true)}finally{$lock.Dispose()};$lockPass=($lockResult.StartsWith('FAIL|')-and(Get-Sha256 $settingsPath)-eq$fileBefore-and[string]$excel.Run($macro+'NxFocusOverlayAppearanceSignature')-ceq$appearanceBefore);Add-Case 'settings_lock_conflict' $lockPass @{result=$lockResult}
    $tempPath=$settingsPath+'.tmp';if(Test-Path $tempPath){throw 'Unexpected temp path before write failure test'};[void](New-Item -ItemType Directory -Path $tempPath);try{$writeResult=[string]$excel.Run($macro+'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#00B050',70,$true,$true)}finally{if(-not ([IO.Path]::GetFullPath($tempPath).StartsWith([IO.Path]::GetFullPath($isolatedProfileRoot),[StringComparison]::OrdinalIgnoreCase))){throw 'Temp cleanup escaped isolated profile'};Remove-Item -LiteralPath $tempPath -Force};$writePass=($writeResult.StartsWith('FAIL|')-and(Get-Sha256 $settingsPath)-eq$fileBefore-and[string]$excel.Run($macro+'NxFocusOverlayAppearanceSignature')-ceq$appearanceBefore);Add-Case 'settings_write_failure' $writePass @{result=$writeResult}
    [IO.File]::WriteAllText($settingsPath,'corrupt',[Text.Encoding]::UTF8);$recovered=[string]$excel.Run($macro+'NxFocusLoadColors');$corruptPass=($recovered.StartsWith('#FF0000')-and(Test-Path $backupPath));Add-Case 'settings_corrupt_recovery' $corruptPass @{reloaded=$recovered};Copy-Item -LiteralPath $backupPath -Destination $settingsPath -Force;[void]$excel.Run($macro+'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#FF0000',60,$true,$true)

    [void]$workbook.Activate();[void]$sheet.Activate();$vbaParityResult=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $macro 'vba_parity_internal' 'B2' 'B2' 'B2' 'NX_FOCUS_OVERLAY_' $false;$script:VbaParity=[pscustomobject]@{backend=[string]$excel.Run($macro+'NxFocusControllerBackend');edges=$vbaParityResult;alpha=$vbaParityResult.alpha;cell=$vbaParityResult.cell}
    $workbook.Saved=$true;$beforeFocus=Get-UserFingerprint $workbook $sheet $userRange;$process=Get-Process -Id ([int]$binding.pid);$gdiBefore=[NxFocusP0Api]::GetGuiResources($process.Handle,0);$userBefore=[NxFocusP0Api]::GetGuiResources($process.Handle,1);$windowCountBefore=[int]$excel.Run($macro+'NxFocusOverlayWindowCount');$timings=New-Object 'Collections.Generic.List[double]'
    for($index=0;$index-lt($MoveCount+$WarmupCount);$index++){$cellObject=$sheet.Cells.Item(2+($index%37),2+(($index*7)%19));try{[void]$cellObject.Select();$watch=[Diagnostics.Stopwatch]::StartNew();[void]$excel.Run($macro+'NxFocusOverlayRefresh');$watch.Stop();if($index-ge$WarmupCount){$timings.Add($watch.Elapsed.TotalMilliseconds)}}finally{Release-ComObject $cellObject}}
    $refreshP95=Get-P95 $timings.ToArray();$gdiAfter=[NxFocusP0Api]::GetGuiResources($process.Handle,0);$userAfter=[NxFocusP0Api]::GetGuiResources($process.Handle,1);$process.Dispose();$gdiDelta=$gdiAfter-$gdiBefore;$userDelta=$userAfter-$userBefore;$windowCountAfter=[int]$excel.Run($macro+'NxFocusOverlayWindowCount');Add-Case 'five_hundred_moves_p95' ($refreshP95-le50-and$windowCountBefore-eq8-and$windowCountAfter-eq8) @{refresh_p95_ms=$refreshP95;windows_before=$windowCountBefore;windows_after=$windowCountAfter};Add-Case 'gdi_handle_delta' ($gdiDelta-le8-and$userDelta-le8) @{gdi_delta=$gdiDelta;user_delta=$userDelta}
    $after=Get-UserFingerprint $workbook $sheet $userRange;Add-Case 'user_state_preserved' ($beforeFocus.saved-eq$after.saved-and$beforeFocus.formats-eq$after.formats-and$beforeFocus.shapes-eq$after.shapes-and$beforeFocus.names-eq$after.names) @{before=$beforeFocus;after=$after}
    [void]$excel.Run($macro+'NxFocusDisable');Start-Sleep -Milliseconds 200;$residualWindows=[NxFocusP0Api]::CountWindows([int]$binding.pid,'NX_FOCUS_OVERLAY_');$residualTimer=[int]$excel.Run($macro+'NxFocusOverlayTimerCount');Add-Case 'residual_overlay_windows' ($residualWindows-eq0) @{windows=$residualWindows};Add-Case 'residual_timers' ($residualTimer-eq0) @{vba_timer=$residualTimer}
}catch{$diagnostic=$_.Exception.Message}finally{
    if($null-ne$secondWorkbook){try{$secondWorkbook.Close($false)}catch{}}
    if($null-ne$workbook){try{$workbook.Close($false)}catch{}}
    if($null-ne$addin){try{$addin.Close($false)}catch{}}
    if($null-ne$binding){$vbaCleanup=Close-OwnedExcel $excel $binding @($userName,$userShape,$userRule,$userRange,$secondSheet,$sheet,$secondWorkbook,$workbook,$addin,$books) 'FocusP0 VBA Excel'}
}

$dllCleanup=$null;$registration='NOT_STARTED';$dllDiagnostic=$null;$registryBeforeHash='';$registryAfterHash='';$registrationAttempted=$false
$dllMoveP95=0.0;$dllGdiDelta=0;$dllUserDelta=0;$dllPrivateMemoryDelta=0
if($IncludeDll -and [string]::IsNullOrEmpty($diagnostic)){
    $hostRoot=Join-Path $EvidenceRoot 'NxHost';$x86=Join-Path $hostRoot 'x86/NxHost32.dll';$x64=Join-Path $hostRoot 'x64/NxHost64.dll';$registrationScript=Join-Path $SourceRoot 'build/Register-NxHost.ps1'
    if([string]::IsNullOrWhiteSpace($NxHostRoot)){
        & (Join-Path $SourceRoot 'build/Build-NxHost.ps1') -OutputRoot $hostRoot -Configuration Release
    }else{
        $NxHostRoot=[IO.Path]::GetFullPath($NxHostRoot);$suppliedX86=Join-Path $NxHostRoot 'NxHost32.dll';$suppliedX64=Join-Path $NxHostRoot 'NxHost64.dll'
        if(-not(Test-Path -LiteralPath $suppliedX86 -PathType Leaf)-or-not(Test-Path -LiteralPath $suppliedX64 -PathType Leaf)){throw 'FocusP0 supplied NxHost artifacts rejected'}
        [void](New-Item -ItemType Directory -Path (Split-Path -Parent $x86) -Force);[void](New-Item -ItemType Directory -Path (Split-Path -Parent $x64) -Force)
        Copy-Item -LiteralPath $suppliedX86 -Destination $x86;Copy-Item -LiteralPath $suppliedX64 -Destination $x64
        if((Get-Sha256 $x86)-cne(Get-Sha256 $suppliedX86)-or(Get-Sha256 $x64)-cne(Get-Sha256 $suppliedX64)){throw 'FocusP0 supplied NxHost artifact copy mismatch'}
        $global:LASTEXITCODE=0
    }
    if($LASTEXITCODE-ne0){$dllDiagnostic='NxHost build failed'}else{
        try{
            $registrationAttempted=$true
            $registerOutput=@(& $registrationScript -Action Register -OptIn -X86Dll $x86 -X64Dll $x64)
            if(-not($registerOutput|Where-Object{[string]$_-ceq'PASS|Register'})){throw 'NxHost register failed'}
            $registryBeforeHash=Get-MarkerValue $registerOutput 'REGISTRY_SNAPSHOT_BEFORE|';if([string]::IsNullOrEmpty($registryBeforeHash)){throw 'NxHost registry pre-snapshot missing'};$registration='REGISTERED'
            if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'DLL parity requires zero Excel processes'}
            $baseline=Get-ExcelProcessBaseline;$excel=$null;$binding=$null;$books=$null;$addin=$null;$workbook=$null;$worksheets=$null;$sheet=$null;$excelWindow=$null;$hostComAddin=$null;$comAddins=$null;$secondWorkbook=$null;$secondSheet=$null
            $dllResidue=999;$vbaFallbackResidue=999;$dllTimerResidue=999
            try{
                $interactive=Start-ExactInteractiveExcel $baseline 'FocusP0 DLL Excel'
                $excel=$interactive.excel;$binding=$interactive.binding;$interactive=$null
                if(-not[bool]$binding.owned){throw 'FocusP0 DLL Excel ownership rejected'}
                $excel.Visible=$true;$excel.DisplayAlerts=$false;$books=$excel.Workbooks;$addin=$books.Open($artifact,$false,$true);$workbook=$books.Add();$worksheets=$workbook.Worksheets;$sheet=$worksheets.Item(1);$sheet.Name='FocusP0Dll'
                $excel.WindowState=-4137;$excelWindow=$excel.ActiveWindow;$excelWindow.Zoom=100;$excelWindow.ScrollRow=1;$excelWindow.ScrollColumn=1
                $dllMacro="'Product.xlam'!";[void]$sheet.Range('B2').Select();[void]$excel.Run($dllMacro+'NxFocusEnable',255);Start-Sleep -Milliseconds 500
                $backend=[string]$excel.Run($dllMacro+'NxFocusControllerBackend');$backendDiagnostic=[string]$excel.Run($dllMacro+'NxFocusControllerBackendDiagnostic')
                if($backend-cne'dll'-or-not[string]::IsNullOrEmpty($backendDiagnostic)){throw ('Healthy DLL backend not active: '+$backend+' '+$backendDiagnostic)}
                $dllSettingsBefore=Get-Sha256 $settingsPath

                $dllMatrix=[ordered]@{}
                $dllMatrix['dll_zoom_100']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_zoom_100' 'B2' 'B2' 'B2' 'NX_FOCUS_DLL_' $false
                $excelWindow.Zoom=75;$dllMatrix['dll_zoom_75']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_zoom_75' 'B2' 'B2' 'B2' 'NX_FOCUS_DLL_' $false
                $excelWindow.Zoom=125;$dllMatrix['dll_zoom_125']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_zoom_125' 'B2' 'B2' 'B2' 'NX_FOCUS_DLL_' $false
                $excelWindow.Zoom=100;$dllMatrix['dll_single_cell']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_single_cell' 'B2' 'B2' 'B2' 'NX_FOCUS_DLL_' $false
                $dllMatrix['dll_multi_cell']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_multi_cell' 'B2:D4' 'B2' 'D4' 'NX_FOCUS_DLL_' $false
                $dllMerge=$sheet.Range('B2:C3');[void]$dllMerge.Merge();try{$dllMatrix['dll_merged_cell']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_merged_cell' 'B2' 'B2' 'B2' 'NX_FOCUS_DLL_' $false}finally{[void]$dllMerge.UnMerge();Release-ComObject $dllMerge}
                $excelWindow.ScrollColumn=5;$excelWindow.ScrollRow=20;$dllMatrix['dll_scroll']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_scroll' 'H25' 'H25' 'H25' 'NX_FOCUS_DLL_' $false
                $dllViewportDrag=Invoke-ViewportDragCase $excel $sheet ([int]$binding.pid) $dllMacro 'NX_FOCUS_DLL_';Add-Case 'smooth_scroll_drag_uia_edges' ([bool]$dllViewportDrag.passed) $dllViewportDrag
                $excelWindow.ScrollColumn=1;$excelWindow.ScrollRow=1;[void]$sheet.Range('C3').Select();$excelWindow.FreezePanes=$true;$dllMatrix['dll_freeze_panes']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_freeze_panes' 'D4' 'D4' 'D4' 'NX_FOCUS_DLL_' $false
                $excelWindow.FreezePanes=$false;$excelWindow.SplitColumn=2;$excelWindow.SplitRow=2;$dllMatrix['dll_split_window']=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_split_window' 'D4' 'D4' 'D4' 'NX_FOCUS_DLL_' $false;$excelWindow.SplitColumn=0;$excelWindow.SplitRow=0

                $sheet.Rows.Item(10).Hidden=$true;[void]$sheet.Range('B10').Select();Wait-Refresh $excel $dllMacro 'NX_FOCUS_DLL_';$dllHiddenRow=([NxFocusP0Api]::FindVisibleWindows([int]$binding.pid,'NX_FOCUS_DLL_').Count-eq0);$sheet.Rows.Item(10).Hidden=$false
                $sheet.Columns.Item(2).Hidden=$true;[void]$sheet.Range('B2').Select();Wait-Refresh $excel $dllMacro 'NX_FOCUS_DLL_';$dllHiddenColumn=([NxFocusP0Api]::FindVisibleWindows([int]$binding.pid,'NX_FOCUS_DLL_').Count-eq0);$sheet.Columns.Item(2).Hidden=$false
                $dllFilterRange=$sheet.Range('A1:B5');$dllFilterValues=New-Object 'object[,]' 5,2;$dllFilterValues[0,0]='K';$dllFilterValues[0,1]='V';$dllFilterValues[1,0]='keep';$dllFilterValues[1,1]=1;$dllFilterValues[2,0]='drop';$dllFilterValues[2,1]=2;$dllFilterValues[3,0]='keep';$dllFilterValues[3,1]=3;$dllFilterValues[4,0]='keep';$dllFilterValues[4,1]=4;$dllFilterRange.Value2=$dllFilterValues;[void]$dllFilterRange.AutoFilter(1,'keep');$dllFilterResult=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_filter' 'B2' 'B2' 'B2' 'NX_FOCUS_DLL_' $false;$sheet.AutoFilterMode=$false;Release-ComObject $dllFilterRange
                $sheet.Protect();$dllProtectedResult=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_protected' 'B2' 'B2' 'B2' 'NX_FOCUS_DLL_' $false;$sheet.Unprotect();$dllMatrix['dll_hidden_filtered_protected']=[pscustomobject]@{passed=($dllHiddenRow-and$dllHiddenColumn-and$dllFilterResult.passed-and$dllProtectedResult.passed);hidden_row=$dllHiddenRow;hidden_column=$dllHiddenColumn;filter=$dllFilterResult;protected=$dllProtectedResult}

                $secondWorkbook=$books.Add();$secondSheet=$secondWorkbook.Worksheets.Item(1);[void]$secondWorkbook.Activate();[void]$secondSheet.Activate();$dllSecond=Invoke-EdgeCase $excel $secondSheet ([int]$binding.pid) $dllMacro 'dll_second_window' 'C3' 'C3' 'C3' 'NX_FOCUS_DLL_' $false;[void]$workbook.Activate();[void]$sheet.Activate();$dllFirst=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_first_window' 'B2' 'B2' 'B2' 'NX_FOCUS_DLL_' $false;$dllMatrix['dll_multiple_workbooks']=[pscustomobject]@{passed=($dllSecond.passed-and$dllFirst.passed);second=$dllSecond;first=$dllFirst};$secondWorkbook.Close($false);Release-ComObject $secondSheet;$secondSheet=$null;Release-ComObject $secondWorkbook;$secondWorkbook=$null

                $pairMap=[ordered]@{dll_zoom_100='zoom_100';dll_zoom_75='zoom_75';dll_zoom_125='zoom_125';dll_single_cell='single_cell';dll_multi_cell='multi_cell';dll_merged_cell='merged_cell';dll_scroll='scroll';dll_freeze_panes='freeze_panes';dll_split_window='split_window';dll_hidden_filtered_protected='hidden_filtered_protected';dll_multiple_workbooks='multiple_workbooks'}
                $parityRows=New-Object 'Collections.Generic.List[object]';$matrixPass=$true;$alphaPass=$true
                foreach($dllName in $pairMap.Keys){$vbaName=[string]$pairMap[$dllName];$dllResult=$dllMatrix[$dllName];$vbaResult=$script:VbaMatrix[$vbaName];$rowPass=([bool]$dllResult.passed-and[bool]$vbaResult.passed);if($dllResult.PSObject.Properties.Name-contains'max_error'-and$vbaResult.PSObject.Properties.Name-contains'max_error'){$rowPass=$rowPass-and([int]$dllResult.max_error-le$PixelTolerance)-and([Math]::Abs([int]$dllResult.max_error-[int]$vbaResult.max_error)-le$PixelTolerance);$alphaPass=$alphaPass-and(($dllResult.alpha-join',')-ceq($vbaResult.alpha-join','))};if(-not$rowPass){$matrixPass=$false};$parityRows.Add([pscustomobject]@{dll=$dllName;vba=$vbaName;passed=$rowPass;dll_max_error=if($dllResult.PSObject.Properties.Name-contains'max_error'){$dllResult.max_error}else{$null};vba_max_error=if($vbaResult.PSObject.Properties.Name-contains'max_error'){$vbaResult.max_error}else{$null}})}
                $dllSettingsAfter=Get-Sha256 $settingsPath
                $script:DllParity=[pscustomobject]@{backend=$backend;matrix=$dllMatrix;parity=$parityRows.ToArray();settings_before=$dllSettingsBefore;settings_after=$dllSettingsAfter}

                [void]$workbook.Activate();[void]$sheet.Activate();$excelWindow.Zoom=100;$excelWindow.ScrollRow=1;$excelWindow.ScrollColumn=1
                $dllProcess=Get-Process -Id ([int]$binding.pid);$dllGdiBefore=[NxFocusP0Api]::GetGuiResources($dllProcess.Handle,0);$dllUserBefore=[NxFocusP0Api]::GetGuiResources($dllProcess.Handle,1);$dllProcess.Refresh();$dllMemoryBefore=$dllProcess.PrivateMemorySize64;$dllWindowsBefore=[NxFocusP0Api]::CountWindows([int]$binding.pid,'NX_FOCUS_DLL_');$dllTimings=New-Object 'Collections.Generic.List[double]'
                for($index=0;$index-lt($DllMoveCount+$WarmupCount);$index++){$dllCell=$sheet.Cells.Item(2+($index%37),2+(($index*7)%19));try{[void]$dllCell.Select();$watch=[Diagnostics.Stopwatch]::StartNew();[void]$excel.Run($dllMacro+'NxFocusControllerRefreshActiveSelection');$watch.Stop();if($index-ge$WarmupCount){$dllTimings.Add($watch.Elapsed.TotalMilliseconds)}}finally{Release-ComObject $dllCell}}
                $dllMoveP95=Get-P95 $dllTimings.ToArray();$dllGdiAfter=[NxFocusP0Api]::GetGuiResources($dllProcess.Handle,0);$dllUserAfter=[NxFocusP0Api]::GetGuiResources($dllProcess.Handle,1);$dllProcess.Refresh();$dllMemoryAfter=$dllProcess.PrivateMemorySize64;$dllProcess.Dispose();$dllGdiDelta=$dllGdiAfter-$dllGdiBefore;$dllUserDelta=$dllUserAfter-$dllUserBefore;$dllPrivateMemoryDelta=$dllMemoryAfter-$dllMemoryBefore;$dllWindowsAfter=[NxFocusP0Api]::CountWindows([int]$binding.pid,'NX_FOCUS_DLL_')
                $dllStressPass=($dllMoveP95-le50-and$dllGdiDelta-le8-and$dllUserDelta-le8-and$dllWindowsBefore-eq8-and$dllWindowsAfter-eq8);Add-Case 'dll_five_hundred_moves_p95' $dllStressPass @{dll_move_count=$DllMoveCount;dll_move_p95_ms=$dllMoveP95;gdi_delta=$dllGdiDelta;user_delta=$dllUserDelta;private_memory_delta=$dllPrivateMemoryDelta;windows_before=$dllWindowsBefore;windows_after=$dllWindowsAfter}
                Add-Case 'dll_vba_geometry_parity' ($backend-ceq'dll'-and$matrixPass-and$dllSettingsBefore-ceq$dllSettingsAfter) @{vba=$script:VbaMatrix;dll=$script:DllParity}
                Add-Case 'dll_vba_color_alpha_parity' $alphaPass @{vba_alpha=$script:VbaParity.alpha;dll_matrix=$dllMatrix}

                $comAddins=$excel.COMAddIns;$hostComAddin=$comAddins.Item('LH.NxHost.Connect');[void]$excel.Run($dllMacro+'NxFocusSuspendForOperation');$hostComAddin.Connect=$false;Start-Sleep -Milliseconds 150;[void]$excel.Run($dllMacro+'NxFocusResumeAfterOperation');Start-Sleep -Milliseconds 150;$fallbackBackend=[string]$excel.Run($dllMacro+'NxFocusControllerBackend');$fallbackDiagnostic=[string]$excel.Run($dllMacro+'NxFocusControllerBackendDiagnostic');$fallbackEdges=Invoke-EdgeCase $excel $sheet ([int]$binding.pid) $dllMacro 'dll_failure_visible_fallback' 'B2' 'B2' 'B2' 'NX_FOCUS_OVERLAY_' $false;Add-Case 'dll_failure_visible_fallback' ($fallbackBackend-ceq'vba-fallback'-and$fallbackDiagnostic.StartsWith('DLL_RESUME_FAILED|')-and$fallbackEdges.passed) @{backend=$fallbackBackend;diagnostic=$fallbackDiagnostic;edges=$fallbackEdges}
                [void]$excel.Run($dllMacro+'NxFocusDisable');Start-Sleep -Milliseconds 200;$dllResidue=[NxFocusP0Api]::CountWindows([int]$binding.pid,'NX_FOCUS_DLL_');$vbaFallbackResidue=[NxFocusP0Api]::CountWindows([int]$binding.pid,'NX_FOCUS_OVERLAY_');$dllTimerResidue=[int]$excel.Run($dllMacro+'NxFocusOverlayTimerCount')
            } finally {
                if($null-ne$secondWorkbook){try{$secondWorkbook.Close($false)}catch{}}
                if($null-ne$workbook){try{$workbook.Close($false)}catch{}}
                if($null-ne$addin){try{$addin.Close($false)}catch{}}
                Release-ComObject $hostComAddin;$hostComAddin=$null;Release-ComObject $comAddins;$comAddins=$null;Release-ComObject $excelWindow;$excelWindow=$null;Release-ComObject $secondSheet;$secondSheet=$null;Release-ComObject $secondWorkbook;$secondWorkbook=$null;Release-ComObject $sheet;$sheet=$null;Release-ComObject $worksheets;$worksheets=$null;Release-ComObject $workbook;$workbook=$null;Release-ComObject $addin;$addin=$null;Release-ComObject $books;$books=$null
                if($null-ne$binding){$dllCleanup=Close-OwnedExcel $excel $binding @() 'FocusP0 DLL Excel';$excel=$null;$binding=$null}
            }
            $dllStopPass=($dllResidue-eq0-and$vbaFallbackResidue-eq0-and$dllTimerResidue-eq0-and$null-ne$dllCleanup-and$dllCleanup.exit_mode-ceq'NATURAL');Add-Case 'dll_stop_zero_windows' $dllStopPass @{dll_windows=$dllResidue;vba_fallback_windows=$vbaFallbackResidue;vba_timer=$dllTimerResidue;dll_cleanup=$dllCleanup}
        }catch{$dllDiagnostic=$_.Exception.Message}finally{
            if($registrationAttempted){
                try{$unregisterOutput=@(& $registrationScript -Action Unregister -X86Dll $x86 -X64Dll $x64);if(-not($unregisterOutput|Where-Object{[string]$_-ceq'PASS|Unregister'})){throw 'NxHost unregister failed'};$registryAfterHash=Get-MarkerValue $unregisterOutput 'REGISTRY_SNAPSHOT_AFTER|';if([string]::IsNullOrEmpty($registryAfterHash)-or$registryBeforeHash-cne$registryAfterHash){throw 'registry snapshot mismatch after restore'};$registration='RESTORED'}catch{$registration='RESTORE_FAILED';if([string]::IsNullOrEmpty($dllDiagnostic)){$dllDiagnostic=$_.Exception.Message}else{$dllDiagnostic=$dllDiagnostic+' | '+$_.Exception.Message}}
            }
        }
        Add-Case 'registry_snapshot_restored' ($registration-ceq'RESTORED'-and-not[string]::IsNullOrEmpty($registryBeforeHash)-and$registryBeforeHash-ceq$registryAfterHash) @{before=$registryBeforeHash;after=$registryAfterHash;registration=$registration}
    }
}

$env:LHEXCEL_PROFILE_ROOT=$originalProfileOverride
$required=@('zoom_75_uia_edges','zoom_100_uia_edges','zoom_125_uia_edges','single_cell_uia_edges','multi_cell_uia_edges','merged_cell_uia_edges','scroll_uia_edges','smooth_scroll_drag_vba_uia_edges','freeze_panes_uia_edges','split_window_uia_edges','hidden_filtered_protected','multiple_workbooks_windows','settings_normal_restart','settings_corrupt_recovery','settings_existing_backup','settings_lock_conflict','settings_write_failure','five_hundred_moves_p95','gdi_handle_delta','user_state_preserved','residual_overlay_windows','residual_timers','dll_vba_geometry_parity','dll_vba_color_alpha_parity','dll_five_hundred_moves_p95','smooth_scroll_drag_uia_edges','dll_failure_visible_fallback','dll_stop_zero_windows','registry_snapshot_restored')
foreach($name in $required){if(@($script:Cases|Where-Object{$_.name-eq$name}).Count-eq0){Add-Case $name $false @{not_reached=$true;diagnostic=$diagnostic;dll_diagnostic=$dllDiagnostic}}}
$ordered=@($script:Cases|Sort-Object{[Array]::IndexOf($required,[string]$_.name)});for($i=0;$i-lt$ordered.Count;$i++){$ordered[$i].ordinal=$i+1};$passed=@($ordered|Where-Object{$_.status-eq'PASS'}).Count;$vbaExit=if($null-ne$vbaCleanup){[string]$vbaCleanup.exit_mode}else{'NOT_STARTED'};$dllExit=if($null-ne$dllCleanup){[string]$dllCleanup.exit_mode}else{'NOT_STARTED'};$allPass=($ordered.Count-eq$required.Count-and$passed-eq$required.Count-and[string]::IsNullOrEmpty($diagnostic)-and[string]::IsNullOrEmpty($dllDiagnostic)-and$registration-eq'RESTORED'-and$vbaExit-eq'NATURAL'-and$dllExit-eq'NATURAL')
$receipt=[ordered]@{schema_version=1;suite='FocusP0';status=if($allPass){'PASS'}else{'DIAGNOSTIC'};run_id=$RunId;artifact_sha256=Get-Sha256 $artifact;windows_scale_percent=200;pixel_tolerance=$PixelTolerance;cases=$ordered;measurements=@{move_count=$MoveCount;refresh_p95_ms=$refreshP95;gdi_handle_delta=$gdiDelta;user_handle_delta=$userDelta;dll_move_count=$DllMoveCount;dll_move_p95_ms=$dllMoveP95;dll_gdi_handle_delta=$dllGdiDelta;dll_user_handle_delta=$dllUserDelta;dll_private_memory_delta=$dllPrivateMemoryDelta};vba_cleanup=$vbaCleanup;dll_cleanup=$dllCleanup;registration=$registration;registry_snapshot_before=$registryBeforeHash;registry_snapshot_after=$registryAfterHash;diagnostic=$diagnostic;dll_diagnostic=$dllDiagnostic}
$receiptPath=Join-Path $EvidenceRoot 'FocusP0.json';Write-Json $receiptPath $receipt
if($allPass){Write-Output ('PASS|FocusP0|'+$passed+'/'+$required.Count);Write-Output $receiptPath;exit 0}
Write-Output ('DIAGNOSTIC|FocusP0|'+$passed+'/'+$required.Count);Write-Output $receiptPath;exit 24
