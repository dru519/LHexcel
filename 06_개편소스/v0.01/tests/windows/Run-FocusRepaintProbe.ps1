param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

if (-not ([Runtime.InteropServices.Marshal].GetMethod('GetActiveObject'))) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class NxFocusRot {
  [DllImport("ole32.dll", CharSet=CharSet.Unicode, PreserveSig=false)] static extern void CLSIDFromProgID(string progId, out Guid clsid);
  [DllImport("oleaut32.dll", PreserveSig=false)] static extern void GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object value);
  public static object GetExcel() { Guid id; CLSIDFromProgID("Excel.Application",out id); object value; GetActiveObject(ref id,IntPtr.Zero,out value); return value; }
}
'@
    function Start-ExactInteractiveExcel([object]$Baseline,[string]$Label,[int]$TimeoutMs=30000) {
        $excelPath=Get-ExactRegisteredExcelExecutable;$launch=$null;$candidate=$null;$candidateBinding=$null;$transferred=$false
        try {
            $launch=Start-Process -FilePath $excelPath -ArgumentList @('/x') -WindowStyle Hidden -PassThru
            $deadline=[DateTime]::UtcNow.AddMilliseconds($TimeoutMs)
            while([DateTime]::UtcNow-lt$deadline){
                try{$candidate=[NxFocusRot]::GetExcel();$candidateBinding=Get-ExactExcelProcessOwnership $candidate $Baseline $Label;if([bool]$candidateBinding.owned-and[int]$candidateBinding.pid-eq$launch.Id-and[bool]$candidate.UserControl){$result=[pscustomobject]@{excel=$candidate;binding=$candidateBinding};$candidate=$null;$candidateBinding=$null;$transferred=$true;return $result}}
                catch{} finally {if(-not$transferred){if($null-ne$candidateBinding-and$null-ne$candidateBinding.process){$candidateBinding.process.Dispose()};if($null-ne$candidate-and[Runtime.InteropServices.Marshal]::IsComObject($candidate)){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($candidate)}catch{}};$candidateBinding=$null;$candidate=$null}}
                Start-Sleep -Milliseconds 100
            }
            throw ($Label+' interactive COM binding timed out')
        } finally {
            $cleanupFailure=$null;if(-not$transferred-and$null-ne$launch){try{if(-not$launch.HasExited){$cleanupFailure=$Label+' unbound Excel preserved for review; pid='+$launch.Id}}catch{$cleanupFailure=$_.Exception.Message}}
            if($null-ne$launch){$launch.Dispose()};if(-not[string]::IsNullOrWhiteSpace($cleanupFailure)){throw $cleanupFailure}
        }
    }
}

if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing evidence preserved' }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Close user Excel before probe' }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$localRoot = Join-Path ([IO.Path]::GetTempPath()) ('focus-repaint-' + [guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $localRoot)
$localXlam = Join-Path $localRoot 'Product.xlam'
Copy-Item -LiteralPath $ProductXlam -Destination $localXlam
$productHash = (Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash

Add-Type -AssemblyName System.Drawing.Common
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class NxFocusScreen {
  [StructLayout(LayoutKind.Sequential)] public struct Rect { public int L,T,R,B; }
  [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
  [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr value);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd,System.Text.StringBuilder text,int count);
  [DllImport("user32.dll")] static extern IntPtr FindWindowEx(IntPtr parent,IntPtr after,string cls,string caption);
  [DllImport("user32.dll")] static extern bool RedrawWindow(IntPtr hwnd,IntPtr rect,IntPtr region,uint flags);
  [DllImport("user32.dll")] static extern bool InvalidateRect(IntPtr hwnd,IntPtr rect,bool erase);
  [DllImport("user32.dll")] static extern bool UpdateWindow(IntPtr hwnd);
  public static Rect Bounds(long hwnd) { Rect r; if(!GetWindowRect(new IntPtr(hwnd),out r)) throw new InvalidOperationException("Excel rectangle unavailable"); return r; }
  public static IntPtr BeginDpi() { return SetThreadDpiAwarenessContext(new IntPtr(-4)); }
  public static void EndDpi(IntPtr old) { SetThreadDpiAwarenessContext(old); }
  static IntPtr Grid(IntPtr root) { IntPtr child=IntPtr.Zero; while((child=FindWindowEx(root,child,null,null))!=IntPtr.Zero) { var s=new System.Text.StringBuilder(64);GetClassName(child,s,64);if(s.ToString().Equals("EXCEL7",StringComparison.OrdinalIgnoreCase))return child;IntPtr nested=Grid(child);if(nested!=IntPtr.Zero)return nested;}return IntPtr.Zero; }
  public static void Refresh(long hwnd,string method) { IntPtr grid=Grid(new IntPtr(hwnd));if(grid==IntPtr.Zero)throw new InvalidOperationException("EXCEL7 grid not found");if(method=="redraw"){if(!RedrawWindow(grid,IntPtr.Zero,IntPtr.Zero,0x181))throw new InvalidOperationException("RedrawWindow failed");}else{if(!InvalidateRect(grid,IntPtr.Zero,false)||!UpdateWindow(grid))throw new InvalidOperationException("Invalidate/Update failed");} }
}
'@

function Capture-Window([long]$hwnd,[string]$path) {
    $old=[NxFocusScreen]::BeginDpi();$bitmap=$null;$graphics=$null
    try {$b=[NxFocusScreen]::Bounds($hwnd);$bitmap=[Drawing.Bitmap]::new($b.R-$b.L,$b.B-$b.T);$graphics=[Drawing.Graphics]::FromImage($bitmap);$graphics.CopyFromScreen($b.L,$b.T,0,0,$bitmap.Size);$bitmap.Save($path,[Drawing.Imaging.ImageFormat]::Png)}
    finally {if($null-ne$graphics){$graphics.Dispose()};if($null-ne$bitmap){$bitmap.Dispose()};[NxFocusScreen]::EndDpi($old)}
}
function Sample-Pixel([string]$path,[int]$absoluteX,[int]$absoluteY,[long]$hwnd) {
    $b=[NxFocusScreen]::Bounds($hwnd);$bitmap=[Drawing.Bitmap]::new($path)
    try {$x=$absoluteX-$b.L;$y=$absoluteY-$b.T;if($x-lt2-or$y-lt2-or$x-ge$bitmap.Width-2-or$y-ge$bitmap.Height-2){throw 'Sample outside capture'};$rs=@();$gs=@();$bs=@();for($dy=-2;$dy-le2;$dy++){for($dx=-2;$dx-le2;$dx++){$c=$bitmap.GetPixel($x+$dx,$y+$dy);$rs+=$c.R;$gs+=$c.G;$bs+=$c.B}};$rs=@($rs|Sort-Object);$gs=@($gs|Sort-Object);$bs=@($bs|Sort-Object);return @([int]$rs[12],[int]$gs[12],[int]$bs[12])}
    finally {$bitmap.Dispose()}
}

function Release-Com([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch { }
    }
}
function Stats([double[]]$values) {
    $sorted = @($values | Sort-Object)
    [ordered]@{ count=$sorted.Count; median_ms=[math]::Round($sorted[[int][math]::Floor($sorted.Count/2)],3); p95_ms=[math]::Round($sorted[[int][math]::Ceiling($sorted.Count*.95)-1],3) }
}
function Select-Range([string]$address) {
    $r=$sheet.Range($address)
    try {
        [void]$r.Select()
    } finally { Release-Com $r }
}
function Invoke-Refresh([string]$method) {
    if([string]::IsNullOrEmpty($method)){return}
    switch($method){
        'calculate' {$sheet.Calculate()}
        'redraw' {[NxFocusScreen]::Refresh([long]$excel.Hwnd,'redraw')}
        'invalidate_update' {[NxFocusScreen]::Refresh([long]$excel.Hwnd,'invalidate_update')}
        'zoom_self' {$z=$window.Zoom;$window.Zoom=$z}
        'scroll_self' {$sr=$window.ScrollRow;$sc=$window.ScrollColumn;$window.ScrollRow=$sr;$window.ScrollColumn=$sc}
        'smallscroll' {$window.SmallScroll(1,0,0,0);$window.SmallScroll(0,1,0,0)}
    }
}
function Cell-Point([string]$address) {
    $r=$sheet.Range($address)
    try {
        $originX=[double]$window.PointsToScreenPixelsX(0);$originY=[double]$window.PointsToScreenPixelsY(0)
        $scaleX=([double]$window.PointsToScreenPixelsX(72)-$originX)/72.0
        $scaleY=([double]$window.PointsToScreenPixelsY(72)-$originY)/72.0
        return @([int][math]::Round($originX+(([double]$r.Left+([double]$r.Width*.72))*$scaleX)),[int][math]::Round($originY+(([double]$r.Top+([double]$r.Height*.72))*$scaleY)))
    } finally { Release-Com $r }
}
function Pixel-Distance([int[]]$a,[int[]]$b) { [math]::Sqrt([math]::Pow($a[0]-$b[0],2)+[math]::Pow($a[1]-$b[1],2)+[math]::Pow($a[2]-$b[2],2)) }
function Capture-State([string]$method,[string]$state,[hashtable]$expected) {
    Start-Sleep -Milliseconds 300
    $path=Join-Path $EvidenceRoot ($method+'-'+$state+'.png')
    Capture-Window ([long]$excel.Hwnd) $path
    $samples=[ordered]@{}
    foreach($address in $expected.Keys){$p=Cell-Point $address;$samples[$address]=Sample-Pixel $path $p[0] $p[1] ([long]$excel.Hwnd)}
    return [ordered]@{ screenshot=$path; samples=$samples; expected=$expected }
}

$methods=@('calculate','color_self','stop_false')
$vbaApi=@'
#If VBA7 Then
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal parent As LongPtr, ByVal after As LongPtr, ByVal className As String, ByVal caption As String) As LongPtr
Private Declare PtrSafe Function RedrawWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal rect As LongPtr, ByVal region As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hwnd As LongPtr, ByVal rect As LongPtr, ByVal erase As Long) As Long
Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal parent As Long, ByVal after As Long, ByVal className As String, ByVal caption As String) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, ByVal rect As Long, ByVal region As Long, ByVal flags As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal rect As Long, ByVal erase As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If
'@
$vbaHelper=@'
Private Sub NxProbeGridRefresh(ByVal invalidateOnly As Boolean)
#If VBA7 Then
    Dim desk As LongPtr, grid As LongPtr
#Else
    Dim desk As Long, grid As Long
#End If
    desk = FindWindowEx(Application.hwnd, 0, "XLDESK", vbNullString)
    grid = FindWindowEx(desk, 0, "EXCEL7", vbNullString)
    If grid = 0 Then Err.Raise 5, , "EXCEL7 grid not found"
    If invalidateOnly Then
        If InvalidateRect(grid, 0, 0) = 0 Then Err.Raise 5, , "InvalidateRect failed"
        If UpdateWindow(grid) = 0 Then Err.Raise 5, , "UpdateWindow failed"
    Else
        If RedrawWindow(grid, 0, 0, &H181) = 0 Then Err.Raise 5, , "RedrawWindow failed"
    End If
End Sub

'@
$cfHelper=@'
Private Sub NxProbeTouchRules(ByVal sheet As Worksheet, ByVal selection As CNxFocusSelection, ByVal mode As String)
    Dim rule As Object
    Dim formula As String
    Dim selectedArea As String
    For Each rule In sheet.Cells.FormatConditions
        formula = ExpressionFormula(rule)
        If IsManagedFormula(formula) Then
            Select Case mode
                Case "color_self": rule.Interior.Color = rule.Interior.Color
                Case "stop_false": rule.StopIfTrue = False
                Case "modify_same": rule.Modify Type:=xlExpression, Formula1:=formula
                Case "same_applies": rule.ModifyAppliesToRange rule.AppliesTo
                Case "literal_formula"
                    selectedArea = "AND(ROW()>=" & CStr(selection.TopRow) & ",ROW()<=" & CStr(selection.BottomRow) & _
                        ",COLUMN()>=" & CStr(selection.LeftColumn) & ",COLUMN()<=" & CStr(selection.RightColumn) & ")"
                    If InStr(1, formula, NX_FOCUS_RULE_V002_ROW, vbTextCompare) > 0 Then
                        If mSettings.Shape = NX_FOCUS_SHAPE_VERTICAL Then
                            formula = "=AND(N(""" & NX_FOCUS_RULE_V002_ROW & """)=0,FALSE"
                        Else
                            formula = "=AND(N(""" & NX_FOCUS_RULE_V002_ROW & """)=0,AND(ROW()>=" & CStr(selection.TopRow) & ",ROW()<=" & CStr(selection.BottomRow) & ")"
                        End If
                    Else
                        If mSettings.Shape = NX_FOCUS_SHAPE_HORIZONTAL Then
                            formula = "=AND(N(""" & NX_FOCUS_RULE_V002_COLUMN & """)=0,FALSE"
                        Else
                            formula = "=AND(N(""" & NX_FOCUS_RULE_V002_COLUMN & """)=0,AND(COLUMN()>=" & CStr(selection.LeftColumn) & ",COLUMN()<=" & CStr(selection.RightColumn) & ")"
                        End If
                    End If
                    If Not mSettings.HighlightSelectedArea Then formula = formula & ",NOT(" & selectedArea & ")"
                    formula = formula & ")"
                    rule.Modify Type:=xlExpression, Formula1:=formula
            End Select
        End If
    Next rule
End Sub

'@
$excel=$null;$books=$null;$addin=$null;$book=$null;$sheet=$null;$window=$null;$binding=$null;$cleanup=$null
$project=$null;$components=$null;$component=$null;$classModule=$null;$userRange=$null;$macro=''
$failure='';$stage='start';$results=New-Object 'Collections.Generic.List[object]';$captures=New-Object 'Collections.Generic.List[object]'
$previousProfile=$env:LHEXCEL_PROFILE_ROOT;$env:LHEXCEL_PROFILE_ROOT=Join-Path $localRoot 'profile'
$script:currentMethod=''
try {
    $baseline=Get-ExcelProcessBaseline
    $interactive=Start-ExactInteractiveExcel $baseline 'focus repaint Excel'
    $excel=$interactive.excel;$binding=$interactive.binding;$interactive=$null
    if(-not$binding.owned){throw 'Excel ownership not established'}
    $excel.Visible=$true;$excel.DisplayAlerts=$false
    $books=$excel.Workbooks;$book=$books.Add();$addin=$books.Open($localXlam,$false,$true)
    $macro="'"+$addin.Name.Replace("'","''")+"'!"
    $project=$addin.VBProject;$components=$project.VBComponents
    $classPath=Join-Path $SourceRoot 'src/vba/features/data/focus/CNxFocusRuleEngine.cls'
    $baseCode=Get-Content -LiteralPath $classPath -Raw -Encoding UTF8
    $baseCode=$baseCode.Substring($baseCode.IndexOf('Option Explicit'))
    $component=$components.Item('CNxFocusRuleEngine');$classModule=$component.CodeModule
    $sheet=$book.Worksheets.Item(1);$sheet.Name='FocusRepaint';$book.Activate();$sheet.Activate();$window=$excel.ActiveWindow;$window.Zoom=100
    $excel.WindowState=-4137
    $userRange=$sheet.Range('K1:K3');$userRule=$userRange.FormatConditions.Add(1,3,'=0');$userRule.Interior.Color=65535;Release-Com $userRule
    $sheet.Columns('A:K').ColumnWidth=11;$sheet.Rows('1:20').RowHeight=24
    foreach($method in $methods){
        $stage='method_'+$method
        $script:currentMethod=$method
        [void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)
        $code=$baseCode
        switch($method){
            'color_self' {$code=$code.Replace('Private Sub InvalidateRefresh()',$cfHelper+'Private Sub InvalidateRefresh()').Replace('    sheet.Calculate','    NxProbeTouchRules sheet, selection, "color_self"')}
            'stop_false' {$code=$code.Replace('Private Sub InvalidateRefresh()',$cfHelper+'Private Sub InvalidateRefresh()').Replace('    sheet.Calculate','    NxProbeTouchRules sheet, selection, "stop_false"')}
            'modify_same' {$code=$code.Replace('Private Sub InvalidateRefresh()',$cfHelper+'Private Sub InvalidateRefresh()').Replace('    sheet.Calculate','    NxProbeTouchRules sheet, selection, "modify_same"')}
            'literal_formula' {$code=$code.Replace('Private Sub InvalidateRefresh()',$cfHelper+'Private Sub InvalidateRefresh()').Replace('    sheet.Calculate','    NxProbeTouchRules sheet, selection, "literal_formula"')}
            'same_applies' {$code=$code.Replace('Private Sub InvalidateRefresh()',$cfHelper+'Private Sub InvalidateRefresh()').Replace('    sheet.Calculate','    NxProbeTouchRules sheet, selection, "same_applies"')}
        }
        $classModule.DeleteLines(1,$classModule.CountOfLines);$classModule.AddFromString($code)
        Select-Range 'A1';[void]$excel.Run($macro+'NxFocusControllerSetEnabled',$true)
        $screenStates=New-Object 'Collections.Generic.List[object]'
        $screenStates.Add((Capture-State $method 'A1' @{B1='highlight';A3='highlight';E3='plain'}))
        Select-Range 'E8';$screenStates.Add((Capture-State $method 'E8' @{C8='highlight';E3='highlight';B1='plain';A3='plain'}))
        Select-Range 'I13';$screenStates.Add((Capture-State $method 'I13' @{C13='highlight';I8='highlight';C8='plain';E3='plain'}))
        Select-Range 'B3:D5';$screenStates.Add((Capture-State $method 'B3-D5' @{F4='highlight';C8='highlight';C13='highlight';I8='plain'}))
        $allScreen=$true;$screenChecks=New-Object 'Collections.Generic.List[object]'
        foreach($state in $screenStates){
            foreach($address in $state.expected.Keys){
                $rgb=[int[]]$state.samples[$address];$distance=Pixel-Distance $rgb @(255,255,255)
                $want=[string]$state.expected[$address];$ok=if($want-eq'highlight'){$distance-ge18}else{$distance-lt18}
                if(-not$ok){$allScreen=$false}
                $screenChecks.Add([ordered]@{state=[IO.Path]::GetFileNameWithoutExtension($state.screenshot);cell=$address;expected=$want;rgb=$rgb;white_distance=[math]::Round($distance,2);pass=$ok})
            }
        }
        $timings=New-Object 'Collections.Generic.List[object]'
        foreach($scenario in @('empty','manual_90000','automatic_90000')){
            [void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)
            $sheet.Cells.ClearContents()
            $excel.Calculation=-4135
            if($scenario-ne'empty'){$grid=$sheet.Range('A1:AX1800');$grid.Formula='=ROW()+COLUMN()+NOW()*0';Release-Com $grid}
            if($scenario-eq'automatic_90000'){$excel.Calculation=-4105}
            $sheet.Calculate();Select-Range 'B2:C3';[void]$excel.Run($macro+'NxFocusControllerSetEnabled',$true)
            $moves=New-Object 'Collections.Generic.List[double]';$addresses=@('B2:C3','F6:G7','D4:E5','H8:I9','C12:D13','K15:L16')
            for($i=0;$i-lt4;$i++){$r=$sheet.Range($addresses[$i%$addresses.Count]);$watch=[Diagnostics.Stopwatch]::StartNew();[void]$r.Select();$watch.Stop();Release-Com $r;if($i-ge1){$moves.Add($watch.Elapsed.TotalMilliseconds)}}
            $timings.Add([ordered]@{scenario=$scenario;latency=(Stats $moves.ToArray())})
        }
        [void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)
        $sheet.Cells.ClearContents();$excel.Calculation=-4135
        $rand=$sheet.Range('Z1');$rand.Formula='=RAND()';$rand.Calculate()
        Select-Range 'A1';[void]$excel.Run($macro+'NxFocusControllerSetEnabled',$true);[void]$excel.Run($macro+'NxFocusControllerRefreshActiveSelection',$false)
        $book.Saved=$true;$randBefore=[double]$rand.Value2;Start-Sleep -Milliseconds 1100;Select-Range 'E8';Select-Range 'I13';Select-Range 'B3:D5'
        $managed=[int]$excel.Run($macro+'NxFocusManagedRuleCount',$sheet);$userCount=[int]$userRange.FormatConditions.Count
        $undoCell=$sheet.Range('AA1');$undoCell.Value2='UNDO_PROBE';$undoCell.Select();$excel.CommandBars.ExecuteMso('ClearContents');$undoBefore=[bool]$excel.CommandBars.GetEnabledMso('Undo');Select-Range 'E8';$undoAfter=[bool]$excel.CommandBars.GetEnabledMso('Undo')
        $side=[ordered]@{rand_unchanged=([double]$rand.Value2-eq$randBefore);saved_true=[bool]$book.Saved;user_cf_count=$userCount;managed_cf_count=$managed;user_cf_intact=(($userCount-$managed)-eq1);undo_created=$undoBefore;undo_after_move=$undoAfter;calculation_manual=([int]$excel.Calculation-eq-4135)}
        Release-Com $undoCell
        Release-Com $rand
        $results.Add([ordered]@{method=$method;screen_pass=$allScreen;screen_checks=$screenChecks.ToArray();timings=$timings.ToArray();side_effects=$side})
    }
    $stage='complete'
} catch { $failure=$_.Exception.ToString() } finally {
    if($null-ne$excel){try{[void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)}catch{}}
    if($null-ne$addin){try{$addin.Close($false)}catch{}}
    if($null-ne$book){try{$book.Close($false)}catch{}}
    Release-Com $classModule;Release-Com $component;Release-Com $components;Release-Com $project;Release-Com $userRange;Release-Com $window;Release-Com $sheet;Release-Com $addin;Release-Com $book;Release-Com $books
    if($null-ne$excel){try{$excel.Quit()}catch{};Release-Com $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null-ne$binding){$cleanup=Stop-ExactProcessAfterGrace $binding.process 'focus repaint Excel' 15000 -Detailed;$binding.process.Dispose()}
    $env:LHEXCEL_PROFILE_ROOT=$previousProfile
}
$remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
for($i=0;$i-lt20-and$remaining.Count-gt0;$i++){Start-Sleep -Milliseconds 100;$remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue)}
$afterHash=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash
$passed=([string]::IsNullOrEmpty($failure)-and$null-ne$cleanup-and$cleanup.exit_mode-eq'NATURAL'-and$productHash-eq$afterHash-and$remaining.Count-eq0)
$receipt=[ordered]@{status=$(if($passed){'PASS'}else{'FAIL'});stage=$stage;failure=$failure;product_sha256=$productHash;product_sha256_after=$afterHash;excel_version=$(if($null-ne$binding){$binding.executable}else{$null});results=$results.ToArray();cleanup=$cleanup;remaining_pids=@($remaining|ForEach-Object{$_.Id});local_copy=$localXlam}
$receipt|ConvertTo-Json -Depth 14|Set-Content -LiteralPath (Join-Path $EvidenceRoot 'focus-repaint.json') -Encoding UTF8
try{Remove-Item -LiteralPath $localRoot -Recurse -Force}catch{}
Write-Output ($receipt|ConvertTo-Json -Depth 6)
if(-not$passed){exit 1}
