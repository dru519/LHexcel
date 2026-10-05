param([Parameter(Mandatory=$true)]$Excel,[Parameter(Mandatory=$true)]$Sheet,[string]$Prefix,[string]$EvidenceRoot,[switch]$DocumentOnly,[switch]$ServicesOnly)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes,WindowsBase,System.Drawing,Accessibility
Add-Type -ReferencedAssemblies UIAutomationClient,UIAutomationTypes,WindowsBase,System.Drawing,Accessibility -TypeDefinition @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Automation;
public static class R65Ui {
    public const string PaneName="\uB0B4\uC5D1\uC140 \uD0D0\uC0C9\uCC3D";
    public const string Books="\uBB38\uC11C \uBAA9\uB85D", Sheets="\uC2DC\uD2B8 \uBAA9\uB85D";
    public const string BookSearch="\uBB38\uC11C \uCC3E\uAE30", SheetSearch="\uC2DC\uD2B8 \uCC3E\uAE30";
    public const string Collapse="\uD0D0\uC0C9\uCC3D \uC811\uAE30", Expand="\uD0D0\uC0C9\uCC3D \uD3BC\uCE58\uAE30";
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd,uint msg,IntPtr wp,IntPtr lp);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd,uint msg,IntPtr wp,IntPtr lp);
    [DllImport("user32.dll")] public static extern bool IsChild(IntPtr parent,IntPtr child);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] private static extern bool GetCursorPos(out POINT p);
    [DllImport("user32.dll")] private static extern IntPtr WindowFromPoint(POINT p);
    [DllImport("user32.dll")] private static extern void mouse_event(uint flags,uint x,uint y,uint data,UIntPtr extra);
    [DllImport("user32.dll")] private static extern void keybd_event(byte key,byte scan,uint flags,UIntPtr extra);
    [StructLayout(LayoutKind.Sequential)] private struct POINT {public int X,Y;}
    public static uint OwnerPid(long hwnd) {uint pid;GetWindowThreadProcessId(new IntPtr(hwnd),out pid);return pid;}
    [DllImport("user32.dll",CharSet=CharSet.Unicode,EntryPoint="SendMessageW")] private static extern IntPtr SetText(IntPtr hwnd,uint msg,IntPtr wp,string text);
    [DllImport("user32.dll",CharSet=CharSet.Unicode,EntryPoint="SendMessageW")] private static extern IntPtr ReadText(IntPtr hwnd,uint msg,IntPtr wp,System.Text.StringBuilder text);
    private delegate bool EnumProc(IntPtr hwnd,IntPtr value);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent,EnumProc callback,IntPtr value);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd,System.Text.StringBuilder text,int size);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd,System.Text.StringBuilder text,int size);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr hwnd,uint id,ref Guid iid,[MarshalAs(UnmanagedType.Interface)] out object value);
    public static string NativeDump(long owner) {
        var lines=new System.Text.StringBuilder();
        EnumChildWindows(new IntPtr(owner),(hwnd,p)=>{
            var cls=new System.Text.StringBuilder(256);var text=new System.Text.StringBuilder(1024);
            GetClassName(hwnd,cls,256);GetWindowText(hwnd,text,1024);
            if(cls.ToString().StartsWith("WindowsForms")) {
                Guid iid=new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");object value;
                int hr=AccessibleObjectFromWindow(hwnd,0xFFFFFFFC,ref iid,out value);
                string detail="";
                if(value!=null) {
                    try {var accessible=(Accessibility.IAccessible)value; detail="role="+accessible.get_accRole(0)+"|name="+accessible.get_accName(0);}
                    catch(Exception error) {detail=error.HResult.ToString("X8");}
                    finally {if(Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}
                }
                lines.AppendLine(hwnd+"|"+cls+"|"+text+"|acc="+hr.ToString("X8")+"|"+detail);
            }
            return true;
        },IntPtr.Zero);return lines.ToString();
    }
    public static AutomationElement Root(long hwnd) {return AutomationElement.FromHandle(new IntPtr(hwnd));}
    public static AutomationElement Find(AutomationElement root,string name) {return root.FindFirst(TreeScope.Descendants,new PropertyCondition(AutomationElement.NameProperty,name));}
    public static AutomationElement Pane(AutomationElement root) {
        foreach(AutomationElement item in root.FindAll(TreeScope.Descendants,new PropertyCondition(AutomationElement.NameProperty,PaneName)))
            if(item.Current.BoundingRectangle.Height>100) return item;
        return null;
    }
    public static double PaneWidth(long owner) {
        var pane=Pane(Root(owner));
        if(pane==null)throw new Exception("Owned document pane unavailable");
        return pane.Current.BoundingRectangle.Width;
    }
    // Exercise the native MSAA contract even where this host's UIA WinForms proxy
    // exposes generic Pane nodes. Do not infer controls from screen coordinates.
    private static Accessibility.IAccessible Accessible(IntPtr hwnd) {
        Guid iid=new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");object value;
        return AccessibleObjectFromWindow(hwnd,0xFFFFFFFC,ref iid,out value)==0 ? value as Accessibility.IAccessible : null;
    }
    public static AutomationElement NativeControl(long owner,string name) {
        IntPtr result=IntPtr.Zero;
        EnumChildWindows(new IntPtr(owner),(hwnd,p)=>{
            if(!IsWindowVisible(hwnd)) return true;
            var accessible=Accessible(hwnd);
            try {if(accessible!=null && accessible.get_accName(0)==name){result=hwnd;return false;}}
            finally {if(accessible!=null && Marshal.IsComObject(accessible))Marshal.ReleaseComObject(accessible);}
            return true;
        },IntPtr.Zero);
        return result==IntPtr.Zero ? null : AutomationElement.FromHandle(result);
    }
    public static void InvokeNamed(long owner,string name) {
        bool invoked=false;
        EnumChildWindows(new IntPtr(owner),(hwnd,p)=>{
            if(!IsWindowVisible(hwnd)) return true;
            var accessible=Accessible(hwnd);
            try {
                if(accessible==null)return true;
                if(accessible.get_accName(0)==name){accessible.accDoDefaultAction(0);invoked=true;return false;}
                if(Convert.ToInt32(accessible.get_accRole(0))!=22)return true;
                for(int child=1;child<=accessible.accChildCount;child++) {
                    object value=null;
                    try {value=accessible.get_accChild(child);}catch(ArgumentException){}
                    var nested=value as Accessibility.IAccessible;
                    try {
                        if(nested!=null && nested.get_accName(0)==name){nested.accDoDefaultAction(0);invoked=true;return false;}
                        if(nested==null && accessible.get_accName(child)==name){accessible.accDoDefaultAction(child);invoked=true;return false;}
                    } catch(ArgumentException) {} finally {if(value!=null && Marshal.IsComObject(value))Marshal.ReleaseComObject(value);}
                }
            } finally {if(accessible!=null && Marshal.IsComObject(accessible))Marshal.ReleaseComObject(accessible);}
            return true;
        },IntPtr.Zero);
        if(!invoked)throw new Exception("Named native action unavailable: "+name);
    }
    public static string Dump(AutomationElement root) {
        var text=new System.Text.StringBuilder();
        foreach(AutomationElement item in root.FindAll(TreeScope.Descendants,Condition.TrueCondition))
            text.AppendLine(item.Current.ControlType.ProgrammaticName+"|"+item.Current.Name+"|"+item.Current.BoundingRectangle);
        return text.ToString();
    }
    public static int Count(AutomationElement list) {return SendMessage(new IntPtr(list.Current.NativeWindowHandle),0x18B,IntPtr.Zero,IntPtr.Zero).ToInt32();}
    public static void Filter(AutomationElement box,string query) {
        if(SetText(new IntPtr(box.Current.NativeWindowHandle),0x000C,IntPtr.Zero,query)==IntPtr.Zero)throw new Exception("Native edit rejected");
    }
    public static string Value(AutomationElement box) {
        var text=new System.Text.StringBuilder(1024);ReadText(new IntPtr(box.Current.NativeWindowHandle),0x000D,new IntPtr(text.Capacity),text);return text.ToString();
    }
    public static void Key(AutomationElement list,long owner,int key) {
        IntPtr handle=new IntPtr(list.Current.NativeWindowHandle);
        if(!IsChild(new IntPtr(owner),handle)) throw new Exception("Foreign UI rejected");
        var accessible=Accessible(handle);
        try {if(accessible!=null) accessible.accSelect(1,0);}
        finally {if(accessible!=null && Marshal.IsComObject(accessible))Marshal.ReleaseComObject(accessible);}
        PostMessage(handle,0x100,new IntPtr(key),IntPtr.Zero);PostMessage(handle,0x101,new IntPtr(key),IntPtr.Zero);
    }
    public static void NativeClickAndEnter(AutomationElement list,long owner) {
        IntPtr handle=new IntPtr(list.Current.NativeWindowHandle), window=new IntPtr(owner);
        if(!IsChild(window,handle))throw new Exception("Foreign list rejected");
        POINT prior;GetCursorPos(out prior);
        try {
            SetForegroundWindow(window);System.Threading.Thread.Sleep(100);
            if(GetForegroundWindow()!=window)throw new Exception("Owned Excel is not foreground");
            var box=list.Current.BoundingRectangle;
            var point=new POINT {X=(int)box.Left+24,Y=(int)box.Top+15};
            if(WindowFromPoint(point)!=handle)throw new Exception("Owned list is occluded");
            SetCursorPos(point.X,point.Y);
            mouse_event(2,0,0,0,UIntPtr.Zero);mouse_event(4,0,0,0,UIntPtr.Zero);
            System.Threading.Thread.Sleep(100);
            keybd_event(0x0D,0,0,UIntPtr.Zero);keybd_event(0x0D,0,2,UIntPtr.Zero);
        } finally {SetCursorPos(prior.X,prior.Y);}
    }
    public static bool Wheel(AutomationElement list,long owner) {
        IntPtr handle=new IntPtr(list.Current.NativeWindowHandle),window=new IntPtr(owner);
        if(!IsChild(window,handle))throw new Exception("Foreign list rejected");
        POINT prior;GetCursorPos(out prior);
        try {
            SetForegroundWindow(window);System.Threading.Thread.Sleep(100);
            if(GetForegroundWindow()!=window)throw new Exception("Owned Excel is not foreground");
            var box=list.Current.BoundingRectangle;
            var point=new POINT{X=(int)box.Left+24,Y=(int)box.Top+15};
            if(WindowFromPoint(point)!=handle)throw new Exception("Owned list is occluded");
            SetCursorPos(point.X,point.Y);mouse_event(2,0,0,0,UIntPtr.Zero);mouse_event(4,0,0,0,UIntPtr.Zero);
            System.Threading.Thread.Sleep(100);
            long before=SendMessage(handle,0x18E,IntPtr.Zero,IntPtr.Zero).ToInt64();
            mouse_event(0x800,0,0,unchecked((uint)-360),UIntPtr.Zero);System.Threading.Thread.Sleep(150);
            return SendMessage(handle,0x18E,IntPtr.Zero,IntPtr.Zero).ToInt64()>before;
        } finally {SetCursorPos(prior.X,prior.Y);}
    }
    public static void Invoke(AutomationElement button) {((InvokePattern)button.GetCurrentPattern(InvokePattern.Pattern)).Invoke();}
    public static void Capture(AutomationElement pane,string path) {
        var r=pane.Current.BoundingRectangle;
        using(var bitmap=new Bitmap((int)r.Width,(int)r.Height)) {using(var g=Graphics.FromImage(bitmap))g.CopyFromScreen((int)r.Left,(int)r.Top,0,0,bitmap.Size);bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);}
    }
}
'@
$cases=New-Object 'Collections.Generic.List[object]'
function TraceStage([string]$Name) {
    [IO.File]::AppendAllText((Join-Path $EvidenceRoot 'r65-stages.log'),
        ((Get-Date).ToString('o')+'|'+$Name+[Environment]::NewLine),(New-Object Text.UTF8Encoding($false)))
}
function Check([string]$Name,[bool]$Ok,[object]$Details) {
    $cases.Add([pscustomobject]@{name=$Name;pass=$Ok;details=$Details})
    if(-not $Ok){throw ('R65 case failed: '+$Name)}
    Write-Output ('PASS|'+$Name)
}
function Probe([string]$Name) {
    TraceStage ('probe.begin.'+$Name)
    $answer=[string]$Excel.Run(($Prefix+'NxR65Probe'),$Name)
    TraceStage ('probe.end.'+$Name)
    Check $Name ($answer.StartsWith('PASS|'+$Name)) $answer
}
$watcher=$null
$cell=$null
try {
    if(-not $ServicesOnly) {
    if($DocumentOnly) {
        $fixture=$Sheet.Parent
        for($i=1;$i -le 80;$i++) {
            $extra=$fixture.Worksheets.Add([Type]::Missing,$fixture.Worksheets.Item($fixture.Worksheets.Count))
            $extra.Name=('Sheet {0:000}' -f $i)
            if($i -eq 2){$extra.Visible=0};if($i -eq 3){$extra.Visible=2}
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($extra)
        }
        $Sheet.Activate()
        Check 'document-prepare' $true 'COM fixture without VBA instrumentation'
    } else { Probe 'document-prepare' }
    [void]$Excel.Run(($Prefix+'NxHostHideNavigator'))
    Check 'document.show' ([bool]$Excel.Run(($Prefix+'NxHostShowDocumentNavigator'))) ''
    Start-Sleep -Milliseconds 350
    $owner=[long]$Excel.Hwnd; $root=[R65Ui]::Root($owner)
    Start-Sleep -Milliseconds 2000
    $pane=[R65Ui]::Pane($root)
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'document-ui.txt'),[R65Ui]::Dump($root),(New-Object Text.UTF8Encoding($false)))
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'document-native.txt'),[R65Ui]::NativeDump($owner),(New-Object Text.UTF8Encoding($false)))
    [R65Ui]::Capture($root,(Join-Path $EvidenceRoot 'document-full.png'))
    Check 'document.ctp' ($null -ne $pane) ''
    $books=[R65Ui]::NativeControl($owner,[R65Ui]::Books);$sheets=[R65Ui]::NativeControl($owner,[R65Ui]::Sheets)
    $search=[R65Ui]::NativeControl($owner,[R65Ui]::SheetSearch)
    Check 'document.two-lists-present' ($null -ne $books -and $null -ne $sheets) ''
    Check 'document.two-lists' ([R65Ui]::Count($sheets) -eq 80) ([R65Ui]::Count($sheets))
    $initial=[string]$Excel.ActiveSheet.Name
    [R65Ui]::Filter($search,'Sheet 002'); Start-Sleep -Milliseconds 150
    [R65Ui]::Key($sheets,$owner,0x24);[R65Ui]::Key($sheets,$owner,0x0D);Start-Sleep -Milliseconds 200
    Check 'document.hidden-blocked' ([string]$Excel.ActiveSheet.Name -eq $initial) ''
    [R65Ui]::Filter($search,'Sheet 003');Start-Sleep -Milliseconds 150
    Check 'document.very-hidden-excluded' ([R65Ui]::Count($sheets) -eq 0) ''
    [R65Ui]::Filter($search,'Sheet 020');Start-Sleep -Milliseconds 150
    Check 'document.search-exact' ([R65Ui]::Count($sheets) -eq 1) @{text=[R65Ui]::Value($search);count=[R65Ui]::Count($sheets)}
    [R65Ui]::Key($sheets,$owner,0x24);Start-Sleep -Milliseconds 150
    Check 'document.selection-not-activation' ([string]$Excel.ActiveSheet.Name -eq $initial) ''
    [R65Ui]::NativeClickAndEnter($sheets,$owner);Start-Sleep -Milliseconds 300
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'document-after-enter.txt'),[R65Ui]::NativeDump($owner),(New-Object Text.UTF8Encoding($false)))
    Check 'document.enter-moves' ([string]$Excel.ActiveSheet.Name -eq 'Sheet 020') ([string]$Excel.ActiveSheet.Name)
    $expanded=[R65Ui]::PaneWidth($owner)
    [R65Ui]::InvokeNamed($owner,[R65Ui]::Collapse);Start-Sleep -Milliseconds 200
    $collapsed=[R65Ui]::PaneWidth($owner)
    Check 'document.collapse' ($collapsed -lt $expanded) @{expanded=$expanded;collapsed=$collapsed}
    [R65Ui]::InvokeNamed($owner,[R65Ui]::Expand);Start-Sleep -Milliseconds 200
    Check 'document.restore-state' ([R65Ui]::Value($search) -eq 'Sheet 020') ''
    [R65Ui]::Filter($search,'')
    Check 'document.wheel-after-focus' ([R65Ui]::Wheel($sheets,$owner)) 'Foreground-guarded native click and wheel'
    [R65Ui]::Filter($search,'Sheet 020')
    [R65Ui]::InvokeNamed($owner,(-join ([char[]]@(0xACE0,0xC815))))
    [R65Ui]::Filter($search,'Sheet 021')
    Start-Sleep -Milliseconds 600
    Check 'document.unpinned-query' ([R65Ui]::Count($sheets) -eq 1) @{count=[R65Ui]::Count($sheets);text=[R65Ui]::Value($search);width=[R65Ui]::PaneWidth($owner)}
    [R65Ui]::NativeClickAndEnter($sheets,$owner);Start-Sleep -Milliseconds 250
    Check 'document.unpinned-move-collapses' ([string]$Excel.ActiveSheet.Name -eq 'Sheet 021' -and [R65Ui]::PaneWidth($owner) -lt $expanded) @{sheet=[string]$Excel.ActiveSheet.Name;width=[R65Ui]::PaneWidth($owner)}
    [R65Ui]::InvokeNamed($owner,[R65Ui]::Expand)
    Start-Sleep -Milliseconds 250
    [R65Ui]::InvokeNamed($owner,(-join ([char[]]@(0xACE0,0xC815))))
    [R65Ui]::Filter($search,'Sheet 020')
    $primary=$Sheet.Parent;$secondary=$null
    try {
        $secondary=$Excel.Workbooks.Add()
        Check 'document.second-window-show' ([bool]$Excel.Run(($Prefix+'NxHostShowDocumentNavigator'))) ''
        $secondOwner=[long]$Excel.Hwnd
        $secondSearch=[R65Ui]::NativeControl($secondOwner,[R65Ui]::SheetSearch)
        Check 'document.second-window-independent' ($secondOwner -ne $owner -and [R65Ui]::Value($secondSearch) -eq '') ''
        $primary.Activate()
        [void]$Excel.Run(($Prefix+'NxHostShowDocumentNavigator'))
        Check 'document.first-window-retains-query' ([R65Ui]::Value($search) -eq 'Sheet 020') ''
    } finally {
        if($null -ne $secondary){$secondary.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($secondary);$secondary=$null}
        $primary.Activate()
    }
    $pane=[R65Ui]::Pane([R65Ui]::Root($owner))
    [R65Ui]::Capture($pane,(Join-Path $EvidenceRoot 'document-pane.png'))
    if($DocumentOnly){ return }
    }
    [void]$Sheet.Activate(); $cell=$Sheet.Range('B2');$cell.Select()
    Probe 'compare-prepare'
    $profile=$env:LHEXCEL_PROFILE_ROOT
    $baseBefore=(Get-FileHash -LiteralPath (Join-Path $profile 'base.xlsx')).Hash
    $otherBefore=(Get-FileHash -LiteralPath (Join-Path $profile 'other.xlsx')).Hash
    Probe 'compare-values'; Probe 'compare-formats'; Probe 'compare-cancel'
    Check 'compare.original-files-preserved' ($baseBefore -eq (Get-FileHash -LiteralPath (Join-Path $profile 'base.xlsx')).Hash -and $otherBefore -eq (Get-FileHash -LiteralPath (Join-Path $profile 'other.xlsx')).Hash) ''
    # Keep the named worksheet lookup independent of PowerShell 5.1's ANSI default.
    $detailSheetName = -join @([char]0xC140,[char]0xCC28,[char]0xC774)
    foreach($mode in @('values','formats')) {
        $output=$Excel.Workbooks.Open((Join-Path $profile ('compare-'+$mode+'.xlsx')),0,$true)
        # Snapshot sheets precede the report; formatting may extend UsedRange.
        $detail=$output.Worksheets.Item($detailSheetName)
        $reportData=$detail.Range('A1:E4').Value2
        $validReport=($reportData[2,2] -eq 'B2' -and $reportData[3,2] -eq 'D4' -and
            [string]$reportData[2,4] -eq '21' -and [string]$reportData[2,5] -eq '42' -and
            [string]$reportData[3,4] -eq '63' -and [string]$reportData[3,5] -eq '3' -and
            [string]::IsNullOrEmpty([string]($reportData[4,1])))
        Check ('compare.report-'+$mode) $validReport 'Named detail sheet: B2 21->42; D4 63->3; no third difference'
        $output.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($detail);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($output)
    }
    Probe 'compare-large-prepare'
    $largeRoot=Join-Path $profile 'large'
    $largeBase=(Get-FileHash -LiteralPath (Join-Path $largeRoot 'base.xlsx')).Hash
    $largeOther=(Get-FileHash -LiteralPath (Join-Path $largeRoot 'other.xlsx')).Hash
    Probe 'compare-large'
    $output=$Excel.Workbooks.Open((Join-Path $largeRoot 'compare-large.xlsx'),0,$true)
    $detail=$output.Worksheets.Item($detailSheetName)
    $reportData=$detail.Range('A1:E4').Value2
    Check 'compare.200000-cells-result' ($reportData[2,2] -eq 'B2' -and $reportData[3,2] -eq 'D4' -and
        [string]$reportData[2,4] -eq '21' -and [string]$reportData[2,5] -eq '42' -and
        [string]$reportData[3,4] -eq '63' -and [string]$reportData[3,5] -eq '3' -and
        [string]::IsNullOrEmpty([string]($reportData[4,1]))) '200000 cells, exact two differences in named detail sheet'
    $output.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($detail);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($output)
    Check 'compare.large-originals-preserved' ($largeBase -eq (Get-FileHash -LiteralPath (Join-Path $largeRoot 'base.xlsx')).Hash -and $largeOther -eq (Get-FileHash -LiteralPath (Join-Path $largeRoot 'other.xlsx')).Hash) ''
    TraceStage 'picture.activate-sheet.begin'
    [void]$Sheet.Activate()
    TraceStage 'picture.activate-sheet.end'
    $cell.Select()
    TraceStage 'picture.select-cell.end'
    $image=New-Object Drawing.Bitmap 400,100
    $g=[Drawing.Graphics]::FromImage($image);$g.Clear([Drawing.Color]::Orange);$g.Dispose()
    $image.Save((Join-Path $profile 'preview.png'),[Drawing.Imaging.ImageFormat]::Png);$image.Dispose()
    TraceStage 'picture.fixture-image.saved'
    $excelPid=[R65Ui]::OwnerPid([long]$Excel.Hwnd)
    TraceStage 'picture.owner-pid.read'
    $helper=Join-Path $PSScriptRoot 'R65PictureWatcher.cs'
    $watcher=Start-Job -ArgumentList $excelPid,$EvidenceRoot,$helper -ScriptBlock {
        param($OwnedPid,$Evidence,$Helper)
        $ErrorActionPreference='Stop'
        Add-Type -AssemblyName System.Drawing
        Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition ([IO.File]::ReadAllText($Helper))
        [R65PictureWatcher]::Run($OwnedPid,(Join-Path $Evidence 'picture-preview.png'),25)
    }
    TraceStage 'picture.watcher.started'
    Probe 'picture'
    if($null -eq (Wait-Job $watcher -Timeout 30)) {throw 'Picture watcher timeout'}
    $observed=@(Receive-Job $watcher -Wait -AutoRemoveJob);$watcher=$null
    Check 'picture.native-dialog' (@($observed | Where-Object {$_ -like 'PASS|Win32*'}).Count -eq 1) $observed
} catch {
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'r65-failure.txt'),
        ($_.Exception.ToString()+[Environment]::NewLine+$_.ScriptStackTrace+[Environment]::NewLine+$_.InvocationInfo.PositionMessage),
        (New-Object Text.UTF8Encoding($false)))
    throw
} finally {
    if($null -ne $cell -and [Runtime.InteropServices.Marshal]::IsComObject($cell)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell);$cell=$null}
    if($null -ne $watcher){Stop-Job $watcher -ErrorAction SilentlyContinue;Remove-Job $watcher -Force -ErrorAction SilentlyContinue}
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'r65-enhancement.json'),($cases.ToArray()|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
}
