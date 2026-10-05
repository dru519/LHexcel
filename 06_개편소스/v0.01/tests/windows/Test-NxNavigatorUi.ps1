param([Parameter(Mandatory=$true)]$Excel,[Parameter(Mandatory=$true)]$Sheet,[string]$Prefix,[string]$EvidenceRoot,[switch]$ExpectFallback)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes,WindowsBase,System.Drawing
if (-not ('NxNavigatorUiProbe' -as [type])) {
Add-Type -ReferencedAssemblies UIAutomationClient,UIAutomationTypes,WindowsBase,System.Drawing -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using System.Drawing;
public static class NxNavigatorUiProbe {
    delegate bool EnumProc(IntPtr hwnd,IntPtr value);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback,IntPtr value);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr parent,EnumProc callback,IntPtr value);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr hwnd,System.Text.StringBuilder text,int capacity);
    [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr hwnd,uint objectId,ref Guid iid,[MarshalAs(UnmanagedType.Interface)] out object value);
    public static object NativeWindow(int processId) {
        object result=null;
        EnumWindows((top,unused)=>{
            uint pid;GetWindowThreadProcessId(top,out pid);if(pid!=processId)return true;
            EnumChildWindows(top,(child,state)=>{
                var name=new System.Text.StringBuilder(256);GetClassName(child,name,name.Capacity);
                if(name.ToString()!="EXCEL7")return true;
                Guid dispatch=new Guid("00020400-0000-0000-C000-000000000046");object value;
                if(AccessibleObjectFromWindow(child,0xFFFFFFF0,ref dispatch,out value)==0){result=value;return false;}
                return true;
            },IntPtr.Zero);return result==null;
        },IntPtr.Zero);
        return result;
    }
    [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT rectangle);
    [DllImport("user32.dll")] static extern bool IsChild(IntPtr parent,IntPtr child);
    [DllImport("user32.dll")] static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr GetWindow(IntPtr hwnd,uint command);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr hwnd,uint message,IntPtr wParam,IntPtr lParam);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hwnd,uint message,IntPtr wParam,IntPtr lParam);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
    [DllImport("user32.dll")] static extern bool GetCursorPos(out POINT point);
    [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(POINT point);
    [DllImport("user32.dll")] static extern void mouse_event(uint flags,uint x,uint y,uint data,UIntPtr extra);
    [DllImport("user32.dll")] static extern void keybd_event(byte key,byte scan,uint flags,UIntPtr extra);
    [StructLayout(LayoutKind.Sequential)] struct POINT { public int X,Y; }
    [StructLayout(LayoutKind.Sequential)] struct RECT { public int L,T,R,B; }
    public static AutomationElement Pane(long excelHwnd) {
        var root=AutomationElement.FromHandle(new IntPtr(excelHwnd));
        foreach(AutomationElement item in root.FindAll(TreeScope.Descendants,new PropertyCondition(AutomationElement.NameProperty,"\uB0B4\uC5D1\uC140 Navigator")))
            if(Search(item)!=null) return item;
        foreach(AutomationElement item in AutomationElement.RootElement.FindAll(TreeScope.Children,new PropertyCondition(AutomationElement.NameProperty,"\uB0B4\uC5D1\uC140 Navigator")))
            if(GetWindow(new IntPtr(item.Current.NativeWindowHandle),4)==new IntPtr(excelHwnd) && Search(item)!=null) return item;
        return null;
    }
    public static AutomationElement Search(AutomationElement pane) {
        return pane.FindFirst(TreeScope.Descendants,new PropertyCondition(AutomationElement.ControlTypeProperty,ControlType.Edit));
    }
    public static bool Docked(AutomationElement pane,long excelHwnd) {
        return pane!=null && IsChild(new IntPtr(excelHwnd),new IntPtr(ContentHandle(pane)));
    }
    public static int ContentHandle(AutomationElement pane) { return Results(pane).Current.NativeWindowHandle; }
    public static bool WindowExists(int hwnd) { return IsWindow(new IntPtr(hwnd)); }
    public static bool ControlsInside(AutomationElement pane) {
        var content=TreeWalker.ControlViewWalker.GetParent(Search(pane));
        var area=pane.Current.BoundingRectangle;
        foreach(AutomationElement child in content.FindAll(TreeScope.Children,Condition.TrueCondition)) {
            var box=child.Current.BoundingRectangle;
            if(box.Width>0 && box.Height>0 && (box.Left<area.Left || box.Top<area.Top || box.Right>area.Right || box.Bottom>area.Bottom)) return false;
        }
        return true;
    }
    public static string Status(AutomationElement pane) {
        foreach(AutomationElement item in pane.FindAll(TreeScope.Descendants,Condition.TrueCondition)) {
            string value=item.Current.Name;
            if(value.StartsWith("\uC2E4\uD589 \uAC00\uB2A5") || value.StartsWith("\uC2E4\uD589 \uBD88\uAC00")) return value;
        }
        return "";
    }
    public static void Filter(AutomationElement pane,string query) {
        ((ValuePattern)Search(pane).GetCurrentPattern(ValuePattern.Pattern)).SetValue(query);
    }
    public static void SelectFirst(AutomationElement pane) {
        var list=Results(pane);
        if(list==null || list.Current.NativeWindowHandle==0) throw new Exception("Route list missing");
        // MSAA SelectionItem.Select changes LB_SETCURSEL without LBN_SELCHANGE.
        // Deliver a Home key to this fixture's observed list HWND, not global input.
        SendMessage(new IntPtr(list.Current.NativeWindowHandle),0x100,new IntPtr(0x24),IntPtr.Zero);
        SendMessage(new IntPtr(list.Current.NativeWindowHandle),0x101,new IntPtr(0x24),IntPtr.Zero);
    }
    public static void ExecuteSelected(AutomationElement pane) {
        IntPtr hwnd=new IntPtr(ContentHandle(pane));
        if(!PostMessage(hwnd,0x100,new IntPtr(0x0D),IntPtr.Zero)) throw new Exception("Fixture input rejected");
        PostMessage(hwnd,0x101,new IntPtr(0x0D),IntPtr.Zero);
    }
    private static AutomationElement Results(AutomationElement pane) {
        return pane.FindFirst(TreeScope.Descendants,new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty,ControlType.List),
            new PropertyCondition(AutomationElement.NameProperty,"\uAE30\uB2A5 \uBAA9\uB85D")));
    }
    public static string Selected(AutomationElement pane) {
        foreach(AutomationElement item in Results(pane).FindAll(TreeScope.Descendants,new PropertyCondition(AutomationElement.ControlTypeProperty,ControlType.ListItem))) {
            if(((SelectionItemPattern)item.GetCurrentPattern(SelectionItemPattern.Pattern)).Current.IsSelected) return item.Current.Name;
        }
        return "";
    }
    public static string ExerciseInput(AutomationElement pane,long excelHwnd) {
        var list=Results(pane);IntPtr hwnd=new IntPtr(list.Current.NativeWindowHandle);
        IntPtr owner=new IntPtr(excelHwnd), top=new IntPtr(pane.Current.NativeWindowHandle);
        POINT original;GetCursorPos(out original);
        try {
            SetForegroundWindow(IsChild(owner,hwnd)?owner:top);
            System.Threading.Thread.Sleep(150);
            IntPtr foreground=GetForegroundWindow();
            if(foreground!=owner && foreground!=top) throw new Exception("Fixture is not foreground; input was not sent");
            var box=list.Current.BoundingRectangle;POINT at=new POINT {X=(int)box.Left+40,Y=(int)box.Top+16};
            if(WindowFromPoint(at)!=hwnd) throw new Exception("Fixture list is occluded; input was not sent");
            SetCursorPos(at.X,at.Y);
            mouse_event(2,0,0,0,UIntPtr.Zero);mouse_event(4,0,0,0,UIntPtr.Zero);
            System.Threading.Thread.Sleep(100);
            string first=Selected(pane);if(first.Length==0)throw new Exception("Mouse click did not select a route");
            keybd_event(0x28,0,0,UIntPtr.Zero);keybd_event(0x28,0,2,UIntPtr.Zero);
            System.Threading.Thread.Sleep(100);
            if(Selected(pane)==first)throw new Exception("Arrow input did not change selection");
            long before=SendMessage(hwnd,0x18E,IntPtr.Zero,IntPtr.Zero).ToInt64();
            mouse_event(0x800,0,0,unchecked((uint)-360),UIntPtr.Zero);
            System.Threading.Thread.Sleep(100);
            long after=SendMessage(hwnd,0x18E,IntPtr.Zero,IntPtr.Zero).ToInt64();
            if(after<=before)throw new Exception("Mouse wheel did not scroll the route list");
            return "Foreground-guarded native click/Down/wheel injection; top index "+before+" -> "+after;
        } finally {SetCursorPos(original.X,original.Y);}
    }
    public static long[] Bounds(AutomationElement pane,long excelHwnd) {
        var box=pane.Current.BoundingRectangle; RECT owner; GetWindowRect(new IntPtr(excelHwnd),out owner);
        return new long[]{(long)box.Left,(long)box.Top,(long)box.Width,(long)box.Height,owner.L,owner.T,owner.R-owner.L,owner.B-owner.T,GetDpiForWindow(new IntPtr(excelHwnd)),pane.Current.NativeWindowHandle};
    }
    public static string Dump(AutomationElement pane) {
        var output=new System.Text.StringBuilder();
        foreach(AutomationElement item in pane.FindAll(TreeScope.Descendants,Condition.TrueCondition))
            output.AppendLine(item.Current.ControlType.ProgrammaticName+"|"+item.Current.Name+"|"+item.Current.BoundingRectangle);
        return output.ToString();
    }
    public static long[] ClientBounds(AutomationElement pane) {
        var content=TreeWalker.ControlViewWalker.GetParent(Search(pane));
        var box=content.Current.BoundingRectangle;
        return new long[]{(long)box.Left,(long)box.Top,(long)box.Width,(long)box.Height};
    }
    public static void Capture(AutomationElement pane,string path) {
        var box=pane.Current.BoundingRectangle;
        using(var bitmap=new Bitmap((int)box.Width,(int)box.Height)) {
            using(var graphics=Graphics.FromImage(bitmap)) graphics.CopyFromScreen((int)box.Left,(int)box.Top,0,0,bitmap.Size);
            bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
    }
}
'@
}
$rows=New-Object 'Collections.Generic.List[object]'
$pane=$null
$originalWindow=@{width=$Excel.Width;height=$Excel.Height;left=$Excel.Left;top=$Excel.Top;state=$Excel.WindowState}
$available=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('7Iuk7ZaJIOqwgOuKpQ=='))
$unavailable=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('7Iuk7ZaJIOu2iOqwgA=='))
function Check([string]$name,[bool]$ok,$details) {
    $rows.Add(@{name=$name;pass=$ok;details=$details})
    if(-not $ok){throw ('UI regression: '+$name)}
}
function Wait-Status([string]$expected) {
    $watch=[Diagnostics.Stopwatch]::StartNew();$observed=''
    do {
        $observed=[NxNavigatorUiProbe]::Status($pane)
        if($observed.StartsWith($expected)){return @{text=$observed;observed_ms=$watch.ElapsedMilliseconds}}
        Start-Sleep -Milliseconds 25
    } while($watch.ElapsedMilliseconds -lt 3000)
    throw ('Status update timeout: '+$observed)
}
try {
    $pane=$null;$deadline=[DateTime]::UtcNow.AddSeconds(5)
    do {$pane=[NxNavigatorUiProbe]::Pane([long]$Excel.Hwnd);if($null -eq $pane){Start-Sleep -Milliseconds 100}}while($null -eq $pane -and [DateTime]::UtcNow -lt $deadline)
    Check 'pane_created' ($null -ne $pane) 'UI Automation'
    $bounds=[NxNavigatorUiProbe]::Bounds($pane,[long]$Excel.Hwnd)
    $rows.Add(@{name='native_bounds';pass=$true;details=$bounds})
    [NxNavigatorUiProbe]::Filter($pane,'NX-DRAW-TITLE-TABLE')
    Start-Sleep -Milliseconds 150
    [NxNavigatorUiProbe]::SelectFirst($pane)
    $selected=[NxNavigatorUiProbe]::Selected($pane)
    Check 'available_focus_off' ($selected.Length -gt 0) (Wait-Status $available)
    $Sheet.Protect();$cell=$Sheet.Range('C3');$cell.Select();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
    $state=Wait-Status $unavailable
    Check 'protected_refresh_focus_off' ([NxNavigatorUiProbe]::Selected($pane) -eq $selected) $state
    $Sheet.Unprotect();$cell=$Sheet.Range('D4');$cell.Select();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
    $state=Wait-Status $available
    Check 'unprotected_refresh_focus_off' ([NxNavigatorUiProbe]::Selected($pane) -eq $selected) $state
    $hwnd=[NxNavigatorUiProbe]::ContentHandle($pane)
    [void]$Excel.Run(($Prefix+'NxHostHideNavigator'))
    [void]$Excel.Run(($Prefix+'NxHostShowNavigator'))
    $pane=$null;$deadline=[DateTime]::UtcNow.AddSeconds(5)
    do {$pane=[NxNavigatorUiProbe]::Pane([long]$Excel.Hwnd);if($null -eq $pane){Start-Sleep -Milliseconds 100}}while($null -eq $pane -and [DateTime]::UtcNow -lt $deadline)
    Check 'reshown_created' ($null -ne $pane) 'UIA after show'
    Check 'reuse_selection' ([NxNavigatorUiProbe]::Selected($pane) -eq $selected -and [NxNavigatorUiProbe]::ContentHandle($pane) -eq $hwnd) (Wait-Status $available)
    Check 'dpi_width' ($bounds[2] -ge (280*$bounds[8]/96)) $bounds
    $client=[NxNavigatorUiProbe]::ClientBounds($pane)
    $docked=[NxNavigatorUiProbe]::Docked($pane,[long]$Excel.Hwnd)
    Check 'ctp_native_child' ($docked -ne [bool]$ExpectFallback) $(if($ExpectFallback){'Expected fallback; Pane lookup requires native owner HWND'}else{'IsChild(Excel, Navigator)'})
    Check 'dpi_client_exact' ([Math]::Abs($client[2]-300*$bounds[8]/96) -le 20 -and $client[3] -gt 300*$bounds[8]/96) $client
    Check 'inside_excel' ($bounds[0] -ge $bounds[4] -and $bounds[1] -ge $bounds[5] -and ($bounds[0]+$bounds[2]) -le ($bounds[4]+$bounds[6]) -and ($bounds[1]+$bounds[3]) -le ($bounds[5]+$bounds[7])) $bounds
    [NxNavigatorUiProbe]::Capture($pane,(Join-Path $EvidenceRoot 'navigator.png'))
    $Excel.WindowState=-4143;$Excel.Width=600;$Excel.Height=410
    $deadline=[DateTime]::UtcNow.AddSeconds(3)
    do {
        Start-Sleep -Milliseconds 100
        $resized=[NxNavigatorUiProbe]::Bounds($pane,[long]$Excel.Hwnd)
        $inside=$resized[0] -ge $resized[4] -and $resized[1] -ge $resized[5] -and ($resized[0]+$resized[2]) -le ($resized[4]+$resized[6]) -and ($resized[1]+$resized[3]) -le ($resized[5]+$resized[7])
    }while(-not $inside -and [DateTime]::UtcNow -lt $deadline)
    Check 'resize_inside_excel' $inside $resized
    Check 'resize_selection_preserved' ([NxNavigatorUiProbe]::Selected($pane) -eq $selected) (Wait-Status $available)
    Check 'resize_controls_inside' ([NxNavigatorUiProbe]::ControlsInside($pane)) 'Actual child control bounds'
    [NxNavigatorUiProbe]::Capture($pane,(Join-Path $EvidenceRoot 'navigator-small.png'))
    $primary=$Sheet.Parent;$second=$null;$secondSheets=$null;$secondSheet=$null;$secondClosed=$false
    try {
        $books=$Excel.Workbooks
        try {$second=$books.Add()}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($books)}
        $secondSheets=$second.Worksheets;$secondSheet=$secondSheets.Item(1)
        $secondSheet.Name='NxSecondFixture';$second.Activate();$secondSheet.Activate()
        [void]$Excel.Run(($Prefix+'NxHostShowNavigator'))
        $secondPane=$null;$deadline=[DateTime]::UtcNow.AddSeconds(5)
        do {$secondPane=[NxNavigatorUiProbe]::Pane([long]$Excel.Hwnd);if($null -eq $secondPane){Start-Sleep -Milliseconds 100}}while($null -eq $secondPane -and [DateTime]::UtcNow -lt $deadline)
        Check 'second_window_pane' ($null -ne $secondPane -and ([NxNavigatorUiProbe]::Docked($secondPane,[long]$Excel.Hwnd) -ne [bool]$ExpectFallback)) 'Independent Excel window'
        Check 'second_window_independent_selection' ([NxNavigatorUiProbe]::Selected($secondPane) -eq '') 'New window starts with its own selection'
        $secondContent=[NxNavigatorUiProbe]::ContentHandle($secondPane)
        $primary.Activate();$Sheet.Activate()
        [void]$Excel.Run(($Prefix+'NxHostShowNavigator'))
        Start-Sleep -Milliseconds 200
        $pane=[NxNavigatorUiProbe]::Pane([long]$Excel.Hwnd)
        Check 'primary_window_preserved' ([NxNavigatorUiProbe]::Selected($pane) -eq $selected -and [NxNavigatorUiProbe]::ContentHandle($pane) -eq $hwnd) 'First window selection retained'
        $second.Close($false)
        $secondClosed=$true
        $deadline=[DateTime]::UtcNow.AddSeconds(3)
        while([NxNavigatorUiProbe]::WindowExists($secondContent) -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 75}
        Check 'closed_window_content_destroyed' (-not [NxNavigatorUiProbe]::WindowExists($secondContent)) 'Closed window has no navigator HWND'
        [void]$Excel.Run(($Prefix+'NxHostShowNavigator'))
        $pane=[NxNavigatorUiProbe]::Pane([long]$Excel.Hwnd)
        Check 'primary_survives_other_close' ([NxNavigatorUiProbe]::ContentHandle($pane) -eq $hwnd -and [NxNavigatorUiProbe]::Selected($pane) -eq $selected) 'Other window close preserves first pane'
    } finally {
        if($null -ne $second -and -not $secondClosed){$second.Close($false)}
        foreach($value in @($secondSheet,$secondSheets,$second)){if($null -ne $value){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($primary)
    }
    $primary=$Sheet.Parent;$firstWindow=$Excel.ActiveWindow;$extraWindow=$null;$extraClosed=$false
    try {
        $extraWindow=$primary.NewWindow()
        $extraWindow.Activate()
        [void]$Excel.Run(($Prefix+'NxHostShowNavigator'))
        $extraPane=$null;$deadline=[DateTime]::UtcNow.AddSeconds(5)
        do{$extraPane=[NxNavigatorUiProbe]::Pane([long]$Excel.Hwnd);if($null -eq $extraPane){Start-Sleep -Milliseconds 100}}while($null -eq $extraPane -and [DateTime]::UtcNow -lt $deadline)
        Check 'same_workbook_new_window' ($null -ne $extraPane -and [NxNavigatorUiProbe]::ContentHandle($extraPane) -ne $hwnd) 'Same document has separate window content'
        $extraHandle=[NxNavigatorUiProbe]::ContentHandle($extraPane)
        $extraWindow.Close();$extraClosed=$true
        $firstWindow.Activate()
        [void]$Excel.Run(($Prefix+'NxHostShowNavigator'))
        Start-Sleep -Milliseconds 200
        $pane=[NxNavigatorUiProbe]::Pane([long]$Excel.Hwnd)
        Check 'same_workbook_close_preserves_primary' ([NxNavigatorUiProbe]::ContentHandle($pane) -eq $hwnd -and [NxNavigatorUiProbe]::Selected($pane) -eq $selected) 'First window state preserved'
        Check 'same_workbook_extra_content_destroyed' (-not [NxNavigatorUiProbe]::WindowExists($extraHandle)) 'Extra window content disposed'
    } finally {
        if($null -ne $extraWindow){if(-not $extraClosed){$extraWindow.Close()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($extraWindow)}
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($firstWindow)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($primary)
    }
    if(-not $ExpectFallback) {
        $other=$null;$otherOwner=$null;$otherBooks=$null;$otherBook=$null;$otherProduct=$null;$otherAddins=$null;$otherHost=$null
        try {
            # Bind to the /x process's own EXCEL7 native object model, not the global ROT.
            $otherBaseline=Get-ExcelProcessBaseline
            $copyBook=$Sheet.Parent
            $copyPath=Join-Path $EvidenceRoot 'separate-process-fixture.xlsx'
            try{$copyBook.SaveCopyAs($copyPath)}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($copyBook)}
            $otherLaunch=Start-Process -FilePath (Get-ExactRegisteredExcelExecutable) -ArgumentList @('/x',('"'+$copyPath+'"')) -WindowStyle Hidden -PassThru
            $deadline=[DateTime]::UtcNow.AddSeconds(20)
            do{
                $nativeWindow=[NxNavigatorUiProbe]::NativeWindow($otherLaunch.Id)
                if($null -ne $nativeWindow){try{$other=$nativeWindow.Application}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($nativeWindow)}}
                if($null -eq $other){Start-Sleep -Milliseconds 100}
            }while($null -eq $other -and [DateTime]::UtcNow -lt $deadline)
            if($null -eq $other){throw ('Unbound second Excel preserved: '+$otherLaunch.Id)}
            $otherOwner=Get-ExactExcelProcessOwnership $other $otherBaseline 'navigator separate Excel process'
            if(-not $otherOwner.owned){throw 'Second Excel ownership rejected'}
            if($otherOwner.pid -ne $otherLaunch.Id){throw 'Second process binding mismatch'}
            $other.Visible=$true;$other.DisplayAlerts=$false
            # COM automation does not automatically connect UI add-ins as normal startup does.
            $otherAddins=$other.COMAddIns;$otherAddins.Update();$otherHost=$otherAddins.Item('LH.NxHost.Connect');$otherHost.Connect=$true
            $otherBooks=$other.Workbooks;$otherBook=$otherBooks.Item('separate-process-fixture.xlsx')
            $otherProduct=$otherBooks.Open((Join-Path $EvidenceRoot 'Product.xlam'),$false,$true)
            $otherBook.Activate()
            $otherPrefix="'"+$otherProduct.Name+"'!"
            Check 'separate_process_show' ([bool]$other.Run(($otherPrefix+'NxHostShowNavigator'))) 'Explicit COM-add-in connection in automation-created instance'
            $otherPane=$null;$deadline=[DateTime]::UtcNow.AddSeconds(5)
            do{$otherPane=[NxNavigatorUiProbe]::Pane([long]$other.Hwnd);if($null -eq $otherPane){Start-Sleep -Milliseconds 100}}while($null -eq $otherPane -and [DateTime]::UtcNow -lt $deadline)
            Check 'separate_process_pane' ($null -ne $otherPane -and [NxNavigatorUiProbe]::Docked($otherPane,[long]$other.Hwnd)) 'Owned second Excel process has a docked navigator'
            Check 'separate_process_independent_content' ([NxNavigatorUiProbe]::ContentHandle($otherPane) -ne $hwnd) 'Content HWND differs across Excel processes'
        } finally {
            if($null -ne $otherProduct){$otherProduct.Close($false)}
            if($null -ne $otherBook){$otherBook.Close($false)}
            if($null -ne $otherHost){$otherHost.Connect=$false}
            foreach($value in @($otherProduct,$otherBook,$otherBooks,$otherHost,$otherAddins)){if($null -ne $value){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)}}
            $otherHost=$null;$otherAddins=$null
            $otherProduct=$null;$otherBook=$null;$otherBooks=$null
            if($null -ne $other -and $null -ne $otherOwner -and $otherOwner.owned){
                $other.Quit();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($other);$other=$null
                [GC]::Collect();[GC]::WaitForPendingFinalizers()
                $otherExit=Stop-ExactProcessAfterGrace -Process $otherOwner.process -Label 'navigator separate Excel' -GraceMs 15000 -Detailed
                Check 'separate_process_natural_exit' ($otherExit.exit_mode -eq 'NATURAL') $otherExit.exit_mode
            }
        }
        $Sheet.Activate()
        [void]$Excel.Run(($Prefix+'NxHostShowNavigator'))
        $pane=[NxNavigatorUiProbe]::Pane([long]$Excel.Hwnd)
        Check 'primary_survives_other_process_exit' ([NxNavigatorUiProbe]::ContentHandle($pane) -eq $hwnd -and [NxNavigatorUiProbe]::Selected($pane) -eq $selected) 'First Excel process keeps navigator state'
    }
    [NxNavigatorUiProbe]::Filter($pane,'')
    Start-Sleep -Milliseconds 100
    $inputEvidence=[NxNavigatorUiProbe]::ExerciseInput($pane,[long]$Excel.Hwnd)
    Check 'native_click_keyboard_wheel' ($inputEvidence.Length -gt 0) $inputEvidence
    $cell=$Sheet.Range('E5')
    try {
        $cell.Value2=0.25;$cell.NumberFormat='0.00';$cell.Select()
        [NxNavigatorUiProbe]::Filter($pane,'NX-CMD-NUMBER-RB-EDIT-NUMBERFORMAT-PERCENTAGE')
        [NxNavigatorUiProbe]::SelectFirst($pane)
        [void](Wait-Status $available)
        [NxNavigatorUiProbe]::ExecuteSelected($pane)
        $deadline=[DateTime]::UtcNow.AddSeconds(5)
        do{Start-Sleep -Milliseconds 100;$format=[string]$cell.NumberFormat}while($format -ne '0.0%' -and [DateTime]::UtcNow -lt $deadline)
        Check 'ctp_ui_executes_owned_sheet_command' ($format -eq '0.0%' -and $cell.Value2 -eq 0.25) 'Enter on observed navigator list formats fixture E5; numeric value preserved'
    } finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)}
} finally {
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'navigator-ui.json'),($rows.ToArray()|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
    if($null -ne $pane){[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'navigator-ui-tree.txt'),[NxNavigatorUiProbe]::Dump($pane),(New-Object Text.UTF8Encoding($false)))}
    $Excel.Width=$originalWindow.width;$Excel.Height=$originalWindow.height;$Excel.Left=$originalWindow.left;$Excel.Top=$originalWindow.top;$Excel.WindowState=$originalWindow.state
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'navigator-ui.json'),($rows.ToArray()|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
}
