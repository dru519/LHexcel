using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

public static class R65PictureWatcher {
    private delegate bool EnumProc(IntPtr hwnd,IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumProc callback,IntPtr state);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr owner,EnumProc callback,IntPtr state);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd,StringBuilder text,int size);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool PostMessage(IntPtr hwnd,uint message,IntPtr wp,IntPtr lp);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd,out RECT rect);
    [DllImport("user32.dll")] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr value);
    [StructLayout(LayoutKind.Sequential)] private struct RECT { public int L,T,R,B; }
    private static string Text(IntPtr hwnd) { var text=new StringBuilder(256);GetWindowText(hwnd,text,text.Capacity);return text.ToString(); }
    private static IntPtr Window(uint owned) {
        IntPtr result=IntPtr.Zero;
        EnumWindows((hwnd,state)=>{uint pid;GetWindowThreadProcessId(hwnd,out pid);
            if(pid==owned && IsWindowVisible(hwnd) && Text(hwnd)=="\uB0B4\uC5D1\uC140 - \uADF8\uB9BC \uBBF8\uB9AC\uBCF4\uAE30"){result=hwnd;return false;}return true;
        },IntPtr.Zero);
        return result;
    }
    public static string Run(uint owned,string path,int seconds) {
        DateTime deadline=DateTime.UtcNow.AddSeconds(seconds);
        bool accepted=false;
        try {
            while(DateTime.UtcNow<deadline) {
                IntPtr window=Window(owned),button=IntPtr.Zero;
                if(window!=IntPtr.Zero) {
                    EnumChildWindows(window,(hwnd,state)=>{
                        if(IsWindowVisible(hwnd) && IsWindowEnabled(hwnd) && Text(hwnd)=="\uD655\uC778"){button=hwnd;return false;}return true;
                    },IntPtr.Zero);
                    if(button!=IntPtr.Zero) {
                        // Allow the Windows opening animation and thumbnail repaint to finish.
                        Thread.Sleep(300);
                        IntPtr prior=SetThreadDpiAwarenessContext(new IntPtr(-4));
                        try {
                            RECT r;GetWindowRect(window,out r);
                            using(var bitmap=new Bitmap(r.R-r.L,r.B-r.T)) {
                                using(var graphics=Graphics.FromImage(bitmap))graphics.CopyFromScreen(r.L,r.T,0,0,bitmap.Size);
                                bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
                            }
                        } finally {if(prior!=IntPtr.Zero)SetThreadDpiAwarenessContext(prior);}
                        accepted=PostMessage(button,0xF5,IntPtr.Zero,IntPtr.Zero);
                        if(!accepted)throw new Exception("Owned picture accept input rejected");
                        return "PASS|Win32 owned caption, enabled child button, screenshot and asynchronous BM_CLICK";
                    }
                }
                Thread.Sleep(100);
            }
            throw new Exception("Owned picture preview did not become ready");
        } finally {
            if(!accepted) {IntPtr window=Window(owned);if(window!=IntPtr.Zero)PostMessage(window,0x10,IntPtr.Zero,IntPtr.Zero);}
        }
    }
}
