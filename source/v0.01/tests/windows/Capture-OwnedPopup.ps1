# Read-only native capture; reject windows outside the exact test Excel PID.
Add-Type -AssemblyName System.Drawing
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition @'
using System;
using System.Drawing;
using System.Runtime.InteropServices;
public static class NxPopupCapture {
    [StructLayout(LayoutKind.Sequential)] struct Rect {public int L,T,R,B;}
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern IntPtr FindWindow(string cls,string caption);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd,out Rect rect);
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hwnd,IntPtr dc,uint flags);
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    public static void Save(string caption,int owner,string path) {
        IntPtr old=SetThreadDpiAwarenessContext(new IntPtr(-4));
        try {
            IntPtr hwnd=FindWindow(null,caption); uint pid;
            GetWindowThreadProcessId(hwnd,out pid);
            if(hwnd==IntPtr.Zero || pid!=(uint)owner) throw new InvalidOperationException("Owned popup not found");
            Rect r; if(!GetWindowRect(hwnd,out r) || r.R<=r.L || r.B<=r.T) throw new InvalidOperationException("Popup rectangle invalid");
            using(var image=new Bitmap(r.R-r.L,r.B-r.T)) {
                using(var g=Graphics.FromImage(image)) {
                    IntPtr dc=g.GetHdc();
                    try {if(!PrintWindow(hwnd,dc,2)) throw new InvalidOperationException("Native popup capture failed");}
                    finally {g.ReleaseHdc(dc);}
                }
                image.Save(path,System.Drawing.Imaging.ImageFormat.Png);
            }
        } finally {SetThreadDpiAwarenessContext(old);}
    }
}
'@
