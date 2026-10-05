using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;

internal static class NxSetupBrandWindowTests {
    [STAThread] static int Main(string[] args) {
        Application.EnableVisualStyles();
        var assembly=Assembly.LoadFrom(args[0]);
        var type=assembly.GetType("LH.NxSetup.SetupWindow",true);
        using(var window=(Form)Activator.CreateInstance(type,new object[]{"v0.01_r101",new string('b',64),new Dictionary<string,byte[]>(),false})) {
            window.Show(); window.PerformLayout(); Application.DoEvents();
            int marks=0;
            foreach(Control control in window.Controls) if(control is Label) {
                foreach(Control child in control.Controls) if(child is PictureBox) {
                    if(((PictureBox)child).Image==null || !control.ClientRectangle.Contains(child.Bounds)) throw new Exception("Brand mark missing or clipped");
                    marks++;
                }
            }
            if(marks!=1 || window.Icon==null) throw new Exception("Brand mark or application icon missing");
            using(var bitmap=new Bitmap(window.Width,window.Height)) {
                window.DrawToBitmap(bitmap,new Rectangle(Point.Empty,bitmap.Size)); bitmap.Save(args[1]);
            }
            window.Close();
        }
        Console.WriteLine("PASS|SetupBrand|embedded icon and native window layout");
        return 0;
    }
}
