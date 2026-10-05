using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using LH.NxSetup;

internal static class NxSetupWindowTests
{
    [STAThread]
    static int Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        using(var window=new SetupWindow("v0.01_r93",new string('b',64),new Dictionary<string,byte[]>())) {
            window.Show(); window.PerformLayout(); Application.DoEvents();
            foreach(Control child in window.Controls) if(child is FlowLayoutPanel) {
                if(child.Controls.Count!=6) throw new Exception("Missing action");
                foreach(Control button in child.Controls) {
                    if(!child.ClientRectangle.Contains(button.Bounds)) throw new Exception("Clipped button: "+button.Text);
                    if(TextRenderer.MeasureText(button.Text,button.Font).Width>button.ClientSize.Width) throw new Exception("Clipped label: "+button.Text);
                }
            }
            using(var image=new Bitmap(window.Width,window.Height)) {
                window.DrawToBitmap(image,new Rectangle(Point.Empty,image.Size));
                image.Save(args[0]);
            }
            window.Close();
        }
        Console.WriteLine("PASS|SetupWindow|6 actions|native WinForms layout");
        return 0;
    }
}
