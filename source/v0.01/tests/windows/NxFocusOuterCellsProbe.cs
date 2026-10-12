using System;
using System.Drawing;
using System.Reflection;

// Synthetic geometry only; native Excel screen coordinates require separate validation.
internal static class NxFocusOuterCellsProbe {
    static int Main(string[] args) {
        try {
            var assembly = Assembly.LoadFrom(args[0]);
            var payload = assembly.GetType("LH.NxHost.FocusPayload", true);
            foreach (string style in new[] {"outer-cells", "wide-stripe", "outline"}) {
                foreach (string shape in new[] {"criss-cross", "horizontal", "vertical"}) {
                    object[] input = {"version=1;shape=" + shape + ";style=" + style + ";color=#5FC8D8;intensity=80;selected=0", null};
                    if (!(bool)payload.GetMethod("TryParse").Invoke(null, input)) throw new Exception("payload rejected");
                }
            }
            Type geometryType = assembly.GetType("LH.NxHost.FocusGeometryResult", true);
            object geometry = Activator.CreateInstance(geometryType, true);
            var selected = new Rectangle(40, 30, 20, 30);
            geometryType.GetField("Selection").SetValue(geometry, selected);
            var bands = (Rectangle[])geometryType.GetField("OuterBands").GetValue(geometry);
            bands[0] = new Rectangle(0,20,100,10); bands[1] = new Rectangle(0,60,100,15);
            bands[2] = new Rectangle(25,0,15,100); bands[3] = new Rectangle(60,0,25,100);
            var method = assembly.GetType("LH.NxHost.FocusOverlayController",true).GetMethod("BuildOuterSegments", BindingFlags.Static | BindingFlags.NonPublic);
            foreach (string shape in new[] {"criss-cross", "horizontal", "vertical"}) {
                var segments = (Rectangle[])method.Invoke(null, new object[] {geometry, shape});
                for (int x=0; x<100; x++) for (int y=0; y<100; y++) {
                    var p = new Point(x,y); int count=0;
                    foreach (var r in segments) if (r.Contains(p)) count++;
                    bool expected = shape == "horizontal" ? bands[0].Contains(p) || bands[1].Contains(p) :
                        shape == "vertical" ? bands[2].Contains(p) || bands[3].Contains(p) :
                        (bands[0].Contains(p) || bands[1].Contains(p) || bands[2].Contains(p) || bands[3].Contains(p));
                    if (count != (expected ? 1 : 0)) throw new Exception(shape + " gap/overlap at " + p);
                }
            }
            Console.WriteLine("PASS|9 payload combinations; 30000 pixel coverage checks; no overlap");
            return 0;
        } catch(Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
