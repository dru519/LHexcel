using System;
using System.IO;
using System.Linq;
using System.Reflection;

internal static class ProtectedLoaderProbe {
    private static bool CoreLoaded() {
        return AppDomain.CurrentDomain.GetAssemblies().Any(a => a.GetName().Name.StartsWith("NxCore",StringComparison.Ordinal));
    }
    [STAThread]
    private static int Main(string[] args) {
        string mode=args[0], path=Path.GetFullPath(args[1]);
        try {
            if(mode=="binding_collision") {
                string other=Path.Combine(Path.GetDirectoryName(path),"other","NxCore64.dll");
                Assembly.LoadFrom(other);
            }
            Assembly bootstrap=Assembly.LoadFrom(path);
            try {
                object bridge=Activator.CreateInstance(bootstrap.GetType("LH.NxHost.BridgeService",true));
                if(mode!="normal")throw new Exception("Invalid core was accepted");
                string version=(string)bridge.GetType().GetProperty("InterfaceVersion").GetValue(bridge,null);
                if(String.IsNullOrEmpty(version)||!CoreLoaded())throw new Exception("Normal bridge failed");
                foreach(string name in new[]{"HwpxExportService","PicturePreviewService","WorkbookCompareService","WorkbookCompareResultsService"})
                    Activator.CreateInstance(bootstrap.GetType("LH.NxHost."+name,true));
                Console.WriteLine("PASS|normal|verified_core_services");
            } catch(TargetInvocationException) {
                if(mode=="normal")throw;
                if(mode!="binding_collision" && CoreLoaded())throw new Exception("Rejected core was loaded");
                Console.WriteLine("PASS|"+mode+"|core_rejected");
            }
            return 0;
        } catch(Exception error) {
            Console.WriteLine("FAIL|"+mode+"|"+error.GetBaseException().Message);
            return 1;
        }
    }
}
