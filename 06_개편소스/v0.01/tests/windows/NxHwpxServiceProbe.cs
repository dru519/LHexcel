using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;

internal static class NxHwpxServiceProbe {
    static Type type, contract;
    static object Call(object service, string name, params object[] args) {
        return type.GetMethod(name).Invoke(service, args);
    }
    static string Constant(string name) { return (string)contract.GetField(name).GetValue(null); }
    static object[] Identity() {
        return new object[] { Constant("ProductTitle"), "nx-hwpx-v1", Constant("BridgeHash"), Constant("FeatureHash"), Constant("CommandHash"), Environment.Is64BitProcess ? "x64" : "x86" };
    }
    static void Assert(bool value, string message) { if (!value) throw new Exception(message); }
    static void Reject(Action action, string message) {
        try { action(); } catch (TargetInvocationException) { return; }
        throw new Exception("Expected rejection: " + message);
    }
    static string Wait(object service, string job) {
        var watch = Stopwatch.StartNew();
        for (;;) {
            string state = (string)Call(service, "GetStatus", job);
            if (state != "running") return state;
            if (watch.ElapsedMilliseconds > 60000) throw new Exception("Job did not settle");
            Thread.Sleep(2);
        }
    }
    static int Main(string[] args) {
        try {
            var assembly = Assembly.LoadFrom(Path.GetFullPath(args[0]));
            type = assembly.GetType("LH.NxHost.HwpxExportService", true);
            contract = assembly.GetType("LH.NxHost.NxHostContract", true);
            string root = Path.GetFullPath(args[3]);
            Assert(!Directory.Exists(root), "Fresh evidence directory required");
            Directory.CreateDirectory(root);
            Environment.SetEnvironmentVariable("LHEXCEL_PROFILE_ROOT", root);
            string payload = File.ReadAllText(args[1]), template = Path.GetFullPath(args[2]);
            object service = Activator.CreateInstance(type);
            Reject(delegate { Call(service, "Start", payload, template); }, "no handshake");
            for (int i = 0; i < 6; i++) {
                var identity = Identity(); identity[i] = "mismatch";
                Assert(!(bool)Call(service, "Connect", identity), "Mismatch accepted at " + i);
                Reject(delegate { Call(service, "Start", payload, template); }, "mismatch blocked operations");
            }
            Assert((bool)Call(service, "Connect", Identity()), "Connect");
            Reject(delegate { Call(service, "Start", payload, "\\\\server\\template.hwpx"); }, "network template");
            Reject(delegate { Call(service, "Start", new string('x', 16777217), template); }, "payload size");
            for (int i = 0; i < 12; i++) {
                string id = (string)Call(service, "Start", payload, template);
                Assert(Wait(service, id) == "succeeded", "Success state");
                string path = (string)Call(service, "GetResult", id);
                Assert(File.Exists(path) && path.StartsWith(root + "\\", StringComparison.OrdinalIgnoreCase), "Owned output");
                Reject(delegate { Call(service, "GetStatus", Guid.NewGuid().ToString()); }, "foreign job");
                Call(service, "Cancel", id);
                Assert(File.Exists(path), "Late cancel removed completed result");
                Assert(Directory.GetFiles(Path.GetDirectoryName(path), "*.part").Length == 0, "Success residue");
            }
            string bad = (string)Call(service, "Start", "{}", template);
            Assert(Wait(service, bad) == "failed", "Invalid input state");
            Reject(delegate { Call(service, "GetResult", bad); }, "Failed result");
            var large = new StringBuilder("{\"schema_version\":1,\"rows\":16384,\"columns\":1,\"cells\":[");
            for (int i = 1; i <= 16384; i++) { if (i > 1) large.Append(','); large.Append("{\"row\":" + i + ",\"column\":1,\"text\":\"cancel-test\"}"); }
            large.Append("]}");
            string cancelId = (string)Call(service, "Start", large.ToString(), template);
            Reject(delegate { Call(service, "Start", payload, template); }, "reentrant start");
            object second = Activator.CreateInstance(type); Assert((bool)Call(second, "Connect", Identity()), "Second connection");
            Reject(delegate { Call(second, "Start", payload, template); }, "cross-service concurrency limit");
            Call(service, "Cancel", cancelId);
            Assert(Wait(service, cancelId) == "cancelled", "Cancellation");
            Assert(Directory.GetFiles(Path.Combine(root, "Temp", "Hwpx"), "*.part").Length == 0, "Cancel residue");
            string retry = (string)Call(second, "Start", payload, template);
            Assert(Wait(second, retry) == "succeeded", "Capacity restored after cancellation");
            Console.WriteLine("PASS|Handshake6|MissingHandshake|Size|LocalPaths|Repeat12|ForeignJob|Failure|Cancel|Capacity|NoResidue");
            return 0;
        } catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
    }
}
