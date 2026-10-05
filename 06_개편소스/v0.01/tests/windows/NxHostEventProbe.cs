using System;
using System.Reflection;

internal static class NxHostEventProbe {
    [STAThread]
    static int Main(string[] args) {
        var assembly = Assembly.LoadFrom(args[0]);
        var type = assembly.GetType("LH.NxHost.ExcelEventBroker", true);
        int refreshes = 0;
        object broker = Activator.CreateInstance(type, new object[] { null, new Action(delegate { refreshes++; }) });
        var queue = type.GetMethod("Queue", BindingFlags.NonPublic | BindingFlags.Instance);
        var tick = type.GetMethod("OnTimer", BindingFlags.NonPublic | BindingFlags.Instance);
        try {
            for (int i = 0; i < 500; i++) queue.Invoke(broker, null);
            tick.Invoke(broker, new object[] { null, EventArgs.Empty });
            if (refreshes != 1) throw new Exception("Navigator must refresh once with focus OFF; actual=" + refreshes);
            tick.Invoke(broker, new object[] { null, EventArgs.Empty });
            if (refreshes != 1) throw new Exception("Idle tick duplicated refresh");
            ((IDisposable)broker).Dispose();
            queue.Invoke(broker, null);
            tick.Invoke(broker, new object[] { null, EventArgs.Empty });
            if (refreshes != 1) throw new Exception("Disposed broker refreshed");
            Console.WriteLine("PASS|FocusOffRefresh|Coalesced500|Idle|Disposed");
            return 0;
        } catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { ((IDisposable)broker).Dispose(); }
    }
}
