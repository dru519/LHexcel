using System;
using System.Runtime.InteropServices;

internal static class NxNavigatorComProbe {
    [STAThread]
    static int Main() {
        object target = null;
        IntPtr unknown = IntPtr.Zero;
        string[] names = { "IDispatch", "IOleObject", "IOleInPlaceObject", "IOleControl", "IViewObject", "IPersistStreamInit" };
        string[] ids = { "00020400-0000-0000-C000-000000000046", "00000112-0000-0000-C000-000000000046",
            "00000113-0000-0000-C000-000000000046", "B196B288-BAB4-101A-B69C-00AA00341D07",
            "0000010D-0000-0000-C000-000000000046", "7FD52380-4E07-101B-AE2D-08002B2EC713" };
        try {
            Console.WriteLine("process_bits=" + (IntPtr.Size * 8));
            target = Activator.CreateInstance(Type.GetTypeFromProgID("LH.NxHost.NavigatorPane", true));
            unknown = Marshal.GetIUnknownForObject(target);
            bool complete = true;
            for (int index = 0; index < names.Length; index++) {
                Guid iid = new Guid(ids[index]);
                IntPtr pointer = IntPtr.Zero;
                try {
                    int hr = Marshal.QueryInterface(unknown, ref iid, out pointer);
                    Console.WriteLine(names[index] + "|" + hr.ToString("X8") + "|" + (hr == 0 && pointer != IntPtr.Zero));
                    complete &= hr == 0 && pointer != IntPtr.Zero;
                } finally { if (pointer != IntPtr.Zero) Marshal.Release(pointer); }
            }
            return complete ? 0 : 2;
        } catch (Exception error) {
            Console.WriteLine("activation_failed|" + error.GetType().FullName + "|" + error.HResult.ToString("X8"));
            return 1;
        } finally {
            if (unknown != IntPtr.Zero) Marshal.Release(unknown);
            IDisposable disposable = target as IDisposable;
            if (disposable != null) disposable.Dispose();
            if (target != null && Marshal.IsComObject(target)) Marshal.FinalReleaseComObject(target);
        }
    }
}
