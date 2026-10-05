using System;
using Microsoft.Win32;
using System.Runtime.InteropServices;

namespace LH.NxHost {
    internal static class Registration {
        private const string Addins = @"Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect";
        public static void Register(string dll) { using (RegistryKey k = Registry.CurrentUser.CreateSubKey(Addins)) { k.SetValue("FriendlyName", "내엑셀 NxHost"); k.SetValue("Description", "내엑셀 COM host"); k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord); k.SetValue("CommandLineSafe", 1, RegistryValueKind.DWord); k.SetValue("Manifest", dll); } }
        public static void Unregister() { Registry.CurrentUser.DeleteSubKeyTree(Addins, false); }
    }
}
