using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace LH.NxHost {
    [ComVisible(true), Guid("F0E8CF30-0F9E-4D25-9F33-0B2FBEF8A2D1"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IBridgeService {
        string InterfaceVersion { get; }
        bool Ping(string productTitle, string interfaceVersion, string bridgeHash, string featureHash, string commandHash, string excelBitness, object callback);
        bool ExecuteRequest(string featureId, string commandId, string payload);
    }

    [ComVisible(true), Guid("B5C7A9E3-4CB0-4C85-A6FC-5F9F3BAA0E58"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface INxHostReverseClient {
        bool ExecuteRequest(string featureId, string commandId, string payload);
        string GetRouteAvailability(string routeKey);
    }

    [ComVisible(true), Guid("9A90294C-81C8-4AEE-A2D8-14B078A61CD2"), ProgId("LH.NxHost.BridgeService"), ClassInterface(ClassInterfaceType.None)]
    public sealed class BridgeService : IBridgeService {
        public string InterfaceVersion { get { return NxHostContract.InterfaceVersion; } }
        private bool handshakeComplete;
        private INxHostReverseClient callback;
        private static INxHostReverseClient activeCallback;
        private static readonly HashSet<string> FeatureAllowlist = new HashSet<string>(StringComparer.Ordinal) { "NX-HOST-BRIDGE", "NX-UTIL-NAVIGATOR", "NX-UTIL-DOCUMENT-NAVIGATOR", "NX-DATA-FOCUS-CELL" };
        private static readonly HashSet<string> CommandAllowlist = new HashSet<string>(StringComparer.Ordinal) { "PING", "DISCONNECT", "SHOW_NAVIGATOR", "HIDE_NAVIGATOR", "SHOW_DOCUMENT_NAVIGATOR", "HIDE_DOCUMENT_NAVIGATOR", "FOCUS_START", "FOCUS_UPDATE", "FOCUS_STOP" };
        public bool Ping(string productTitle, string interfaceVersion, string bridgeHash, string featureHash, string commandHash, string excelBitness, object reverseClient) {
            handshakeComplete = string.Equals(productTitle, NxHostContract.ProductTitle, StringComparison.Ordinal) &&
                string.Equals(interfaceVersion, NxHostContract.InterfaceVersion, StringComparison.Ordinal) &&
                string.Equals(bridgeHash, NxHostContract.BridgeHash, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(featureHash, NxHostContract.FeatureHash, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(commandHash, NxHostContract.CommandHash, StringComparison.OrdinalIgnoreCase) &&
                (excelBitness == "x86" || excelBitness == "x64") && excelBitness == NxHostContract.ProcessBitness;
            callback = handshakeComplete ? ReverseClientAdapter.TryCreate(reverseClient) : null;
            handshakeComplete = handshakeComplete && callback != null;
            activeCallback = handshakeComplete ? callback : null;
            return handshakeComplete;
        }
        public bool ExecuteRequest(string featureId, string commandId, string payload) {
            if (!handshakeComplete) return false;
            if (!FeatureAllowlist.Contains(featureId) || !CommandAllowlist.Contains(commandId) || payload == null || payload.Length > 512 || payload.IndexOf('\0') >= 0 || payload.IndexOf('\r') >= 0 || payload.IndexOf('\n') >= 0) return false;
            if (commandId == "PING") return featureId == "NX-HOST-BRIDGE";
            if (commandId == "DISCONNECT" && featureId == "NX-HOST-BRIDGE") { NxHostRuntime.StopFocus(); handshakeComplete = false; callback = null; activeCallback = null; return true; }
            return NxHostRuntime.Execute(featureId, commandId, payload);
        }
        internal static bool ExecuteReverse(string featureId, string commandId, string payload) { return activeCallback != null && activeCallback.ExecuteRequest(featureId, commandId, payload); }
        // Separate, closed navigation callback; never route display tokens as feature IDs.
        internal static Array GetDocuments() {
            var client = activeCallback as ReverseClientAdapter;
            return client == null ? null : client.GetDocuments();
        }
        internal static bool MoveDocument(string token) {
            var client = activeCallback as ReverseClientAdapter;
            return client != null && client.MoveDocument(token);
        }
        internal static string GetReverseAvailability(string routeKey) {
            if (activeCallback == null || string.IsNullOrEmpty(routeKey) || routeKey.Length > 160) return "";
            return activeCallback.GetRouteAvailability(routeKey) ?? "";
        }
    }

    internal static class NxHostContract {
        public const string InterfaceVersion = "nx-bridge-v1";
        public const string ProductTitle = "내엑셀 v0.01_r86";
        public const string BridgeHash = "79027b72f740152d898b2b5a954f7f27bba9e85ed4f3317a49666cd9969f5401";
        public const string FeatureHash = "fb4c29845c44d115061adae4f375d7ede8b6ce87717973120fe9ce80743686db";
        public const string CommandHash = "8bf73de7e93494322cf58f2a9e5648ea983e3ce0cbba40b0821364152eb3e809";
        public static string ProcessBitness { get { return IntPtr.Size == 4 ? "x86" : "x64"; } }
    }
    internal sealed class ReverseClientAdapter : INxHostReverseClient {
        private readonly object target;
        private readonly System.Reflection.MethodInfo executeMethod;
        private readonly System.Reflection.MethodInfo availabilityMethod;
        private ReverseClientAdapter(object value, System.Reflection.MethodInfo execute, System.Reflection.MethodInfo availability) {
            target = value;
            executeMethod = execute;
            availabilityMethod = availability;
        }
        public static INxHostReverseClient TryCreate(object value) {
            if (value == null) return null;
            var execute = value.GetType().GetMethod("ExecuteRequest", new[] { typeof(string), typeof(string), typeof(string) });
            var availability = value.GetType().GetMethod("GetRouteAvailability", new[] { typeof(string) });
            if ((execute != null && availability != null) || Marshal.IsComObject(value))
                return new ReverseClientAdapter(value, execute, availability);
            return null;
        }
        public bool ExecuteRequest(string featureId, string commandId, string payload) {
            try {
                object result = executeMethod != null
                    ? executeMethod.Invoke(target, new object[] { featureId, commandId, payload })
                    : target.GetType().InvokeMember("ExecuteRequest", System.Reflection.BindingFlags.InvokeMethod, null, target, new object[] { featureId, commandId, payload });
                return result is bool && (bool)result;
            } catch { return false; }
        }
        public string GetRouteAvailability(string routeKey) {
            try {
                object result = availabilityMethod != null
                    ? availabilityMethod.Invoke(target, new object[] { routeKey })
                    : target.GetType().InvokeMember("GetRouteAvailability", System.Reflection.BindingFlags.InvokeMethod, null, target, new object[] { routeKey });
                return result as string ?? "";
            } catch { return ""; }
        }
        internal Array GetDocuments() {
            try { return target.GetType().InvokeMember("GetDocumentSnapshot", System.Reflection.BindingFlags.InvokeMethod, null, target, new object[0]) as Array; }
            catch { return null; }
        }
        internal bool MoveDocument(string token) {
            int value;
            if (!Int32.TryParse(token, out value) || value <= 0) return false;
            try { return true.Equals(target.GetType().InvokeMember("MoveDocument", System.Reflection.BindingFlags.InvokeMethod, null, target, new object[] { token })); }
            catch { return false; }
        }
    }
}
