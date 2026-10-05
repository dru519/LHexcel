using System;
using System.Collections.Generic;
using System.Drawing;

namespace LH.NxHost {
    internal static class NxHostRuntime {
        private static NxHostAddIn addIn;
        private static readonly FocusOverlayController overlay = new FocusOverlayController();
        private static FocusPayload focusPayload;
        private static bool focusActive;
        private static string focusSignature = String.Empty;
        private static Action<bool> focusPollingChanged;

        public static void Attach(NxHostAddIn value) { addIn = value; }
        internal static void SetFocusPolling(Action<bool> callback) { focusPollingChanged = callback; }
        public static void Detach(NxHostAddIn value) {
            if (!ReferenceEquals(addIn, value)) return;
            StopFocus();
            overlay.Dispose();
            focusPollingChanged = null;
            addIn = null;
        }

        public static bool Execute(string featureId, string commandId, string payload) {
            if (featureId == "NX-UTIL-DOCUMENT-NAVIGATOR" && commandId == "SHOW_DOCUMENT_NAVIGATOR") return addIn != null && addIn.CreatePane(true);
            if (featureId == "NX-UTIL-DOCUMENT-NAVIGATOR" && commandId == "HIDE_DOCUMENT_NAVIGATOR") return addIn == null || addIn.HidePane(true);
            if (featureId == "NX-UTIL-NAVIGATOR" && commandId == "SHOW_NAVIGATOR") return addIn != null && addIn.CreatePane();
            if (featureId == "NX-UTIL-NAVIGATOR" && commandId == "HIDE_NAVIGATOR") return addIn != null && addIn.HidePane();
            if (featureId != "NX-DATA-FOCUS-CELL") return false;
            if (commandId == "FOCUS_STOP") { StopFocus(); return true; }
            if (commandId != "FOCUS_START" && commandId != "FOCUS_UPDATE") return false;
            FocusPayload parsed;
            if (!FocusPayload.TryParse(payload, out parsed)) return false;
            focusPayload = parsed;
            focusActive = true;
            if (focusPollingChanged != null) focusPollingChanged(true);
            return RefreshFocus();
        }

        internal static bool RefreshFocus() {
            if (!focusActive || focusPayload == null || addIn == null || addIn.Application == null) return false;
            FocusGeometryResult geometry;
            if (!FocusGeometryBuilder.TryBuild(addIn.Application, out geometry)) {
                // A hidden/off-screen/non-range selection is a valid non-rendering state,
                // not an engine failure that should force a VBA backend downgrade.
                overlay.Hide();
                focusSignature = FocusGeometryBuilder.GetStateSignature(addIn.Application);
                return true;
            }
            if (!overlay.Show(geometry, focusPayload)) return false;
            focusSignature = geometry.Signature;
            return true;
        }

        internal static void PollFocus() {
            if (!focusActive || addIn == null || addIn.Application == null) return;
            string current = FocusGeometryBuilder.GetStateSignature(addIn.Application);
            if (!String.Equals(current, focusSignature, StringComparison.Ordinal)) RefreshFocus();
        }

        internal static void HideFocus() { overlay.Hide(); }
        internal static void StopFocus() { overlay.Reset(); focusPayload = null; focusActive = false; focusSignature = String.Empty; if (focusPollingChanged != null) focusPollingChanged(false); }
        internal static bool FocusActive { get { return focusActive; } }
        internal static int FocusWindowCount { get { return overlay.WindowCount; } }
        internal static string FocusSnapshot { get { return overlay.Snapshot; } }
    }

    internal sealed class FocusPayload {
        public string Shape;
        public string Style;
        public Color Color;
        public int Intensity;
        public bool Selected;

        public static bool TryParse(string payload, out FocusPayload result) {
            result = null;
            if (String.IsNullOrEmpty(payload) || payload.Length > 512) return false;
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string part in payload.Split(';')) {
                int split = part.IndexOf('=');
                if (split <= 0 || split == part.Length - 1) return false;
                string key = part.Substring(0, split);
                if (map.ContainsKey(key)) return false;
                map[key] = part.Substring(split + 1);
            }
            if (map.Count != 6 || !map.ContainsKey("version") || map["version"] != "1" ||
                !map.ContainsKey("shape") || !map.ContainsKey("style") || !map.ContainsKey("color") ||
                !map.ContainsKey("intensity") || !map.ContainsKey("selected")) return false;
            string shape = map["shape"];
            if (shape != "criss-cross" && shape != "horizontal" && shape != "vertical") return false;
            if (map["style"] != "wide-stripe") return false;
            int intensity;
            if (!Int32.TryParse(map["intensity"], out intensity) || intensity < 0 || intensity > 100) return false;
            bool selected;
            if (map["selected"] == "0") selected = false;
            else if (map["selected"] == "1") selected = true;
            else if (!Boolean.TryParse(map["selected"], out selected)) return false;
            Color color;
            try { color = ColorTranslator.FromHtml(map["color"]); }
            catch { return false; }
            result = new FocusPayload { Shape = shape, Style = map["style"], Color = color, Intensity = intensity, Selected = selected };
            return true;
        }
    }
}
