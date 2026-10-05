using System;
using System.Drawing;
using System.Windows.Forms;

namespace LH.NxHost {
    internal sealed class FocusOverlayController : IDisposable {
        private readonly SegmentForm[] forms = new SegmentForm[8];

        public bool Show(FocusGeometryResult geometry, FocusPayload payload) {
            if (geometry == null || payload == null || geometry.Owner == IntPtr.Zero) return false;
            EnsureForms();
            for (int index = 0; index < forms.Length; index++) {
                bool enabled = SegmentEnabled((FocusSegment)index, payload);
                forms[index].ApplyAppearance(payload.Color, payload.Intensity);
                forms[index].SetOwner(geometry.Owner);
                if (enabled && geometry.Visible[index]) forms[index].ShowSegment(geometry.Segments[index]);
                else forms[index].HideSegment();
            }
            return true;
        }

        public void Hide() {
            for (int index = 0; index < forms.Length; index++) if (forms[index] != null) forms[index].HideSegment();
        }

        public int WindowCount {
            get { int count = 0; for (int index = 0; index < forms.Length; index++) if (forms[index] != null && !forms[index].IsDisposed) count++; return count; }
        }

        public string Snapshot {
            get {
                var parts = new string[forms.Length];
                for (int index = 0; index < forms.Length; index++) parts[index] = forms[index] == null ? SegmentToken((FocusSegment)index) + ":MISSING" : forms[index].Snapshot;
                return String.Join("|", parts);
            }
        }

        public void Reset() {
            for (int index = 0; index < forms.Length; index++) {
                if (forms[index] == null) continue;
                forms[index].HideSegment();
                forms[index].Dispose();
                forms[index] = null;
            }
        }

        public void Dispose() { Reset(); }

        private void EnsureForms() {
            for (int index = 0; index < forms.Length; index++) if (forms[index] == null || forms[index].IsDisposed) forms[index] = new SegmentForm(SegmentToken((FocusSegment)index));
        }

        private static bool SegmentEnabled(FocusSegment segment, FocusPayload payload) {
            if (segment == FocusSegment.RowLeft || segment == FocusSegment.RowRight) return payload.Shape == "criss-cross" || payload.Shape == "horizontal";
            if (segment == FocusSegment.ColumnTop || segment == FocusSegment.ColumnBottom) return payload.Shape == "criss-cross" || payload.Shape == "vertical";
            return payload.Selected;
        }

        internal static string SegmentToken(FocusSegment segment) {
            switch (segment) {
                case FocusSegment.RowLeft: return "ROW_LEFT";
                case FocusSegment.RowRight: return "ROW_RIGHT";
                case FocusSegment.ColumnTop: return "COLUMN_TOP";
                case FocusSegment.ColumnBottom: return "COLUMN_BOTTOM";
                case FocusSegment.SelectTop: return "SELECT_TOP";
                case FocusSegment.SelectRight: return "SELECT_RIGHT";
                case FocusSegment.SelectBottom: return "SELECT_BOTTOM";
                case FocusSegment.SelectLeft: return "SELECT_LEFT";
                default: throw new ArgumentOutOfRangeException("segment");
            }
        }

        private sealed class SegmentForm : Form {
            private const int WS_EX_TRANSPARENT = 0x20;
            private const int WS_EX_LAYERED = 0x80000;
            private const int WS_EX_NOACTIVATE = 0x8000000;
            private const int WS_EX_TOOLWINDOW = 0x80;
            private readonly string token;
            private IntPtr owner;
            private Color color;
            private int intensity;

            internal SegmentForm(string value) {
                token = value;
                Text = "NX_FOCUS_DLL_" + value;
                Name = Text;
                FormBorderStyle = FormBorderStyle.None;
                ShowInTaskbar = false;
                StartPosition = FormStartPosition.Manual;
                Enabled = false;
            }

            protected override bool ShowWithoutActivation { get { return true; } }
            protected override CreateParams CreateParams {
                get {
                    CreateParams value = base.CreateParams;
                    value.ExStyle |= WS_EX_TRANSPARENT | WS_EX_LAYERED | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
                    return value;
                }
            }

            internal void ApplyAppearance(Color value, int percent) {
                color = value;
                intensity = Math.Max(0, Math.Min(100, percent));
                BackColor = color;
                Opacity = intensity / 100.0;
            }

            internal void SetOwner(IntPtr value) {
                if (value == IntPtr.Zero) throw new ArgumentException("Excel owner is required", "value");
                owner = value;
                if (IsHandleCreated) FocusNativeMethods.SetOwner(Handle, owner);
            }

            internal void ShowSegment(Rectangle bounds) {
                if (bounds.Width <= 0 || bounds.Height <= 0) { HideSegment(); return; }
                SetBounds(bounds.Left, bounds.Top, bounds.Width, bounds.Height, BoundsSpecified.All);
                if (!Visible) Show();
                FocusNativeMethods.SetOwner(Handle, owner);
            }

            internal void HideSegment() { if (Visible) Hide(); }

            internal string Snapshot {
                get {
                    return token + ":" + (Visible ? Bounds.Left + "," + Bounds.Top + "," + Bounds.Width + "," + Bounds.Height : "HIDDEN") +
                        ";" + color.ToArgb() + ";" + (int)Math.Floor(intensity * 255.0 / 100.0 + 0.5);
                }
            }
        }
    }
}
