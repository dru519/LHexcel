using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace LH.NxHost {
    internal enum FocusSegment {
        RowLeft = 0,
        RowRight = 1,
        ColumnTop = 2,
        ColumnBottom = 3,
        SelectTop = 4,
        SelectRight = 5,
        SelectBottom = 6,
        SelectLeft = 7
    }

    internal sealed class FocusGeometryResult {
        public readonly Rectangle[] Segments = new Rectangle[8];
        public readonly bool[] Visible = new bool[8];
        public Rectangle Grid;
        public Rectangle Selection;
        public IntPtr Owner;
        public int Dpi;
        public int Zoom;
        public int PaneIndex;
        public string Signature;

        public void Set(FocusSegment segment, int left, int top, int right, int bottom) {
            if (right <= left || bottom <= top) return;
            int index = (int)segment;
            Segments[index] = Rectangle.FromLTRB(left, top, right, bottom);
            Visible[index] = true;
        }
    }

    internal static class FocusGeometryBuilder {
        private const double RowHeaderLogicalPixels = 31.0;
        private const double ColumnHeaderLogicalPixels = 22.5;

        public static bool TryBuild(Excel.Application app, out FocusGeometryResult result) {
            result = null;
            if (app == null) return false;
            Excel.Window window = null;
            Excel.Pane pane = null;
            Excel.Range visible = null;
            Excel.Range selection = null;
            Excel.Range logical = null;
            Excel.Areas areas = null;
            try {
                window = app.ActiveWindow;
                selection = app.Selection as Excel.Range;
                if (window == null || selection == null || window.Hwnd == 0) return false;
                areas = selection.Areas;
                if (areas == null || areas.Count != 1) return false;
                logical = selection;
                pane = window.ActivePane;
                if (pane == null) return false;
                visible = pane.VisibleRange;
                if (visible == null) return false;

                double visibleLeft = Convert.ToDouble(visible.Left);
                double visibleTop = Convert.ToDouble(visible.Top);
                double visibleRight = visibleLeft + Convert.ToDouble(visible.Width);
                double visibleBottom = visibleTop + Convert.ToDouble(visible.Height);
                double selectionLeftPoints = Convert.ToDouble(logical.Left);
                double selectionTopPoints = Convert.ToDouble(logical.Top);
                double selectionRightPoints = selectionLeftPoints + Convert.ToDouble(logical.Width);
                double selectionBottomPoints = selectionTopPoints + Convert.ToDouble(logical.Height);
                if (selectionRightPoints <= visibleLeft || selectionBottomPoints <= visibleTop ||
                    selectionLeftPoints >= visibleRight || selectionTopPoints >= visibleBottom) return false;
                selectionLeftPoints = Math.Max(selectionLeftPoints, visibleLeft);
                selectionTopPoints = Math.Max(selectionTopPoints, visibleTop);
                selectionRightPoints = Math.Min(selectionRightPoints, visibleRight);
                selectionBottomPoints = Math.Min(selectionBottomPoints, visibleBottom);

                int dpi = FocusNativeMethods.GetDpiForWindow(new IntPtr(window.Hwnd));
                if (dpi <= 0) dpi = 96;
                int zoom = Convert.ToInt32(window.Zoom);
                if (zoom <= 0) return false;
                double pointScale = dpi / 72.0 * zoom / 100.0;
                int paneIndex = pane.Index;
                bool rightPane;
                bool bottomPane;
                GetPanePosition(window, paneIndex, out rightPane, out bottomPane);
                int divider = window.FreezePanes ? 0 : Round(2.0 * dpi / 96.0);
                int rowHeader = window.DisplayHeadings ? Round(RowHeaderLogicalPixels * dpi / 96.0) : 0;
                int columnHeader = window.DisplayHeadings ? Round(ColumnHeaderLogicalPixels * dpi / 96.0) : 0;
                int gridLeft = window.PointsToScreenPixelsX(0) + Round(visibleLeft * pointScale);
                int gridTop = window.PointsToScreenPixelsY(0) + Round(visibleTop * pointScale);
                if (rightPane) gridLeft += rowHeader + Round(Convert.ToDouble(window.SplitHorizontal) * pointScale) + divider;
                if (bottomPane) gridTop += columnHeader + Round(Convert.ToDouble(window.SplitVertical) * pointScale) + divider;

                Rectangle excel7;
                if (!FocusNativeMethods.TryGetExcel7Bounds(new IntPtr(window.Hwnd), out excel7) &&
                    !FocusNativeMethods.TryGetExcel7Bounds(new IntPtr(app.Hwnd), out excel7)) return false;
                int gridRight = Math.Min(excel7.Right, gridLeft + Round(Convert.ToDouble(visible.Width) * pointScale));
                int gridBottom = Math.Min(excel7.Bottom, gridTop + Round(Convert.ToDouble(visible.Height) * pointScale));
                gridLeft = Math.Max(gridLeft, excel7.Left);
                gridTop = Math.Max(gridTop, excel7.Top);
                if (gridRight <= gridLeft || gridBottom <= gridTop) return false;

                int selectionLeft = gridLeft + Round((selectionLeftPoints - visibleLeft) * pointScale);
                int selectionTop = gridTop + Round((selectionTopPoints - visibleTop) * pointScale);
                int selectionRight = selectionLeft + Round((selectionRightPoints - selectionLeftPoints) * pointScale);
                int selectionBottom = selectionTop + Round((selectionBottomPoints - selectionTopPoints) * pointScale);
                selectionLeft = Math.Max(selectionLeft, gridLeft);
                selectionTop = Math.Max(selectionTop, gridTop);
                selectionRight = Math.Min(selectionRight, gridRight);
                selectionBottom = Math.Min(selectionBottom, gridBottom);
                if (selectionRight <= selectionLeft || selectionBottom <= selectionTop) return false;

                var geometry = new FocusGeometryResult();
                geometry.Grid = Rectangle.FromLTRB(gridLeft, gridTop, gridRight, gridBottom);
                geometry.Selection = Rectangle.FromLTRB(selectionLeft, selectionTop, selectionRight, selectionBottom);
                geometry.Owner = new IntPtr(window.Hwnd);
                geometry.Dpi = dpi;
                geometry.Zoom = zoom;
                geometry.PaneIndex = paneIndex;
                geometry.Signature = BuildStateSignature(window, pane, visible, logical);
                geometry.Set(FocusSegment.RowLeft, gridLeft, selectionTop, selectionLeft, selectionBottom);
                geometry.Set(FocusSegment.RowRight, selectionRight, selectionTop, gridRight, selectionBottom);
                geometry.Set(FocusSegment.ColumnTop, selectionLeft, gridTop, selectionRight, selectionTop);
                geometry.Set(FocusSegment.ColumnBottom, selectionLeft, selectionBottom, selectionRight, gridBottom);
                int outline = Math.Max(1, Round(2.0 * dpi / 96.0));
                geometry.Set(FocusSegment.SelectTop, selectionLeft, Math.Max(gridTop, selectionTop - outline), selectionRight, selectionTop);
                geometry.Set(FocusSegment.SelectRight, selectionRight, selectionTop, Math.Min(gridRight, selectionRight + outline), selectionBottom);
                geometry.Set(FocusSegment.SelectBottom, selectionLeft, selectionBottom, selectionRight, Math.Min(gridBottom, selectionBottom + outline));
                geometry.Set(FocusSegment.SelectLeft, Math.Max(gridLeft, selectionLeft - outline), selectionTop, selectionLeft, selectionBottom);
                result = geometry;
                return true;
            } catch {
                return false;
            } finally {
                Release(areas);
                if (!Object.ReferenceEquals(logical, selection)) Release(logical);
                Release(selection);
                Release(visible);
                Release(pane);
                Release(window);
            }
        }

        public static string GetStateSignature(Excel.Application app) {
            if (app == null) return String.Empty;
            Excel.Window window = null;
            Excel.Pane pane = null;
            Excel.Range visible = null;
            Excel.Range selection = null;
            try {
                window = app.ActiveWindow;
                selection = app.Selection as Excel.Range;
                if (window == null || selection == null) return String.Empty;
                pane = window.ActivePane;
                visible = pane == null ? null : pane.VisibleRange;
                if (pane == null || visible == null) return String.Empty;
                return BuildStateSignature(window, pane, visible, selection);
            } catch { return String.Empty; }
            finally { Release(selection); Release(visible); Release(pane); Release(window); }
        }

        private static string BuildStateSignature(Excel.Window window, Excel.Pane pane, Excel.Range visible, Excel.Range selection) {
            // VisibleRange.Address is cell-granular. During Excel smooth scrolling or
            // a window drag the physical grid origin can move while that address stays
            // unchanged, so the polling loop must also track the real viewport pixels.
            int screenOriginX = window.PointsToScreenPixelsX(0);
            int screenOriginY = window.PointsToScreenPixelsY(0);
            Rectangle excel7;
            if (!FocusNativeMethods.TryGetExcel7Bounds(new IntPtr(window.Hwnd), out excel7)) {
                excel7 = Rectangle.Empty;
            }
            return window.Hwnd + "|" + pane.Index + "|" + window.Zoom + "|" + window.SplitColumn + "|" +
                window.SplitRow + "|" + window.SplitHorizontal + "|" + window.SplitVertical + "|" +
                window.FreezePanes + "|" + visible.Address[false, false, Excel.XlReferenceStyle.xlA1, true, Type.Missing] + "|" +
                selection.Address[false, false, Excel.XlReferenceStyle.xlA1, true, Type.Missing] + "|" +
                screenOriginX + "|" + screenOriginY + "|" + excel7.Left + "|" + excel7.Top + "|" +
                excel7.Right + "|" + excel7.Bottom;
        }

        private static void GetPanePosition(Excel.Window window, int index, out bool right, out bool bottom) {
            right = false;
            bottom = false;
            bool columns = window.SplitColumn > 0 || Convert.ToDouble(window.SplitHorizontal) > 0;
            bool rows = window.SplitRow > 0 || Convert.ToDouble(window.SplitVertical) > 0;
            if (columns && rows) { right = index == 2 || index == 4; bottom = index == 3 || index == 4; }
            else if (columns) right = index == 2;
            else if (rows) bottom = index == 2;
        }

        private static int Round(double value) { return value >= 0 ? (int)Math.Floor(value + 0.5) : -(int)Math.Floor(-value + 0.5); }
        private static void Release(object value) { if (value != null && Marshal.IsComObject(value)) try { Marshal.FinalReleaseComObject(value); } catch { } }
    }

    internal static class FocusNativeMethods {
        internal delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr parameter);
        [DllImport("user32.dll")] internal static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder value, int maximum);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd, out NativeRect rect);
        [DllImport("user32.dll")] internal static extern int GetDpiForWindow(IntPtr hwnd);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)] internal static extern IntPtr SetWindowLongPtr64(IntPtr hwnd, int index, IntPtr value);
        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)] internal static extern IntPtr SetWindowLong32(IntPtr hwnd, int index, IntPtr value);
        [StructLayout(LayoutKind.Sequential)] private struct NativeRect { public int Left, Top, Right, Bottom; }

        internal static IntPtr SetOwner(IntPtr hwnd, IntPtr owner) { return IntPtr.Size == 8 ? SetWindowLongPtr64(hwnd, -8, owner) : SetWindowLong32(hwnd, -8, owner); }
        internal static bool TryGetExcel7Bounds(IntPtr parent, out Rectangle bounds) {
            IntPtr found = IntPtr.Zero;
            EnumChildWindows(parent, delegate(IntPtr hwnd, IntPtr parameter) {
                var name = new StringBuilder(64);
                GetClassName(hwnd, name, name.Capacity);
                if (String.Equals(name.ToString(), "EXCEL7", StringComparison.Ordinal)) { found = hwnd; return false; }
                return true;
            }, IntPtr.Zero);
            NativeRect rect;
            if (found == IntPtr.Zero || !GetWindowRect(found, out rect)) { bounds = Rectangle.Empty; return false; }
            bounds = Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom);
            return bounds.Width > 0 && bounds.Height > 0;
        }
    }
}
