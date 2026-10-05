using System;
using System.Runtime.InteropServices;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using FormsTimer = System.Windows.Forms.Timer;

namespace LH.NxHost {
    internal sealed class ExcelEventBroker : IDisposable {
        private Excel.Application excel;
        private readonly Action refresh;
        private readonly FormsTimer timer;
        private int pending;
        private bool disposed;
        private bool focusPolling;

        public ExcelEventBroker(Excel.Application application, Action callback) {
            excel = application;
            refresh = callback;
            timer = new FormsTimer { Interval = 75 };
            timer.Tick += OnTimer;
            Hook();
        }

        private void Hook() {
            if (excel == null) return;
            excel.SheetSelectionChange += OnSheetSelectionChange;
            excel.SheetActivate += OnSheetActivate;
            excel.WindowActivate += OnWindowActivate;
            excel.WindowDeactivate += OnWindowDeactivate;
            excel.WindowResize += OnWindowResize;
            excel.WorkbookBeforeSave += OnWorkbookBeforeSave;
            excel.WorkbookBeforeClose += OnWorkbookBeforeClose;
        }

        // Excel connection-point callbacks own these borrowed COM event arguments.
        // Do not cache or FinalRelease them here: another subscriber can share the same RCW.
        private void OnSheetSelectionChange(object sh, Excel.Range target) { Queue(); }
        private void OnSheetActivate(object sh) { Queue(); }
        private void OnWindowActivate(Excel.Workbook wb, Excel.Window wn) { Queue(); }
        private void OnWindowDeactivate(Excel.Workbook wb, Excel.Window wn) { NxHostRuntime.HideFocus(); Queue(); }
        private void OnWindowResize(Excel.Workbook wb, Excel.Window wn) { Queue(); }
        private void OnWorkbookBeforeSave(Excel.Workbook wb, bool saveAsUi, ref bool cancel) { NxHostRuntime.HideFocus(); Queue(); }
        private void OnWorkbookBeforeClose(Excel.Workbook wb, ref bool cancel) { NxHostRuntime.HideFocus(); Queue(); }
        private void Queue() { if (!disposed) { Interlocked.Exchange(ref pending, 1); if (!timer.Enabled) timer.Start(); } }

        internal void SetFocusPolling(bool enabled) {
            if (disposed) return;
            focusPolling = enabled;
            if (enabled && !timer.Enabled) timer.Start();
            if (!enabled && Interlocked.CompareExchange(ref pending, 0, 0) == 0) timer.Stop();
        }

        private void OnTimer(object sender, EventArgs e) {
            if (disposed) return;
            // Navigation receives Excel events even when the independent focus overlay is off.
            if (Interlocked.Exchange(ref pending, 0) != 0 && refresh != null) refresh();
            NxHostRuntime.PollFocus();
            if (!focusPolling && Interlocked.CompareExchange(ref pending, 0, 0) == 0) timer.Stop();
        }

        public void Dispose() {
            if (disposed) return;
            disposed = true;
            if (excel != null) {
                excel.SheetSelectionChange -= OnSheetSelectionChange;
                excel.SheetActivate -= OnSheetActivate;
                excel.WindowActivate -= OnWindowActivate;
                excel.WindowDeactivate -= OnWindowDeactivate;
                excel.WindowResize -= OnWindowResize;
                excel.WorkbookBeforeSave -= OnWorkbookBeforeSave;
                excel.WorkbookBeforeClose -= OnWorkbookBeforeClose;
            }
            timer.Stop();
            timer.Tick -= OnTimer;
            timer.Dispose();
            if (excel != null && Marshal.IsComObject(excel)) try { Marshal.FinalReleaseComObject(excel); } catch { }
            excel = null;
        }
    }
}
