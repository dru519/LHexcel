using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Extensibility;
using Microsoft.Office.Core;

namespace LH.NxHost {
    [ComVisible(true), Guid("4A4B4F02-76A6-4F22-BD4B-85E0308EC91D"), ProgId("LH.NxHost.Connect"), ClassInterface(ClassInterfaceType.None)]
    public sealed class NxHostAddIn : IDTExtensibility2, ICustomTaskPaneConsumer {
        private Excel.Application application;
        private ICTPFactory ctpFactory;
        private ExcelEventBroker events;
        private readonly Dictionary<int, NavigatorSession> sessions = new Dictionary<int, NavigatorSession>();
        private readonly Dictionary<int, NavigatorSession> documentSessions = new Dictionary<int, NavigatorSession>();
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom) {
            application = Application as Excel.Application;
            NxHostRuntime.Attach(this);
            if (application != null) {
                events = new ExcelEventBroker(application, OnExcelRefresh);
                NxHostRuntime.SetFocusPolling(events.SetFocusPolling);
            }
        }
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom) { Dispose(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { Dispose(); }
        public void CTPFactoryAvailable(ICTPFactory CTPFactoryInst) { ctpFactory = CTPFactoryInst; }
        public bool CreatePane(bool documents = false) {
            if (application == null) return false;
            PruneClosedSessions();
            int key = application.Hwnd;
            var panes = documents ? documentSessions : sessions;
            NavigatorSession session;
            if (!panes.TryGetValue(key, out session)) {
                Excel.Window owner = application.ActiveWindow;
                if (owner == null) return false;
                session = new NavigatorSession(application, owner, key, ctpFactory, documents);
                panes.Add(key, session);
            }
            if (session.Show()) return true;
            panes.Remove(key);
            session.Dispose();
            return false;
        }
        public bool HidePane(bool documents = false) {
            if (application == null) return true;
            NavigatorSession session;
            return !(documents ? documentSessions : sessions).TryGetValue(application.Hwnd, out session) || session.Hide();
        }
        private void PruneClosedSessions() {
            foreach (var panes in new[] { sessions, documentSessions })
            foreach (int key in new List<int>(panes.Keys)) {
                if (panes[key].IsAlive) continue;
                NavigatorSession session = panes[key];
                panes.Remove(key);
                session.Dispose();
            }
        }
        private void OnExcelRefresh() {
            NxHostRuntime.RefreshFocus();
            PruneClosedSessions();
            if (application == null) return;
            try {
                NavigatorSession session;
                if (sessions.TryGetValue(application.Hwnd, out session)) session.Refresh();
                if (documentSessions.TryGetValue(application.Hwnd, out session)) session.Refresh();
            } catch (COMException) { }
        }
        private void Dispose() {
            WorkbookCompareResultsService.CloseActive();
            NxHostRuntime.SetFocusPolling(null);
            if (events != null) { events.Dispose(); events = null; }
            foreach (NavigatorSession session in sessions.Values) session.Dispose();
            sessions.Clear();
            foreach (NavigatorSession session in documentSessions.Values) session.Dispose();
            documentSessions.Clear();
            NxHostRuntime.Detach(this);
            ReleaseCom(ctpFactory); ctpFactory = null;
            ReleaseCom(application); application = null;
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
        }
        internal static void ReleaseCom(object value) {
            if (value != null && Marshal.IsComObject(value)) try { Marshal.FinalReleaseComObject(value); } catch { }
        }
        internal Excel.Application Application { get { return application; } }
    }

    // Each Excel SDI window owns exactly one UI and its search/selection state.
    internal sealed class NavigatorSession : IDisposable, INavigatorRequestHandler, IDocumentNavigatorHandler {
        private const string NavigatorProgId = "LH.NxHost.NavigatorPane";
        private readonly string NavigatorTitle;
        private readonly bool documents;
        private int expandedWidth;
        private Excel.Application application; // Borrowed from the add-in; never released here.
        private Excel.Window activeWindow;
        private readonly ICTPFactory ctpFactory; // Borrowed; add-in releases after all sessions.
        private readonly IntPtr ownerHandle;
        private NavigatorPane pane;
        private CustomTaskPane customTaskPane;
        private Form navigatorWindow;
        private bool disposing;
        internal NavigatorSession(Excel.Application app, Excel.Window owner, int hwnd, ICTPFactory factory, bool documentMode = false) {
            application = app; activeWindow = owner; ownerHandle = new IntPtr(hwnd); ctpFactory = factory;
            documents = documentMode;
            NavigatorTitle = documents ? "내엑셀 탐색창" : "내엑셀 Navigator";
            TracePane(factory == null ? "factory_missing" : "factory_available", null);
        }
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hwnd);
        internal bool IsAlive { get { return !disposing && IsWindow(ownerHandle); } }
        internal bool Show() {
            if (application == null || !IsAlive) return false;
            if (customTaskPane != null) { customTaskPane.Visible = true; Refresh(); TracePane("ctp_reshown", null); return true; }
            if (navigatorWindow != null && !navigatorWindow.IsDisposed) {
                navigatorWindow.Show();
                PlaceNavigatorWindow(false);
                navigatorWindow.BringToFront();
                return navigatorWindow.Visible;
            }
            if (ctpFactory != null) {
                CustomTaskPane candidate = null;
                NavigatorPane content = null;
                try {
                    TracePane("create_ctp", null);
                    candidate = ctpFactory.CreateCTP(NavigatorProgId, NavigatorTitle, activeWindow);
                    TracePane("content_control", null);
                    object candidateControl = candidate.ContentControl;
                    content = candidateControl as NavigatorPane;
                    // A verified bootstrap hosts the core pane as its direct child.
                    var bootstrap = candidateControl as UserControl;
                    if (content == null && bootstrap != null) {
                        foreach (Control child in bootstrap.Controls) {
                            content = child as NavigatorPane;
                            if (content != null) break;
                        }
                    }
                    if (content != null) {
                        content.EnableHostedDpi();
                        candidate.DockPosition = documents ? MsoCTPDockPosition.msoCTPDockPositionLeft : MsoCTPDockPosition.msoCTPDockPositionRight;
                        // Raw Office.Core ICTP uses physical pixels (not VSTO's point wrapper).
                        candidate.Width = (int)((documents ? 216 : 300) * GetDpiForWindow(ownerHandle) / 96);
                        if (documents) content.SetDocumentHandler(this);
                        else content.SetRequestHandler(this);
                        candidate.Visible = true;
                        pane = content;
                        customTaskPane = candidate;
                        candidate = null; // Transfer ownership only after successful display.
                        TracePane("ctp_shown", null);
                        return true;
                    }
                    TracePane("content_type_mismatch", null);
                } catch (Exception error) { TracePane("ctp_failed", error); }
                finally {
                    if (candidate != null) {
                        try { candidate.Delete(); } catch { }
                        ReleaseCom(candidate);
                        if (content != null) content.Dispose();
                    }
                }
            }
            if (ctpFactory == null) TracePane("factory_missing", null);
            // Document mode falls back to the established VBA navigator, not another DLL window.
            return !documents && CreateFallbackWindow();
        }
        internal bool Hide() {
            if (customTaskPane != null) {
                customTaskPane.Visible = false;
                return !customTaskPane.Visible;
            }
            if (navigatorWindow == null || navigatorWindow.IsDisposed) return true;
            navigatorWindow.Hide();
            return !navigatorWindow.Visible;
        }
        private bool CreateFallbackWindow() {
            IntPtr excelHandle = ownerHandle;
            Form window = new Form();
            window.SuspendLayout();
            window.AutoScaleDimensions = new SizeF(96F, 96F);
            window.AutoScaleMode = AutoScaleMode.Dpi;
            window.Font = new Font("맑은 고딕", 9F);
            window.Text = NavigatorTitle;
            window.ShowInTaskbar = false;
            window.StartPosition = FormStartPosition.Manual;
            window.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            window.MinimizeBox = false;
            window.MaximizeBox = false;
            window.ClientSize = new Size(300, 460);
            window.MinimumSize = new Size(260, 340);
            NavigatorPane content = new NavigatorPane();
            /* The fallback form owns DPI scaling; the child inherits that scale. */
            content.AutoScaleMode = AutoScaleMode.Inherit;
            content.Dock = DockStyle.Fill;
            content.SetRequestHandler(this);
            window.Controls.Add(content);
            window.FormClosing += NavigatorWindowClosing;
            window.Shown += NavigatorWindowShown;
            pane = content;
            navigatorWindow = window;
            try {
                window.ResumeLayout(true);
                window.Show(new ExcelWindowOwner(excelHandle));
                TracePane("fallback_shown", null);
                return window.Visible;
            } catch {
                window.FormClosing -= NavigatorWindowClosing;
                window.Shown -= NavigatorWindowShown;
                navigatorWindow = null;
                pane = null;
                window.Dispose();
                return false;
            }
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct WindowRect { public int Left, Top, Right, Bottom; }
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out WindowRect rectangle);
        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);
        private void NavigatorWindowShown(object sender, EventArgs e) { PlaceNavigatorWindow(true); }
        private void PlaceNavigatorWindow(bool alignRight) {
            if (application == null || navigatorWindow == null || !navigatorWindow.Visible) return;
            try {
            IntPtr owner = ownerHandle;
            Rectangle area = Screen.FromHandle(owner).WorkingArea;
            WindowRect rectangle;
            if (GetWindowRect(owner, out rectangle))
                area = Rectangle.Intersect(area, Rectangle.FromLTRB(rectangle.Left, rectangle.Top, rectangle.Right, rectangle.Bottom));
            if (area.Width <= 0 || area.Height <= 0) return;
            float scale = navigatorWindow.CurrentAutoScaleDimensions.Width / 96F;
            navigatorWindow.MinimumSize = new Size(Math.Min(area.Width, (int)(260 * scale)), Math.Min(area.Height, (int)(340 * scale)));
            navigatorWindow.Size = new Size(Math.Min(area.Width, navigatorWindow.Width), Math.Min(area.Height, navigatorWindow.Height));
            int left = alignRight ? area.Right - navigatorWindow.Width - (int)(16 * scale) : navigatorWindow.Left;
            int top = alignRight ? area.Top + (int)(72 * scale) : navigatorWindow.Top;
            navigatorWindow.Location = new Point(Math.Max(area.Left, Math.Min(area.Right - navigatorWindow.Width, left)), Math.Max(area.Top, Math.Min(area.Bottom - navigatorWindow.Height, top)));
            } catch (COMException error) { TracePane("placement_failed", error); }
        }
        internal void TracePane(string stage, Exception error) {
            // Opt-in test diagnostics only. Never log document content or exception messages.
            if (Environment.GetEnvironmentVariable("LHEXCEL_NXHOST_DIAGNOSTICS") != "1") return;
            try {
                string root = Environment.GetEnvironmentVariable("LHEXCEL_PROFILE_ROOT");
                if (string.IsNullOrEmpty(root) || !System.IO.Path.IsPathRooted(root) || !System.IO.Directory.Exists(root)) return;
                string path = System.IO.Path.Combine(root, "navigator-diagnostics.txt");
                if (System.IO.File.Exists(path) && new System.IO.FileInfo(path).Length > 16384) return;
                string dimensions = navigatorWindow == null ? "" : "|client=" + navigatorWindow.ClientSize + "|scale=" + navigatorWindow.AutoScaleDimensions + "|current=" + navigatorWindow.CurrentAutoScaleDimensions;
                if (customTaskPane != null) dimensions += "|ctp_dock=" + customTaskPane.DockPosition + "|ctp_width=" + customTaskPane.Width + "|visible=" + customTaskPane.Visible;
                System.IO.File.AppendAllText(path, stage + "|hr=" + (error == null ? "0" : error.HResult.ToString("X8")) + dimensions + Environment.NewLine);
            } catch (System.IO.IOException) { } catch (UnauthorizedAccessException) { } catch (ArgumentException) { } catch (COMException) { }
        }
        private void NavigatorWindowClosing(object sender, FormClosingEventArgs e) {
            if (!disposing && e.CloseReason == CloseReason.UserClosing) {
                e.Cancel = true;
                if (navigatorWindow != null) navigatorWindow.Hide();
            }
        }
        internal void Refresh() { if (pane != null) pane.RefreshPending(); PlaceNavigatorWindow(false); }
        public Array GetDocuments() {
            if (!IsAlive || application == null) return null;
            try { return new IntPtr(application.Hwnd) == ownerHandle ? BridgeService.GetDocuments() : null; }
            catch (COMException) { return null; }
        }
        public bool MoveDocument(string token) {
            if (!IsAlive || application == null) return false;
            try { return new IntPtr(application.Hwnd) == ownerHandle && BridgeService.MoveDocument(token); }
            catch (COMException) { return false; }
        }
        public void SetCollapsed(bool collapsed) {
            if (customTaskPane == null) return;
            try {
                int scale = (int)GetDpiForWindow(ownerHandle);
                if (collapsed) { expandedWidth = customTaskPane.Width; customTaskPane.Width = 32 * scale / 96; }
                else customTaskPane.Width = Math.Max(196 * scale / 96, expandedWidth);
                TracePane(collapsed ? "document_collapsed" : "document_expanded", null);
            } catch (COMException) { }
        }
        public bool ExecuteRequest(string featureId, string commandId, string payload) {
            if (!IsAlive || activeWindow == null) return false;
            try {
                activeWindow.Activate();
                if (new IntPtr(application.Hwnd) != ownerHandle) return false;
                return BridgeService.ExecuteReverse(featureId, commandId, payload);
            } catch (COMException) { return false; }
        }
        public string GetAvailability(string routeKey) {
            if (!IsAlive) return "unavailable|Excel 창이 닫혔습니다.";
            try {
                if (new IntPtr(application.Hwnd) != ownerHandle) return "unavailable|해당 Excel 창을 먼저 활성화하세요.";
                return BridgeService.GetReverseAvailability(routeKey);
            } catch (COMException) { return "unavailable|Excel 상태를 확인할 수 없습니다."; }
        }
        public void CloseNavigator() { Hide(); }
        public void Dispose() {
            if (disposing) return;
            disposing = true;
            if (customTaskPane != null) {
                try { customTaskPane.Visible = false; customTaskPane.Delete(); } catch (COMException) { }
                finally { ReleaseCom(customTaskPane); customTaskPane = null; }
            }
            if (navigatorWindow != null) {
                navigatorWindow.FormClosing -= NavigatorWindowClosing;
                navigatorWindow.Shown -= NavigatorWindowShown;
                navigatorWindow.Close(); navigatorWindow.Dispose(); navigatorWindow = null;
            }
            if (pane != null) { pane.Dispose(); pane = null; }
            ReleaseCom(activeWindow); activeWindow = null;
            application = null;
        }
        private static void ReleaseCom(object value) { NxHostAddIn.ReleaseCom(value); }
    }
    internal sealed class ExcelWindowOwner : IWin32Window {
        private readonly IntPtr handle;
        internal ExcelWindowOwner(IntPtr value) { handle = value; }
        public IntPtr Handle { get { return handle; } }
    }
}
