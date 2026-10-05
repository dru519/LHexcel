using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace LH.NxHost {
    [ComVisible(true)]
    [Guid("A13D5ED6-B0F2-43C0-9499-E53E6CE38DC0"), ProgId("LH.NxHost.NavigatorPane"), ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(INavigatorPaneControl))]
    public sealed class NavigatorPane : UserControl, INavigatorPaneControl {
        private readonly TextBox search = new TextBox { Dock = DockStyle.Top };
        private readonly ComboBox category = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly ListBox results = new ListBox { Dock = DockStyle.Fill };
        private readonly Label description = new Label { Dock = DockStyle.Bottom, Height = 56, AutoEllipsis = true };
        private readonly Label requirements = new Label { Dock = DockStyle.Bottom, Height = 22, AutoEllipsis = true };
        private readonly Label status = new Label { Dock = DockStyle.Bottom, Height = 34, AutoEllipsis = true, Text = "0개 항목" };
        private readonly ToolTip tips = new ToolTip();
        private INavigatorRequestHandler handler;
        private bool refreshQueued;
        private DocumentNavigatorControl documents;

        public NavigatorPane() {
            Dock = DockStyle.Fill;
            Padding = new Padding(6);
            MinimumSize = new Size(240, 300);
            Font = new Font("맑은 고딕", 9F);
            search.AccessibleName = "기능 찾기";
            category.AccessibleName = "기능 분류";
            results.AccessibleName = "기능 목록";
            Controls.Add(results);
            Controls.Add(description);
            Controls.Add(requirements);
            Controls.Add(status);
            Controls.Add(category);
            Controls.Add(search);
            category.Items.Add(new NavigatorCategory("", "전체"));
            foreach (NavigatorCategory item in NxGeneratedNavigatorCatalog.Categories) category.Items.Add(item);
            category.SelectedIndex = 0;
            tips.SetToolTip(search, "기능명, 분류명 또는 ID로 찾습니다.");
            search.TextChanged += SearchChanged;
            search.KeyDown += KeyPressed;
            category.SelectedIndexChanged += CategoryChanged;
            category.KeyDown += KeyPressed;
            results.SelectedIndexChanged += ResultSelected;
            results.DoubleClick += ResultChosen;
            results.KeyDown += KeyPressed;
            Render("");
        }

        public void SetRequestHandler(INavigatorRequestHandler value) {
            handler = value;
            RenderDetails();
        }
        internal void SetDocumentHandler(IDocumentNavigatorHandler value) {
            foreach (Control child in Controls) child.Visible = false;
            Padding = Padding.Empty;
            documents = new DocumentNavigatorControl(value) { Dock = DockStyle.Fill };
            Controls.Add(documents);
            documents.BringToFront();
            documents.RefreshSnapshot();
        }

        internal void EnableHostedDpi() {
            // The CTP is an unmanaged parent, so this control owns its layout scale.
            // The separate Form fallback continues to own scaling via Inherit.
            SuspendLayout();
            MinimumSize = Size.Empty; // A docked pane must fit the Office-owned client area.
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ResumeLayout(true);
        }

        public void RefreshPending() {
            // Availability changes do not change this immutable catalogue.
            // Coalesce Excel events and never erase the user's route/scroll.
            if (IsDisposed || Disposing || !IsHandleCreated || refreshQueued) return;
            refreshQueued = true;
            try {
                BeginInvoke((MethodInvoker)(() => {
                    refreshQueued = false;
                    if (!IsDisposed && !Disposing) {
                        if (documents != null) documents.RefreshSnapshot();
                        else RenderDetails();
                    }
                }));
            } catch (InvalidOperationException) { refreshQueued = false; }
        }

        private void SearchChanged(object sender, EventArgs e) { Render(search.Text); }
        private void CategoryChanged(object sender, EventArgs e) { Render(search.Text); }
        private void ResultSelected(object sender, EventArgs e) { RenderDetails(); }

        private void Render(string term) {
            string query = (term ?? "").Trim();
            NavigatorItem prior = results.SelectedItem as NavigatorItem;
            string priorKey = prior == null ? "" : prior.RouteKey;
            int priorTop = results.Items.Count == 0 ? 0 : results.TopIndex;
            NavigatorCategory selected = category.SelectedItem as NavigatorCategory;
            results.BeginUpdate();
            try {
                results.Items.Clear();
                foreach (NavigatorItem item in NxGeneratedNavigatorCatalog.Items) {
                    if (query.Length > 0 && item.SearchText.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) < 0) continue;
                    if (selected != null && selected.Id.Length > 0 && item.CategoryId != selected.Id) continue;
                    results.Items.Add(item);
                    if (item.RouteKey == priorKey) results.SelectedIndex = results.Items.Count - 1;
                }
                if (results.Items.Count > 0) results.TopIndex = Math.Min(priorTop, results.Items.Count - 1);
            } finally { results.EndUpdate(); }
            description.Text = "기능을 선택하면 설명이 표시됩니다.";
            requirements.Text = "요구 입력: -";
            status.Text = results.Items.Count + "개 항목";
            RenderDetails();
        }

        private void RenderDetails() {
            NavigatorItem item = results.SelectedItem as NavigatorItem;
            if (item == null) return;
            description.Text = item.Description;
            tips.SetToolTip(description, item.Description);
            requirements.Text = NxGeneratedNavigatorCatalog.DisplayText(item.InputContext) + " · " + NxGeneratedNavigatorCatalog.DisplayText(item.ExecutionGrade);
            tips.SetToolTip(requirements, requirements.Text);
            string availability = "";
            try { if (handler != null) availability = handler.GetAvailability(item.RouteKey) ?? ""; }
            catch (COMException) { availability = ""; }
            catch (InvalidOperationException) { availability = ""; }
            if (availability.Length == 0) availability = "unavailable|현재 Excel 상태를 확인할 수 없습니다.";
            string[] parts = availability.Split(new[] { '|' }, 2);
            string caption = parts[0] == "available" ? "실행 가능" : parts[0] == "needs_input" ? "입력 필요" : "실행 불가";
            status.Text = parts.Length > 1 && parts[1].Length > 0 ? caption + " · " + parts[1] : caption;
        }

        private void ResultChosen(object sender, EventArgs e) { ExecuteSelected(); }

        private void ExecuteSelected() {
            NavigatorItem item = results.SelectedItem as NavigatorItem;
            if (item == null || handler == null) return;
            string availability = handler.GetAvailability(item.RouteKey);
            if (availability.StartsWith("unavailable|", StringComparison.Ordinal)) {
                RenderDetails();
                return;
            }
            status.Text = handler.ExecuteRequest(item.FeatureId, item.CommandId, "show=1")
                ? "실행 요청을 전달했습니다."
                : "실행할 수 없습니다.";
        }

        private void KeyPressed(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Enter) {
                ExecuteSelected();
                e.Handled = true;
                e.SuppressKeyPress = true;
            } else if (e.KeyCode == Keys.Escape) {
                if (handler != null) handler.CloseNavigator();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                if (documents != null) { documents.Dispose(); documents = null; }
                search.TextChanged -= SearchChanged;
                search.KeyDown -= KeyPressed;
                category.SelectedIndexChanged -= CategoryChanged;
                category.KeyDown -= KeyPressed;
                results.SelectedIndexChanged -= ResultSelected;
                results.DoubleClick -= ResultChosen;
                results.KeyDown -= KeyPressed;
                tips.Dispose();
            }
            handler = null;
            refreshQueued = false;
            base.Dispose(disposing);
        }
    }

    // Office requires IDispatch when retrieving the hosted control. Expose no
    // general Control members: execution still uses the closed request handler.
    [ComVisible(true), Guid("6E5992D2-0EC7-4553-8B3A-0B449A0B67EF"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface INavigatorPaneControl { }

    public interface INavigatorRequestHandler {
        bool ExecuteRequest(string featureId, string commandId, string payload);
        string GetAvailability(string routeKey);
        void CloseNavigator();
    }

    internal sealed class NavigatorCategory {
        internal readonly string Id;
        internal readonly string Title;
        internal NavigatorCategory(string id, string title) { Id = id; Title = title; }
        public override string ToString() { return Title; }
    }

    internal sealed class NavigatorItem {
        internal readonly string ItemType;
        internal readonly string Id;
        internal readonly string Title;
        internal readonly string Description;
        internal readonly string CategoryId;
        internal readonly string SearchText;
        internal readonly string InputContext;
        internal readonly string SelectionMode;
        internal readonly string LaunchSurface;
        internal readonly string MutationScope;
        internal readonly string FeedbackMode;
        internal readonly string ExecutionGrade;
        internal string RouteKey { get { return ItemType + ":" + Id; } }
        internal string FeatureId { get { return ItemType == "feature" ? Id : ""; } }
        internal string CommandId { get { return ItemType == "command" ? Id : ""; } }
        internal NavigatorItem(string itemType, string id, string title, string description, string categoryId, string searchText,
            string inputContext, string selectionMode, string launchSurface, string mutationScope, string feedbackMode, string executionGrade) {
            ItemType = itemType;
            Id = id;
            Title = title;
            Description = description;
            CategoryId = categoryId;
            SearchText = searchText;
            InputContext = inputContext;
            SelectionMode = selectionMode;
            LaunchSurface = launchSurface;
            MutationScope = mutationScope;
            FeedbackMode = feedbackMode;
            ExecutionGrade = executionGrade;
        }
        public override string ToString() { return Title; }
    }
}
