using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace LH.NxHost {
    internal interface IDocumentNavigatorHandler {
        Array GetDocuments();
        bool MoveDocument(string token);
        void SetCollapsed(bool collapsed);
    }
    internal sealed class DocumentEntry {
        internal string Token, Parent, Title;
        internal bool Available, Active;
        public override string ToString() { return Title + (Active ? " (현재)" : "") + (Available ? "" : " (숨김)"); }
    }
    // No Excel objects, file paths, values or formulas are retained in this UI.
    internal sealed class DocumentNavigatorControl : UserControl {
        private readonly IDocumentNavigatorHandler handler;
        private readonly TextBox bookSearch = new TextBox { Dock = DockStyle.Top, AccessibleName = "문서 찾기" };
        private readonly TextBox sheetSearch = new TextBox { Dock = DockStyle.Top, AccessibleName = "시트 찾기" };
        private readonly ListBox books = NewList("문서 목록"), sheets = NewList("시트 목록");
        private readonly Button bookClear = new Button(), sheetClear = new Button();
        private readonly Label bookHeading = new Label(), sheetHeading = new Label();
        private readonly SplitContainer split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterWidth = 6, TabStop = true };
        private readonly ToolStrip tools = new ToolStrip { Dock = DockStyle.Top, GripStyle = ToolStripGripStyle.Hidden, RenderMode = ToolStripRenderMode.System, AccessibleName = "문서 탐색 도구" };
        private readonly Label status = new Label { Dock = DockStyle.Bottom, Height = 32, AutoEllipsis = true };
        private readonly Label noBooks = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Text = "열린 문서가 없습니다" };
        private readonly Label noSheets = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Text = "표시할 시트가 없습니다" };
        private readonly Button rail = new Button { Dock = DockStyle.Fill, Text = "›", AccessibleName = "탐색창 펼치기", Visible = false };
        private readonly ToolStripButton pin = new ToolStripButton("고정") { CheckOnClick = true, Checked = true };
        private readonly ToolTip tips = new ToolTip();
        private readonly List<Image> toolbarImages = new List<Image>();
        private List<DocumentEntry> snapshot = new List<DocumentEntry>();
        private string selectedBook = "", selectedSheet = "";
        private bool rendering, collapsed, settingSplit;
        private double ratio = 0.5;
        internal DocumentNavigatorControl(IDocumentNavigatorHandler value) {
            handler = value;
            Font = new Font("맑은 고딕", 9F);
            BackColor = SystemColors.Window;
            status.Height = Font.Height + 10;
            status.TextAlign = ContentAlignment.MiddleLeft;
            status.Padding = new Padding(4,0,4,0);
            split.SplitterWidth = Math.Max(6, Font.Height / 2);
            var refresh = new ToolStripButton("새로 고침") { AccessibleName = "새로 고침", ToolTipText = "열린 문서와 시트 목록 새로 고침" };
            var fold = new ToolStripButton("접기") { AccessibleName = "탐색창 접기", ToolTipText = "탐색창 접기 (Esc)", Alignment = ToolStripItemAlignment.Right };
            SetToolbarImage(refresh,"refresh"); SetToolbarImage(pin,"pin"); SetToolbarImage(fold,"fold");
            pin.AccessibleName = "고정"; pin.Alignment = ToolStripItemAlignment.Right;
            pin.ToolTipText = "고정됨 · 클릭하면 이동 후 자동으로 접힙니다";
            tools.Items.AddRange(new ToolStripItem[] { refresh, pin, fold });
            pin.CheckedChanged += (s,e) => {
                pin.ToolTipText = pin.Checked ? "고정됨 · 클릭하면 이동 후 자동으로 접힙니다" : "자동 접기 · 클릭하면 탐색창을 고정합니다";
                Trace(pin.Checked ? "pinned" : "unpinned");
            };
            FillPanel(split.Panel1, books, bookSearch, bookClear, bookHeading, "통합문서");
            FillPanel(split.Panel2, sheets, sheetSearch, sheetClear, sheetHeading, "시트");
            split.Panel1.Controls.Add(noBooks); noBooks.BringToFront();
            split.Panel2.Controls.Add(noSheets); noSheets.BringToFront();
            Controls.Add(split); Controls.Add(status); Controls.Add(tools); Controls.Add(rail);
            refresh.Click += (s,e) => RefreshSnapshot(); fold.Click += (s,e) => Collapse(true);
            rail.Click += (s,e) => Collapse(false);
            bookSearch.TextChanged += (s,e) => RenderBooks();
            sheetSearch.TextChanged += (s,e) => RenderSheets();
            books.SelectedIndexChanged += (s,e) => {
                if (rendering) return;
                var item = books.SelectedItem as DocumentEntry;
                if (item == null || item.Token == selectedBook) return;
                selectedBook = item.Token; selectedSheet = ""; sheetSearch.Clear(); RenderSheets();
            };
            sheets.SelectedIndexChanged += (s,e) => {
                if (!rendering && sheets.SelectedItem is DocumentEntry) selectedSheet = ((DocumentEntry)sheets.SelectedItem).Token;
            };
            books.DoubleClick += (s,e) => Activate(books); sheets.DoubleClick += (s,e) => Activate(sheets);
            books.KeyDown += ListKey; sheets.KeyDown += ListKey;
            bookSearch.KeyDown += (s,e) => SearchKey(books,e);
            sheetSearch.KeyDown += (s,e) => SearchKey(sheets,e); tools.KeyDown += EscapeKey;
            books.MouseMove += ShowTip; sheets.MouseMove += ShowTip;
            split.SplitterMoved += (s,e) => {
                if (settingSplit || split.Height <= split.SplitterWidth) return;
                ratio = Math.Max(.25, Math.Min(.75, (double)split.SplitterDistance / (split.Height - split.SplitterWidth)));
                PlaceSplitter();
            };
            split.DoubleClick += (s,e) => { ratio = .5; PlaceSplitter(); };
            split.KeyDown += (s,e) => {
                if (e.KeyCode == Keys.Home) ratio = .5;
                else if (e.KeyCode == Keys.Up) ratio -= .05;
                else if (e.KeyCode == Keys.Down) ratio += .05;
                else return;
                e.Handled = e.SuppressKeyPress = true; PlaceSplitter();
            };
            split.SizeChanged += (s,e) => PlaceSplitter();
            // Excel owns the message loop: WinForms message filters are not a reliable
            // input hook in a CTP. Use control events and the native list wheel instead.
            Leave += (s,e) => {
                if (!pin.Checked && IsHandleCreated) BeginInvoke((MethodInvoker)(() => {
                    if (!IsDisposed && !ContainsFocus && !pin.Checked) Collapse(true);
                }));
            };
        }
        private static ListBox NewList(string title) {
            var list = new ListBox { Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, IntegralHeight = false, HorizontalScrollbar = false, AccessibleName = title, DrawMode = DrawMode.OwnerDrawFixed };
            list.FontChanged += (s,e) => list.ItemHeight = list.Font.Height + Math.Max(8,list.Font.Height/3);
            list.DrawItem += (s,e) => {
                if (e.Index < 0) return;
                var item = list.Items[e.Index] as DocumentEntry;
                if (item == null) return;
                bool selected = (e.State & DrawItemState.Selected) != 0;
                bool contrast = SystemInformation.HighContrast;
                Color background = selected ? (contrast ? SystemColors.Highlight : Color.FromArgb(218,239,225)) : SystemColors.Window;
                using(var brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush,e.Bounds);
                Color color = !item.Available ? SystemColors.GrayText : selected && contrast ? SystemColors.HighlightText : SystemColors.WindowText;
                int inset = Math.Max(5,e.Font.Height/4);
                if(item.Active) using(var brush = new SolidBrush(contrast ? color : Color.FromArgb(33,115,70)))
                    e.Graphics.FillRectangle(brush,e.Bounds.X,e.Bounds.Y+3,Math.Max(3,inset/2),e.Bounds.Height-6);
                var text = new Rectangle(e.Bounds.X+inset,e.Bounds.Y,e.Bounds.Width-inset*2,e.Bounds.Height);
                TextRenderer.DrawText(e.Graphics, item.Title+(item.Available ? "" : " (숨김)"), e.Font, text, color, TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
                e.DrawFocusRectangle();
            };
            return list;
        }
        private void FillPanel(Control panel, ListBox list, TextBox search, Button clear, Label heading, string title) {
            int gap = Math.Max(4,Font.Height/4);
            search.Font = Font;
            panel.Padding = new Padding(gap,0,gap,gap);
            heading.Text = title; heading.Dock = DockStyle.Top; heading.Height = Font.Height+gap*2;
            heading.TextAlign = ContentAlignment.MiddleLeft;
            var searchRow = new Panel { Dock=DockStyle.Top, Height=search.PreferredHeight+gap, Padding=new Padding(0,0,0,gap) };
            search.Dock = DockStyle.Fill;
            search.AccessibleDescription = title+" 이름으로 검색합니다. 아래쪽 화살표로 목록에 이동합니다.";
            search.HandleCreated += (s,e) => SendMessage(search.Handle,0x1501,new IntPtr(1),title+" 검색");
            clear.Text = "×"; clear.AccessibleName = title+" 검색 지우기";
            clear.Dock = DockStyle.Right; clear.Width = Math.Max(24,Font.Height+gap); clear.FlatStyle = FlatStyle.Flat;
            clear.FlatAppearance.BorderSize = 0; clear.Enabled = false;
            clear.Click += (s,e) => { search.Clear(); search.Focus(); };
            clear.KeyDown += EscapeKey;
            search.TextChanged += (s,e) => clear.Enabled = search.Text.Length>0;
            tips.SetToolTip(clear,title+" 검색 지우기"); tips.SetToolTip(search,search.AccessibleDescription);
            searchRow.Controls.Add(search); searchRow.Controls.Add(clear);
            panel.Controls.Add(list); panel.Controls.Add(searchRow); panel.Controls.Add(heading);
        }
        private void SetToolbarImage(ToolStripButton button,string kind) {
            var image = new Bitmap(48,48);
            string theme=SystemColors.Control.GetBrightness()<0.5F ? "reverse" : "ink";
            string asset=kind=="refresh" ? "refresh" : kind=="fold" ? "chevron" : null;
            bool loaded=false;
            if(asset!=null) using(var stream=typeof(DocumentNavigatorControl).Assembly.GetManifestResourceStream("NxHost.Brand."+asset+"-"+theme+".png")) {
                if(stream!=null) using(var source=Image.FromStream(stream)) {
                    image.Dispose(); image=new Bitmap(source); loaded=true;
                    if(kind=="fold") image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                }
            }
            if(!loaded)
            using(var graphics = Graphics.FromImage(image)) using(var pen = new Pen(SystemColors.ControlText,1.4F)) {
                graphics.SmoothingMode = SmoothingMode.AntiAlias; graphics.ScaleTransform(3,3);
                if(kind=="pin") {
                    graphics.DrawLines(pen,new[]{new PointF(5,2),new PointF(11,2),new PointF(10,7),new PointF(12,10),new PointF(4,10),new PointF(6,7),new PointF(5,2)});
                    graphics.DrawLine(pen,8,10,8,14);
                } else if(kind=="refresh") {
                    graphics.DrawArc(pen,3,3,10,10,35,290);
                    graphics.DrawLines(pen,new[]{new PointF(10,2),new PointF(13,5),new PointF(10,6)});
                } else graphics.DrawLines(pen,new[]{new PointF(10,3),new PointF(5,8),new PointF(10,13)});
            }
            var prior=button.Image;
            toolbarImages.Add(image); button.Image=image; button.Tag=kind; button.DisplayStyle=ToolStripItemDisplayStyle.Image;
            if(prior!=null && toolbarImages.Remove(prior)) prior.Dispose();
        }
        protected override void OnSystemColorsChanged(EventArgs e) {
            base.OnSystemColorsChanged(e);
            if(tools==null || toolbarImages==null) return;
            foreach(ToolStripItem item in tools.Items) {
                var button=item as ToolStripButton;
                if(button!=null && button.Tag is string) SetToolbarImage(button,(string)button.Tag);
            }
            books.Invalidate(); sheets.Invalidate();
        }
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern IntPtr SendMessage(IntPtr hwnd,uint message,IntPtr value,string text);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr hwnd);
        protected override void OnHandleCreated(EventArgs e) {
            base.OnHandleCreated(e);
            ApplyToolbarDpi((int)GetDpiForWindow(Handle));
        }
        protected override void WndProc(ref Message message) {
            base.WndProc(ref message);
            if(message.Msg==0x02E3 && IsHandleCreated) ApplyToolbarDpi((int)GetDpiForWindow(Handle));
        }
        private void ApplyToolbarDpi(int dpi) {
            if(tools==null) return;
            // This child is created after Office's unmanaged CTP parent scales.
            // ToolStrip otherwise keeps 16 physical pixels even at 200 percent.
            double scale=Math.Max(96,Math.Min(768,dpi))/96.0;
            int imageEdge=(int)Math.Round(16*scale), buttonEdge=(int)Math.Round(24*scale);
            tools.ImageScalingSize=new Size(imageEdge,imageEdge);
            foreach(ToolStripItem item in tools.Items) {
                item.AutoSize=false; item.Size=new Size(buttonEdge,buttonEdge);
            }
            Trace("toolbar_dpi="+dpi+"|image="+imageEdge+"|button="+buttonEdge+"|font="+Font.Height);
        }
        internal void RefreshSnapshot() {
            Array data = handler.GetDocuments();
            if (data == null) { status.Text = "목록을 새로 고칠 수 없습니다"; return; }
            var next = new List<DocumentEntry>();
            try {
                if (data.Length != 0) {
                    if (data.Rank != 2 || data.GetLength(1) != 5 || data.GetLength(0) > 10000) throw new ArgumentException();
                    int row = data.GetLowerBound(0), column = data.GetLowerBound(1);
                    var tokens = new HashSet<string>(StringComparer.Ordinal);
                    for (int n = 0; n < data.GetLength(0); n++) {
                        var item = new DocumentEntry { Token = Convert.ToString(data.GetValue(row+n,column)),
                            Parent = Convert.ToString(data.GetValue(row+n,column+1)), Title = Convert.ToString(data.GetValue(row+n,column+2)),
                            Available = Convert.ToBoolean(data.GetValue(row+n,column+3)), Active = Convert.ToBoolean(data.GetValue(row+n,column+4)) };
                        int token;
                        if (!Int32.TryParse(item.Token, out token) || token <= 0 || !tokens.Add(item.Token) || item.Title.Length > 255) throw new ArgumentException();
                        next.Add(item);
                    }
                }
            } catch { status.Text = "목록을 새로 고칠 수 없습니다"; return; }
            snapshot = next;
            if (!snapshot.Exists(x => x.Parent == "" && x.Token == selectedBook)) {
                var first = snapshot.Find(x => x.Parent == "" && x.Active) ?? snapshot.Find(x => x.Parent == "");
                selectedBook = first == null ? "" : first.Token; selectedSheet = "";
            }
            RenderBooks(); RenderSheets();
        }
        private void RenderBooks() {
            Fill(books, x => x.Parent == "" && Matches(x, bookSearch.Text), selectedBook);
            bookHeading.Text = "통합문서 · "+books.Items.Count;
            noBooks.Text = bookSearch.Text.Trim().Length == 0 ? "열린 문서가 없습니다" : "일치하는 문서가 없습니다";
            noBooks.Visible = books.Items.Count == 0;
        }
        private void RenderSheets() {
            if (!snapshot.Exists(x => x.Parent == selectedBook && x.Token == selectedSheet)) {
                var first = snapshot.Find(x => x.Parent == selectedBook && x.Active) ?? snapshot.Find(x => x.Parent == selectedBook);
                selectedSheet = first == null ? "" : first.Token;
            }
            Fill(sheets, x => x.Parent == selectedBook && x.Parent != "" && Matches(x, sheetSearch.Text), selectedSheet);
            sheetHeading.Text = "시트 · "+sheets.Items.Count;
            noSheets.Text = sheetSearch.Text.Trim().Length == 0 ? "표시할 시트가 없습니다" : "일치하는 시트가 없습니다";
            noSheets.Visible = sheets.Items.Count == 0;
            int count = snapshot.FindAll(x => x.Parent == selectedBook && x.Parent != "").Count;
            var book = snapshot.Find(x => x.Parent == "" && x.Token == selectedBook);
            status.Text = (book==null ? "문서 선택" : book.Title)+" · "+count+"개 시트";
            tips.SetToolTip(status,status.Text+"\nEnter 또는 더블클릭하여 이동");
        }
        private static bool Matches(DocumentEntry item, string query) { return item.Title.IndexOf(query.Trim(), StringComparison.CurrentCultureIgnoreCase) >= 0; }
        private void Fill(ListBox list, Predicate<DocumentEntry> filter, string selected) {
            int top = list.Items.Count == 0 ? 0 : list.TopIndex;
            rendering = true; list.BeginUpdate();
            try {
                list.Items.Clear();
                foreach (var item in snapshot) if (filter(item)) {
                    list.Items.Add(item);
                    if (item.Token == selected) list.SelectedIndex = list.Items.Count - 1;
                }
                if (list.Items.Count > 0) list.TopIndex = Math.Min(top, list.Items.Count - 1);
            } finally { list.EndUpdate(); rendering = false; }
        }
        private void Activate(ListBox list) {
            var item = list.SelectedItem as DocumentEntry;
            Trace(item == null ? "activation_no_selection" : item.Available ? "activation_requested" : "activation_hidden");
            if (item == null || !item.Available) return;
            if (!handler.MoveDocument(item.Token)) { Trace("activation_rejected"); status.Text = "대상이 닫혔거나 변경되었습니다. 새로 고침하세요."; }
            else { Trace("activation_complete"); RefreshSnapshot(); if (!pin.Checked) Collapse(true); }
        }
        private void Trace(string stage) {
            var session = handler as NavigatorSession;
            if (session != null) session.TracePane("document_" + stage, null);
        }
        private void ListKey(object sender, KeyEventArgs e) {
            if (e.KeyCode == Keys.Escape) { EscapeKey(sender,e); return; }
            if (e.KeyCode != Keys.Enter) return;
            Trace("enter_key");
            e.Handled = e.SuppressKeyPress = true; Activate((ListBox)sender);
        }
        private void SearchKey(ListBox list,KeyEventArgs e) {
            if(e.KeyCode==Keys.Down) {
                e.Handled=e.SuppressKeyPress=true;
                if(list.Items.Count>0) { if(list.SelectedIndex<0)list.SelectedIndex=0; list.Focus(); }
            } else EscapeKey(list,e);
        }
        private void EscapeKey(object sender, KeyEventArgs e) {
            if (e.KeyCode != Keys.Escape) return;
            e.Handled = e.SuppressKeyPress = true; Collapse(true);
        }
        private void ShowTip(object sender, MouseEventArgs e) {
            var list = (ListBox)sender;
            int index = list.IndexFromPoint(e.Location);
            tips.SetToolTip(list, index < 0 ? "" : ((DocumentEntry)list.Items[index]).Title);
        }
        private void PlaceSplitter() {
            int height = split.Height - split.SplitterWidth;
            if (height < 8 || settingSplit) return;
            settingSplit = true;
            try {
                ratio = Math.Max(.25, Math.Min(.75, ratio));
                split.Panel1MinSize = 0; split.Panel2MinSize = 0;
                split.SplitterDistance = (int)(height * ratio);
            } finally { settingSplit = false; }
        }
        private void Collapse(bool value) {
            if (collapsed == value) return;
            collapsed = value;
            split.Visible = status.Visible = tools.Visible = !value;
            rail.Visible = value; handler.SetCollapsed(value);
            if (value) rail.Focus(); else bookSearch.Focus();
        }
        protected override void Dispose(bool disposing) {
            if (disposing) { tips.Dispose(); snapshot.Clear(); foreach(var image in toolbarImages)image.Dispose(); toolbarImages.Clear(); }
            base.Dispose(disposing);
        }
    }
}
