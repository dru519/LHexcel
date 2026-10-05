using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace LH.NxHost {
    [ComVisible(true), Guid("C812F023-E912-49AA-98B7-75DC86501901"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IPicturePreviewService {
        [DispId(1)] bool Connect(string productTitle, string interfaceVersion, string bridgeHash, string featureHash, string commandHash, string bitness);
        [DispId(2)] bool ShowPreview(object rows, string sizeMode, string excelHwnd);
        [DispId(3)] object RenderSummary(string summary, int width, int height);
        [DispId(4)] object RenderCells(object cells, int width, int height);
    }
    [ComVisible(true), Guid("C812F023-E912-49AA-98B7-75DC86501902"), ProgId("LH.NxHost.PicturePreviewService"), ClassInterface(ClassInterfaceType.None)]
    public sealed class PicturePreviewService : IPicturePreviewService {
        private bool connected;
        private static int showing;
        public bool Connect(string title, string version, string bridge, string feature, string command, string bitness) {
            return connected = title == NxHostContract.ProductTitle && version == "nx-picture-preview-v1" &&
                bridge == NxHostContract.BridgeHash && feature == NxHostContract.FeatureHash &&
                command == NxHostContract.CommandHash && bitness == NxHostContract.ProcessBitness;
        }
        public bool ShowPreview(object rows, string sizeMode, string excelHwnd) {
            if (!connected) throw new InvalidOperationException("PICTURE_NOT_CONNECTED");
            if (sizeMode != "FIT" && sizeMode != "ORIGINAL") throw new ArgumentException("PICTURE_MODE");
            long handle;
            if (!Int64.TryParse(excelHwnd, out handle)) throw new ArgumentException("PICTURE_OWNER");
            IntPtr owner = new IntPtr(handle);
            uint pid; GetWindowThreadProcessId(owner, out pid);
            if (!IsWindow(owner) || pid != (uint)Process.GetCurrentProcess().Id) throw new ArgumentException("PICTURE_OWNER");
            var entries = PicturePreviewEntry.Parse(rows as Array);
            if (Interlocked.CompareExchange(ref showing, 1, 0) != 0) throw new InvalidOperationException("PICTURE_BUSY");
            try {
                IntPtr popup = GetLastActivePopup(owner);
                uint popupPid; GetWindowThreadProcessId(popup, out popupPid);
                if (IsWindow(popup) && popupPid == pid) owner = popup;
                using (var form = new PicturePreviewForm(entries, sizeMode, (int)GetDpiForWindow(owner))) return form.ShowDialog(new ExcelWindowOwner(owner)) == DialogResult.OK;
            } finally { Interlocked.Exchange(ref showing, 0); }
        }
        public object RenderSummary(string summary, int width, int height) {
            if (!connected) throw new InvalidOperationException("PICTURE_NOT_CONNECTED");
            using (var bitmap = WorksheetPreviewRenderer.Summary(summary, width, height)) return PreviewPicture.Wrap(bitmap);
        }
        public object RenderCells(object cells, int width, int height) {
            if (!connected) throw new InvalidOperationException("PICTURE_NOT_CONNECTED");
            using (var bitmap = WorksheetPreviewRenderer.Cells(cells as Array, width, height)) return PreviewPicture.Wrap(bitmap);
        }
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll")] private static extern IntPtr GetLastActivePopup(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr hwnd);
    }
    internal sealed class PreviewPicture : AxHost {
        private PreviewPicture() : base("") { }
        internal static object Wrap(Image image) { return GetIPictureDispFromPicture(image); }
    }
    internal sealed class PicturePreviewEntry : IDisposable {
        internal string Path, Address, Failure = "";
        internal double CellWidth, CellHeight, WidthPt, HeightPt;
        internal Bitmap Thumbnail;
        internal RectangleF Placement(string sizeMode) {
            double fit=sizeMode=="FIT" ? Math.Min(CellWidth/WidthPt,CellHeight/HeightPt) : 1;
            float width=(float)(WidthPt*fit),height=(float)(HeightPt*fit);
            return new RectangleF(sizeMode=="FIT" ? (float)(CellWidth-width)/2 : 0,
                sizeMode=="FIT" ? (float)(CellHeight-height)/2 : 0,width,height);
        }
        internal static List<PicturePreviewEntry> Parse(Array rows) {
            if (rows == null || rows.Rank != 2 || rows.GetLength(1) != 4 || rows.GetLength(0) < 1 || rows.GetLength(0) > 500) throw new ArgumentException("PICTURE_ROWS");
            var result = new List<PicturePreviewEntry>();
            for (int i=0; i<rows.GetLength(0); i++) {
                int r=i+rows.GetLowerBound(0), c=rows.GetLowerBound(1);
                string path = Convert.ToString(rows.GetValue(r,c));
                string address = Convert.ToString(rows.GetValue(r,c+1));
                double width = Convert.ToDouble(rows.GetValue(r,c+2)), height = Convert.ToDouble(rows.GetValue(r,c+3));
                if (path.Length > 1024 || !System.IO.Path.IsPathRooted(path) || System.IO.Path.GetPathRoot(path).Length < 3 ||
                    path.StartsWith(@"\\?\") || path.StartsWith(@"\\.\") || path.IndexOf(':',2) >= 0 || path.IndexOf('\0') >= 0 ||
                    address.Length > 1024 || !(width > 0 && width <= 1000000) || !(height > 0 && height <= 1000000)) throw new ArgumentException("PICTURE_INPUT");
                string ext = System.IO.Path.GetExtension(path).ToLowerInvariant();
                if (Array.IndexOf(new[] {".png",".jpg",".jpeg",".bmp",".gif",".tif",".tiff"}, ext) < 0) throw new ArgumentException("PICTURE_TYPE");
                result.Add(new PicturePreviewEntry { Path=System.IO.Path.GetFullPath(path), Address=address, CellWidth=width, CellHeight=height });
            }
            return result;
        }
        internal void Load() {
            try {
                using (var stream = new FileStream(Path, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                    if (stream.Length < 4 || stream.Length > 32 * 1024 * 1024) throw new InvalidDataException();
                    using (var image = Image.FromStream(stream, false, true)) {
                        if ((long)image.Width * image.Height > 16000000 || image.Width < 1 || image.Height < 1) throw new InvalidDataException();
                        float dpiX = image.HorizontalResolution, dpiY = image.VerticalResolution;
                        WidthPt = image.Width * 72.0 / (dpiX >= 1 && dpiX <= 10000 ? dpiX : 96);
                        HeightPt = image.Height * 72.0 / (dpiY >= 1 && dpiY <= 10000 ? dpiY : 96);
                        double fit = Math.Min(160.0 / image.Width, 120.0 / image.Height);
                        Thumbnail = new Bitmap((int)Math.Max(1,image.Width*fit),(int)Math.Max(1,image.Height*fit));
                        using (var graphics = Graphics.FromImage(Thumbnail)) {
                            graphics.Clear(Color.White);
                            graphics.DrawImage(image, new Rectangle(0,0,Thumbnail.Width,Thumbnail.Height));
                        }
                    }
                }
            } catch (Exception error) {
                if (!(error is ArgumentException || error is IOException || error is UnauthorizedAccessException || error is OutOfMemoryException || error is ExternalException)) throw;
                Dispose(); Failure = "손상·미지원·크기 제한 또는 읽기 실패";
            }
        }
        public void Dispose() { if (Thumbnail != null) { Thumbnail.Dispose(); Thumbnail = null; } }
    }
    internal sealed class PicturePreviewForm : Form {
        private readonly List<PicturePreviewEntry> entries;
        private readonly string mode;
        private readonly double dpiScale;
        private readonly ListView list = new ListView { Dock=DockStyle.Fill, View=View.Details, FullRowSelect=true, MultiSelect=false, HideSelection=false, AccessibleName="그림과 대상 셀 매핑" };
        private readonly Panel canvas = new Panel { Dock=DockStyle.Fill, BackColor=Color.White, AccessibleName="셀 맞춤 미리보기" };
        private readonly Label status = new Label { Dock=DockStyle.Bottom, Height=44, AutoEllipsis=true };
        private readonly Label modeSummary = new Label { Dock=DockStyle.Top, AutoEllipsis=true, TextAlign=ContentAlignment.MiddleLeft };
        private readonly Label selectionInfo = new Label { Dock=DockStyle.Bottom, AutoEllipsis=true, TextAlign=ContentAlignment.MiddleLeft };
        private readonly ToolTip tips = new ToolTip();
        private readonly ComboBox viewMode = new ComboBox { Dock=DockStyle.Top, DropDownStyle=ComboBoxStyle.DropDownList, AccessibleName="미리보기 크기 비교" };
        private readonly Button accept = new Button { Text="확인", Width=82, Enabled=false, DialogResult=DialogResult.OK };
        private readonly System.Windows.Forms.Timer loader = new System.Windows.Forms.Timer { Interval=20 };
        private readonly ImageList images = new ImageList { ImageSize=new Size(64,48), ColorDepth=ColorDepth.Depth32Bit };
        private int loaded;
        internal PicturePreviewForm(List<PicturePreviewEntry> value, string sizeMode, int dpi = 96) {
            entries=value; mode=sizeMode;
            dpiScale=Math.Max(96,dpi)/96.0;
            // Excel owns DPI awareness. Scale pixel geometry explicitly; point fonts
            // already render at the host DPI and must not be scaled a second time.
            AutoScaleMode=AutoScaleMode.None;
            Font=new Font("맑은 고딕",9); Text="내엑셀 - 그림 미리보기";
            ClientSize=new Size(D(760),D(460)); MinimumSize=new Size(D(560),D(360)); StartPosition=FormStartPosition.CenterParent;
            ShowInTaskbar=false; MinimizeBox=false; MaximizeBox=false;
            Padding=new Padding(D(10)); BackColor=SystemColors.Control;
            status.Height=Font.Height*2+D(12);
            modeSummary.Text=mode=="FIT" ? "셀에 맞춤 (FIT) · 비율을 유지해 대상 셀 안에 배치합니다." : "원본 크기 (ORIGINAL) · 원래 크기로 배치해 셀을 넘을 수 있습니다.";
            modeSummary.Text="실제 삽입: "+modeSummary.Text;
            modeSummary.Height=Font.Height+D(16);
            selectionInfo.Height=Font.Height*3+D(12); selectionInfo.Text="목록에서 그림을 선택하세요.";
            var split=new SplitContainer { Dock=DockStyle.Fill, Size=new Size(D(760),D(400)), SplitterDistance=D(430) };
            split.SplitterWidth=D(8); split.Panel1MinSize=D(200); split.Panel2MinSize=D(140);
            split.Panel1.Controls.Add(list); split.Panel2.Controls.Add(canvas);
            var listHeading=new Label { Text="그림 목록 · "+entries.Count+"개", Dock=DockStyle.Top, Height=Font.Height+D(8), TextAlign=ContentAlignment.MiddleLeft };
            var canvasHeading=new Label { Text="배치 미리보기", Dock=DockStyle.Top, Height=listHeading.Height, TextAlign=ContentAlignment.MiddleLeft };
            split.Panel1.Controls.Add(listHeading); split.Panel2.Controls.Add(canvasHeading);
            viewMode.Items.AddRange(new object[]{"보기 비교: 셀에 맞춤 (FIT)","보기 비교: 원본 크기 (ORIGINAL)"});
            viewMode.SelectedIndex=mode=="FIT" ? 0 : 1;
            split.Panel2.Controls.Add(viewMode); viewMode.BringToFront();
            tips.SetToolTip(viewMode,"배치 차이만 비교합니다. 실제 삽입 설정은 이전 창에서 변경하세요.");
            viewMode.SelectedIndexChanged+=(s,e)=>{UpdateSelectionInfo();canvas.Invalidate();};
            canvas.BorderStyle=BorderStyle.FixedSingle;
            int buttonHeight=Math.Max(D(28),Font.Height+D(8));
            accept.Size=new Size(D(82),buttonHeight); accept.Margin=new Padding(D(4));
            var actions=new FlowLayoutPanel { Dock=DockStyle.Bottom, Height=buttonHeight+D(12), FlowDirection=FlowDirection.RightToLeft };
            var close=new Button { Text="취소", Size=accept.Size, Margin=accept.Margin, DialogResult=DialogResult.Cancel };
            actions.Controls.Add(close); actions.Controls.Add(accept); CancelButton=close; AcceptButton=accept;
            Controls.Add(split); Controls.Add(selectionInfo); Controls.Add(status); Controls.Add(actions); Controls.Add(modeSummary);
            list.Columns.Add("파일명",D(150)); list.Columns.Add("대상 셀",D(150)); list.Columns.Add("확인",D(110)); list.SmallImageList=images; list.ShowItemToolTips=true;
            images.ImageSize=new Size(Math.Min(256,D(48)),Math.Min(256,D(36)));
            foreach (var entry in entries) list.Items.Add(new ListViewItem(new[] {System.IO.Path.GetFileName(entry.Path),entry.Address,"확인 대기"}) {ToolTipText=System.IO.Path.GetFileName(entry.Path)+" → "+entry.Address});
            list.SelectedIndexChanged += (s,e) => {
                UpdateSelectionInfo();
                canvas.Invalidate();
            };
            canvas.Paint += DrawPreview;
            canvas.Resize += (s,e)=>canvas.Invalidate();
            loader.Tick += LoadNext;
            Shown += (s,e) => { Size=new Size(Math.Min(Width,Screen.FromControl(this).WorkingArea.Width),Math.Min(Height,Screen.FromControl(this).WorkingArea.Height)); loader.Start(); };
        }
        private int D(double value) {return (int)Math.Round(value*dpiScale);}
        private string ViewMode { get {return viewMode.SelectedIndex==0 ? "FIT" : "ORIGINAL";} }
        private void UpdateSelectionInfo() {
            if(list.SelectedIndices.Count!=1){selectionInfo.Text="목록에서 그림을 선택하세요.";tips.SetToolTip(selectionInfo,"");return;}
            var entry=entries[list.SelectedIndices[0]];
            string text=System.IO.Path.GetFileName(entry.Path)+" → "+entry.Address;
            if(entry.Thumbnail!=null) {
                var picture=entry.Placement(ViewMode);
                text+=String.Format("\r\n셀 {0:0.#} × {1:0.#}pt · 그림 {2:0.#} × {3:0.#}pt",entry.CellWidth,entry.CellHeight,picture.Width,picture.Height);
                text+=ViewMode=="FIT" ? String.Format(" · 좌우 여백 {0:0.#}pt / 상하 여백 {1:0.#}pt",picture.X,picture.Y) :
                    (picture.Width>entry.CellWidth || picture.Height>entry.CellHeight ? " · 셀 바깥으로 넘칩니다." : " · 셀 왼쪽 위에 배치됩니다.");
                if(ViewMode!=mode) text+="\r\n비교용 보기입니다. 실제 삽입 설정은 변경되지 않았습니다.";
            }
            selectionInfo.Text=text;tips.SetToolTip(selectionInfo,text);
        }
        private void LoadNext(object sender, EventArgs args) {
            if (loaded >= entries.Count) { loader.Stop(); return; }
            var budget=Stopwatch.StartNew();int batch=0;
            do {
            var entry=entries[loaded]; entry.Load();
            if (entry.Thumbnail != null) {
                using(var tile=new Bitmap(images.ImageSize.Width,images.ImageSize.Height)) {
                    double fit=Math.Min((double)tile.Width/entry.Thumbnail.Width,(double)tile.Height/entry.Thumbnail.Height);
                    int width=(int)(entry.Thumbnail.Width*fit),height=(int)(entry.Thumbnail.Height*fit);
                    using(var graphics=Graphics.FromImage(tile)){graphics.Clear(Color.White);graphics.DrawImage(entry.Thumbnail,new Rectangle((tile.Width-width)/2,(tile.Height-height)/2,width,height));}
                    images.Images.Add(tile);
                }
                list.Items[loaded].ImageIndex=images.Images.Count-1;
            }
            list.Items[loaded].SubItems[2].Text=entry.Failure.Length == 0 ? "확인됨" : entry.Failure;
            loaded++;
            status.Text=loaded+" / "+entries.Count+"개 확인 · 원본 셀은 변경하지 않습니다.\n확인을 눌러 돌아간 뒤 실제 삽입 직전에 파일과 대상을 다시 검사합니다.";
            if (list.SelectedIndices.Count == 0) list.Items[loaded-1].Selected=true;
            batch++;
            } while(loaded<entries.Count && batch<8 && budget.ElapsedMilliseconds<12);
            UpdateSelectionInfo();
            canvas.Invalidate();
            if (loaded == entries.Count) { loader.Stop(); accept.Enabled=entries.TrueForAll(x=>x.Failure.Length==0); }
        }
        private void DrawPreview(object sender, PaintEventArgs args) {
            if (list.SelectedIndices.Count != 1) return;
            var entry=entries[list.SelectedIndices[0]];
            if (entry.Thumbnail == null) { TextRenderer.DrawText(args.Graphics,entry.Failure,Font,canvas.ClientRectangle,Color.Firebrick); return; }
            var placement=entry.Placement(ViewMode);
            double extentWidth=Math.Max(entry.CellWidth,placement.Right),extentHeight=Math.Max(entry.CellHeight,placement.Bottom);
            double scale=Math.Min((canvas.Width-D(64))/extentWidth,(canvas.Height-D(76))/extentHeight);
            if (scale <= 0) return;
            var cell=new RectangleF(D(40),Font.Height+D(30),(float)(entry.CellWidth*scale),(float)(entry.CellHeight*scale));
            // The COM contract supplies the target extent, not neighbouring cell sizes.
            // Keep the surrounding sheet illustrative and identify the real target address.
            using(var grid=new Pen(Color.FromArgb(220,220,220))) {
                for(int x=D(40);x<canvas.Width;x+=D(64))args.Graphics.DrawLine(grid,x,cell.Top,x,canvas.Height);
                for(int y=(int)cell.Top;y<canvas.Height;y+=D(20))args.Graphics.DrawLine(grid,D(40),y,canvas.Width,y);
            }
            args.Graphics.FillRectangle(Brushes.White,cell);
            args.Graphics.FillRectangle(Brushes.WhiteSmoke,0,cell.Top,D(40),canvas.Height-cell.Top);
            args.Graphics.FillRectangle(Brushes.WhiteSmoke,0,cell.Top-D(22),canvas.Width,D(22));
            TextRenderer.DrawText(args.Graphics,entry.Address,Font,new Rectangle(D(40),(int)cell.Top-D(22),Math.Max(1,canvas.Width-D(40)),D(22)),Color.Black,TextFormatFlags.EndEllipsis|TextFormatFlags.VerticalCenter);
            var picture=new RectangleF(cell.X+(float)(placement.X*scale),cell.Y+(float)(placement.Y*scale),(float)(placement.Width*scale),(float)(placement.Height*scale));
            args.Graphics.DrawImage(entry.Thumbnail,picture);
            if(placement.Right>entry.CellWidth || placement.Bottom>entry.CellHeight)
                using(var pen=new Pen(Color.Firebrick,D(1))) {pen.DashStyle=System.Drawing.Drawing2D.DashStyle.Dash;args.Graphics.DrawRectangle(pen,picture.X,picture.Y,picture.Width,picture.Height);}
            using(var selection=new Pen(Color.FromArgb(33,115,70),D(2)))args.Graphics.DrawRectangle(selection,cell.X,cell.Y,cell.Width,cell.Height);
            TextRenderer.DrawText(args.Graphics,"초록 테두리: 대상 범위 · 주변 격자는 예시",Font,new Point(D(8),D(4)),Color.Black);
        }
        protected override void Dispose(bool disposing) {
            if (disposing) { loader.Stop(); loader.Dispose(); images.Dispose(); tips.Dispose(); foreach(var entry in entries) entry.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
