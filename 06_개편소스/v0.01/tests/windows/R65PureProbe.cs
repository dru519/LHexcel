using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using LH.NxHost;

internal static class R65PureProbe {
    private static int count;
    private static void Check(bool value,string name) { if(!value) throw new Exception(name); Console.WriteLine("PASS|"+name); count++; }
    private static T Field<T>(object value,string name) { return (T)value.GetType().GetField(name,BindingFlags.NonPublic|BindingFlags.Instance).GetValue(value); }
    [STAThread] private static int Main(string[] args) {
        try {
            Directory.CreateDirectory(args[0]);
            Application.EnableVisualStyles();
            CompareCases();
            PictureCases(args[0]);
            PreviewCases(args[0]);
            DocumentCases();
            Console.WriteLine("PASS|total="+count); return 0;
        } catch(Exception error) { Console.WriteLine("FAIL|"+error); return 1; }
    }
    private static bool Connect(WorkbookCompareService service) { return service.Connect(NxHostContract.ProductTitle,"nx-compare-v1",NxHostContract.BridgeHash,NxHostContract.FeatureHash,NxHostContract.CommandHash,NxHostContract.ProcessBitness); }
    private static object[,] Rows(int count) {
        var result=new object[count,5];
        for(int r=0;r<count;r++) { result[r,0]="0"; result[r,1]=""; result[r,2]="5:"+r; result[r,3]=""; result[r,4]=""; }
        return result;
    }
    private static void PreviewCases(string folder) {
        using(var image=WorksheetPreviewRenderer.Summary("비교 표본\nA1: 10 → 20\nA2: 같음 → 같음",300,100)) {
            Check(image.Width==600 && image.Height==200,"preview high resolution");
            Check(image.GetPixel(590,50).ToArgb()==Color.FromArgb(255,199,206).ToArgb(),"preview differing value red");
            Check(image.GetPixel(590,96).ToArgb()==Color.White.ToArgb(),"preview equal value white");
            image.Save(Path.Combine(folder,"comparison-preview.png"));
        }
        var cells=new object[,] {{0,0,80,24,"셀",16777215,0,9,true,2},{0,24,80,24,"강조",14342874,0,9,false,2}};
        using(var image=WorksheetPreviewRenderer.Cells(cells,160,80)) {
            Check(image.GetPixel(150,90).ToArgb()==ColorTranslator.FromOle(14342874).ToArgb(),"preview sample fill");
            image.Save(Path.Combine(folder,"cell-preview.png"));
            Check(PreviewPicture.Wrap(image)!=null,"preview COM picture conversion");
        }
        bool rejected=false;
        try {WorksheetPreviewRenderer.Cells(new object[1,2],100,80);}catch(ArgumentException){rejected=true;}
        Check(rejected,"preview rejects malformed matrix");
    }
    private static void Wait(WorkbookCompareService service,string job) {
        var deadline=DateTime.UtcNow.AddSeconds(5);
        while(service.GetStatus(job)=="running" && DateTime.UtcNow<deadline) Thread.Sleep(1);
        Check(service.GetStatus(job)!="running","compare finishes bounded");
    }
    private static void CompareCases() {
        var service=new WorkbookCompareService();
        bool rejected=false;
        try { service.Start(Rows(1),Rows(1)); } catch(InvalidOperationException) {rejected=true;}
        Check(rejected,"compare rejects unconnected");
        Check(!service.Connect("wrong","nx-compare-v1","","","","x64"),"compare rejects wrong identity");
        Check(Connect(service),"compare exact handshake");
        var left=Rows(7); var right=Rows(7);
        right[1,0]="1"; right[1,1]="8:=2";
        left[2,0]=right[2,0]="1"; left[2,1]="8:=1"; right[2,1]="8:=2";
        right[3,2]="#EMPTY"; right[4,3]="$A$1:$B$1"; right[5,4]="format";
        right[6,0]="1"; right[6,2]="8:1"; right[6,3]="merge"; right[6,4]="format";
        string job=service.Start(left,right); right[0,2]="mutated after Start";
        Wait(service,job); var result=(string[])service.GetResult(job);
        Check(result[0]=="","compare immutable caller snapshot");
        Check(result[1]=="수식/상수" && result[2]=="수식" && result[3]=="값" && result[4]=="병합" && result[5]=="서식","compare reason parity");
        Check(result[6]=="수식/상수, 값, 병합, 서식","compare reason order");
        Check(service.GetProgress(job)==100,"compare progress completes");
        rejected=false; try { service.GetResult("foreign"); } catch(ArgumentException) {rejected=true;}
        Check(rejected,"compare rejects foreign job");
        rejected=false; try { service.Start(new object[1,4],Rows(1)); } catch(InvalidDataException) {rejected=true;}
        Check(rejected,"compare rejects shape");
        var invalid=Rows(1); invalid[0,2]=new object(); rejected=false;
        try {service.Start(invalid,Rows(1));} catch(InvalidDataException) {rejected=true;}
        Check(rejected,"compare rejects non-string worker data");
        var cts=new CancellationTokenSource();cts.Cancel(); rejected=false;
        try { WorkbookCompareService.Compare(new[]{"0","","","",""},new[]{"0","","","",""},cts.Token,p=>{}); } catch(OperationCanceledException) {rejected=true;}
        Check(rejected,"compare cancellation checked before work");
        var large=Rows(200000);
        var timer=System.Diagnostics.Stopwatch.StartNew(); job=service.Start(large,large); Wait(service,job);timer.Stop();
        Check(((string[])service.GetResult(job)).Length==200000,"compare 200000 cells");
        Console.WriteLine("METRIC|compare_200000_copy_and_compute_ms="+timer.Elapsed.TotalMilliseconds);
    }
    private static void PictureCases(string root) {
        string png=Path.Combine(root,"valid.png"), broken=Path.Combine(root,"broken.png");
        using(var bitmap=new Bitmap(400,100)) { using(var g=Graphics.FromImage(bitmap)) g.Clear(Color.Orange); bitmap.Save(png,System.Drawing.Imaging.ImageFormat.Png); }
        File.WriteAllBytes(broken,new byte[]{137,80,78,71,13,10,26,10});
        object[,] data={{png,"$A$1",100.0,40.0},{broken,"$A$2",100.0,40.0}};
        var entries=PicturePreviewEntry.Parse(data);
        entries[0].Load();entries[1].Load();
        Check(entries[0].Thumbnail!=null && entries[0].Failure=="","picture decodes valid");
        Check(entries[1].Thumbnail==null && entries[1].Failure!="","picture identifies truncated payload");
        Check(Math.Abs((double)entries[0].Thumbnail.Width/entries[0].Thumbnail.Height-4)<.01,"picture thumbnail preserves aspect");
        var fitPlacement=entries[0].Placement("FIT");
        Check(Math.Abs(fitPlacement.Width-100)<.01 && Math.Abs(fitPlacement.Height-25)<.01 && Math.Abs(fitPlacement.Y-7.5)<.01,"FIT centered with measured padding");
        var originalPlacement=entries[0].Placement("ORIGINAL");
        Check(originalPlacement.X==0 && originalPlacement.Y==0 && originalPlacement.Width>100,"ORIGINAL overflow geometry retained");
        using(var stream=new FileStream(png,FileMode.Open,FileAccess.ReadWrite,FileShare.None)) Check(stream.Length>0,"picture releases source handle");
        foreach(var entry in entries) entry.Dispose();
        Check(entries[0].Thumbnail==null,"picture disposes bitmap");
        using(var form=new PicturePreviewForm(PicturePreviewEntry.Parse(new object[,]{{png,"'[Long workbook.xlsx]Sheet1'!$A$1",100.0,40.0}}),"FIT",192)) {
            Check(Field<Label>(form,"modeSummary").Text.Contains("셀에 맞춤"),"picture explains size mode");
            Check(Field<Label>(form,"selectionInfo").AutoEllipsis,"picture long mapping uses ellipsis");
            form.Show(); Application.DoEvents();
            var list=Field<ListView>(form,"list"); list.Items[0].Selected=true; Application.DoEvents();
            typeof(PicturePreviewForm).GetMethod("LoadNext",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(form,new object[]{null,EventArgs.Empty});
            var choices=Field<ComboBox>(form,"viewMode");
            Check(Field<Label>(form,"selectionInfo").Text.Contains("여백"),"FIT padding summary visible");
            choices.SelectedIndex=1;
            Check(Field<Label>(form,"selectionInfo").Text.Contains("넘칩니다") && Field<Label>(form,"selectionInfo").Text.Contains("변경되지"),"alternate view explains overflow without changing insertion");
            Check(Field<Label>(form,"modeSummary").Text.Contains("셀에 맞춤"),"actual insertion mode retained");
            Check(Field<Label>(form,"selectionInfo").Text.Contains("$A$1"),"picture selected mapping shown");
            list.SelectedIndices.Clear();
            Check(Field<Label>(form,"selectionInfo").Text=="목록에서 그림을 선택하세요.","picture cleared selection resets mapping");
            Check(Field<Button>(form,"accept").Bottom<=Field<Button>(form,"accept").Parent.ClientSize.Height,"picture actions stay inside footer");
            bool escapeHandled=(bool)typeof(Form).GetMethod("ProcessDialogKey",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(form,new object[]{Keys.Escape});
            Check(escapeHandled && form.DialogResult==DialogResult.Cancel,"picture Escape chooses cancel");
            form.Close();
        }
        data[0,0]="relative.png"; bool rejected=false; try{PicturePreviewEntry.Parse(data);}catch(ArgumentException){rejected=true;}
        Check(rejected,"picture rejects relative path");
        data[0,0]=png;data[0,2]=Double.NaN;rejected=false;try{PicturePreviewEntry.Parse(data);}catch(ArgumentException){rejected=true;}
        Check(rejected,"picture rejects invalid geometry");
    }
    private sealed class Handler : IDocumentNavigatorHandler {
        internal int Moved;
        internal bool Collapsed;
        internal Array Data;
        public Array GetDocuments(){return Data;}
        public bool MoveDocument(string token){Moved++;return token!="999";}
        public void SetCollapsed(bool collapsed){Collapsed=collapsed;}
    }
    private static void DocumentCases() {
        var data=new object[103,5];
        data[0,0]="1";data[0,1]="";data[0,2]="Alpha";data[0,3]=true;data[0,4]=true;
        data[1,0]="2";data[1,1]="";data[1,2]="Beta";data[1,3]=true;data[1,4]=false;
        for(int i=2;i<103;i++){data[i,0]=(i+1).ToString();data[i,1]="1";data[i,2]="Sheet "+i;data[i,3]=i!=2;data[i,4]=i==3;}
        var handler=new Handler{Data=data};
        using(var form=new Form { ClientSize=new Size(430,700),ShowInTaskbar=false })
        using(var control=new DocumentNavigatorControl(handler){Dock=DockStyle.Fill}) {
            form.Controls.Add(control);form.Show();Application.DoEvents();control.RefreshSnapshot();
            var books=Field<ListBox>(control,"books");var sheets=Field<ListBox>(control,"sheets");
            var bs=Field<TextBox>(control,"bookSearch");var ss=Field<TextBox>(control,"sheetSearch");
            var split=Field<SplitContainer>(control,"split");
            var pin=Field<ToolStripButton>(control,"pin");
            Check(pin.DisplayStyle==ToolStripItemDisplayStyle.Image && pin.Image!=null && pin.Checked,"navigator pin uses checked icon");
            Check(pin.AccessibleName=="고정" && pin.ToolTipText.Length>0,"navigator pin remains accessible");
            var iconScale=typeof(DocumentNavigatorControl).GetMethod("ApplyToolbarDpi",BindingFlags.Instance|BindingFlags.NonPublic);
            Check(iconScale!=null,"navigator has hosted toolbar DPI boundary");
            iconScale.Invoke(control,new object[]{192});
            Check(Field<ToolStrip>(control,"tools").ImageScalingSize==new Size(32,32) && pin.Size==new Size(48,48),"navigator physical icons fit 200 percent");
            iconScale.Invoke(control,new object[]{96});
            var oldPinImage=pin.Image;
            typeof(Control).GetMethod("OnSystemColorsChanged",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(control,new object[]{EventArgs.Empty});
            Check(pin.Image!=null && pin.Image!=oldPinImage && Field<List<Image>>(control,"toolbarImages").Count==3,"navigator icons refresh system colors without accumulation");
            Check(bs.Height>=bs.PreferredHeight && ss.Height>=ss.PreferredHeight,"navigator search font fits row");
            Check(books.Items.Count==2 && sheets.Items.Count==101,"navigator split lists");
            books.SelectedIndex=1;Check(handler.Moved==0 && sheets.Items.Count==0,"navigator click selects without activating");
            books.SelectedIndex=0; sheets.TopIndex=30; sheets.SelectedIndex=32;
            split.SplitterDistance=(int)((split.Height-split.SplitterWidth)*.7); double prior=(double)split.SplitterDistance/(split.Height-split.SplitterWidth);
            control.RefreshSnapshot();
            Check(sheets.TopIndex==30 && sheets.SelectedIndex==32,"navigator retains selection and scroll");
            Check(Math.Abs((double)split.SplitterDistance/(split.Height-split.SplitterWidth)-prior)<.02,"navigator retains split ratio");
            bs.Text="Beta";Check(books.Items.Count==1 && sheets.Items.Count==101,"navigator book search preserves selected document");
            Field<Button>(control,"bookClear").PerformClick();
            Check(bs.Text=="" && books.Items.Count==2,"navigator clear book search");
            ss.Text="Sheet 12";Check(sheets.Items.Count==1,"navigator independent sheet search");
            Field<Button>(control,"sheetClear").PerformClick();
            Check(ss.Text=="" && sheets.Items.Count==101,"navigator clear sheet search");
            var escape=new KeyEventArgs(Keys.Escape);
            typeof(Control).GetMethod("OnKeyDown",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(Field<Button>(control,"sheetClear"),new object[]{escape});
            Check(handler.Collapsed && escape.Handled,"navigator clear button supports escape");
            typeof(DocumentNavigatorControl).GetMethod("Collapse",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(control,new object[]{false});
            bs.Clear();ss.Clear();sheets.SelectedIndex=0;
            typeof(DocumentNavigatorControl).GetMethod("Activate",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(control,new object[]{sheets});
            Check(handler.Moved==0,"navigator hidden sheet cannot activate");
            sheets.SelectedIndex=1;typeof(DocumentNavigatorControl).GetMethod("Activate",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(control,new object[]{sheets});
            Check(handler.Moved==1,"navigator explicit activation");
            var old=sheets.Items.Count; handler.Data=null;control.RefreshSnapshot();
            Check(sheets.Items.Count==old,"navigator refresh failure retains snapshot");
            typeof(DocumentNavigatorControl).GetMethod("Collapse",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(control,new object[]{true});
            Check(handler.Collapsed && Field<Button>(control,"rail").Visible,"navigator collapse rail");
            form.Close();
        }
    }
}
