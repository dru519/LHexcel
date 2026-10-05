using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace LH.NxHost {
    [ComVisible(true), Guid("C812F023-E912-49AA-98B7-75DC86501905"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IWorkbookCompareResultsService {
        [DispId(1)] bool Connect(string title,string version,string bridge,string feature,string command,string bitness);
        [DispId(2)] void ShowResults(object rows,string reportName,string excelHwnd,object callback);
        [DispId(3)] void Close();
    }
    [ComVisible(true), Guid("C812F023-E912-49AA-98B7-75DC86501906"), ProgId("LH.NxHost.WorkbookCompareResultsService"), ClassInterface(ClassInterfaceType.None)]
    public sealed class WorkbookCompareResultsService : IWorkbookCompareResultsService {
        private bool connected;
        private static CompareResultsForm active;
        public bool Connect(string title,string version,string bridge,string feature,string command,string bitness) {
            return connected=title==NxHostContract.ProductTitle && version=="nx-compare-results-v1" &&
                bridge==NxHostContract.BridgeHash && feature==NxHostContract.FeatureHash &&
                command==NxHostContract.CommandHash && bitness==NxHostContract.ProcessBitness;
        }
        public void ShowResults(object rows,string reportName,string excelHwnd,object callback) {
            if(!connected) throw new InvalidOperationException("COMPARE_RESULTS_NOT_CONNECTED");
            long hwnd;
            if(!Int64.TryParse(excelHwnd,out hwnd) || callback==null || reportName==null || reportName.Length>1024) throw new ArgumentException("COMPARE_RESULTS_INPUT");
            IntPtr owner=new IntPtr(hwnd); uint pid;
            GetWindowThreadProcessId(owner,out pid);
            if(!IsWindow(owner) || pid!=(uint)Process.GetCurrentProcess().Id) throw new ArgumentException("COMPARE_RESULTS_OWNER");
            var snapshot=CompareResultsSnapshot.Parse(rows);
            CloseActive();
            var form=new CompareResultsForm(snapshot,reportName,callback,owner,(int)GetDpiForWindow(owner));
            active=form;
            form.FormClosed+=(s,e)=>{if(Object.ReferenceEquals(active,form)) active=null;};
            try { form.Show(new ExcelWindowOwner(owner)); }
            catch { form.Dispose(); if(Object.ReferenceEquals(active,form)) active=null; throw; }
        }
        public void Close() { if(connected) CloseActive(); }
        internal static void CloseActive() {
            var form=active; active=null;
            if(form!=null) { try { form.Close(); } finally { form.Dispose(); } }
        }
        internal static void KeepVisibleAfterNavigation(Form form) {
            IntPtr window=GetForegroundWindow();uint pid;
            GetWindowThreadProcessId(window,out pid);
            var name=new StringBuilder(64);GetClassName(window,name,name.Capacity);
            if(pid!=(uint)Process.GetCurrentProcess().Id || name.ToString()!="XLMAIN")return;
            // Keep the report as the lifetime owner. Do not activate/reorder that owner
            // or become global TopMost when revealing the results above a source window.
            SetWindowPos(form.Handle,IntPtr.Zero,0,0,0,0,0x0001|0x0002|0x0010|0x0200);
        }
        internal static Rectangle OwnerBounds(IntPtr owner) {
            NativeRect rect;
            return GetWindowRect(owner,out rect) ? Rectangle.FromLTRB(rect.Left,rect.Top,rect.Right,rect.Bottom) : Screen.FromHandle(owner).WorkingArea;
        }
        [StructLayout(LayoutKind.Sequential)] private struct NativeRect { public int Left,Top,Right,Bottom; }
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd,out NativeRect rect);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll",CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd,StringBuilder name,int count);
        [DllImport("user32.dll")] private static extern bool SetWindowPos(IntPtr hwnd,IntPtr after,int x,int y,int width,int height,uint flags);
        [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr hwnd);
    }

    // No workbook, range, path execution, or COM object crosses this snapshot boundary.
    internal sealed class CompareResultsSnapshot {
        internal readonly string[] Cells;
        internal int Count { get { return Cells.Length/9; } }
        private CompareResultsSnapshot(string[] cells) { Cells=cells; }
        internal string Get(int row,int col) { return Cells[row*9+col]; }
        internal static CompareResultsSnapshot Parse(object value) {
            if(value==null || value is DBNull) return new CompareResultsSnapshot(new string[0]);
            var source=value as Array;
            if(source==null || source.Rank!=2 || source.GetLength(1)!=9 || source.GetLength(0)>200000) throw new ArgumentException("COMPARE_RESULTS_SHAPE");
            int count=source.GetLength(0),r0=source.GetLowerBound(0),c0=source.GetLowerBound(1);
            var cells=new string[count*9]; long chars=0;
            for(int r=0;r<count;r++) {
                for(int c=0;c<9;c++) {
                    var text=source.GetValue(r+r0,c+c0) as string;
                    if(text==null || text.Length>131072) throw new ArgumentException("COMPARE_RESULTS_VALUE");
                    chars+=text.Length;
                    if(chars>16777216) throw new ArgumentException("COMPARE_RESULTS_SIZE");
                    cells[r*9+c]=text;
                }
                if(cells[r*9].Length<1 || cells[r*9].Length>31 || cells[r*9+1].Length>10) throw new ArgumentException("COMPARE_RESULTS_LOCATION");
                foreach(int col in new[]{7,8}) if(cells[r*9+col]!="0" && cells[r*9+col]!="1") throw new ArgumentException("COMPARE_RESULTS_SIDE");
                string reason=cells[r*9+2];
                if(reason=="기준 문서에 없음" || reason=="비교 문서에 없음") continue;
                foreach(string part in reason.Split(new[]{", "},StringSplitOptions.None))
                    if(part!="값" && part!="수식" && part!="수식/상수" && part!="서식" && part!="병합") throw new ArgumentException("COMPARE_RESULTS_REASON");
            }
            return new CompareResultsSnapshot(cells);
        }
        internal List<int> Filter(string kind) {
            if(Array.IndexOf(new[]{"전체","값","수식","서식","병합","시트"},kind)<0) throw new ArgumentException("COMPARE_RESULTS_FILTER");
            var found=new List<int>();
            for(int i=0;i<Count;i++) {
                string reason=Get(i,2);
                if(kind=="전체" || (kind=="시트" ? reason.EndsWith("문서에 없음",StringComparison.Ordinal) : reason.IndexOf(kind,StringComparison.Ordinal)>=0)) found.Add(i);
            }
            return found;
        }
        internal static string Display(string text) {
            if(text=="#EMPTY") return "(빈 셀)";
            int colon=text.IndexOf(':'); int type;
            if(colon>0 && colon<=2 && Int32.TryParse(text.Substring(0,colon),out type) && type>=0 && type<=20) return text.Substring(colon+1);
            return text;
        }
    }

    internal sealed class CompareResultsForm : Form {
        private readonly CompareResultsSnapshot snapshot;
        private object callback;
        private readonly IntPtr owner;
        private readonly double scale;
        private List<int> visible;
        private bool navigating,disposed;
        private readonly ComboBox filter=new ComboBox { DropDownStyle=ComboBoxStyle.DropDownList,AccessibleName="차이 종류" };
        private readonly ListView list=new ListView { Dock=DockStyle.Fill,View=View.Details,VirtualMode=true,FullRowSelect=true,MultiSelect=false,HideSelection=false,AccessibleName="비교 차이 목록" };
        private readonly TextBox baseDetail=DetailBox("기준 데이터의 값과 수식");
        private readonly TextBox compareDetail=DetailBox("비교 데이터의 값과 수식");
        private readonly Label status=new Label { Dock=DockStyle.Bottom,AutoEllipsis=true,TextAlign=ContentAlignment.MiddleLeft };
        private readonly Button previous=new Button { Text="이전 차이" },next=new Button { Text="다음 차이" };
        private readonly Button left=new Button { Text="기준 원본으로 이동" },right=new Button { Text="비교 원본으로 이동" };
        private readonly Timer lifetime=new Timer { Interval=500 };
        internal CompareResultsForm(CompareResultsSnapshot data,string name,object handler,IntPtr hwnd,int dpi) {
            snapshot=data;callback=handler;owner=hwnd;scale=Math.Max(96,dpi)/96.0;visible=data.Filter("전체");
            AutoScaleMode=AutoScaleMode.None; Font=new Font("맑은 고딕",9); Text="내엑셀 - 비교 결과 탐색";
            ClientSize=new Size(D(780),D(460)); MinimumSize=new Size(D(610),D(350));
            StartPosition=FormStartPosition.Manual;ShowInTaskbar=false;MinimizeBox=false;Padding=new Padding(D(10));
            int line=Font.Height+D(12),button=Math.Max(D(28),line);
            var heading=new Label { Dock=DockStyle.Top,Height=line,AutoEllipsis=true,Text=name+" · 보고서 내용을 탐색합니다." };
            var toolbar=new FlowLayoutPanel { Dock=DockStyle.Top,Height=button+D(12),WrapContents=false };
            filter.Items.AddRange(new object[]{"전체","값","수식","서식","병합","시트"});filter.Width=D(96);filter.SelectedIndex=0;
            foreach(var b in new[]{previous,next}){b.Size=new Size(D(86),button);b.Margin=new Padding(D(4));}
            toolbar.Controls.Add(filter);toolbar.Controls.Add(previous);toolbar.Controls.Add(next);
            var actions=new FlowLayoutPanel { Dock=DockStyle.Bottom,Height=button+D(12),FlowDirection=FlowDirection.RightToLeft,WrapContents=false };
            var close=new Button { Text="닫기",Size=new Size(D(76),button),Margin=new Padding(D(4)) };
            foreach(var b in new[]{left,right}){b.Size=new Size(D(145),button);b.Margin=close.Margin;}
            actions.Controls.Add(close);actions.Controls.Add(right);actions.Controls.Add(left);
            var detail=new TableLayoutPanel { Dock=DockStyle.Bottom,Height=Font.Height*6+D(18),ColumnCount=2,RowCount=2,BackColor=Color.White };
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,50));detail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,50));
            detail.RowStyles.Add(new RowStyle(SizeType.Absolute,line));detail.RowStyles.Add(new RowStyle(SizeType.Percent,100));
            detail.Controls.Add(new Label { Text="기준 데이터",Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleLeft },0,0);
            detail.Controls.Add(new Label { Text="비교 데이터",Dock=DockStyle.Fill,TextAlign=ContentAlignment.MiddleLeft },1,0);
            detail.Controls.Add(baseDetail,0,1);detail.Controls.Add(compareDetail,1,1);
            status.Height=Font.Height*2+D(10);
            Controls.Add(list);Controls.Add(detail);Controls.Add(status);Controls.Add(actions);Controls.Add(toolbar);Controls.Add(heading);
            string[] names={"시트","셀","차이","기준 값","비교 값"};int[] widths={120,70,150,200,200};
            for(int i=0;i<names.Length;i++)list.Columns.Add(names[i],D(widths[i]));
            list.RetrieveVirtualItem+=(s,e)=>{
                int row=visible[e.ItemIndex];
                var values=new string[5];for(int c=0;c<5;c++){values[c]=CompareResultsSnapshot.Display(snapshot.Get(row,c));if(values[c].Length>256)values[c]=values[c].Substring(0,256)+"…";}
                e.Item=new ListViewItem(values);
            };
            list.SelectedIndexChanged+=(s,e)=>RefreshSelection();
            filter.SelectedIndexChanged+=(s,e)=>ApplyFilter();
            previous.Click+=(s,e)=>MoveDifference(-1);next.Click+=(s,e)=>MoveDifference(1);
            left.Click+=(s,e)=>Navigate(0);right.Click+=(s,e)=>Navigate(1);
            close.Click+=(s,e)=>Close();CancelButton=close;
            lifetime.Tick+=(s,e)=>CheckLifetime();
            Shown+=(s,e)=>{
                var area=Screen.FromHandle(owner).WorkingArea;
                Size=new Size(Math.Min(Width,area.Width),Math.Min(Height,area.Height));
                var parent=WorkbookCompareResultsService.OwnerBounds(owner);
                Location=new Point(Math.Max(area.Left,Math.Min(area.Right-Width,parent.Left+(parent.Width-Width)/2)),
                    Math.Max(area.Top,Math.Min(area.Bottom-Height,parent.Top+(parent.Height-Height)/2)));
                ApplyFilter();lifetime.Start();list.Focus();
            };
        }
        private int D(int value){return (int)Math.Round(value*scale);}
        private static TextBox DetailBox(string name) {
            return new TextBox { Dock=DockStyle.Fill,ReadOnly=true,Multiline=true,ScrollBars=ScrollBars.Both,WordWrap=false,BackColor=Color.White,AccessibleName=name };
        }
        private string DetailText(int row,int side) {
            if(snapshot.Get(row,7+side)=="0")return "해당 데이터가 없습니다.";
            string value=CompareResultsSnapshot.Display(snapshot.Get(row,3+side));
            string formula=CompareResultsSnapshot.Display(snapshot.Get(row,5+side));
            return "값: "+value+(String.IsNullOrEmpty(formula) ? "" : "\r\n수식: "+formula);
        }
        private int Selected { get { return list.SelectedIndices.Count==1 ? list.SelectedIndices[0] : -1; } }
        private object Call(string method,params object[] args) {
            return callback.GetType().InvokeMember(method,BindingFlags.InvokeMethod,null,callback,args);
        }
        private void ApplyFilter() {
            list.SelectedIndices.Clear(); list.VirtualListSize=0;
            visible=snapshot.Filter(Convert.ToString(filter.SelectedItem));
            list.VirtualListSize=visible.Count;list.Invalidate();
            if(visible.Count>0) Select(0);else RefreshSelection();
        }
        private void Select(int index) {
            list.SelectedIndices.Clear();list.SelectedIndices.Add(index);list.EnsureVisible(index);RefreshSelection();
        }
        private void MoveDifference(int delta) {
            if(visible.Count==0)return;
            int index=Selected;if(index<0)index=0;else index=Math.Max(0,Math.Min(visible.Count-1,index+delta));
            Select(index);
        }
        private void RefreshSelection() {
            int index=Selected;bool selected=index>=0 && index<visible.Count;
            previous.Enabled=selected && index>0;next.Enabled=selected && index<visible.Count-1;
            left.Enabled=selected && snapshot.Get(visible[index],7)=="1";
            right.Enabled=selected && snapshot.Get(visible[index],8)=="1";
            status.Text=visible.Count+" / "+snapshot.Count+"건 · "+(selected ? (index+1)+"번째 차이" : "선택 없음")+"\n현재 원본 내용은 비교 당시와 다를 수 있습니다. 셀 값·서식은 변경하지 않습니다.";
            baseDetail.Text=selected ? DetailText(visible[index],0) : "";
            compareDetail.Text=selected ? DetailText(visible[index],1) : "";
        }
        private void Navigate(int side) {
            int index=Selected;if(navigating || index<0 || index>=visible.Count || snapshot.Get(visible[index],7+side)!="1")return;
            navigating=true;
            try {
                status.Text=Convert.ToString(Call("Navigate",visible[index]+1,side));
                WorkbookCompareResultsService.KeepVisibleAfterNavigation(this);
            }
            catch(COMException) { status.Text="Excel이 다른 작업 중입니다. 작업을 마친 뒤 다시 이동하세요."; }
            catch(TargetInvocationException) { status.Text="이동할 수 없습니다. 보고서를 다시 열어 확인하세요."; }
            catch(MissingMemberException) { Close(); }
            catch(InvalidComObjectException) { Close(); }
            finally { navigating=false; }
        }
        private void CheckLifetime() {
            if(navigating || disposed)return;
            if(!WorkbookCompareResultsService.IsWindow(owner)){Close();return;}
            try { if(!Convert.ToBoolean(Call("IsAlive")))Close(); }
            catch(COMException error) {
                // Retry only genuine Excel busy/retry-later responses, not a broken dispatch contract.
                if(error.ErrorCode!=unchecked((int)0x80010001) && error.ErrorCode!=unchecked((int)0x8001010A))Close();
            }
            catch(TargetInvocationException) { Close(); }
            catch(MissingMemberException) { Close(); }
            catch(InvalidComObjectException) { Close(); }
        }
        protected override void Dispose(bool disposing) {
            if(disposing && !disposed) {
                disposed=true;lifetime.Stop();lifetime.Dispose();
                object prior=callback;callback=null;
                if(prior!=null)try{prior.GetType().InvokeMember("ReleaseSession",BindingFlags.InvokeMethod,null,prior,new object[0]);}catch(COMException){}catch(TargetInvocationException){}catch(MissingMemberException){}catch(InvalidComObjectException){}
                // Do not FinalRelease a borrowed VBA CCW; dropping this reference is enough.
            }
            base.Dispose(disposing);
        }
    }
}
