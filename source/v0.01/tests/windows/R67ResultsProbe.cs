using System;
using System.Reflection;
using System.Windows.Forms;
using LH.NxHost;

public sealed class ResultsCallback {
    public int Released, Navigated, LastIndex, LastSide, AliveCalls;
    public bool Alive=true;
    public bool IsAlive(){AliveCalls++;return Alive;}
    public string Navigate(int index,int side){Navigated++;LastIndex=index;LastSide=side;return index+":"+side;}
    public void ReleaseSession(){Released++;}
}
public static class R67ResultsProbe {
    static int count;
    static void Check(bool value,string name){if(!value)throw new Exception(name);count++;Console.WriteLine("PASS|"+name);}
    static void Reject(Action action,string name){try{action();}catch(ArgumentException){Check(true,name);return;}throw new Exception("accepted "+name);}
    static Array Rows(int count=3) {
        var data=Array.CreateInstance(typeof(object),new[]{count,9},new[]{3,2});
        for(int r=0;r<count;r++) {
            string[] row={"검토","B2",r%3==0?"수식/상수, 값":r%3==1?"서식":"병합","8:=1+1","5:2","8:=A1","8:=B1","1","1"};
            for(int c=0;c<9;c++)data.SetValue(row[c],r+3,c+2);
        }
        return data;
    }
    [STAThread] public static int Main() {
        try {
            var input=Rows();var data=CompareResultsSnapshot.Parse(input);
            Check(data.Count==3 && data.Get(0,0)=="검토","safearray-lower-bounds");
            input.SetValue("changed",3,2);Check(data.Get(0,0)=="검토","immutable-capture");
            Check(data.Filter("전체").Count==3,"filter-all");
            Check(data.Filter("수식").Count==1 && data.Filter("수식")[0]==0,"formula-constant-filter");
            Check(data.Filter("값")[0]==0 && data.Filter("서식")[0]==1 && data.Filter("병합")[0]==2,"filter-stable-index");
            Check(data.Filter("시트").Count==0,"empty-filter");
            Check(CompareResultsSnapshot.Parse(null).Count==0,"empty-report");
            Reject(()=>CompareResultsSnapshot.Parse(new string[3]),"reject-rank");
            Reject(()=>CompareResultsSnapshot.Parse(new string[1,8]),"reject-columns");
            var invalid=Rows();invalid.SetValue(2,3,5);Reject(()=>CompareResultsSnapshot.Parse(invalid),"reject-non-string");
            invalid=Rows();invalid.SetValue("2",3,9);Reject(()=>CompareResultsSnapshot.Parse(invalid),"reject-side");
            invalid=Rows();invalid.SetValue("run command",3,4);Reject(()=>CompareResultsSnapshot.Parse(invalid),"reject-reason");
            invalid=Rows();invalid.SetValue(new string('x',131073),3,5);Reject(()=>CompareResultsSnapshot.Parse(invalid),"reject-cell-size");
            invalid=Rows(130);for(int r=3;r<133;r++)invalid.SetValue(new string('x',131072),r,5);
            Reject(()=>CompareResultsSnapshot.Parse(invalid),"reject-total-size");
            Check(CompareResultsSnapshot.Parse(Rows(200000)).Count==200000,"max-rows");
            Reject(()=>CompareResultsSnapshot.Parse(new object[200001,9]),"reject-max-overflow");
            Reject(()=>data.Filter("unknown"),"reject-unknown-filter");
            Check(CompareResultsSnapshot.Display("8:=SUM(A1)")=="=SUM(A1)" && CompareResultsSnapshot.Display("#EMPTY")=="(빈 셀)","display-literal");
            Check(CompareResultsSnapshot.Display("https://example.invalid")=="https://example.invalid","display-does-not-interpret-uri");
            var callback=new ResultsCallback();
            var form=new CompareResultsForm(data,"보고서.xlsx",callback,IntPtr.Zero,192);
            Check(form.Font.Name=="맑은 고딕" && Math.Abs(form.Font.SizeInPoints-9)<0.1,"font-9pt-single-scale");
            Check(!form.TopMost && !form.ShowInTaskbar && form.CancelButton!=null,"owned-window-policy");
            Check(callback.Navigated==0,"no-automatic-navigation");
            var listHandle=((ListView)typeof(CompareResultsForm).GetField("list",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form)).Handle;
            var combo=(ComboBox)typeof(CompareResultsForm).GetField("filter",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form);
            combo.SelectedItem="서식";
            var baseDetail=(TextBox)typeof(CompareResultsForm).GetField("baseDetail",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form);
            var compareDetail=(TextBox)typeof(CompareResultsForm).GetField("compareDetail",BindingFlags.NonPublic|BindingFlags.Instance).GetValue(form);
            Check(baseDetail.ReadOnly && compareDetail.ReadOnly && baseDetail.BackColor==System.Drawing.Color.White && compareDetail.BackColor==System.Drawing.Color.White,"side-by-side-read-only-white");
            Check(baseDetail.Text=="값: =1+1\r\n수식: =A1" && compareDetail.Text=="값: 2\r\n수식: =B1","side-by-side-filtered-values");
            combo.SelectedItem="시트";
            Check(baseDetail.Text=="" && compareDetail.Text=="","empty-filter-clears-detail");
            combo.SelectedItem="서식";
            typeof(CompareResultsForm).GetMethod("Navigate",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(form,new object[]{1});
            Check(callback.Navigated==1 && callback.LastIndex==2 && callback.LastSide==1,"filtered-row-dispatch-original-index:"+callback.Navigated+":"+callback.LastIndex+":"+callback.LastSide);
            form.Dispose();form.Dispose();Check(callback.Released==1,"release-exactly-once");
            using(var owner=new Form()) {
                var hwnd=owner.Handle;
                var invalidCallback=new object();
                var broken=new CompareResultsForm(data,"보고서.xlsx",invalidCallback,hwnd,96);
                typeof(CompareResultsForm).GetMethod("CheckLifetime",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(broken,null);
                broken.Dispose();Check(broken.IsDisposed,"missing-callback-fails-closed");
                var handler=new ResultsCallback { Alive=false };
                var expired=new CompareResultsForm(data,"보고서.xlsx",handler,hwnd,96);
                typeof(CompareResultsForm).GetMethod("CheckLifetime",BindingFlags.NonPublic|BindingFlags.Instance).Invoke(expired,null);
                expired.Dispose();Check(handler.AliveCalls==1 && handler.Released==1,"lifetime-check-and-release");
            }
            Console.WriteLine("PASS|total="+count);return 0;
        } catch(Exception error){Console.Error.WriteLine(error);return 1;}
    }
}
