using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Web.Script.Serialization;
using LH.NxHost;

// Same public DTO and independent oracle for both baseline and candidate service sources.
internal static class R66CompareProbe {
    private static int checks;
    private static readonly MethodInfo Capture = typeof(WorkbookCompareService).GetMethod("Capture", BindingFlags.Static | BindingFlags.NonPublic);
    private static readonly MethodInfo Compare = typeof(WorkbookCompareService).GetMethod("Compare", BindingFlags.Static | BindingFlags.NonPublic);
    private static void Check(bool valid, string name) { if (!valid) throw new Exception(name); checks++; Console.WriteLine("PASS|" + name); }
    private static Array Rows(int rows) {
        var data = Array.CreateInstance(typeof(object), new[]{rows,5}, new[]{3,-2});
        for (int r=0;r<rows;r++) { data.SetValue("0",r+3,-2); data.SetValue("",r+3,-1); data.SetValue("5:"+r,r+3,0); data.SetValue("",r+3,1); data.SetValue("",r+3,2); }
        return data;
    }
    private static object Copy(Array data) { return Capture.Invoke(null,new object[]{data,0L}); }
    private static string[] Run(object a, object b, CancellationToken token, Action<int> progress) { return (string[])Compare.Invoke(null,new object[]{a,b,token,progress}); }
    private static string[] Oracle(Array a, Array b) {
        var result=new string[a.GetLength(0)];
        for(int i=0;i<result.Length;i++) {
            var reasons=new List<string>(); int row=i+3;
            if(!Equals(a.GetValue(row,-2),b.GetValue(row,-2))) reasons.Add("수식/상수");
            else if(Equals(a.GetValue(row,-2),"1") && !Equals(a.GetValue(row,-1),b.GetValue(row,-1))) reasons.Add("수식");
            if(!Equals(a.GetValue(row,0),b.GetValue(row,0))) reasons.Add("값");
            if(!Equals(a.GetValue(row,1),b.GetValue(row,1))) reasons.Add("병합");
            if(!Equals(a.GetValue(row,2),b.GetValue(row,2))) reasons.Add("서식");
            result[i]=String.Join(", ",reasons.ToArray());
        }
        return result;
    }
    private static void Reject(Array data,string code) {
        bool rejected=false;
        try { Copy(data); } catch(TargetInvocationException e) { rejected=e.InnerException is InvalidDataException && e.InnerException.Message==code; }
        Check(rejected,"reject "+code);
    }
    private static void Cases() {
        var a=Rows(4096); var b=Rows(4096);
        for(int i=0;i<4096;i++) {
            if(i%2==0) b.SetValue("1",i+3,-2);
            if(i%3==0) {a.SetValue("1",i+3,-2);b.SetValue("1",i+3,-2);b.SetValue("8:=SUM(A1)",i+3,-1);}
            if(i%5==0) b.SetValue("8:=untrusted",i+3,0);
            if(i%7==0) b.SetValue("$A$1:$B$1",i+3,1);
            if(i%11==0) b.SetValue("0.00",i+3,2);
        }
        var expected=Oracle(a,b); var left=Copy(a);var right=Copy(b);b.SetValue("post-capture mutation",3,0);
        int previous=-1;bool monotonic=true;
        var actual=Run(left,right,CancellationToken.None,p=>{monotonic &= p>=previous && p<=100;previous=p;});
        Check(String.Join("|",actual)==String.Join("|",expected),"4096 independent reason parity; nonzero array lower bounds; immutable snapshot");
        Check(monotonic && previous==100,"monotonic progress reaches 100");
        var cts=new CancellationTokenSource();bool cancelled=false;int last=-1;
        try {Run(left,right,cts.Token,p=>{last=p;if(p>=10)cts.Cancel();});} catch(TargetInvocationException e) {cancelled=e.InnerException is OperationCanceledException;}
        Check(cancelled && last>=10 && last<25,"mid-comparison cancellation within one 256-row stride");cts.Dispose();
        Reject(Array.CreateInstance(typeof(object),1,4),"COMPARE_SHAPE");
        var invalid=Rows(1);invalid.SetValue(new object(),3,0);Reject(invalid,"COMPARE_VALUE");
        invalid=Rows(1);invalid.SetValue("2",3,-2);Reject(invalid,"COMPARE_FORMULA_FLAG");
        invalid=Rows(1);invalid.SetValue(new string('a',131073),3,0);Reject(invalid,"COMPARE_VALUE");
        invalid=Rows(130);for(int i=0;i<130;i++)invalid.SetValue(new string('a',131072),i+3,0);Reject(invalid,"COMPARE_SIZE");
        var service=new WorkbookCompareService();
        Check(service.Connect(NxHostContract.ProductTitle,"nx-compare-v1",NxHostContract.BridgeHash,NxHostContract.FeatureHash,NxHostContract.CommandHash,NxHostContract.ProcessBitness),"public COM handshake unchanged");
        string job=service.Start(a,a);var deadline=DateTime.UtcNow.AddSeconds(10);
        while(service.GetStatus(job)=="running" && DateTime.UtcNow<deadline)Thread.Sleep(1);
        Check(service.GetStatus(job)=="succeeded" && ((string[])service.GetResult(job)).Length==4096,"public DTO result shape unchanged");
    }
    private static object Measure(int rows) {
        var data=Rows(rows);Copy(data); // JIT warm-up, not included.
        var milliseconds=new double[5];var bytes=new long[5];
        for(int i=0;i<5;i++) MeasureOnce(data,rows,i,out milliseconds[i],out bytes[i]);
        Array.Sort(milliseconds);Array.Sort(bytes);
        return new {rows=rows,repeats=5,median_capture_compute_ms=milliseconds[2],median_retained_snapshot_bytes=bytes[2],times_ms=milliseconds,snapshot_bytes=bytes};
    }
    [System.Runtime.CompilerServices.MethodImpl(System.Runtime.CompilerServices.MethodImplOptions.NoInlining)]
    private static void MeasureOnce(Array data,int rows,int repeat,out double milliseconds,out long bytes) {
        GC.Collect();GC.WaitForPendingFinalizers();GC.Collect();long before=GC.GetTotalMemory(true);
        var clock=Stopwatch.StartNew();var a=Copy(data);var b=Copy(data);clock.Stop();
        bytes=GC.GetTotalMemory(true)-before;double capture=clock.Elapsed.TotalMilliseconds;
        clock.Restart();var result=Run(a,b,CancellationToken.None,p=>{});clock.Stop();
        milliseconds=capture+clock.Elapsed.TotalMilliseconds;
        Check(result.Length==rows && Array.TrueForAll(result,v=>v==""),"benchmark equal "+rows+" repeat "+repeat);
        GC.KeepAlive(a);GC.KeepAlive(b);GC.KeepAlive(data);
    }
    private static int Main(string[] args) {
        try {Cases();var metrics=new[]{Measure(10000),Measure(200000)};
            File.WriteAllText(args[0],new JavaScriptSerializer().Serialize(new {status="PASS",checks=checks,metrics=metrics,process_bitness=IntPtr.Size*8}));
            Console.WriteLine("PASS|checks="+checks);return 0;
        } catch(Exception e) {Console.WriteLine("FAIL|"+e);return 1;}
    }
}
