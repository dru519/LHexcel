using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace LH.NxHost {
    [ComVisible(true), Guid("C812F023-E912-49AA-98B7-75DC86501903"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IWorkbookCompareService {
        [DispId(1)] bool Connect(string title, string version, string bridge, string feature, string command, string bitness);
        [DispId(2)] string Start(object left, object right);
        [DispId(3)] string GetStatus(string job);
        [DispId(4)] object GetResult(string job);
        [DispId(5)] int GetProgress(string job);
        [DispId(6)] void Cancel(string job);
    }
    [ComVisible(true), Guid("C812F023-E912-49AA-98B7-75DC86501904"), ProgId("LH.NxHost.WorkbookCompareService"), ClassInterface(ClassInterfaceType.None)]
    public sealed class WorkbookCompareService : IWorkbookCompareService {
        private static readonly SemaphoreSlim Capacity = new SemaphoreSlim(1,1);
        private readonly object gate = new object();
        private bool connected;
        private string id="", state="idle";
        private string[] output;
        private CancellationTokenSource cancellation;
        private int progress;
        public bool Connect(string title,string version,string bridge,string feature,string command,string bitness) {
            lock(gate) {
                if (state=="running") return false;
                return connected=title==NxHostContract.ProductTitle && version=="nx-compare-v1" &&
                    bridge==NxHostContract.BridgeHash && feature==NxHostContract.FeatureHash &&
                    command==NxHostContract.CommandHash && bitness==NxHostContract.ProcessBitness;
            }
        }
        public string Start(object left,object right) {
            lock(gate) {
                if (!connected || state=="running") throw new InvalidOperationException("COMPARE_NOT_READY");
                // Copy to plain strings on the caller/Excel thread. Workers see no COM objects.
                long characters=0;
                string[] a=Capture(left as Array,ref characters), b=Capture(right as Array,ref characters);
                if (a.Length!=b.Length) throw new InvalidDataException("COMPARE_SHAPE");
                if (!Capacity.Wait(0)) throw new InvalidOperationException("COMPARE_BUSY");
                if (cancellation!=null) cancellation.Dispose();
                cancellation=new CancellationTokenSource(TimeSpan.FromSeconds(60));
                CancellationToken token=cancellation.Token;
                id=Guid.NewGuid().ToString("D"); state="running"; output=null; progress=0;
                try {
                    Task.Factory.StartNew(() => {
                        string[] result=null;
                        string completed="succeeded";
                        try {
                            result=Compare(a,b,token,p=>Interlocked.Exchange(ref progress,p));
                        } catch(OperationCanceledException) { completed="cancelled"; }
                        catch { completed="failed"; }
                        finally {
                            // Publish terminal state only after capacity is reusable.
                            Capacity.Release();
                            lock(gate) {
                                if(token.IsCancellationRequested) completed="cancelled";
                                output=completed=="succeeded" ? result : null;
                                cancellation.Dispose(); cancellation=null;
                                state=completed;
                            }
                        }
                    }, CancellationToken.None,TaskCreationOptions.None,TaskScheduler.Default);
                } catch { state="failed"; Capacity.Release(); throw; }
                return id;
            }
        }
        public string GetStatus(string job) { lock(gate) { Require(job); return state; } }
        public object GetResult(string job) { lock(gate) { Require(job); if(state!="succeeded") throw new InvalidOperationException("COMPARE_RESULT_NOT_READY"); return output.Clone(); } }
        public int GetProgress(string job) { lock(gate) { Require(job); return Interlocked.CompareExchange(ref progress,0,0); } }
        public void Cancel(string job) { lock(gate) { Require(job); if(state=="running") cancellation.Cancel(); } }
        private void Require(string job) { if (String.IsNullOrEmpty(job) || job!=id) throw new ArgumentException("COMPARE_JOB"); }
        internal static string[] Capture(Array source,ref long characters) {
            if (source==null || source.Rank!=2 || source.GetLength(1)!=5 || source.GetLength(0)<1 || source.GetLength(0)>200000) throw new InvalidDataException("COMPARE_SHAPE");
            int rows=source.GetLength(0);
            // One immutable snapshot buffer instead of one tiny array per cell.
            var result=new string[rows*5];
            int row=source.GetLowerBound(0), col=source.GetLowerBound(1);
            for(int r=0;r<rows;r++) {
                int offset=r*5;
                for(int c=0;c<5;c++) {
                    var value=source.GetValue(r+row,c+col) as string;
                    if(value==null || value.Length>131072) throw new InvalidDataException("COMPARE_VALUE");
                    characters+=value.Length;
                    if(characters>16777216) throw new InvalidDataException("COMPARE_SIZE");
                    result[offset+c]=value;
                }
                if(result[offset]!="0" && result[offset]!="1") throw new InvalidDataException("COMPARE_FORMULA_FLAG");
            }
            return result;
        }
        internal static string[] Compare(string[] left,string[] right,CancellationToken token,Action<int> update) {
            var result=new string[left.Length/5];
            for(int i=0;i<result.Length;i++) {
                if((i&255)==0) { token.ThrowIfCancellationRequested(); update(i*100/result.Length); }
                int offset=i*5; string reason="";
                if(left[offset]!=right[offset]) reason="수식/상수";
                else if(left[offset]=="1" && left[offset+1]!=right[offset+1]) reason="수식";
                if(left[offset+2]!=right[offset+2]) reason=Append(reason,"값");
                if(left[offset+3]!=right[offset+3]) reason=Append(reason,"병합");
                if(left[offset+4]!=right[offset+4]) reason=Append(reason,"서식");
                result[i]=reason;
            }
            token.ThrowIfCancellationRequested(); update(100); return result;
        }
        private static string Append(string text,string next) { return text.Length==0 ? next : text+", "+next; }
    }
}
