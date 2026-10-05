using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace LH.NxHost {
    [ComVisible(true), Guid("2C1FC9B6-2471-4A82-9080-10331E120801"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IHwpxExportService {
        [DispId(1)] string InterfaceVersion { get; }
        [DispId(2)] bool Connect(string productTitle, string interfaceVersion, string bridgeHash, string featureHash, string commandHash, string excelBitness);
        [DispId(3)] string Start(string payloadJson, string templatePath);
        [DispId(4)] string GetStatus(string jobId);
        [DispId(5)] string GetResult(string jobId);
        [DispId(6)] string GetError(string jobId);
        [DispId(7)] void Cancel(string jobId);
        [DispId(8)] string ConnectionStatus { get; }
    }

    [ComVisible(true), Guid("1434C649-18AA-4435-B0D2-2BD81B548A01"), ProgId("LH.NxHost.HwpxExportService"), ClassInterface(ClassInterfaceType.None)]
    public sealed class HwpxExportService : IHwpxExportService {
        private static readonly SemaphoreSlim Capacity = new SemaphoreSlim(1, 1);
        private readonly object gate = new object();
        private bool connected;
        private string connectionStatus = "not-connected", id = "", state = "idle", output = "", error = "";
        private CancellationTokenSource cancellation;
        public string InterfaceVersion { get { return "nx-hwpx-v1"; } }
        public string ConnectionStatus { get { lock (gate) return connectionStatus; } }
        public bool Connect(string productTitle, string interfaceVersion, string bridgeHash, string featureHash, string commandHash, string excelBitness) {
            lock (gate) {
                if (state == "running") return false;
                connected = false;
                connectionStatus = productTitle != NxHostContract.ProductTitle ? "product-mismatch" :
                    interfaceVersion != InterfaceVersion ? "interface-mismatch" :
                    excelBitness != (Environment.Is64BitProcess ? "x64" : "x86") ? "architecture-mismatch" :
                    bridgeHash != NxHostContract.BridgeHash || featureHash != NxHostContract.FeatureHash || commandHash != NxHostContract.CommandHash ? "contract-mismatch" : "ready";
                return connected = connectionStatus == "ready";
            }
        }
        public string Start(string payloadJson, string templatePath) {
            lock (gate) {
                if (!connected) throw new InvalidOperationException("HWPX_NOT_CONNECTED");
                if (state == "running") throw new InvalidOperationException("HWPX_BUSY");
                if (String.IsNullOrEmpty(payloadJson) || payloadJson.Length > HwpxTableWriter.MaxPayloadChars) throw new InvalidDataException("INPUT_SIZE");
                if (String.IsNullOrEmpty(templatePath) || templatePath.Length > 260) throw new InvalidDataException("TEMPLATE_PATH");
                CheckLocalPath(templatePath);
                if ((File.GetAttributes(templatePath) & FileAttributes.ReparsePoint) != 0) throw new InvalidDataException("TEMPLATE_REPARSE");
                if (!Capacity.Wait(0)) throw new InvalidOperationException("HWPX_BUSY");
                try {
                    string root = OutputRoot();
                    CheckLocalPath(root);
                    Directory.CreateDirectory(root);
                    CheckLocalPath(root);
                    id = Guid.NewGuid().ToString("D"); output = Path.Combine(root, id + ".hwpx");
                    if (cancellation != null) cancellation.Dispose();
                    cancellation = new CancellationTokenSource(HwpxPictureWriter.IsPicturePayload(payloadJson) ? 120000 : 60000); state = "running"; error = "";
                    // The immutable JSON snapshot is the only workbook data available to this worker.
                    Task.Factory.StartNew(delegate { Generate(payloadJson, templatePath); }, CancellationToken.None, TaskCreationOptions.None, TaskScheduler.Default);
                    return id;
                } catch { state = "failed"; error = "START_FAILED"; Capacity.Release(); throw; }
            }
        }
        private void Generate(string json, string templatePath) {
            string temporary = output + ".part"; bool ownTemporary = false;
            string finalState = "failed", finalError = "GENERATION_FAILED";
            try {
                // Writer owns/removes a partial CreateNew file on failure; never delete a pre-existing file here.
                HwpxTableWriter.Write(json, templatePath, temporary, cancellation.Token);
                ownTemporary = true;
                lock (gate) {
                    cancellation.Token.ThrowIfCancellationRequested();
                    CheckLocalPath(Path.GetDirectoryName(output));
                    File.Move(temporary, output); ownTemporary = false;
                    finalState = "succeeded"; finalError = "";
                }
            } catch (OperationCanceledException) { finalState = "cancelled"; finalError = "CANCELLED"; }
            catch (Exception ex) { finalError = ex is InvalidDataException ? "INVALID_INPUT_OR_TEMPLATE" : "GENERATION_FAILED"; }
            finally {
                if (ownTemporary) try { File.Delete(temporary); } catch { }
                lock (gate) {
                    cancellation.Dispose(); cancellation = null;
                    Capacity.Release();
                    state = finalState; error = finalError;
                }
            }
        }
        private void CheckId(string jobId) {
            if (!connected || String.IsNullOrEmpty(id) || jobId != id) throw new InvalidOperationException("HWPX_JOB_REJECTED");
        }
        public string GetStatus(string jobId) { lock (gate) { CheckId(jobId); return state; } }
        public string GetError(string jobId) { lock (gate) { CheckId(jobId); return error; } }
        public string GetResult(string jobId) {
            lock (gate) { CheckId(jobId); if (state != "succeeded") throw new InvalidOperationException("HWPX_NOT_SUCCEEDED"); return output; }
        }
        public void Cancel(string jobId) { lock (gate) { CheckId(jobId); if (state == "running") cancellation.Cancel(); } }
        private static string OutputRoot() {
            string profile = Environment.GetEnvironmentVariable("LHEXCEL_PROFILE_ROOT");
            if (String.IsNullOrEmpty(profile)) profile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LHexcel");
            if (profile.Length > 195 || profile.IndexOf('\0') >= 0) throw new InvalidDataException("PROFILE_PATH");
            CheckLocalPath(profile);
            return Path.Combine(Path.GetFullPath(profile), "Temp", "Hwpx");
        }
        private static void CheckLocalPath(string path) {
            if (path.Length < 3 || path[1] != ':' || (path[2] != '\\' && path[2] != '/') || path.IndexOf(':', 2) >= 0) throw new InvalidDataException("LOCAL_PATH_REQUIRED");
            for (var dir = new DirectoryInfo(Path.GetFullPath(path)); dir != null; dir = dir.Parent)
                if (dir.Exists && (dir.Attributes & FileAttributes.ReparsePoint) != 0) throw new InvalidDataException("REPARSE_PATH_REJECTED");
        }
    }
}
