using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Win32;

namespace LH.NxSetup
{
    internal sealed partial class NativeInstaller
    {
        public string InstallDirectory { get { return runtime; } }
        public string StartupFile { get { return xlstart; } }
        internal static string RegistrationSummary()
        {
            var lines=new List<string>();
            foreach(var view in Views)using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                using(var key=root.OpenSubKey(@"Software\Classes\CLSID\"+Ids[0]+@"\InprocServer32")){
                    string path=key==null ? null:key.GetValue("CodeBase") as string;
                    lines.Add((view==RegistryView.Registry32 ? "32비트 DLL: ":"64비트 DLL: ")+(path??"등록 없음"));
                }
            return String.Join("\r\n",lines.ToArray());
        }
        string StateRoot { get { return Path.GetDirectoryName(Path.GetDirectoryName(receipt)); } }

        static XElement ReadXml(string path)
        {
            Package.RejectReparseAncestors(path);
            if(new FileInfo(path).Length>1024*1024) throw new InvalidDataException("설치 기록 크기가 올바르지 않습니다.");
            using(var reader=XmlReader.Create(path,new XmlReaderSettings {
                DtdProcessing=DtdProcessing.Prohibit, XmlResolver=null })) return XElement.Load(reader);
        }

        // Only fixed package paths and receipt-verified bytes are used to recover old versions.
        bool HasActiveEntryPoint(string oldVersion)
        {
            string release=oldVersion.Split('_')[1];
            string oldStartup=Path.Combine(Path.GetDirectoryName(xlstart),"내엑셀 "+oldVersion+".xlam");
            Package.RejectReparseAncestors(oldStartup);
            if(File.Exists(oldStartup)) return true;
            string oldRuntime=Path.Combine(Path.GetDirectoryName(runtime),release);
            foreach(var view in Views) using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                foreach(string id in Ids) using(var registered=key.OpenSubKey(Key(@"Software\Classes\CLSID\"+id+@"\InprocServer32",view))) {
                    string location=registered==null ? null : registered.GetValue("CodeBase") as string;
                    Uri uri;
                    if(location!=null && Uri.TryCreate(location,UriKind.Absolute,out uri) && uri.IsFile &&
                       Path.GetFullPath(uri.LocalPath).StartsWith(oldRuntime+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase)) return true;
                }
            return false;
        }

        NativeInstaller PreviousInstallation()
        {
            if(!Directory.Exists(StateRoot)) return null;
            Package.RejectReparseAncestors(StateRoot);
            NativeInstaller found=null;
            foreach(string directory in Directory.GetDirectories(StateRoot)) {
                string release=Path.GetFileName(directory);
                if(!System.Text.RegularExpressions.Regex.IsMatch(release,@"\Ar[0-9]+\z")) continue;
                string path=Path.Combine(directory,"native-install.xml");
                if(!File.Exists(path)) continue;
                XElement doc=ReadXml(path);
                string state=(string)doc.Attribute("state");
                if(state=="RESTORED") continue;
                string oldVersion="v0.01_"+release;
                if(oldVersion==version) continue;
                // Historical receipt/DLL remnants are not an installed entry point.
                // Leave them untouched instead of interpreting their stale state as current.
                if(!HasActiveEntryPoint(oldVersion)) { log(oldVersion+"의 과거 설치 기록은 보존하고 건너뜁니다."); continue; }
                if(found!=null) throw new IOException("활성 설치 기록이 여러 개입니다. 설치 상태를 먼저 확인하세요.");
                if((string)doc.Attribute("version")!=oldVersion ||
                   (state!="ACTIVE" && state!="PREPARING" && state!="REMOVING"))
                    throw new IOException("이전 설치 기록을 확인할 수 없습니다.");
                int oldNumber, newNumber;
                if(!Int32.TryParse(release.Substring(1),out oldNumber) ||
                   !Int32.TryParse(version.Split('_')[1].Substring(1),out newNumber) || oldNumber>=newNumber)
                    throw new IOException("현재 설치보다 이전 버전으로 업데이트할 수 없습니다.");
                string oldHash=(string)doc.Attribute("packageSha256");
                if(!ValidHash(oldHash)) throw new IOException("이전 설치 해시가 올바르지 않습니다.");
                string cache=test ? Path.Combine(testHome,"cache",oldVersion,oldHash) : Package.CachedPath(oldVersion,oldHash);
                var bytes=new Dictionary<string,byte[]>(StringComparer.OrdinalIgnoreCase);
                foreach(string name in new[]{"Product.xlam","NxHost32.dll","NxHost64.dll","NxCore32.dll","NxCore64.dll"}) {
                    string cached=Path.Combine(cache,"payload",name);
                    Package.RejectReparseAncestors(cached);
                    if(!File.Exists(cached) || new FileInfo(cached).Length>64L*1024*1024)
                        throw new IOException("이전 설치 패키지가 없습니다. 기존 EXE로 제거 후 설치하세요.");
                    bytes.Add("payload/"+name,File.ReadAllBytes(cached));
                }
                found=new NativeInstaller(oldVersion,oldHash,cache,bytes,log,testHome,prefix,null);
                found.ReadReceipt(); // Fixed destinations and every cached payload hash must match.
                found.ValidateOwned(true);
            }
            return found;
        }

        public string MaintenanceStatus()
        {
            if(HasPackageRefresh()) return version+"의 수정된 설치 패키지로 업데이트할 수 있습니다.";
            if(File.Exists(receipt) && ReadReceipt()!="RESTORED") return version+" 설치됨 · 복구 또는 제거 가능";
            NativeInstaller old=PreviousInstallation();
            return old==null ? "새로 설치할 수 있습니다." : old.version+" → "+version+" 업데이트 가능";
        }

        public string InstalledVersion()
        {
            if(HasPackageRefresh()) { RecordedPackage(); return version; }
            if(File.Exists(receipt) && ReadReceipt()!="RESTORED") return version;
            NativeInstaller old=PreviousInstallation();
            return old==null ? null : old.version;
        }

        public void RemoveInstalledVersion()
        {
            if(HasPackageRefresh()) {RecordedPackage().Uninstall();return;}
            if(File.Exists(receipt) && ReadReceipt()!="RESTORED") { Uninstall(); return; }
            NativeInstaller old=PreviousInstallation();
            if(old!=null) { old.Uninstall(); return; }
            Uninstall();
        }

        public bool HasPackageRefresh()
        {
            if(!File.Exists(receipt)) return false;
            var doc=ReadXml(receipt);
            return (string)doc.Attribute("state")!="RESTORED" && (string)doc.Attribute("packageSha256")!=packageHash;
        }

        NativeInstaller RecordedPackage()
        {
            var doc=ReadXml(receipt);
            string oldHash=(string)doc.Attribute("packageSha256");
            if((string)doc.Attribute("version")!=version || !ValidHash(oldHash)) throw new IOException("기존 패키지 기록이 올바르지 않습니다.");
            string cache=test ? Path.Combine(testHome,"cache",version,oldHash) : Package.CachedPath(version,oldHash);
            var bytes=new Dictionary<string,byte[]>(StringComparer.OrdinalIgnoreCase);
            foreach(string name in new[]{"Product.xlam","NxHost32.dll","NxHost64.dll","NxCore32.dll","NxCore64.dll"}) {
                string path=Path.Combine(cache,"payload",name);Package.RejectReparseAncestors(path);
                if(!File.Exists(path) || new FileInfo(path).Length>64L*1024*1024)throw new IOException("기존 설치 캐시가 없습니다. 기존 EXE로 제거한 뒤 설치하세요.");
                bytes.Add("payload/"+name,File.ReadAllBytes(path));
            }
            var old=new NativeInstaller(version,oldHash,cache,bytes,log,testHome,prefix,null);
            old.ReadReceipt();old.ValidateOwned(true);return old;
        }

        sealed class Snapshot
        {
            public readonly Dictionary<string,byte[]> Files=new Dictionary<string,byte[]>();
            public readonly Dictionary<RegistryView,Dictionary<string,Dictionary<string,object>>> Registry
                =new Dictionary<RegistryView,Dictionary<string,Dictionary<string,object>>>();
        }

        Snapshot CaptureMaintenance()
        {
            var snapshot=new Snapshot();
            string backup=Path.Combine(Path.GetDirectoryName(receipt),"maintenance-"+DateTime.UtcNow.ToString("yyyyMMddHHmmss")+"-"+Guid.NewGuid().ToString("N"));
            Package.RejectReparseAncestors(backup);
            Directory.CreateDirectory(backup);
            int i=0;
            var manifest=new XElement("maintenance-backup",new XAttribute("version",version));
            foreach(string path in Destinations().Keys.Concat(new[]{receipt})) {
                byte[] bytes=File.Exists(path) ? File.ReadAllBytes(path) : null;
                snapshot.Files.Add(path,bytes);
                string name=(i++).ToString()+".backup";
                manifest.Add(new XElement("file",new XAttribute("path",path),new XAttribute("present",bytes!=null),
                    new XAttribute("backup",name),new XAttribute("sha256",bytes==null ? "" : Package.Hash(bytes))));
                if(bytes!=null) File.WriteAllBytes(Path.Combine(backup,name),bytes);
            }
            foreach(var view in Views) using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)) {
                var rows=new Dictionary<string,Dictionary<string,object>>();
                foreach(var row in RegistryRows(view)) using(var current=key.OpenSubKey(Key(row.Key,view))) {
                    if(current==null) continue;
                    var values=new Dictionary<string,object>();
                    foreach(string name in current.GetValueNames()) values.Add(name,current.GetValue(name));
                    rows.Add(row.Key,values);
                }
                snapshot.Registry.Add(view,rows);
            }
            SaveRegistryBackup();
            manifest.Save(Path.Combine(backup,"manifest.xml"));
            log("복구용 파일 보관: "+backup);
            return snapshot;
        }

        void RestoreMaintenance(Snapshot snapshot)
        {
            foreach(var item in snapshot.Files) {
                Package.RejectReparseAncestors(item.Key);
                if(item.Value==null) { if(File.Exists(item.Key)) File.Delete(item.Key); }
                else {
                    Directory.CreateDirectory(Path.GetDirectoryName(item.Key));
                    File.WriteAllBytes(item.Key,item.Value);
                }
            }
            foreach(var view in Views) using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)) {
                foreach(string root in Roots(view)) key.DeleteSubKeyTree(Key(root,view),false);
                foreach(var row in snapshot.Registry[view]) using(var child=key.CreateSubKey(Key(row.Key,view)))
                    foreach(var value in row.Value) child.SetValue(value.Key,value.Value,
                        value.Value is int ? RegistryValueKind.DWord : RegistryValueKind.String);
            }
        }

        void RestoreMissing()
        {
            // A changed file or foreign registry value is rejected before mutation.
            ValidateRepair();
            foreach(var file in Destinations()) if(!File.Exists(file.Key)) {
                Directory.CreateDirectory(Path.GetDirectoryName(file.Key));
                using(var stream=new FileStream(file.Key,FileMode.CreateNew,FileAccess.Write,FileShare.None))
                    stream.Write(file.Value,0,file.Value.Length);
                Hit("repair-file");
            }
            foreach(var view in Views) using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                foreach(var row in RegistryRows(view)) using(var child=key.CreateSubKey(Key(row.Key,view)))
                    foreach(var value in row.Value) child.SetValue(value.Key,value.Value,
                        value.Value is int ? RegistryValueKind.DWord : RegistryValueKind.String);
            Hit("repair-registry");
            ValidateOwned(false); SaveReceipt("ACTIVE");
        }

        void ValidateRepair()
        {
            repairValidation=true;
            try { ValidateOwned(true); } finally { repairValidation=false; }
        }

        public void Repair()
        {
            using(var mutex=new Mutex(false,@"Local\LHExcel.Setup.Install")) {
                bool owned=false;
                try {
                    try { owned=mutex.WaitOne(0); } catch(AbandonedMutexException) { owned=true; }
                    if(!owned) throw new IOException("다른 내엑셀 설치가 진행 중입니다.");
                    RequireClosed(); Package.VerifyDirectory(packageRoot,files);
                    if(!File.Exists(receipt) || ReadReceipt()=="RESTORED") throw new IOException("복구할 설치 기록이 없습니다.");
                    ValidateRepair();
                    Snapshot before=CaptureMaintenance();
                    try { RestoreMissing(); }
                    catch { RestoreMaintenance(before); throw; }
                    log("누락된 설치 파일과 등록을 복구했습니다. 사용자 설정은 유지됩니다.");
                } finally { if(owned) mutex.ReleaseMutex(); }
            }
        }

        public void Update()
        {
            using(var mutex=new Mutex(false,@"Local\LHExcel.Setup.Install")) {
                bool owned=false;
                try {
                    try { owned=mutex.WaitOne(0); } catch(AbandonedMutexException) { owned=true; }
                    if(!owned) throw new IOException("다른 내엑셀 설치가 진행 중입니다.");
                    RequireClosed(); Package.VerifyDirectory(packageRoot,files);
                    NativeInstaller old=HasPackageRefresh() ? RecordedPackage() : null;
                    if(old==null && File.Exists(receipt) && ReadReceipt()!="RESTORED") { Repair(); return; }
                    if(old==null)old=PreviousInstallation();
                    if(old==null) { Install(); return; }
                    // Reject an unrelated startup copy before taking down the old version.
                    if(old.version!=version && Directory.Exists(runtime)) throw new IOException("새 버전 경로가 이미 있습니다.");
                    foreach(string path in Directory.Exists(Path.GetDirectoryName(xlstart)) ?
                        Directory.GetFiles(Path.GetDirectoryName(xlstart),"내엑셀*.xlam") : new string[0])
                        if(!String.Equals(path,old.xlstart,StringComparison.OrdinalIgnoreCase))
                            throw new IOException("다른 XLSTART 내엑셀 파일이 있습니다. 먼저 확인하세요.");
                    Snapshot previous=old.CaptureMaintenance();
                    log(old.version+"에서 "+version+"로 업데이트합니다.");
                    try {
                        old.Cleanup();
                        Hit("upgrade-removed");
                        Install();
                    } catch(Exception error) {
                        // Install's own rollback removes only the new version; preserve old evidence.
                        try {
                            if(File.Exists(receipt) && ReadReceipt()!="RESTORED") Cleanup();
                            old.RestoreMaintenance(previous);
                        } catch(Exception recovery) {
                            throw new IOException("업데이트 오류: "+error.Message+" / 자동 복원 미완료: "+recovery.Message+". 복구용 파일과 설치 기록을 보존했습니다.",error);
                        }
                        throw new IOException("업데이트 실패로 이전 설치를 복원했습니다: "+error.Message,error);
                    }
                    log("업데이트 완료. 다음 Excel 시작부터 "+version+"를 불러옵니다.");
                } finally { if(owned) mutex.ReleaseMutex(); }
            }
        }
    }
}
