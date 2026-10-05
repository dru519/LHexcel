using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Xml.Linq;
using Microsoft.Win32;

namespace LH.NxSetup
{
    public sealed class NxAssemblyIdentityReader : MarshalByRefObject
    {
        public string[] Read(byte[] bytes)
        {
            AssemblyName name=Assembly.ReflectionOnlyLoad(bytes).GetName();
            return new[]{name.FullName,name.Version.ToString()};
        }
    }

    internal sealed partial class NativeInstaller
    {
        readonly string version, packageHash, packageRoot, runtime, xlstart, receipt, prefix;
        readonly Dictionary<string, byte[]> files;
        readonly Action<string> log;
        readonly Action<string> fault;
        readonly bool test;
        readonly string testHome;
        bool repairValidation;
        readonly Dictionary<RegistryView, string[]> assemblies;
        static readonly string[] Ids = {
            "{4A4B4F02-76A6-4F22-BD4B-85E0308EC91D}", "{9A90294C-81C8-4AEE-A2D8-14B078A61CD2}",
            "{A13D5ED6-B0F2-43C0-9499-E53E6CE38DC0}", "{1434C649-18AA-4435-B0D2-2BD81B548A01}",
            "{C812F023-E912-49AA-98B7-75DC86501902}", "{C812F023-E912-49AA-98B7-75DC86501904}", "{C812F023-E912-49AA-98B7-75DC86501906}" };
        static readonly string[] Progs = {"LH.NxHost.Connect", "LH.NxHost.BridgeService", "LH.NxHost.NavigatorPane", "LH.NxHost.HwpxExportService", "LH.NxHost.PicturePreviewService", "LH.NxHost.WorkbookCompareService", "LH.NxHost.WorkbookCompareResultsService"};
        static readonly string[] Classes = {"LH.NxHost.NxHostAddIn", "LH.NxHost.BridgeService", "LH.NxHost.NavigatorPane", "LH.NxHost.HwpxExportService", "LH.NxHost.PicturePreviewService", "LH.NxHost.WorkbookCompareService", "LH.NxHost.WorkbookCompareResultsService"};
        static readonly RegistryView[] Views = Environment.Is64BitOperatingSystem ?
            new[] {RegistryView.Registry32, RegistryView.Registry64} : new[] {RegistryView.Registry32};
        const string Addin = @"Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect";

        public NativeInstaller(string version, string hash, string root, Dictionary<string, byte[]> files, Action<string> log)
            : this(version, hash, root, files, log, null, null, null) {}

        // Only linked test executable can select an isolated layout; no production CLI override.
        internal NativeInstaller(string version, string hash, string root, Dictionary<string, byte[]> files,
            Action<string> log, string testRoot, string testPrefix, Action<string> fault)
        {
            this.version=version; packageHash=hash; packageRoot=root; this.files=files; this.log=log; this.fault=fault;
            assemblies=new Dictionary<RegistryView,string[]>();
            assemblies[RegistryView.Registry32]=ReadAssemblyIdentity(files["payload/NxHost32.dll"]);
            if(Environment.Is64BitOperatingSystem)
                assemblies[RegistryView.Registry64]=ReadAssemblyIdentity(files["payload/NxHost64.dll"]);
            test=testRoot!=null; testHome=testRoot; prefix=test ? testPrefix : "";
            if (test && (String.IsNullOrEmpty(prefix) || !prefix.StartsWith(@"Software\LHExcelSetupTests\")))
                throw new ArgumentException("Invalid test registry namespace");
            string release=version.Split('_')[1];
            runtime=Path.Combine(testRoot ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LHexcel", "NxHost", release);
            string appdata=testRoot ?? Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            xlstart=Path.Combine(appdata, "Microsoft", "Excel", "XLSTART", "내엑셀 " + version + ".xlam");
            receipt=Path.Combine(testRoot ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "LHexcel", "NxHost", "state", release, "native-install.xml");
        }

        string Key(string path, RegistryView view) { return prefix + (test ? view.ToString()+"\\" : "") + path; }
        static string[] ReadAssemblyIdentity(byte[] bytes)
        {
            AppDomain domain=AppDomain.CreateDomain("NxSetup.Metadata."+Guid.NewGuid().ToString("N"));
            try {
                var reader=(NxAssemblyIdentityReader)domain.CreateInstanceFromAndUnwrap(
                    Assembly.GetExecutingAssembly().Location,typeof(NxAssemblyIdentityReader).FullName);
                return reader.Read(bytes);
            } finally { AppDomain.Unload(domain); }
        }
        bool HasAppsEntry { get { return Int32.Parse(version.Split('_')[1].Substring(1))>=93; } }
        string AppsKey { get { return @"Software\Microsoft\Windows\CurrentVersion\Uninstall\"+
            (Int32.Parse(version.Split('_')[1].Substring(1))>=97 ? "LHExcel" : "LHExcel-"+version); } }
        bool AppsEntryFor(RegistryView view) { return HasAppsEntry && view==(Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32); }
        string[] Roots(RegistryView view) { return Ids.Select(id => @"Software\Classes\CLSID\"+id)
            .Concat(Progs.Select(p => @"Software\Classes\"+p)).Concat(new[]{Addin})
            .Concat(AppsEntryFor(view) ? new[]{AppsKey} : new string[0]).ToArray(); }
        Dictionary<string, byte[]> Destinations()
        {
            return new Dictionary<string, byte[]> {
                {Path.Combine(runtime,"x86","NxHost32.dll"),files["payload/NxHost32.dll"]},
                {Path.Combine(runtime,"x64","NxHost64.dll"),files["payload/NxHost64.dll"]},
                {Path.Combine(runtime,"x86","NxCore32.dll"),files["payload/NxCore32.dll"]},
                {Path.Combine(runtime,"x64","NxCore64.dll"),files["payload/NxCore64.dll"]},
                {xlstart,files["payload/Product.xlam"]} };
        }
        void RequireClosed()
        {
            if (test) return;
            var processes=Process.GetProcessesByName("EXCEL");
            bool running=processes.Length!=0;
            foreach (var process in processes) process.Dispose();
            if (running) throw new InvalidOperationException("Excel 문서를 저장하고 모든 Excel 창을 닫아 주세요.");
        }
        public static ushort Machine(string path)
        {
            using(var reader=new BinaryReader(File.OpenRead(path))) {
                if(reader.ReadUInt16()!=0x5A4D) throw new InvalidDataException("Invalid executable");
                reader.BaseStream.Position=60; int offset=reader.ReadInt32();
                if(offset<64 || offset>reader.BaseStream.Length-6) throw new InvalidDataException("Invalid PE offset");
                reader.BaseStream.Position=offset;
                if(reader.ReadUInt32()!=0x4550) throw new InvalidDataException("Invalid PE signature");
                return reader.ReadUInt16();
            }
        }
        internal static string FindExcel()
        {
            foreach(var hive in new[]{RegistryHive.CurrentUser,RegistryHive.LocalMachine})
                foreach(var view in Views)
                    using(var key=RegistryKey.OpenBaseKey(hive,view))
                    using(var app=key.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")) {
                        string path=app==null ? null : app.GetValue("") as string;
                        if(path!=null && File.Exists(path.Trim('"'))) return path.Trim('"');
                    }
            throw new InvalidOperationException("설치된 Excel 실행 파일을 찾을 수 없습니다.");
        }
        public string Check()
        {
            if(HasPackageRefresh()) {RecordedPackage().Check();return "같은 버전의 수정된 패키지입니다. 업데이트를 선택하세요.";}
            Package.VerifyDirectory(packageRoot,files);
            if(!test) {
                string excel=FindExcel(); ushort machine=Machine(excel);
                if(machine!=0x14c && machine!=0x8664) throw new InvalidOperationException("지원 Excel 아키텍처: x86/x64");
                log("Excel: "+excel+" / "+(machine==0x14c ? "32비트" : "64비트"));
            }
            foreach(var path in Destinations().Keys) Package.RejectReparseAncestors(Path.GetDirectoryName(path));
            Package.RejectReparseAncestors(Path.GetDirectoryName(receipt));
            if(File.Exists(receipt)) {
                string state=ReadReceipt();
                if(state=="ACTIVE") { ValidateOwned(false); return "설치 파일과 등록 상태가 정상입니다. 실제 연결은 Excel에서 확인하세요."; }
                if(state=="PREPARING" || state=="REMOVING") return "중단된 설치 기록이 있습니다. 제거를 실행해 복구할 수 있습니다.";
            }
            return "설치 환경 확인 완료. 이 EXE로 설치한 활성 기록은 없습니다.";
        }
        void SaveReceipt(string state)
        {
            string directory=Path.GetDirectoryName(receipt);
            Directory.CreateDirectory(directory);
            string temp=receipt+"."+Guid.NewGuid().ToString("N")+".tmp";
            var document=new XElement("installation",new XAttribute("version",version),
                new XAttribute("packageSha256",packageHash),new XAttribute("state",state),
                new XAttribute("priorState","absent"));
            foreach(var file in Destinations()) document.Add(new XElement("file",
                new XAttribute("path",file.Key),new XAttribute("sha256",Package.Hash(file.Value))));
            document.Save(temp);
            if(File.Exists(receipt)) File.Replace(temp,receipt,null); else File.Move(temp,receipt);
        }
        string ReadReceipt()
        {
            if((File.GetAttributes(receipt)&FileAttributes.ReparsePoint)!=0) throw new IOException("Invalid receipt path");
            var settings=new System.Xml.XmlReaderSettings { DtdProcessing=System.Xml.DtdProcessing.Prohibit, XmlResolver=null };
            XElement doc;
            using(var reader=System.Xml.XmlReader.Create(receipt,settings)) doc=XElement.Load(reader);
            string state=(string)doc.Attribute("state");
            bool removed=state=="RESTORED";
            if(state!="ACTIVE" && state!="PREPARING" && state!="REMOVING" && !removed)
                throw new InvalidDataException("Invalid installation state");
            string priorHash=(string)doc.Attribute("packageSha256");
            if((string)doc.Attribute("version")!=version || !ValidHash(priorHash) || (!removed && priorHash!=packageHash) ||
               (string)doc.Attribute("priorState")!="absent") throw new InvalidDataException("다른 설치 패키지의 기록입니다. 해당 설치 파일을 사용하세요.");
            var expected=Destinations();
            var rows=doc.Elements("file").ToArray();
            if(rows.Length!=expected.Count || rows.Select(r=>(string)r.Attribute("path")).Distinct().Count()!=rows.Length)
                throw new InvalidDataException("Invalid installation receipt");
            foreach(var row in rows) {
                string path=(string)row.Attribute("path");
                string hash=(string)row.Attribute("sha256");
                if(path==null || !expected.ContainsKey(path) || !ValidHash(hash) || (!removed && hash!=Package.Hash(expected[path])))
                    throw new InvalidDataException("Installation receipt mismatch");
            }
            return state;
        }
        static bool ValidHash(string value)
        {
            return value!=null && value.Length==64 && value.All(c => (c>='0' && c<='9') || (c>='a' && c<='f'));
        }
        Dictionary<string,Dictionary<string,object>> RegistryRows(RegistryView view)
        {
            string dll=Path.Combine(runtime,view==RegistryView.Registry32 ? "x86" : "x64",
                view==RegistryView.Registry32 ? "NxHost32.dll" : "NxHost64.dll");
            // Registration identity comes from the manifest-verified embedded bytes, never an
            // installed or cached DLL that may have been quarantined or replaced.
            string[] assembly=assemblies[view];
            var rows=new Dictionary<string,Dictionary<string,object>>(StringComparer.OrdinalIgnoreCase);
            for(int i=0;i<Ids.Length;i++) {
                string clsid=@"Software\Classes\CLSID\"+Ids[i], prog=@"Software\Classes\"+Progs[i];
                rows[prog]=new Dictionary<string,object>{{"",Classes[i]}};
                rows[prog+@"\CLSID"]=new Dictionary<string,object>{{"",Ids[i]}};
                rows[clsid]=new Dictionary<string,object>{{"",Classes[i]}};
                rows[clsid+@"\ProgId"]=new Dictionary<string,object>{{"",Progs[i]}};
                rows[clsid+@"\Implemented Categories"]=new Dictionary<string,object>();
                rows[clsid+@"\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"]=new Dictionary<string,object>();
                var managed=new Dictionary<string,object>{{"Class",Classes[i]},{"Assembly",assembly[0]},
                    {"RuntimeVersion","v4.0.30319"},{"CodeBase",new Uri(dll).AbsoluteUri}};
                rows[clsid+@"\InprocServer32\"+assembly[1]]=new Dictionary<string,object>(managed);
                managed[""]="mscoree.dll"; managed["ThreadingModel"]="Both";
                rows[clsid+@"\InprocServer32"]=managed;
            }
            rows[Addin]=new Dictionary<string,object>{{"FriendlyName","내엑셀 NxHost"},{"Description","내엑셀 Enhanced DLL host"},
                {"LoadBehavior",3},{"CommandLineSafe",1}};
            if(AppsEntryFor(view)) {
                string installer=Path.Combine(packageRoot,"..","Setup-"+packageHash+".exe");
                installer=Path.GetFullPath(installer);
                rows[AppsKey]=new Dictionary<string,object>{{"DisplayName",Int32.Parse(version.Split('_')[1].Substring(1))>=98 ? "내엑셀" : "내엑셀 "+version},{"DisplayVersion",version},
                    {"Publisher","내엑셀"},{"InstallLocation",runtime},{"NoModify",1},{"NoRepair",1},
                    {"UninstallString","\""+installer+"\" --uninstall"+(Int32.Parse(version.Split('_')[1].Substring(1))>=97 ? "" : " --accept-current-user-changes")}};
                if(Int32.Parse(version.Split('_')[1].Substring(1))>=97)
                    rows[AppsKey]["QuietUninstallString"]="\""+installer+"\" --uninstall --accept-current-user-changes";
            }
            return rows;
        }
        void AssertAbsent()
        {
            if(Directory.Exists(runtime)) throw new IOException("기존 DLL 설치 경로가 있습니다. 해당 버전 제거 후 다시 설치하세요.");
            string start=Path.GetDirectoryName(xlstart);
            if(Directory.Exists(start) && Directory.GetFiles(start,"내엑셀*.xlam").Length!=0)
                throw new IOException("XLSTART에 기존 내엑셀이 있습니다. 해당 버전 제거 후 다시 설치하세요.");
            string scriptState=Path.Combine(Path.GetDirectoryName(receipt),"install-state.json");
            if(File.Exists(scriptState) && File.ReadAllText(scriptState).Contains("ACTIVE"))
                throw new IOException("스크립트 설치 기록이 있습니다. 기존 설치 절차로 제거하세요.");
            foreach(var view in Views)
                using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                    foreach(var root in Roots(view)) using(var existing=key.OpenSubKey(Key(root,view)))
                        if(existing!=null) throw new IOException("기존 내엑셀 등록이 있습니다. 설치 상태를 확인한 뒤 해당 설치를 제거하세요.\r\n"+view+" · HKCU\\"+root);
        }
        bool MatchesOwnedValue(string path,string name,object expected,object actual)
        {
            // Accept only this release's historical label; identity, location and command remain strict.
            if(path==AppsKey && name=="DisplayName" && Object.Equals(expected,"내엑셀") &&
                Object.Equals(actual,"내엑셀 "+version)) return true;
            return Object.Equals(expected,actual);
        }
        void ValidateTree(RegistryKey key, string path, Dictionary<string,Dictionary<string,object>> rows, bool partial)
        {
            if(!rows.ContainsKey(path)) throw new IOException("설치 후 등록 항목이 변경되었습니다.");
            var expected=rows[path];
            if(!partial && key.ValueCount!=expected.Count) throw new IOException("설치 등록 값 개수가 변경되었습니다.");
            foreach(string name in key.GetValueNames()) {
                object value=key.GetValue(name,null,RegistryValueOptions.DoNotExpandEnvironmentNames);
                if((repairValidation || partial) && path==Addin && name=="LoadBehavior" && value is int &&
                   ((int)value==0 || (int)value==2) && key.GetValueKind(name)==RegistryValueKind.DWord) continue;
                if(!expected.ContainsKey(name) || !MatchesOwnedValue(path,name,expected[name],value) ||
                    key.GetValueKind(name)!=(value is int ? RegistryValueKind.DWord : RegistryValueKind.String))
                    throw new IOException("설치 등록 값이 변경되었습니다. 자동 제거를 중단합니다.\r\nHKCU\\"+path+"\r\n값 이름: "+(name.Length==0 ? "(기본값)" : name)+" · 형식: "+key.GetValueKind(name)+
                        "\r\n레지스트리 구분: "+key.View+
                        "\r\n기대값: "+(expected.ContainsKey(name) ? Convert.ToString(expected[name]) : "(없음)")+
                        "\r\n실제값: "+Convert.ToString(value));
            }
            foreach(string child in key.GetSubKeyNames()) using(var nested=key.OpenSubKey(child))
                ValidateTree(nested,path+"\\"+child,rows,partial);
        }
        void ValidateOwned(bool partial)
        {
            foreach(var file in Destinations()) {
                Package.RejectReparseAncestors(Path.GetDirectoryName(file.Key));
                FileAttributes attributes;
                try { attributes=File.GetAttributes(file.Key); }
                catch(FileNotFoundException) { if(partial) continue; throw new IOException("설치 파일이 없습니다."); }
                catch(DirectoryNotFoundException) { if(partial) continue; throw new IOException("설치 파일이 없습니다."); }
                if((attributes&FileAttributes.ReparsePoint)!=0 ||
                    new FileInfo(file.Key).Length!=file.Value.Length || Package.Hash(File.ReadAllBytes(file.Key))!=Package.Hash(file.Value))
                    throw new IOException("설치 파일이 변경되었습니다. 자동 제거를 중단합니다.");
            }
            foreach(var view in Views) using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)) {
                var rows=RegistryRows(view);
                foreach(var root in Roots(view)) using(var existing=key.OpenSubKey(Key(root,view))) {
                    if(existing==null) { if(partial) continue; throw new IOException("COM 등록이 없습니다."); }
                    ValidateTree(existing,root,rows,partial);
                }
                if(!partial) foreach(var row in rows) using(var existing=key.OpenSubKey(Key(row.Key,view)))
                    if(existing==null) throw new IOException("COM 하위 등록이 없습니다.");
            }
        }
        void SaveRegistryBackup()
        {
            string directory=Path.GetDirectoryName(receipt);
            Package.RejectReparseAncestors(directory);
            Directory.CreateDirectory(directory);
            var document=new XElement("registry-backup",new XAttribute("version",version),
                new XAttribute("packageSha256",packageHash));
            foreach(var view in Views) using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                foreach(var row in RegistryRows(view)) using(var existing=key.OpenSubKey(Key(row.Key,view))) {
                    if(existing==null) continue;
                    var item=new XElement("key",new XAttribute("view",view),new XAttribute("path",row.Key));
                    foreach(string name in existing.GetValueNames()) item.Add(new XElement("value",
                        new XAttribute("name",name),new XAttribute("kind",existing.GetValueKind(name)),
                        Convert.ToString(existing.GetValue(name,null,RegistryValueOptions.DoNotExpandEnvironmentNames),
                            System.Globalization.CultureInfo.InvariantCulture)));
                    document.Add(item);
                }
            string path=Path.Combine(directory,"native-registry."+DateTime.UtcNow.ToString("yyyyMMddHHmmss")+"."+
                Guid.NewGuid().ToString("N")+".backup.xml");
            document.Save(path);
        }
        bool Present(string path)
        {
            try { File.GetAttributes(path); return true; }
            catch(FileNotFoundException) { return false; }
            catch(DirectoryNotFoundException) { return false; }
        }
        void AssertRemoved()
        {
            foreach(var file in Destinations()) if(Present(file.Key))
                throw new IOException("설치 파일 일부가 남아 있습니다. 제거 완료로 기록하지 않습니다.");
            foreach(var view in Views) using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                foreach(var root in Roots(view)) using(var existing=key.OpenSubKey(Key(root,view)))
                    if(existing!=null) throw new IOException("COM 등록 일부가 남아 있습니다. 제거 완료로 기록하지 않습니다.");
        }
        void Cleanup()
        {
            // Only paths from our fixed layout, never paths accepted from the receipt.
            ValidateOwned(true);
            SaveRegistryBackup();
            SaveReceipt("REMOVING");
            foreach(var view in Views) using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                foreach(var root in Roots(view)) key.DeleteSubKeyTree(Key(root,view),false);
            Hit("cleanup");
            foreach(var file in Destinations()) if(Present(file.Key)) File.Delete(file.Key);
            foreach(var path in new[]{Path.Combine(runtime,"x86"),Path.Combine(runtime,"x64"),runtime})
                if(Directory.Exists(path) && !Directory.EnumerateFileSystemEntries(path).Any()) Directory.Delete(path);
            AssertRemoved();
            SaveReceipt("RESTORED");
        }
        void Hit(string point) { if(fault!=null) fault(point); }
        public void Install()
        {
            using(var mutex=new Mutex(false,@"Local\LHExcel.Setup.Install")) {
                bool owned=false;
                try {
                    try { owned=mutex.WaitOne(0); } catch(AbandonedMutexException) { owned=true; }
                    if(!owned) throw new IOException("다른 내엑셀 설치가 진행 중입니다.");
                    RequireClosed(); Check();
                    if(File.Exists(receipt) && ReadReceipt()!="RESTORED")
                        throw new IOException("기존 설치 기록이 있습니다. 설치 상태 확인 또는 제거를 실행하세요.");
                    AssertAbsent();
                    // Preserve the old receipt before replacing it, even after a failed earlier attempt.
                    if(File.Exists(receipt)) File.Copy(receipt,Path.Combine(Path.GetDirectoryName(receipt),
                        "native-install."+DateTime.UtcNow.ToString("yyyyMMddHHmmss")+"."+Guid.NewGuid().ToString("N")+".backup.xml"),false);
                    SaveReceipt("PREPARING");
                    try {
                        foreach(var file in Destinations()) {
                            Directory.CreateDirectory(Path.GetDirectoryName(file.Key));
                            using(var stream=new FileStream(file.Key,FileMode.CreateNew,FileAccess.Write,FileShare.None))
                                stream.Write(file.Value,0,file.Value.Length);
                            Hit("file");
                        }
                        foreach(var view in Views) using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                            foreach(var row in RegistryRows(view)) {
                                using(var child=key.CreateSubKey(Key(row.Key,view)))
                                    foreach(var pair in row.Value) child.SetValue(pair.Key,pair.Value,
                                        pair.Value is int ? RegistryValueKind.DWord : RegistryValueKind.String);
                                Hit("registry");
                            }
                        ValidateOwned(false); SaveReceipt("ACTIVE");
                        log("설치 완료. 다음 Excel 시작부터 내엑셀을 불러옵니다.");
                    } catch(Exception primary) {
                        try { Cleanup(); } catch(Exception recovery) {
                            throw new IOException("설치 오류: "+primary.Message+" / 복구 필요: "+recovery.Message,primary);
                        }
                        throw new IOException("설치를 중단하고 이번 설치 상태를 복원했습니다: "+primary.Message,primary);
                    }
                } finally { if(owned) mutex.ReleaseMutex(); }
            }
        }
        public void Uninstall()
        {
            using(var mutex=new Mutex(false,@"Local\LHExcel.Setup.Install")) {
                bool owned=false;
                try {
                    try { owned=mutex.WaitOne(0); } catch(AbandonedMutexException) { owned=true; }
                    if(!owned) throw new IOException("다른 내엑셀 설치가 진행 중입니다.");
                    RequireClosed();
                    Package.RejectReparseAncestors(Path.GetDirectoryName(receipt));
                    if(!File.Exists(receipt)) throw new IOException("이 EXE의 설치 기록이 없습니다.");
                    string state=ReadReceipt();
                    if(state=="RESTORED") {
                        // A repeated removal must never touch shared COM registration:
                        // a later version may now own it.
                        foreach(string path in Destinations().Keys)
                            if(File.Exists(path)) throw new IOException("제거 완료 기록과 설치 파일이 일치하지 않습니다. 설치 상태를 확인하세요.");
                        foreach(var view in Views) if(AppsEntryFor(view)) using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)) {
                            using(var app=root.OpenSubKey(Key(AppsKey,view))) if(app!=null) {
                                var expected=RegistryRows(view)[AppsKey];
                                if(app.SubKeyCount!=0 || app.GetValueNames().Any(n=>!expected.ContainsKey(n)))
                                    throw new IOException("프로그램 목록의 소유권을 확인할 수 없습니다.");
                                foreach(var value in expected)
                                        if(!MatchesOwnedValue(AppsKey,value.Key,value.Value,app.GetValue(value.Key)))
                                        throw new IOException("프로그램 목록이 변경되어 보존했습니다.");
                            }
                            root.DeleteSubKeyTree(Key(AppsKey,view),false);
                        }
                        log("이미 제거된 버전입니다. 해당 버전의 프로그램 목록 정리를 확인했습니다.");
                        return;
                    }
                    if(state!="ACTIVE" && state!="PREPARING" && state!="REMOVING") throw new IOException("제거할 활성 설치가 없습니다.");
                    // Missing owned artifacts may be quarantine/partial-cleanup evidence. Every
                    // artifact that is still present must still match before any mutation.
                    ValidateOwned(true); Cleanup();
                    log("제거 및 등록 정리가 완료됐습니다. 설치 패키지와 기록은 보존했습니다.");
                } finally { if(owned) mutex.ReleaseMutex(); }
            }
        }
    }
}
