using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Microsoft.Win32;

namespace LH.NxSetup
{
    internal sealed partial class NativeInstaller
    {
        string SetupHistoryRoot { get { return test ? Path.Combine(testHome,"cache") :
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"LHexcel","Setup"); } }

        public string[] RemovedHistory()
        {
            if(!Directory.Exists(StateRoot)) return new string[0];
            var result=new List<string>();
            foreach(string directory in Directory.GetDirectories(StateRoot)) {
                string release=Path.GetFileName(directory);
                if(!System.Text.RegularExpressions.Regex.IsMatch(release,@"\Ar[0-9]+\z")) continue;
                string candidate="v0.01_"+release;
                if(candidate==version) continue;
                string path=Path.Combine(directory,"native-install.xml");
                if(!File.Exists(path)) continue;
                var doc=ReadXml(path);
                if((string)doc.Attribute("version")==candidate && (string)doc.Attribute("state")=="RESTORED" && !HasActiveEntryPoint(candidate))
                    result.Add(candidate);
            }
            return result.OrderByDescending(v=>Int32.Parse(v.Split('_')[1].Substring(1))).ToArray();
        }

        static void ValidateHistoryTree(string path)
        {
            Package.RejectReparseAncestors(path);
            if(!Directory.Exists(path)) return;
            foreach(string child in Directory.GetFileSystemEntries(path)) {
                if((File.GetAttributes(child)&FileAttributes.ReparsePoint)!=0) throw new IOException("연결된 폴더가 있어 기록 정리를 중단했습니다: "+child);
                if(Directory.Exists(child)) ValidateHistoryTree(child);
            }
        }

        public string ArchiveHistory(string[] selected, bool includeCache)
        {
            if(selected==null || selected.Length==0) throw new IOException("정리할 이전 버전을 선택하세요.");
            using(var mutex=new Mutex(false,@"Local\LHExcel.Setup.Install")) {
                bool owned=false;
                try {
                    try{owned=mutex.WaitOne(0);}catch(AbandonedMutexException){owned=true;}
                    if(!owned) throw new IOException("다른 내엑셀 설치가 진행 중입니다.");
                    RequireClosed();
                    var allowed=new HashSet<string>(RemovedHistory(),StringComparer.Ordinal);
                    var sources=new List<string>();
                    foreach(string item in selected.Distinct()) {
                        if(!allowed.Contains(item)) throw new IOException("사용 중이거나 제거가 확인되지 않은 기록은 정리할 수 없습니다: "+item);
                        string release=item.Split('_')[1];
                        // Residual runtime or an Apps entry requires removal, not history cleanup.
                        if(Directory.Exists(Path.Combine(Path.GetDirectoryName(runtime),release)))
                            throw new IOException("이전 프로그램 폴더가 남아 있습니다. 먼저 제거 상태를 확인하세요: "+item);
                        foreach(var view in Views) using(var registry=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                            using(var app=registry.OpenSubKey(Key(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\LHExcel-"+item,view)))
                                if(app!=null) throw new IOException("프로그램 목록이 남아 있어 기록을 보존했습니다: "+item);
                        sources.Add(Path.Combine(StateRoot,release));
                        string cache=Path.Combine(SetupHistoryRoot,item);
                        if(includeCache && Directory.Exists(cache)) sources.Add(cache);
                    }
                    foreach(string path in sources) ValidateHistoryTree(path);
                    string archive=Path.Combine(Path.GetDirectoryName(StateRoot),"history-backup",DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N"));
                    Package.RejectReparseAncestors(archive);
                    Directory.CreateDirectory(archive);
                    var moved=new List<KeyValuePair<string,string>>();
                    try {
                        foreach(string path in sources) {
                            string target=Path.Combine(archive,Path.GetFileName(path));
                            Directory.Move(path,target); moved.Add(new KeyValuePair<string,string>(path,target));
                        }
                    } catch {
                        foreach(var pair in moved.AsEnumerable().Reverse()) Directory.Move(pair.Value,pair.Key);
                        throw;
                    }
                    log("선택한 이전 설치 기록을 목록에서 정리했습니다. 복구 보관 위치: "+archive);
                    return archive;
                } finally {if(owned)mutex.ReleaseMutex();}
            }
        }
    }
}
