using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace LH.NxSetup.Reset {
    internal sealed class ResetPlan {
        internal readonly List<string> Files=new List<string>();
        internal readonly List<string> Directories=new List<string>();
        internal readonly List<KeyValuePair<RegistryView,string>> Keys=new List<KeyValuePair<RegistryView,string>>();
        public override string ToString() {
            return "등록 항목 "+Keys.Count+"개 / 파일 "+Files.Count+"개 / 폴더 "+Directories.Count+"개\r\n"+
                String.Join("\r\n",Keys.Select(k=>k.Key+" · HKCU\\"+k.Value).Concat(Files).Concat(Directories));
        }
    }
    internal sealed class ResetEngine {
        readonly string local,roaming,prefix;
        readonly bool isolated;
        internal static readonly string[] Ids={
            "{4A4B4F02-76A6-4F22-BD4B-85E0308EC91D}","{9A90294C-81C8-4AEE-A2D8-14B078A61CD2}",
            "{A13D5ED6-B0F2-43C0-9499-E53E6CE38DC0}","{1434C649-18AA-4435-B0D2-2BD81B548A01}",
            "{C812F023-E912-49AA-98B7-75DC86501902}","{C812F023-E912-49AA-98B7-75DC86501904}",
            "{C812F023-E912-49AA-98B7-75DC86501906}"};
        internal static readonly string[] Progs={"LH.NxHost.Connect","LH.NxHost.BridgeService","LH.NxHost.NavigatorPane",
            "LH.NxHost.HwpxExportService","LH.NxHost.PicturePreviewService","LH.NxHost.WorkbookCompareService","LH.NxHost.WorkbookCompareResultsService"};
        internal static readonly RegistryView[] Views=Environment.Is64BitOperatingSystem ?
            new[]{RegistryView.Registry32,RegistryView.Registry64}:new[]{RegistryView.Registry32};
        internal ResetEngine():this(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),""){}
        internal ResetEngine(string local,string roaming,string prefix) {
            if(prefix!=""&&!prefix.StartsWith(@"Software\LHExcelResetTests\"))throw new ArgumentException("Test scope");
            this.local=Path.GetFullPath(local);this.roaming=Path.GetFullPath(roaming);this.prefix=prefix;isolated=prefix!="";
        }
        string Key(RegistryView view,string path){return prefix+(isolated?view+@"\":"")+path;}
        static void NoLinks(string path) {
            for(string p=Path.GetFullPath(path);!String.IsNullOrEmpty(p);p=Path.GetDirectoryName(p))
                if((File.Exists(p)||Directory.Exists(p))&&(File.GetAttributes(p)&FileAttributes.ReparsePoint)!=0)
                    throw new IOException("연결된 경로는 삭제하지 않습니다: "+p);
        }
        static void Within(string path,string root) {
            string full=Path.GetFullPath(path),parent=Path.GetFullPath(root).TrimEnd('\\')+@"\";
            if(!full.StartsWith(parent,StringComparison.OrdinalIgnoreCase))throw new IOException("삭제 범위 밖 경로: "+path);
            NoLinks(full);
        }
        static bool VersionFolder(string name){return Regex.IsMatch(name,@"\Ar[0-9]+\z");}
        bool RuntimePath(string path) {
            if(String.IsNullOrEmpty(path))return false;
            string parent=Path.Combine(local,"LHexcel","NxHost")+@"\";
            string full=Path.GetFullPath(path);
            if(!full.StartsWith(parent,StringComparison.OrdinalIgnoreCase))return false;
            return Regex.IsMatch(full.Substring(parent.Length),@"\Ar[0-9]+\\x(?:86|64)\\NxHost(?:32|64)\.dll\z",RegexOptions.IgnoreCase);
        }
        void ValidateCom(RegistryKey key) {
            foreach(string name in key.GetValueNames())if(name.Equals("CodeBase",StringComparison.OrdinalIgnoreCase)){
                Uri uri;string value=key.GetValue(name) as string;
                if(!Uri.TryCreate(value,UriKind.Absolute,out uri)||!uri.IsFile||!RuntimePath(uri.LocalPath))
                    throw new IOException("다른 경로의 COM 등록은 보존합니다: "+value);
            }
            foreach(string child in key.GetSubKeyNames())using(var sub=key.OpenSubKey(child))ValidateCom(sub);
        }
        static void AddTree(ResetPlan plan,string path,string boundary) {
            Within(path,boundary);
            foreach(string child in Directory.GetFileSystemEntries(path)){
                NoLinks(child);
                if(Directory.Exists(child))AddTree(plan,child,boundary);
                else {
                    string ext=Path.GetExtension(child).ToLowerInvariant();
                    if(new[]{".xlsx",".xlsm",".xls",".xltx",".xltm",".hwp",".hwpx"}.Contains(ext))
                        throw new IOException("사용자 문서 가능성이 있어 삭제를 중단합니다: "+child);
                    plan.Files.Add(child);
                }
            }
            plan.Directories.Add(path);
        }
        internal ResetPlan Inspect() {
            var plan=new ResetPlan();
            string product=Path.Combine(local,"LHexcel"),host=Path.Combine(product,"NxHost"),setup=Path.Combine(product,"Setup");
            NoLinks(host);NoLinks(setup);
            if(Directory.Exists(host))foreach(string folder in Directory.GetDirectories(host))
                if(VersionFolder(Path.GetFileName(folder))||new[]{"state","history-backup"}.Contains(Path.GetFileName(folder)))
                    AddTree(plan,folder,host);
            if(Directory.Exists(setup)) {
                foreach(string folder in Directory.GetDirectories(setup))
                    if(Regex.IsMatch(Path.GetFileName(folder),@"\Av0\.01_r[0-9]+\z"))AddTree(plan,folder,setup);
                string last=Path.Combine(setup,"last-error.txt");
                if(File.Exists(last)){Within(last,setup);plan.Files.Add(last);}
            }
            string start=Path.Combine(roaming,"Microsoft","Excel","XLSTART");NoLinks(start);
            if(Directory.Exists(start))foreach(string file in Directory.GetFiles(start))
                if(Regex.IsMatch(Path.GetFileName(file),@"\A내엑셀 v0\.01_r[0-9]+\.xlam\z",RegexOptions.IgnoreCase))
                    {Within(file,start);plan.Files.Add(file);}
            foreach(var view in Views)using(var registry=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)){
                for(int i=0;i<Ids.Length;i++){
                    string cls=Key(view,@"Software\Classes\CLSID\"+Ids[i]);
                    using(var key=registry.OpenSubKey(cls)){
                        if(key!=null){ValidateCom(key);plan.Keys.Add(new KeyValuePair<RegistryView,string>(view,cls));}
                    }
                    string prog=Key(view,@"Software\Classes\"+Progs[i]);
                    using(var key=registry.OpenSubKey(prog))if(key!=null){
                        using(var link=key.OpenSubKey("CLSID"))
                            if(link!=null&&!String.Equals(link.GetValue("") as string,Ids[i],StringComparison.OrdinalIgnoreCase))
                                throw new IOException("다른 CLSID를 가리키는 등록은 보존합니다: "+prog);
                        plan.Keys.Add(new KeyValuePair<RegistryView,string>(view,prog));
                    }
                }
                string addin=Key(view,@"Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect");
                using(var key=registry.OpenSubKey(addin))if(key!=null)plan.Keys.Add(new KeyValuePair<RegistryView,string>(view,addin));
                string arp=Key(view,@"Software\Microsoft\Windows\CurrentVersion\Uninstall");
                using(var apps=registry.OpenSubKey(arp))if(apps!=null)foreach(string name in apps.GetSubKeyNames()){
                    if(!Regex.IsMatch(name,@"\ALHExcel(?:-v0\.01_r[0-9]+)?\z",RegexOptions.IgnoreCase))continue;
                    using(var app=apps.OpenSubKey(name)){
                        string display=app.GetValue("DisplayName") as string;
                        if(display==null||!Regex.IsMatch(display,@"\A내엑셀(?: v0\.01_r[0-9]+)?\z"))throw new IOException("프로그램 항목 소유권 불명: "+name);
                        string location=app.GetValue("InstallLocation") as string;
                        if(location==null)throw new IOException("설치 경로 없는 항목: "+name);
                        Within(location,host);
                        if(!VersionFolder(Path.GetFileName(location.TrimEnd('\\'))))throw new IOException("설치 경로 확인 필요: "+location);
                        plan.Keys.Add(new KeyValuePair<RegistryView,string>(view,arp+@"\"+name));
                    }
                }
            }
            return plan;
        }
        internal void Execute(Action<string> log) {
            if(!isolated&&!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("CODEX_WINDOWS_SANDBOX_PACKAGE_FAMILY")))
                throw new IOException("Codex 도구 환경에서는 실제 사용자 등록 삭제를 실행하지 않습니다. 파일 탐색기에서 직접 실행하세요.");
            using(var mutex=new Mutex(false,@"Local\LHExcel.Setup.Install")){
                bool owned=false;try{
                    try{owned=mutex.WaitOne(0);}catch(AbandonedMutexException){owned=true;}
                    if(!owned)throw new IOException("다른 설치 작업이 진행 중입니다.");
                    if(!isolated&&Process.GetProcessesByName("EXCEL").Length!=0)throw new IOException("문서를 저장하고 Excel을 모두 닫아 주세요.");
                    var plan=Inspect();
                    string self=System.Reflection.Assembly.GetExecutingAssembly().Location;
                    if(plan.Files.Any(f=>String.Equals(f,self,StringComparison.OrdinalIgnoreCase)))throw new IOException("삭제 도구를 설치 폴더 밖으로 옮겨 주세요.");
                    // Preflight all files before the first destructive operation.
                    foreach(string file in plan.Files)using(var stream=new FileStream(file,FileMode.Open,FileAccess.ReadWrite,FileShare.None)){}
                    foreach(var entry in plan.Keys)using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,entry.Key)){
                        key.DeleteSubKeyTree(entry.Value,false);log("등록 삭제: "+entry.Key+" · "+entry.Value);
                    }
                    foreach(string file in plan.Files){Within(file,file.StartsWith(local+@"\",StringComparison.OrdinalIgnoreCase)?local:roaming);File.Delete(file);log("파일 삭제: "+file);}
                    foreach(string folder in plan.Directories){NoLinks(folder);Directory.Delete(folder,false);log("폴더 삭제: "+folder);}
                    var remaining=Inspect();
                    if(remaining.Keys.Count+remaining.Files.Count+remaining.Directories.Count!=0)throw new IOException("잔여 항목이 있습니다.\r\n"+remaining);
                    log("PASS: 내엑셀 EXE 설치 항목 0개. 백업하지 않았습니다. 문서·양식·사용자 설정은 보존했습니다.");
                }finally{if(owned)mutex.ReleaseMutex();}
            }
        }
    }
    internal static class ResetWindow {
        internal static void Run(){
            Application.EnableVisualStyles();
            var engine=new ResetEngine();
            var form=new Form{Text="내엑셀 설치 초기화 — 백업 없이 완전 삭제",Width=820,Height=600,StartPosition=FormStartPosition.CenterScreen};
            var text=new TextBox{Dock=DockStyle.Fill,Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Both,WordWrap=false};
            var run=new Button{Dock=DockStyle.Bottom,Height=45,Text="내엑셀 설치 내역 완전 삭제"};
            var home=new Button{Text="초기화면으로 돌아가기",AutoSize=true};
            var finish=new Button{Text="마침",AutoSize=true};
            var completedActions=new FlowLayoutPanel{Dock=DockStyle.Bottom,Height=52,Padding=new Padding(8),FlowDirection=FlowDirection.RightToLeft,Visible=false};
            completedActions.Controls.Add(home);completedActions.Controls.Add(finish);
            finish.Click+=(s,e)=>form.Close();
            home.Click+=(s,e)=>{
                try{
                    // The reset runs from a temporary copy that survives installation removal.
                    using(var process=Process.Start(new ProcessStartInfo(System.Reflection.Assembly.GetExecutingAssembly().Location){UseShellExecute=true})){
                        if(process==null)throw new IOException("설치 초기화면을 열지 못했습니다.");
                    }
                    form.Close();
                }catch(Exception ex){MessageBox.Show(form,ex.Message,"초기화면 열기 오류",MessageBoxButtons.OK,MessageBoxIcon.Error);}
            };
            form.Controls.Add(text);form.Controls.Add(run);form.Controls.Add(completedActions);bool busy=false,finished=false;
            form.FormClosing+=(s,e)=>{if(busy)e.Cancel=true;};
            form.Shown+=(s,e)=>{try{text.Text="Excel을 닫아 주세요. 문서·양식·사용자 설정은 보존합니다.\r\n백업 없이 삭제하며 복원할 수 없습니다.\r\n\r\n"+engine.Inspect();}catch(Exception ex){text.Text=ex.Message;run.Enabled=false;}};
            run.Click+=async(s,e)=>{
                if(finished){form.Close();return;}
                if(MessageBox.Show("내엑셀의 EXE 설치 파일·등록·설치 이력을 백업 없이 삭제합니다.\r\n계속하시겠습니까?","완전 삭제 확인",MessageBoxButtons.YesNo,MessageBoxIcon.Warning)!=DialogResult.Yes)return;
                busy=true;run.Enabled=false;var log=new StringBuilder();
                try{await Task.Run(()=>engine.Execute(line=>log.AppendLine(line)));text.Text=log.ToString();finished=true;}
                catch(Exception ex){text.Text="FAIL: "+ex.Message+"\r\n이미 처리된 항목:\r\n"+log;}
                finally{
                    busy=false;
                    if(finished){run.Visible=false;completedActions.Visible=true;form.AcceptButton=finish;form.CancelButton=finish;}
                    string dir=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"LHexcel","Diagnostics");
                    Directory.CreateDirectory(dir);
                    string path=Path.Combine(dir,"reset-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N")+".txt");
                    File.WriteAllText(path,text.Text,new UTF8Encoding(false));text.AppendText("\r\n결과 파일: "+path);
                }
            };
            Application.Run(form);
        }
    }
}
