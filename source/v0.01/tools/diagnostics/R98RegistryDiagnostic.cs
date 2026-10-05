using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;

namespace LH.NxSetup {
    // Read-only entry point: never calls Stage/Install/Update/Repair/Uninstall.
    internal static class R98RegistryDiagnostic {
        static readonly BindingFlags PrivateInstance=BindingFlags.Instance|BindingFlags.NonPublic;
        [STAThread] static void Main() {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            var window=new Form {Text="내엑셀 r98 등록 정보 진단 — 읽기 전용",Width=820,Height=580,StartPosition=FormStartPosition.CenterScreen};
            var output=new TextBox {Dock=DockStyle.Fill,Multiline=true,ReadOnly=true,ScrollBars=ScrollBars.Both,WordWrap=false};
            var button=new Button {Dock=DockStyle.Bottom,Height=44,Text="등록 정보 검사 (설치·제거하지 않음)"};
            window.Controls.Add(output);window.Controls.Add(button);
            bool busy=false;
            window.FormClosing+=(s,e)=>{if(busy)e.Cancel=true;};
            button.Click+=async(s,e)=>{
                busy=true;button.Enabled=false;output.Text="등록값과 파일 해시를 읽고 있습니다.";
                try { output.Text=await Task.Run(()=>Inspect()); }
                catch(Exception error){output.Text=error.ToString();}
                finally{busy=false;button.Enabled=true;}
            };
            window.Shown+=(s,e)=>button.PerformClick();
            Application.Run(window);
        }
        static byte[] Resource(string name) {
            using(var source=Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            using(var buffer=new MemoryStream()){if(source==null)throw new InvalidDataException(name);source.CopyTo(buffer);return buffer.ToArray();}
        }
        internal static string Inspect() {
            byte[] zip=Resource("NxSetup.Package.zip");
            string version="v0.01_r98",hash=Package.Hash(zip);
            var files=Package.Read(zip,version);
            var messages=new StringBuilder();
            var installer=new NativeInstaller(version,hash,Package.CachedPath(version,hash),files,s=>messages.AppendLine(s));
            var report=new XElement("registry-diagnostic",
                new XAttribute("time",DateTimeOffset.Now.ToString("o")),new XAttribute("readOnly",true),
                new XAttribute("executable",Assembly.GetExecutingAssembly().Location),
                new XAttribute("packageSha256",hash),new XAttribute("process64",Environment.Is64BitProcess),
                new XAttribute("localAppData",Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)));
            var method=typeof(NativeInstaller).GetMethod("RegistryRows",PrivateInstance);
            int mismatches=0;
            foreach(var view in Environment.Is64BitOperatingSystem ? new[]{RegistryView.Registry32,RegistryView.Registry64}:new[]{RegistryView.Registry32}) {
                var rows=(Dictionary<string,Dictionary<string,object>>)method.Invoke(installer,new object[]{view});
                using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                foreach(var row in rows) {
                    if(!row.Value.ContainsKey("CodeBase"))continue;
                    using(var key=root.OpenSubKey(row.Key,false)) {
                        string expected=(string)row.Value["CodeBase"];
                        object actual=key==null ? null:key.GetValue("CodeBase",null,RegistryValueOptions.DoNotExpandEnvironmentNames);
                        string kind=actual==null ? "Missing":key.GetValueKind("CodeBase").ToString();
                        bool same=Object.Equals(expected,actual)&&kind=="String";
                        if(!same)mismatches++;
                        report.Add(new XElement("value",new XAttribute("view",view),new XAttribute("key",row.Key),
                            new XAttribute("name","CodeBase"),new XAttribute("kind",kind),new XAttribute("equal",same),
                            new XElement("expected",expected),new XElement("actual",actual??"(missing)")));
                    }
                }
            }
            string status;
            try {
                typeof(NativeInstaller).GetMethod("ReadReceipt",PrivateInstance).Invoke(installer,new object[0]);
                typeof(NativeInstaller).GetMethod("ValidateOwned",PrivateInstance).Invoke(installer,new object[]{true});
                status="PASS";
                messages.AppendLine("제거 전 소유권 검사: 정상 (실제 제거는 실행하지 않음)");
            } catch(Exception error) {
                status="FAIL";messages.AppendLine((error.InnerException??error).ToString());
            }
            report.Add(new XAttribute("status",status),new XAttribute("codeBaseMismatches",mismatches),new XElement("details",messages.ToString()));
            string directory=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"LHexcel","Diagnostics");
            Directory.CreateDirectory(directory);
            string path=Path.Combine(directory,"r98-registry-"+DateTime.Now.ToString("yyyyMMdd-HHmmss")+"-"+Guid.NewGuid().ToString("N")+".xml");
            using(var stream=new FileStream(path,FileMode.CreateNew,FileAccess.Write))report.Save(stream);
            return "진단 결과: "+status+"\r\nCodeBase 불일치: "+mismatches+"\r\n결과 파일: "+path+"\r\n\r\n"+report.ToString();
        }
    }
}
