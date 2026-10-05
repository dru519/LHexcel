using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Win32;
using LH.NxSetup;

internal static class NxSetupR97Tests
{
    static int count;
    static void Check(bool value,string label) { if(!value) throw new Exception(label); count++; }
    [STAThread]
    static int Main(string[] args)
    {
        string root=Path.Combine(Path.GetTempPath(),"nx97-"+Guid.NewGuid().ToString("N").Substring(0,8));
        string prefix=@"Software\LHExcelSetupTests\"+Guid.NewGuid().ToString("N")+@"\";
        try {
            byte[] zip=File.ReadAllBytes(args[0]); string hash=Package.Hash(zip);
            var files=Package.Read(zip,args[1]);
            string package=Path.Combine(root,"package");
            foreach(var file in files) { string path=Path.Combine(package,file.Key); Directory.CreateDirectory(Path.GetDirectoryName(path)); File.WriteAllBytes(path,file.Value); }
            int prompts=0,launches=0;
            Check(!SetupWindow.OfferExcelLaunch("repair",()=>{prompts++;return true;},()=>launches++),"repair must not launch");
            Check(prompts==0 && launches==0,"repair no prompt");
            Check(!SetupWindow.OfferExcelLaunch("update",()=>{prompts++;return false;},()=>launches++),"decline respected");
            Check(launches==0,"decline no launch");
            Check(SetupWindow.OfferExcelLaunch("update",()=>true,()=>launches++),"accepted launch");
            Check(launches==1,"launch exactly once");
            using(var window=new SetupWindow("v0.01_r97",hash,new Dictionary<string,byte[]>())) {
                var flags=System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic;
                var review=typeof(SetupWindow).GetMethod("ReviewAction",flags);
                var nextControl=(System.Windows.Forms.Button)typeof(SetupWindow).GetField("nextButton",flags).GetValue(window);
                var detail=(System.Windows.Forms.TextBox)typeof(SetupWindow).GetField("output",flags).GetValue(window);
                Check(!detail.Visible,"technical logs initially collapsed");
                Check(nextControl.Text=="다음 >","welcome offers next instead of immediate mutation");
                var utility=(System.Windows.Forms.FlowLayoutPanel)typeof(SetupWindow).GetField("utilities",flags).GetValue(window);
                bool hasReset=false;foreach(System.Windows.Forms.Control control in utility.Controls)
                    if(control.Text=="완전 삭제 (설치 이력 포함)")hasReset=true;
                Check(hasReset,"reset is available outside the blocked maintenance choices");
                var launch=(System.Windows.Forms.Button)typeof(SetupWindow).GetField("launchExcel",flags).GetValue(window);
                Check(launch.Text=="Excel 실행" && launch.Parent==nextControl.Parent,"Excel launch is a separate navigation button");
                Check(nextControl.Parent.Controls.GetChildIndex(nextControl)<nextControl.Parent.Controls.GetChildIndex(launch),"RTL navigation places finish after Excel launch");
                review.Invoke(window,new object[]{"uninstall"});
                Check(nextControl.Text=="제거","uninstall requires explicit review action");
                review.Invoke(window,new object[]{"repair"});
                Check(nextControl.Text=="복구","repair review action");
                review.Invoke(window,new object[]{"check"});
                Check(nextControl.Text=="확인 시작","check review action");
                typeof(SetupWindow).GetField("completed",flags).SetValue(window,true);
                typeof(SetupWindow).GetField("screen",flags).SetValue(window,"done");
                window.Show();System.Windows.Forms.Application.DoEvents();nextControl.PerformClick();
                Check(!window.IsDisposed,"check completion must not exit");
                Check((string)typeof(SetupWindow).GetField("screen",flags).GetValue(window)=="choices","check returns to management");
                typeof(SetupWindow).GetField("pendingAction",flags).SetValue(window,"uninstall");
                typeof(SetupWindow).GetMethod("ShowChoices",flags).Invoke(window,null);
                var chosen=(System.Windows.Forms.RadioButton)typeof(SetupWindow).GetField("checkButton",flags).GetValue(window);
                var installChoice=(System.Windows.Forms.RadioButton)typeof(SetupWindow).GetField("installButton",flags).GetValue(window);
                var repairChoice=(System.Windows.Forms.RadioButton)typeof(SetupWindow).GetField("repairButton",flags).GetValue(window);
                string selectedAction=installChoice.Checked ? "update" : repairChoice.Checked ? "repair" : "check";
                Check((string)typeof(SetupWindow).GetField("pendingAction",flags).GetValue(window)==selectedAction,"return refreshes pending action even when same choice remains checked");
                typeof(SetupWindow).GetMethod("ShowAlreadyRemoved",flags).Invoke(window,null);
                Check(nextControl.Text=="마침","already removed has explicit finish");
                window.Close();
            }
            foreach(float factor in new[]{1.25F,1.5F})using(var window=new SetupWindow(args[1],hash,new Dictionary<string,byte[]>())) {
                var flags=System.Reflection.BindingFlags.Instance|System.Reflection.BindingFlags.NonPublic;
                window.Show();window.Scale(new System.Drawing.SizeF(factor,factor));
                typeof(SetupWindow).GetMethod("ReviewAction",flags).Invoke(window,new object[]{"update"});
                System.Windows.Forms.Application.DoEvents();
                var nextControl=(System.Windows.Forms.Button)typeof(SetupWindow).GetField("nextButton",flags).GetValue(window);
                Check(nextControl.Visible && nextControl.Parent.ClientRectangle.Contains(nextControl.Bounds),"scaled action button remains inside navigation "+factor);
                var summary=(System.Windows.Forms.TextBox)typeof(SetupWindow).GetField("summary",flags).GetValue(window);
                summary.AppendText("\r\n"+new string('x',400));
                Check(summary.Multiline && summary.ScrollBars==System.Windows.Forms.ScrollBars.Vertical && summary.Height>80,"long path/result scrollable at scale "+factor);
                using(var bitmap=new System.Drawing.Bitmap(window.Width,window.Height)) {
                    window.DrawToBitmap(bitmap,new System.Drawing.Rectangle(0,0,window.Width,window.Height));
                    bitmap.Save(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"layout-"+(int)(factor*100)+".png"));
                }
                window.Close();
            }
            using(var empty=new System.Windows.Forms.CheckedListBox())using(var panel=SetupWindow.HistorySelectionTools(empty)) {
                Check(!panel.Controls[0].Enabled && !panel.Controls[1].Enabled,"empty history disables bulk selection");
                Check(panel.Controls[2].Text=="0 / 0개 선택","empty count");
            }
            using(var list=new System.Windows.Forms.CheckedListBox()) {
                list.Items.AddRange(new object[]{"r94","r95","r96"});
                using(var panel=SetupWindow.HistorySelectionTools(list)) {
                    ((System.Windows.Forms.Button)panel.Controls[0]).PerformClick();
                    Check(list.CheckedItems.Count==3,"select all history items");
                    Check(panel.Controls[2].Text=="3 / 3개 선택","selected count accurate");
                    ((System.Windows.Forms.Button)panel.Controls[1]).PerformClick();
                    Check(list.CheckedItems.Count==0,"clear all history items");
                    list.SetItemChecked(1,true);
                    Check(panel.Controls[2].Text=="1 / 3개 선택","individual selection count accurate");
                    for(int i=0;i<200;i++)list.Items.Add("old-"+i);
                    ((System.Windows.Forms.Button)panel.Controls[0]).PerformClick();
                    Check(list.CheckedItems.Count==203,"bulk selection covers scrollable long history");
                    ((System.Windows.Forms.Button)panel.Controls[1]).PerformClick();
                    Check(list.CheckedItems.Count==0,"bulk clear covers scrollable long history");
                }
            }
            var install=new NativeInstaller("v0.01_r97",hash,package,files,Console.WriteLine,root,prefix,null);
            install.Update(); Check(install.InstalledVersion()=="v0.01_r97","installed detected");
            string view=Environment.Is64BitOperatingSystem ? "Registry64" : "Registry32";
            string arp=prefix+view+@"\Software\Microsoft\Windows\CurrentVersion\Uninstall\LHExcel";
            var backup=new Dictionary<string,object>();
            using(var key=Registry.CurrentUser.OpenSubKey(arp)) {
                Check(key!=null,"single product entry");
                Check(((string)key.GetValue("UninstallString")).EndsWith(" --uninstall"),"interactive ARP");
                Check(((string)key.GetValue("QuietUninstallString")).EndsWith(" --uninstall --accept-current-user-changes"),"explicit quiet removal");
                foreach(string name in key.GetValueNames()) backup[name]=key.GetValue(name);
            }
            foreach(string arch in Environment.Is64BitOperatingSystem ? new[]{"Registry32","Registry64"} : new[]{"Registry32"})
                using(var key=Registry.CurrentUser.OpenSubKey(prefix+arch+@"\Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect",true)) key.SetValue("LoadBehavior",2,RegistryValueKind.DWord);
            install.RemoveInstalledVersion(); Check(!File.Exists(install.StartupFile),"disabled add-in startup removed");
            using(var key=Registry.CurrentUser.OpenSubKey(arp)) Check(key==null,"ARP removed");
            // Simulate an old Apps list item after removal, with a newer shared registration.
            string shared=prefix+@"Registry32\Software\Classes\LH.NxHost.Connect";
            using(var key=Registry.CurrentUser.CreateSubKey(shared)) key.SetValue("newer-owner","keep");
            using(var key=Registry.CurrentUser.CreateSubKey(arp)) foreach(var row in backup) key.SetValue(row.Key,row.Value);
            install.Uninstall();
            using(var key=Registry.CurrentUser.OpenSubKey(arp)) Check(key==null,"stale owned ARP removed");
            using(var key=Registry.CurrentUser.OpenSubKey(shared)) Check((string)key.GetValue("newer-owner")=="keep","newer registration preserved");
            install.Uninstall(); count++;
            using(var key=Registry.CurrentUser.CreateSubKey(arp)) { foreach(var row in backup) key.SetValue(row.Key,row.Value); key.SetValue("DisplayVersion","foreign"); }
            bool rejected=false;try{install.Uninstall();}catch(IOException){rejected=true;}
            Check(rejected,"foreign ARP rejected");
            using(var key=Registry.CurrentUser.OpenSubKey(arp)) Check((string)key.GetValue("DisplayVersion")=="foreign","foreign ARP preserved");
            string upgradeRoot=Path.Combine(root,"upgrade"),upgradePrefix=prefix+@"upgrade\";
            string oldCache=Path.Combine(upgradeRoot,"cache","v0.01_r95",hash);
            foreach(var file in files) {string path=Path.Combine(oldCache,file.Key);Directory.CreateDirectory(Path.GetDirectoryName(path));File.WriteAllBytes(path,file.Value);}
            var old=new NativeInstaller("v0.01_r95",hash,oldCache,files,Console.WriteLine,upgradeRoot,upgradePrefix,null);
            var next=new NativeInstaller("v0.01_r97",new string('c',64),package,files,Console.WriteLine,upgradeRoot,upgradePrefix,null);
            old.Install();Check(next.InstalledVersion()=="v0.01_r95","old version discovered");
            next.Update();Check(!File.Exists(old.StartupFile) && File.Exists(next.StartupFile),"upgrade leaves one startup");
            using(var key=Registry.CurrentUser.OpenSubKey(upgradePrefix+view+@"\Software\Microsoft\Windows\CurrentVersion\Uninstall\LHExcel-v0.01_r95")) Check(key==null,"legacy ARP removed on upgrade");
            next.RemoveInstalledVersion();Check(!File.Exists(next.StartupFile),"upgraded startup removed");
            old.Install();next.RemoveInstalledVersion();Check(!File.Exists(old.StartupFile),"new EXE removes installed older version");
            Check(Array.IndexOf(next.RemovedHistory(),"v0.01_r95")>=0,"removed history offered");
            Check(Array.IndexOf(next.RemovedHistory(),"v0.01_r97")<0,"current version history retained");
            bool blocked=false;try{next.ArchiveHistory(new[]{".."},true);}catch(IOException){blocked=true;}
            Check(blocked,"unlisted history rejected");
            old.Install();
            Check(Array.IndexOf(next.RemovedHistory(),"v0.01_r95")<0,"active history excluded");
            next.RemoveInstalledVersion();
            string archived=next.ArchiveHistory(new[]{"v0.01_r95"},true);
            Check(File.Exists(Path.Combine(archived,"r95","native-install.xml")),"receipt recoverable");
            Check(File.Exists(Path.Combine(archived,"v0.01_r95",hash,"payload","Product.xlam")),"cache recoverable");
            Check(!Directory.Exists(Path.Combine(upgradeRoot,"LHexcel","NxHost","state","r95")),"selected record removed from active history");
            Check(File.Exists(Path.Combine(upgradeRoot,"LHexcel","NxHost","state","r97","native-install.xml")),"current record preserved");
            string revisionCache=Path.Combine(upgradeRoot,"cache","v0.01_r97",new string('c',64));
            foreach(var file in files){string path=Path.Combine(revisionCache,file.Key);Directory.CreateDirectory(Path.GetDirectoryName(path));File.WriteAllBytes(path,file.Value);}
            var priorRevision=new NativeInstaller("v0.01_r97",new string('c',64),revisionCache,files,Console.WriteLine,upgradeRoot,upgradePrefix,null);
            priorRevision.Install();
            var refresh=new NativeInstaller("v0.01_r97",new string('d',64),package,files,Console.WriteLine,upgradeRoot,upgradePrefix,null);
            Check(refresh.HasPackageRefresh() && refresh.InstalledVersion()=="v0.01_r97","same release refresh detected");
            refresh.Update();Check(!refresh.HasPackageRefresh() && File.Exists(refresh.StartupFile),"same release refresh applied");
            refresh.RemoveInstalledVersion();Check(!File.Exists(refresh.StartupFile),"refreshed release removable");
            string brandRoot=Path.Combine(root,"brand"),brandPrefix=prefix+"Brand\\";
            var branded=new NativeInstaller("v0.01_r98",hash,package,files,Console.WriteLine,brandRoot,brandPrefix,null);
            branded.Install();
            string brandArp=brandPrefix+view+@"\Software\Microsoft\Windows\CurrentVersion\Uninstall\LHExcel";
            using(var key=Registry.CurrentUser.OpenSubKey(brandArp,true)) {
                Check((string)key.GetValue("DisplayName")=="내엑셀","r98 single product display name");
                Check((string)key.GetValue("DisplayVersion")=="v0.01_r98","version retained separately");
                key.SetValue("DisplayName","내엑셀 v0.01_r98");
            }
            branded.Check();count++;
            branded.Repair();
            using(var key=Registry.CurrentUser.OpenSubKey(brandArp,true)) {
                Check((string)key.GetValue("DisplayName")=="내엑셀","repair normalizes historical label");
                key.SetValue("DisplayName","내엑셀 v0.01_r95");
            }
            rejected=false;try{branded.Uninstall();}catch(IOException){rejected=true;}
            Check(rejected,"foreign release label still rejected");
            using(var key=Registry.CurrentUser.OpenSubKey(brandArp,true))key.SetValue("DisplayName","내엑셀");
            branded.Uninstall();
            string brandCache=Path.Combine(brandRoot,"cache","v0.01_r95",hash);
            foreach(var file in files){string path=Path.Combine(brandCache,file.Key);Directory.CreateDirectory(Path.GetDirectoryName(path));File.WriteAllBytes(path,file.Value);}
            var brandOld=new NativeInstaller("v0.01_r95",hash,brandCache,files,Console.WriteLine,brandRoot,brandPrefix,null);
            brandOld.Install();branded.Update();
            using(var key=Registry.CurrentUser.OpenSubKey(brandPrefix+view+@"\Software\Microsoft\Windows\CurrentVersion\Uninstall\LHExcel-v0.01_r95"))
                Check(key==null,"r95 to r98 replaces old product entry");
            using(var key=Registry.CurrentUser.OpenSubKey(brandArp))
                Check((string)key.GetValue("DisplayName")=="내엑셀","upgraded product has stable name");
            Check(!File.Exists(brandOld.StartupFile) && File.Exists(branded.StartupFile),"no parallel old startup after upgrade");
            branded.Uninstall();
            Console.WriteLine("PASS|R97_SETUP|"+count+"|ISOLATED_WINDOWS_FILES_REGISTRY");return 0;
        } catch(Exception error){Console.Error.WriteLine(error);return 1;}
        finally {foreach(var view in new[]{RegistryView.Registry32,RegistryView.Registry64}) using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view)) key.DeleteSubKeyTree(prefix.TrimEnd('\\'),false);}
    }
}
