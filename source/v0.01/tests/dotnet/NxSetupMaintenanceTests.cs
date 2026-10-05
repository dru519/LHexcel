using System;
using System.IO;
using System.Collections.Generic;
using Microsoft.Win32;
using LH.NxSetup;

internal static class NxSetupMaintenanceTests
{
    static int count;
    static void Require(bool value,string message) { if(!value) throw new Exception(message); count++; }
    static void Reject(Action action) { try { action(); } catch(IOException) { count++; return; } throw new Exception("Expected rejection"); }
    static void Write(string path,byte[] value) { Directory.CreateDirectory(Path.GetDirectoryName(path)); File.WriteAllBytes(path,value); }
    static int Main(string[] args)
    {
        string root=Path.Combine(Path.GetTempPath(),"nx-maintenance-"+Guid.NewGuid().ToString("N"));
        string prefix=@"Software\LHExcelSetupTests\"+Guid.NewGuid().ToString("N")+@"\";
        try {
            byte[] zip=File.ReadAllBytes(args[0]); string hash=Package.Hash(zip), nextHash=new string('b',64);
            var files=Package.Read(zip,args.Length>1 ? args[1] : "v0.01_r92");
            foreach(string fault in new[]{"upgrade-removed","file","registry",null}) {
                string home=Path.Combine(root,(count+1).ToString()), reg=prefix+(count+1)+@"\";
                string cache=Path.Combine(home,"cache","v0.01_r92",hash), current=Path.Combine(home,"current");
                foreach(var f in files) { Write(Path.Combine(cache,f.Key),f.Value); Write(Path.Combine(current,f.Key),f.Value); }
                string settings=Path.Combine(home,"Settings","user.cfg"); Write(settings,new byte[]{1,2,3});
                var old=new NativeInstaller("v0.01_r92",hash,cache,files,Console.WriteLine,home,reg,null);
                old.Install();
                // A historic ACTIVE receipt with one leftover DLL but no startup/COM
                // entry must not block an update or be deleted as if it were active.
                string stale=Path.Combine(home,"LHexcel","NxHost","state","r71","native-install.xml");
                Write(stale,System.Text.Encoding.UTF8.GetBytes("<installation version=\"v0.01_r71\" state=\"ACTIVE\"/>"));
                string leftover=Path.Combine(home,"LHexcel","NxHost","r71","x64","NxHost64.dll");
                Write(leftover,new byte[]{7,1});
                string xlam=old.StartupFile;
                bool injected=false;
                var next=new NativeInstaller("v0.01_r93",nextHash,current,files,Console.WriteLine,home,reg,
                    delegate(string point) { if(!injected && point==fault) { injected=true; throw new IOException("injected "+point); } });
                Require(next.MaintenanceStatus().Contains("업데이트 가능"),"old version discovery");
                if(fault!=null) {
                    Reject(()=>next.Update());
                    Require(old.Check().Contains("정상"),"old restored after "+fault);
                    Require(File.Exists(xlam) && !File.Exists(next.StartupFile),"one old startup after rollback");
                }
                next.Update();
                Require(next.Check().Contains("정상"),"new active");
                Require(File.Exists(stale) && File.ReadAllBytes(leftover)[0]==7,"inactive remnants preserved");
                Require(!File.Exists(xlam) && File.Exists(next.StartupFile),"one new startup");
                Require(File.ReadAllBytes(settings).Length==3,"settings preserved");
                File.Delete(next.StartupFile);
                string addin=reg+@"Registry32\Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect";
                using(var key=Registry.CurrentUser.OpenSubKey(addin,true)) key.SetValue("LoadBehavior",0,RegistryValueKind.DWord);
                next.Repair();
                Require(next.Check().Contains("정상"),"missing file and disabled registration repaired");
                File.AppendAllText(next.StartupFile,"changed");
                Reject(()=>next.Repair());
                Require(File.ReadAllText(next.StartupFile).EndsWith("changed"),"changed file preserved");
                File.WriteAllBytes(next.StartupFile,files["payload/Product.xlam"]);
                using(var key=Registry.CurrentUser.OpenSubKey(addin,true)) key.SetValue("FriendlyName","foreign");
                Reject(()=>next.Repair());
                using(var key=Registry.CurrentUser.OpenSubKey(addin,true)) key.SetValue("FriendlyName","내엑셀 NxHost");
                File.Delete(next.StartupFile);
                var broken=new NativeInstaller("v0.01_r93",nextHash,current,files,Console.WriteLine,home,reg,
                    delegate(string point) { if(point=="repair-file") throw new IOException("repair fault"); });
                Reject(()=>broken.Repair());
                Require(!File.Exists(next.StartupFile),"repair rollback restores missing-file state");
                next.Update(); Require(next.Check().Contains("정상"),"same-version update repairs");
                next.Uninstall();
            }
            Console.WriteLine("PASS|Maintenance|"+count+"|ISOLATED_WINDOWS_FILES_REGISTRY|"+root);
            return 0;
        } catch(Exception e) { Console.Error.WriteLine(e); return 1; }
        finally {
            foreach(var view in new[]{RegistryView.Registry32,RegistryView.Registry64})
                using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                    key.DeleteSubKeyTree(prefix.TrimEnd('\\'),false);
        }
    }
}
