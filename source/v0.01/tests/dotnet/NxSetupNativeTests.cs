using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;
using LH.NxSetup;

internal static class NxSetupNativeTests
{
    static int count;
    static void Expect(bool condition) { if(!condition) throw new Exception("Assertion failed"); count++; }
    static void Reject(Action action) { try { action(); } catch(IOException) { count++; return; } catch(InvalidDataException) { count++; return; } throw new Exception("Expected rejection"); }
    static void CheckHwpxRegistration(string prefix, bool present) {
        foreach(string view in new[]{"Registry32","Registry64"}) {
            if(view=="Registry64" && !Environment.Is64BitOperatingSystem) continue;
            using(var prog=Registry.CurrentUser.OpenSubKey(prefix+view+@"\Software\Classes\LH.NxHost.HwpxExportService\CLSID"))
                Expect(present ? prog!=null && (string)prog.GetValue("")=="{1434C649-18AA-4435-B0D2-2BD81B548A01}" : prog==null);
            using(var cls=Registry.CurrentUser.OpenSubKey(prefix+view+@"\Software\Classes\CLSID\{1434C649-18AA-4435-B0D2-2BD81B548A01}\InprocServer32"))
                Expect(present ? cls!=null && (string)cls.GetValue("Class")=="LH.NxHost.HwpxExportService" : cls==null);
        }
    }
    static string WritePackage(string parent, Dictionary<string,byte[]> files) {
        string root=Path.Combine(parent,"verified-package");
        foreach(var file in files) {
            string path=Path.Combine(root,file.Key.Replace('/',Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllBytes(path,file.Value);
        }
        Package.VerifyDirectory(root,files);
        return root;
    }
    static int Main(string[] args)
    {
        string baseRoot=Path.Combine(Path.GetTempPath(),"nx-native-tests-"+Guid.NewGuid().ToString("N"));
        string prefix=@"Software\LHExcelSetupTests\"+Guid.NewGuid().ToString("N")+@"\";
        try {
            byte[] zip=File.ReadAllBytes(args[0]); string hash=Package.Hash(zip);
            string version=args.Length>1 ? args[1] : "v0.01_r64";
            var files=Package.Read(zip,version);
            string package=WritePackage(baseRoot,files);
            int caseNo=0;
            foreach(string fail in new[]{"file","registry",null}) {
                string root=Path.Combine(baseRoot,(caseNo++).ToString());
                string reg=prefix+caseNo+@"\";
                string settings=Path.Combine(root,@"LHexcel\Settings\user-fixture.cfg");
                Directory.CreateDirectory(Path.GetDirectoryName(settings));File.WriteAllText(settings,"owned-preferences");
                var install=new NativeInstaller(version,hash,package,files,Console.WriteLine,root,reg,
                    delegate(string point) { if(point==fail) throw new IOException("Injected "+point); });
                if(fail!=null) {
                    Reject(()=>install.Install());
                    CheckHwpxRegistration(reg,false);
                    Expect(!File.Exists(Path.Combine(root,@"Microsoft\Excel\XLSTART\내엑셀 "+version+".xlam")));
                    using(var key=Registry.CurrentUser.OpenSubKey(reg+@"Registry32\Software\Classes\LH.NxHost.Connect")) Expect(key==null);
                } else {
                    install.Install();
                    CheckHwpxRegistration(reg,true);
                    Expect(install.Check().Contains("정상"));
                    Reject(()=>install.Install());
                    string xlam=Path.Combine(root,@"Microsoft\Excel\XLSTART\내엑셀 "+version+".xlam");
                    File.AppendAllText(xlam,"changed");
                    Reject(()=>install.Uninstall());
                    using(var key=Registry.CurrentUser.OpenSubKey(reg+@"Registry32\Software\Classes\LH.NxHost.Connect")) Expect(key!=null);
                    File.WriteAllBytes(xlam,files["payload/Product.xlam"]);
                    install.Uninstall();
                    CheckHwpxRegistration(reg,false);
                    Expect(!File.Exists(xlam));
                    install.Install(); install.Uninstall(); count++;
                    // A repackaged release may replace only a fully removed installation.
                    string receipt=Path.Combine(root,@"LHexcel\NxHost\state",version.Split('_')[1],"native-install.xml");
                    string previous=File.ReadAllText(receipt);
                    int backupCount=Directory.GetFiles(Path.GetDirectoryName(receipt),"native-install.*.backup.xml").Length;
                    string newHash=new string('a',64);
                    var replacement=new NativeInstaller(version,newHash,package,files,Console.WriteLine,root,reg,null);
                    Expect(replacement.Check().Contains("활성 기록은 없습니다"));
                    replacement.Install();
                    string[] backups=Directory.GetFiles(Path.GetDirectoryName(receipt),"native-install.*.backup.xml");
                    Expect(backups.Length==backupCount+1 && Array.Exists(backups,p=>File.ReadAllText(p)==previous));
                    Reject(()=>install.Uninstall());
                    replacement.Uninstall();
                    // A present changed registry value blocks removal before files or keys mutate.
                    replacement.Install();
                    using(var key=Registry.CurrentUser.OpenSubKey(reg+@"Registry32\Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect",true))
                        key.SetValue("FriendlyName","changed");
                    Reject(()=>replacement.Uninstall());
                    Expect(File.Exists(xlam));
                    using(var key=Registry.CurrentUser.OpenSubKey(reg+@"Registry32\Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect",true))
                        key.SetValue("FriendlyName","내엑셀 NxHost");
                    replacement.Uninstall();
                    // Missing cache and installed files are recoverable without restoring DLLs.
                    replacement.Install();
                    Directory.Delete(package,true);
                    string installed32=Path.Combine(root,@"LHexcel\NxHost\"+version.Split('_')[1]+@"\x86\NxHost32.dll");
                    File.Delete(installed32);
                    replacement.Uninstall();
                    Expect(!File.Exists(installed32));
                    Expect(Directory.GetFiles(Path.GetDirectoryName(receipt),"native-registry.*.backup.xml").Length>0);
                    // A partial cleanup remains REMOVING and can be resumed without recreating DLLs.
                    package=WritePackage(baseRoot,files);
                    replacement=new NativeInstaller(version,newHash,package,files,Console.WriteLine,root,reg,null);
                    replacement.Install();
                    var interrupted=new NativeInstaller(version,newHash,package,files,Console.WriteLine,root,reg,
                        delegate(string point) { if(point=="cleanup") throw new IOException("Injected cleanup"); });
                    Reject(()=>interrupted.Uninstall());
                    Expect(File.ReadAllText(receipt).Contains("REMOVING"));
                    replacement.Uninstall();
                    // Restored receipts must not authorize removal of user data.
                    File.WriteAllText(xlam,"user");
                    Reject(()=>install.Install());
                    Expect(File.ReadAllText(xlam)=="user");
                    File.Delete(xlam);
                    string restored=File.ReadAllText(receipt);
                    File.WriteAllText(receipt,restored.Replace("RESTORED","ACTIVE"));
                    Reject(()=>install.Install());
                    File.WriteAllText(receipt,restored.Replace("RESTORED","UNKNOWN"));
                    Reject(()=>replacement.Install());
                    File.WriteAllText(receipt,restored);
                }
                Expect(File.ReadAllText(settings)=="owned-preferences");
            }
            // Preserve an unrelated pre-existing XLAM and refuse before creating COM keys.
            string conflict=Path.Combine(baseRoot,"conflict");
            string existing=Path.Combine(conflict,@"Microsoft\Excel\XLSTART\내엑셀 v0.01_r62.xlam");
            Directory.CreateDirectory(Path.GetDirectoryName(existing)); File.WriteAllText(existing,"user");
            var c=new NativeInstaller(version,hash,package,files,Console.WriteLine,conflict,prefix+@"conflict\",null);
            Reject(()=>c.Install()); Expect(File.ReadAllText(existing)=="user");
            Console.WriteLine("PASS|NativeInstallerTests|"+count+"|ISOLATED_HKCU_AND_FILES|"+baseRoot);
            return 0;
        } catch(Exception e) { Console.Error.WriteLine(e); return 1; }
        finally {
            // Fixed unique test namespace only; fixture files remain for inspection.
            foreach(var view in new[]{RegistryView.Registry32,RegistryView.Registry64})
                using(var key=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
                    key.DeleteSubKeyTree(prefix.TrimEnd('\\'),false);
        }
    }
}
