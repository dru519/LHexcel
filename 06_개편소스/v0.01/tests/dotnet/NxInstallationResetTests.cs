using System;using System.IO;using Microsoft.Win32;using LH.NxSetup.Reset;
class NxInstallationResetTests {
 static int count;static void Assert(bool value,string label){if(!value)throw new Exception(label);count++;}
 static void Put(string path){Directory.CreateDirectory(Path.GetDirectoryName(path));File.WriteAllText(path,"fixture");}
 static int Main(){
  string root=Path.Combine(Path.GetTempPath(),"nx-reset-test-"+Guid.NewGuid().ToString("N"));
  string prefix=@"Software\LHExcelResetTests\"+Guid.NewGuid().ToString("N")+@"\";
  var engine=new ResetEngine(Path.Combine(root,"local"),Path.Combine(root,"roaming"),prefix);
  string product=Path.Combine(root,"local","LHexcel"),host=Path.Combine(product,"NxHost");
  string dll=Path.Combine(host,"r95","x86","NxHost32.dll"),receipt=Path.Combine(host,"state","r98","native-install.xml");
  string startup=Path.Combine(root,"roaming","Microsoft","Excel","XLSTART","내엑셀 v0.01_r98.xlam");
  string template=Path.Combine(root,"roaming","Microsoft","Excel","XLSTART","내엑셀 템플릿","양식.xltx");
  try{
   Put(dll);Put(receipt);Put(startup);Put(template);Put(Path.Combine(product,"Setup","v0.01_r95","Setup-fixture.exe"));
   Put(Path.Combine(product,"Settings","user.cfg"));Put(Path.Combine(root,"ordinary.xlsx"));
   foreach(var v in ResetEngine.Views)using(var registry=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,v)){
    using(var k=registry.CreateSubKey(prefix+v+@"\Software\Classes\CLSID\"+ResetEngine.Ids[0]+@"\InprocServer32"))k.SetValue("CodeBase",new Uri(dll).AbsoluteUri);
    foreach(string rev in new[]{"95","98"})using(var k=registry.CreateSubKey(prefix+v+@"\Software\Microsoft\Windows\CurrentVersion\Uninstall\"+(rev=="95"?"LHExcel-v0.01_r95":"LHExcel"))){
     k.SetValue("DisplayName","내엑셀 v0.01_r"+rev);k.SetValue("InstallLocation",Path.Combine(host,"r"+rev));
    }
   }
   Assert(engine.Inspect().Keys.Count==3*ResetEngine.Views.Length,"both old and new product records detected");
   string document=Path.Combine(host,"r95","document.xlsx");Put(document);
   bool blocked=false;try{engine.Execute(s=>{});}catch(IOException){blocked=true;}
   Assert(blocked&&File.Exists(dll),"document in deletion area blocks entire reset");
   File.Delete(document);
   using(var locked=new FileStream(dll,FileMode.Open,FileAccess.Read,FileShare.None)){
    blocked=false;try{engine.Execute(s=>{});}catch(IOException){blocked=true;}
    Assert(blocked,"locked payload blocked");Assert(engine.Inspect().Keys.Count==3*ResetEngine.Views.Length,"no registry changes before file preflight");
   }
   var view=ResetEngine.Views[0];string keyPath=prefix+view+@"\Software\Classes\CLSID\"+ResetEngine.Ids[0]+@"\InprocServer32";
   using(var registry=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))using(var k=registry.OpenSubKey(keyPath,true))k.SetValue("CodeBase","file:///C:/unrelated/NxHost32.dll");
   blocked=false;try{engine.Execute(s=>{});}catch(IOException){blocked=true;}
   Assert(blocked&&File.Exists(dll),"foreign COM path rejected");
   using(var registry=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))using(var k=registry.OpenSubKey(keyPath,true))k.SetValue("CodeBase",new Uri(dll).AbsoluteUri);
   engine.Execute(Console.WriteLine);
   Assert(!File.Exists(dll)&&!File.Exists(receipt)&&!File.Exists(startup),"mixed r95 and r98 installation removed");
   Assert(!Directory.Exists(Path.Combine(product,"Setup","v0.01_r95")),"old installer cache deleted without backup");
   Assert(File.Exists(template)&&File.Exists(Path.Combine(root,"ordinary.xlsx"))&&File.Exists(Path.Combine(product,"Settings","user.cfg")),"documents templates settings retained");
   Assert(engine.Inspect().Keys.Count==0,"all owned registry records removed");
   engine.Execute(s=>{});Assert(engine.Inspect().Files.Count==0,"repeated reset safe");
   Console.WriteLine("PASS|RESET|"+count+"|ISOLATED_ONLY");return 0;
  }catch(Exception e){Console.WriteLine(e);return 1;}
  finally{foreach(var v in ResetEngine.Views)using(var registry=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,v))registry.DeleteSubKeyTree(prefix.TrimEnd('\\'),false);}
 }
}
