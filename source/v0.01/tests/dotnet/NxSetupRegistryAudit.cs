using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using Microsoft.Win32;
using LH.NxSetup;
public static class NxSetupRegistryAudit {
 public static int Main(string[] args){
  var zip=File.ReadAllBytes(args[0]);var hash=Package.Hash(zip);
  var installer=new NativeInstaller("v0.01_r97",hash,Package.CachedPath("v0.01_r97",hash),Package.Read(zip,"v0.01_r97"),Console.WriteLine);
  foreach(var view in new[]{RegistryView.Registry32,RegistryView.Registry64}){
   var rows=(Dictionary<string,Dictionary<string,object>>)typeof(NativeInstaller).GetMethod("RegistryRows",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(installer,new object[]{view});
   using(var root=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))foreach(var row in rows)using(var key=root.OpenSubKey(row.Key)){
    if(key==null){Console.WriteLine("MISSING "+view+" "+row.Key);continue;}
    foreach(string name in key.GetValueNames()){
     object expected;var actual=key.GetValue(name);
     if(!row.Value.TryGetValue(name,out expected)||!Object.Equals(actual,expected)||key.GetValueKind(name)!=(actual is int?RegistryValueKind.DWord:RegistryValueKind.String))Console.WriteLine(view+" "+row.Key+" ["+name+"] expected="+expected+" actual="+actual+" kind="+key.GetValueKind(name));
    }
   }
  }
  try{typeof(NativeInstaller).GetMethod("ValidateOwned",BindingFlags.Instance|BindingFlags.NonPublic).Invoke(installer,new object[]{true});Console.WriteLine("VALIDATE_PASS");}catch(Exception error){Console.WriteLine(error);return 1;}return 0;
 }
}
