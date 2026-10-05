"""Generate thin COM proxies from the existing interfaces, never core references."""
import argparse
import hashlib
import re
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
SERVICES=('BridgeService','HwpxExportService','PicturePreviewService','WorkbookCompareService','WorkbookCompareResultsService')

def interface(text, name):
    match=re.search(r'    \[ComVisible\(true\), Guid\("[^"]+"\), InterfaceType\(ComInterfaceType.InterfaceIsDual\)\]\s+public interface '+name+r'\s*{',text)
    if not match:raise ValueError('Interface metadata changed: '+name)
    start=match.end();depth=1;end=start
    while depth:
        char=text[end];depth+= (char=='{')-(char=='}');end+=1
    return text[match.start():end],text[start:end-1]

def class_attribute(text,name):
    if name == 'NavigatorPane':
        pattern=r'(\[ComVisible\(true\)\]\s*\[Guid\("[^"]+"\), ProgId\("LH.NxHost.NavigatorPane"\), ClassInterface\(ClassInterfaceType.None\)\]\s*\[ComDefaultInterface\(typeof\(INavigatorPaneControl\)\)\])\s*public sealed class NavigatorPane'
        match=re.search(pattern,text)
        if not match:raise ValueError('Navigator COM metadata changed')
        return match.group(1)
    match=re.search(r'(\[ComVisible\(true\), Guid\("[^"]+"\), ProgId\("[^"]+"\), ClassInterface\(ClassInterfaceType.None\)\])\s+public sealed class '+name+r'\b',text)
    if not match:raise ValueError('Class metadata changed: '+name)
    return match.group(1)

def proxies():
    parts=['using System; using System.Runtime.InteropServices; using System.Windows.Forms; using Extensibility; using Microsoft.Office.Core;\nnamespace LH.NxHost {']
    for name in SERVICES:
        text=(ROOT/'src/dotnet/NxHost'/f'{name}.cs').read_text(encoding='utf-8')
        declaration,body=interface(text,'I'+name)
        parts += [declaration,class_attribute(text,name),f'public sealed class {name} : I{name} {{',f'private readonly object core = ProtectedCore.Create("{name}");']
        for line in body.strip().splitlines():
            line=re.sub(r'\[DispId\([0-9]+\)\]\s*','',line.strip())
            prop=re.fullmatch(r'(string|bool|int|object) (\w+)\s*{\s*get;\s*}',line)
            if prop:
                kind,method=prop.groups()
                parts.append(f'public {kind} {method} {{ get {{ return ({kind})ProtectedCore.Get(core,"{method}"); }} }}')
                continue
            method=re.fullmatch(r'(string|bool|int|object|void) (\w+)\(([^)]*)\);',line)
            if not method:raise ValueError('Unsupported COM member: '+line)
            kind,member,parameters=method.groups()
            args=[p.strip().split()[-1] for p in parameters.split(',')] if parameters.strip() else []
            call=f'ProtectedCore.Call(core,"{member}"'+(' ,'+','.join(args) if args else '')+')'
            statement=call+';' if kind=='void' else f'return ({kind})'+call+';'
            parts.append(f'public {kind} {member}({parameters}) {{ {statement} }}')
        parts.append('}')
    addin=(ROOT/'src/dotnet/NxHost/NxHostAddIn.cs').read_text(encoding='utf-8')
    parts += [class_attribute(addin,'NxHostAddIn'),'''
public sealed class NxHostAddIn : IDTExtensibility2, ICustomTaskPaneConsumer {
    private readonly object core = ProtectedCore.Create("NxHostAddIn");
    public void OnConnection(object application, ext_ConnectMode mode, object instance, ref Array custom) {
        ((IDTExtensibility2)core).OnConnection(application,mode,instance,ref custom);
    }
    public void OnDisconnection(ext_DisconnectMode mode, ref Array custom) { ((IDTExtensibility2)core).OnDisconnection(mode,ref custom); }
    public void OnAddInsUpdate(ref Array custom) { ((IDTExtensibility2)core).OnAddInsUpdate(ref custom); }
    public void OnStartupComplete(ref Array custom) { ((IDTExtensibility2)core).OnStartupComplete(ref custom); }
    public void OnBeginShutdown(ref Array custom) { ((IDTExtensibility2)core).OnBeginShutdown(ref custom); }
    public void CTPFactoryAvailable(ICTPFactory factory) { ((ICustomTaskPaneConsumer)core).CTPFactoryAvailable(factory); }
}
''']
    pane=(ROOT/'src/dotnet/NxHost/NavigatorPane.cs').read_text(encoding='utf-8')
    pane_interface=re.search(r'\[ComVisible\(true\), Guid\("[^"]+"\), InterfaceType\(ComInterfaceType.InterfaceIsIDispatch\)\]\s*public interface INavigatorPaneControl { }',pane)
    if not pane_interface:raise ValueError('Navigator dispatch interface changed')
    parts += [pane_interface.group(0),class_attribute(pane,'NavigatorPane'),'''
public sealed class NavigatorPane : UserControl, INavigatorPaneControl {
    public NavigatorPane() {
        AutoScaleMode=AutoScaleMode.None;
        var content=(UserControl)ProtectedCore.Create("NavigatorPane");
        content.Dock=DockStyle.Fill;
        Controls.Add(content);
    }
}
}''']
    return '\n'.join(parts)

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--core',type=Path,required=True)
    parser.add_argument('--mvid',required=True)
    parser.add_argument('--out',type=Path,required=True)
    args=parser.parse_args()
    if args.core.name not in ('NxCore32.dll','NxCore64.dll'):raise ValueError('Unexpected core name')
    if not re.fullmatch(r'[0-9a-f-]{36}',args.mvid):raise ValueError('Invalid module identity')
    args.out.mkdir(parents=True,exist_ok=True)
    (args.out/'Proxies.cs').write_text(proxies(),encoding='utf-8')
    digest=hashlib.sha256(args.core.read_bytes()).hexdigest()
    (args.out/'CorePin.cs').write_text('namespace LH.NxHost { internal static class CorePin {\n'+
        f'internal const string FileName="{args.core.name}", Sha256="{digest}", ModuleId="{args.mvid}";\n'+
        '} }\n',encoding='utf-8')

if __name__=='__main__':main()
