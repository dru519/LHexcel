param([ValidateSet('Audit','Detach','SelfTest')][string]$Mode='Audit', [string]$EvidenceRoot)
# Recovery of this exact owned r71 test fixture only. Never reads/loads the blocked DLL.
# This is not an antivirus exception, a general uninstaller or an automatic rollback.
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
function ExpectedRows([string]$runtime,[string]$arch) {
    $ids=@('4A4B4F02-76A6-4F22-BD4B-85E0308EC91D','9A90294C-81C8-4AEE-A2D8-14B078A61CD2','A13D5ED6-B0F2-43C0-9499-E53E6CE38DC0','1434C649-18AA-4435-B0D2-2BD81B548A01','C812F023-E912-49AA-98B7-75DC86501902','C812F023-E912-49AA-98B7-75DC86501904','C812F023-E912-49AA-98B7-75DC86501906')
    $progs=@('Connect','BridgeService','NavigatorPane','HwpxExportService','PicturePreviewService','WorkbookCompareService','WorkbookCompareResultsService')
    $classes=@('NxHostAddIn','BridgeService','NavigatorPane','HwpxExportService','PicturePreviewService','WorkbookCompareService','WorkbookCompareResultsService')
    $dll=if($arch -eq 'x86'){'NxHost32'}else{'NxHost64'}
    $codebase=([Uri](Join-Path $runtime "$arch/$dll.dll")).AbsoluteUri
    $rows=@{}; $roots=@()
    for($i=0;$i -lt $ids.Count;$i++) {
        $id='{'+$ids[$i]+'}'; $prog='Software\Classes\LH.NxHost.'+$progs[$i]; $cls='Software\Classes\CLSID\'+$id
        $class='LH.NxHost.'+$classes[$i]; $roots+=@($prog,$cls)
        $rows[$prog]=@{''=$class}; $rows[$prog+'\CLSID']=@{''=$id}
        $rows[$cls]=@{''=$class}; $rows[$cls+'\ProgId']=@{''=('LH.NxHost.'+$progs[$i])}
        $rows[$cls+'\Implemented Categories']=@{}
        $rows[$cls+'\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}']=@{}
        $managed=@{Class=$class; Assembly="$dll, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null"; RuntimeVersion='v4.0.30319'; CodeBase=$codebase}
        $rows[$cls+'\InprocServer32\0.0.0.0']=$managed.Clone()
        $managed['']='mscoree.dll'; $managed['ThreadingModel']='Both'; $rows[$cls+'\InprocServer32']=$managed
    }
    $addin='Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect'
    $rows[$addin]=@{FriendlyName='내엑셀 NxHost'; Description='내엑셀 Enhanced DLL host'; LoadBehavior=3; CommandLineSafe=1}
    @{Rows=$rows; Roots=($roots+@($addin))}
}
function ReadTree($key,[string]$path,$result) {
    $values=@{}
    foreach($name in $key.GetValueNames()) {
        $values[$name]=@{Value=$key.GetValue($name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames); Kind=$key.GetValueKind($name).ToString()}
    }
    $result[$path]=$values
    foreach($child in $key.GetSubKeyNames()) {
        $nested=$key.OpenSubKey($child)
        try { ReadTree $nested ($path+'\'+$child) $result } finally { $nested.Dispose() }
    }
}
function AssertRows($actual,$expected) {
    if($actual.Count -ne $expected.Count){throw 'Unexpected registry inventory; preserved'}
    foreach($path in $expected.Keys) {
        if(!$actual.ContainsKey($path) -or $actual[$path].Count -ne $expected[$path].Count){throw "Changed registry node: $path"}
        foreach($name in $expected[$path].Keys) {
            $value=$expected[$path][$name]; $kind=if($value -is [int]){'DWord'}else{'String'}
            if(!$actual[$path].ContainsKey($name) -or $actual[$path][$name].Kind -cne $kind -or $actual[$path][$name].Value -cne $value){throw "Changed registry value: $path / $name"}
        }
    }
}
function Snapshot($base,$roots) {
    $actual=@{}
    foreach($path in $roots) {
        $key=$base.OpenSubKey($path)
        if(!$key){throw "Missing fixture registry root: $path"}
        try {ReadTree $key $path $actual} finally {$key.Dispose()}
    }
    return $actual
}
if($Mode -eq 'SelfTest') {
    $expected=@{'node'=@{A='value';N=3}}
    $actual=@{'node'=@{A=@{Value='value';Kind='String'};N=@{Value=3;Kind='DWord'}}}
    AssertRows $actual $expected
    $rejected=0
    foreach($variant in @('extra','changed','kind','missing')) {
        $bad=@{'node'=@{A=@{Value='value';Kind='String'};N=@{Value=3;Kind='DWord'}}}
        switch($variant){extra {$bad.extra=@{}} changed {$bad.node.A.Value='other'} kind {$bad.node.N.Kind='String'} missing {$bad.node.Remove('A')}}
        try {AssertRows $bad $expected} catch {$rejected++}
    }
    if($rejected -ne 4){throw 'Validation negative test failed'}
    Write-Output 'PASS|RecoveryValidation|5'; return
}
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Excel open; preserved'}
$runtime=Join-Path $env:LOCALAPPDATA 'LHexcel/NxHost/r71'
$receipt=Join-Path $env:LOCALAPPDATA 'LHexcel/NxHost/state/r71/native-install.xml'
$xlstart=Join-Path $env:APPDATA 'Microsoft/Excel/XLSTART/내엑셀 v0.01_r71.xlam'
$held=Join-Path $env:LOCALAPPDATA 'LHExcel/verification/r72-security-hold-20260912/내엑셀 v0.01_r71.xlam'
$hashes=@{}
$hashes[(Join-Path $runtime 'x86/NxHost32.dll')]='a6de9aa52388526d842b13e1821dce5f7d4f10f02d7250e55441010880b89974'
$hashes[(Join-Path $runtime 'x64/NxHost64.dll')]='df921d18d1317a625c8353ae995576a39e7d877d7836f8a36ca61b2eeda20aa6'
$hashes[$xlstart]='964d16ef4373a92815e17877a143aa4986d3284bd95f42f3cac79d2bd5908748'
if((Get-Item -LiteralPath $receipt).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Reparse receipt refused'}
$settings=New-Object System.Xml.XmlReaderSettings
$settings.DtdProcessing=[System.Xml.DtdProcessing]::Prohibit; $settings.XmlResolver=$null
$reader=[System.Xml.XmlReader]::Create($receipt,$settings)
$doc=New-Object System.Xml.XmlDocument
try {$doc.Load($reader)} finally {$reader.Dispose()}
$install=$doc.DocumentElement
if($install.version -cne 'v0.01_r71' -or $install.state -cne 'ACTIVE' -or $install.priorState -cne 'absent' -or $install.packageSha256 -cne 'bee9fa619491374ad8d8e1f72038f204e336c9315fd83ec60156427639ba1dc7'){throw 'Not the owned fixture receipt'}
$seen=@{}
foreach($entry in $install.SelectNodes('file')) {
    if(!$hashes.ContainsKey($entry.path) -or $seen.ContainsKey($entry.path) -or $entry.sha256 -cne $hashes[$entry.path]){throw 'Fixture receipt file mismatch'}
    $seen[$entry.path]=$true
}
if($seen.Count -ne 3 -or (Test-Path -LiteralPath $xlstart) -or (Get-FileHash -LiteralPath $held).Hash.ToLowerInvariant() -cne $hashes[$xlstart]){throw 'XLSTART fixture hold not confirmed'}
$records=@(); $handles=@()
try {
    foreach($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
        if($view -eq [Microsoft.Win32.RegistryView]::Registry64 -and ![Environment]::Is64BitOperatingSystem){continue}
        $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view); $handles+=,$base
        $arch=if($view -eq [Microsoft.Win32.RegistryView]::Registry32){'x86'}else{'x64'}
        $expected=ExpectedRows $runtime $arch; $actual=Snapshot $base $expected.Roots; AssertRows $actual $expected.Rows
        $records+=@{View=$view.ToString(); Base=$base; Expected=$expected; Actual=$actual}
    }
    if($Mode -eq 'Audit'){Write-Output ('PASS|OwnedFixtureAudit|views='+$records.Count); return}
    if(!$EvidenceRoot -or (Test-Path -LiteralPath $EvidenceRoot)){throw 'Fresh backup directory required'}
    New-Item -ItemType Directory -Path $EvidenceRoot | Out-Null
    Copy-Item -LiteralPath $receipt -Destination (Join-Path $EvidenceRoot 'native-install.before.xml')
    foreach($record in $records) {
        $backup=Join-Path $EvidenceRoot ($record.View+'.clixml')
        $record.Actual | Export-Clixml -LiteralPath $backup
        AssertRows (Import-Clixml -LiteralPath $backup) $record.Expected.Rows
    }
    foreach($record in $records){AssertRows (Snapshot $record.Base $record.Expected.Roots) $record.Expected.Rows}
    foreach($record in $records) {
        foreach($path in $record.Expected.Roots){$record.Base.DeleteSubKeyTree($path,$false)}
        foreach($path in $record.Expected.Roots){$remaining=$record.Base.OpenSubKey($path); if($remaining){$remaining.Dispose();throw 'Registration still exists'}}
    }
    @{Status='REGISTRATION_DETACHED';Version='v0.01_r71';DllFiles='PRESERVED_NOT_LOADED';Receipt='PRESERVED_NOT_MARKED_UNINSTALLED';Backup=$EvidenceRoot} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'recovery.json') -Encoding UTF8
    Write-Output 'PASS|RegistrationDetached|DLL files and old receipt preserved'
} finally {foreach($handle in $handles){$handle.Dispose()}}
