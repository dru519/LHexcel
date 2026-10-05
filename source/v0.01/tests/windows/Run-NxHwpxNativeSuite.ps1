param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$NxHostRoot,
    [string]$RunId = ''
)
$ErrorActionPreference='Stop'; Set-StrictMode -Version Latest
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
$sourceRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$ProductXlam=[IO.Path]::GetFullPath($ProductXlam)
$EvidenceRoot=[IO.Path]::GetFullPath($EvidenceRoot)
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Fresh evidence directory required'}
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Pre-existing Excel detected'}
. (Join-Path $sourceRoot 'build/Excel-ProcessLifecycle.ps1')
$office=Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
if($office.Platform -ne 'x64'){throw 'This native harness requires installed x64 Excel; x86 uses the component probe'}
$dll=[IO.Path]::GetFullPath((Join-Path $NxHostRoot 'NxHost64.dll'))
$classId='{1434C649-18AA-4435-B0D2-2BD81B548A01}'
$progId='LH.NxHost.HwpxExportService'
$keys=@("Software\Classes\CLSID\$classId","Software\Classes\$progId")
$registry=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::Registry64)
foreach($path in $keys){$existing=$registry.OpenSubKey($path);if($null -ne $existing){$existing.Dispose();$registry.Dispose();throw 'Existing HWPX service registration must be preserved'}}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$copy=Join-Path $EvidenceRoot 'HwpxProbe.xlam'; Copy-Item -LiteralPath $ProductXlam -Destination $copy
$before=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash
$assembly=[Reflection.AssemblyName]::GetAssemblyName($dll)
$className='LH.NxHost.HwpxExportService'
$codeBase=([Uri]$dll).AbsoluteUri
$priorProfile=$env:LHEXCEL_PROFILE_ROOT; $env:LHEXCEL_PROFILE_ROOT=Join-Path $EvidenceRoot 'profile'
$registered=$false;$excel=$null;$books=$null;$hostBook=$null;$addin=$null;$project=$null;$components=$null;$binding=$null;$cleanup=$null;$failure=$null
$cases=New-Object 'Collections.Generic.List[object]'
function Release-Com([object]$value){if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}}
function Record([string]$name,[string]$result){$cases.Add([pscustomobject]@{name=$name;result=$result;status=if($result.StartsWith('PASS|')){'PASS'}else{'FAIL'}});Write-Output $result;if(-not $result.StartsWith('PASS|')){throw "Failed: $name"}}
try{
    $baseline=Get-ExcelProcessBaseline
    $interactive=Start-ExactInteractiveExcel $baseline 'HWPX DLL Excel'
    $excel=$interactive.excel;$binding=$interactive.binding;$interactive=$null
    if(-not $binding.owned){throw 'Excel ownership rejected'}
    $excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.EnableEvents=$true
    $books=$excel.Workbooks;$hostBook=$books.Add();$addin=$books.Open($copy,$false,$false)
    $project=$addin.VBProject;$components=$project.VBComponents
    $component=$null;$module=$null
    try{
        $code=[IO.File]::ReadAllText((Join-Path $sourceRoot 'tests/vba/hwpx/T_NxHwpxNative.bas'),[Text.Encoding]::UTF8)
        $code=$code.Substring($code.IndexOf('Option Explicit')).Replace("`r`n","`n").Replace("`n","`r`n")
        $component=$components.Add(1);$component.Name='T_NxHwpxNative';$module=$component.CodeModule;[void]$module.AddFromString($code)
    }finally{Release-Com $module;Release-Com $component}
    $prefix="'"+$addin.Name+"'!"
    Record 'prepare' ([string]$excel.Run(($prefix+'NxHwpxNativePrepare'),$EvidenceRoot))
    Record 'no_dll' ([string]$excel.Run(($prefix+'NxHwpxNativeRun'),'unavailable'))
    Record 'legacy' ([string]$excel.Run(($prefix+'NxHwpxNativeLegacy'),$EvidenceRoot))
    # Register only the newly introduced class for this process test. Existing installed host/add-in keys are not touched.
    $registered=$true
    $key=$registry.CreateSubKey($keys[1]);try{$key.SetValue('',$className)}finally{$key.Dispose()}
    $key=$registry.CreateSubKey($keys[1]+'\CLSID');try{$key.SetValue('',$classId)}finally{$key.Dispose()}
    $key=$registry.CreateSubKey($keys[0]);try{$key.SetValue('',$className)}finally{$key.Dispose()}
    $key=$registry.CreateSubKey($keys[0]+'\ProgId');try{$key.SetValue('',$progId)}finally{$key.Dispose()}
    $key=$registry.CreateSubKey($keys[0]+'\InprocServer32')
    try{$key.SetValue('','mscoree.dll');$key.SetValue('ThreadingModel','Both');$key.SetValue('Class',$className);$key.SetValue('Assembly',$assembly.FullName);$key.SetValue('RuntimeVersion','v4.0.30319');$key.SetValue('CodeBase',$codeBase)}finally{$key.Dispose()}
    Record 'dll' ([string]$excel.Run(($prefix+'NxHwpxNativeRun'),'native'))
    Record 'invalid_restores_state' ([string]$excel.Run(($prefix+'NxHwpxNativeRun'),'invalid'))
    Record 'retry' ([string]$excel.Run(($prefix+'NxHwpxNativeRun'),'native'))
    Record 'custom_status' ([string]$excel.Run(($prefix+'NxHwpxNativeRun'),'custom'))
    $faultComponent=$null;$faultModule=$null;$faultLine=0;$original=''
    try{
        $faultComponent=$components.Item('NxHwpxController');$faultModule=$faultComponent.CodeModule
        $matches=@(1..$faultModule.CountOfLines | Where-Object {$faultModule.Lines($_,1).Trim() -eq 'started = Timer'})
        if($matches.Count -ne 1){throw 'Unique cancellation fault point missing'}
        $faultLine=$matches[0];$original=$faultModule.Lines($faultLine,1)
        $faultModule.ReplaceLine($faultLine,'    Err.Raise 18, "Native test-only cancellation", "User interrupt"')
        [void]$excel.Run(($prefix+'NxHwpxNativeReloadPayload'),$EvidenceRoot)
        Record 'cancel_restores_state' ([string]$excel.Run(($prefix+'NxHwpxNativeRun'),'cancel'))
    }finally{if($faultLine -gt 0){$faultModule.ReplaceLine($faultLine,$original)};Release-Com $faultModule;Release-Com $faultComponent}
    Start-Sleep -Milliseconds 250
    $exports=@(Get-ChildItem -LiteralPath (Join-Path $env:LHEXCEL_PROFILE_ROOT 'Temp/Hwpx') -Filter '*.hwpx')
    if($exports.Count -ne 3){throw 'Expected exactly three native exports'}
    if(@(Get-ChildItem -LiteralPath (Join-Path $env:LHEXCEL_PROFILE_ROOT 'Temp/Hwpx') -Filter '*.part').Count){throw 'Native partial residue'}
}catch{$failure=$_.Exception.ToString()}
finally{
    if($null -ne $addin){try{[void]$excel.Run(("'"+$addin.Name+"'!NxHwpxNativeDispose"))}catch{};try{$addin.Close($false)}catch{}}
    if($null -ne $hostBook){try{$hostBook.Close($false)}catch{}}
    Release-Com $hostBook;Release-Com $components;Release-Com $project;Release-Com $addin;Release-Com $books
    if($null -ne $excel){try{$excel.Quit()}catch{};Release-Com $excel}
    $hostBook=$null;$components=$null;$project=$null;$addin=$null;$books=$null;$excel=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null -ne $binding -and $null -ne $binding.process){$cleanup=Stop-ExactProcessAfterGrace $binding.process 'HWPX DLL Excel' 15000 -Detailed;$binding.process.Dispose()}
    if($registered){foreach($path in $keys){$registry.DeleteSubKeyTree($path,$false)}}
    $registry.Dispose();$env:LHEXCEL_PROFILE_ROOT=$priorProfile
}
$settle=[Diagnostics.Stopwatch]::StartNew()
do{$remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id);if($remaining.Count -eq 0){break};Start-Sleep -Milliseconds 100}while($settle.ElapsedMilliseconds -lt 2000)
$after=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash
$ok=$null -eq $failure -and $cases.Count -eq 8 -and $remaining.Count -eq 0 -and $null -ne $cleanup -and $cleanup.exit_mode -eq 'NATURAL' -and $before -eq $after
$receipt=[ordered]@{suite='NxHwpxNative';status=if($ok){'PASS'}else{'FAIL'};run_id=$RunId;cases=@($cases.ToArray());failure=$failure;cleanup=$cleanup;remaining_excel=$remaining;product_sha256=$before;product_sha256_after=$after;dll_sha256=(Get-FileHash -LiteralPath $dll -Algorithm SHA256).Hash;registration_scope='temporary new HWPX class only; removed';hancom_rendering='NOT_RUN'}
[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'NxHwpxNative.json'),(($receipt|ConvertTo-Json -Depth 12)+"`n"),(New-Object Text.UTF8Encoding($false)))
if(-not $ok){throw ('Native HWPX suite failed: '+$failure)}
Write-Output 'PASS|NxHwpxNative|8/8'
