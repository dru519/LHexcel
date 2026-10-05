param(
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$ArtifactPath,
    [string]$NxHostRoot = '',
    [string]$RunId = '',
    [string]$SourceDigest = '',
    [string]$SnapshotDigest = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }
if (-not [Environment]::Is64BitOperatingSystem) { throw 'NxHost Navigator validation requires 64-bit Windows' }

$SourceRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$EvidenceRoot=[IO.Path]::GetFullPath($EvidenceRoot)
$ArtifactPath=[IO.Path]::GetFullPath($ArtifactPath)
if([string]::IsNullOrWhiteSpace($RunId)){$RunId=[Guid]::NewGuid().ToString().ToLowerInvariant()}
if($RunId -notmatch '^[0-9a-f-]{36}$'){throw 'RunId must be a lowercase GUID'}
if(-not [string]::IsNullOrWhiteSpace($SourceDigest) -and $SourceDigest -notmatch '^[0-9a-f]{64}$'){throw 'SourceDigest must be a lowercase SHA-256'}
if(-not [string]::IsNullOrWhiteSpace($SnapshotDigest) -and $SnapshotDigest -notmatch '^[0-9a-f]{64}$'){throw 'SnapshotDigest must be a lowercase SHA-256'}
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Existing NxHost Navigator evidence root is rejected'}
if(-not(Test-Path -LiteralPath $ArtifactPath -PathType Leaf)-or[IO.Path]::GetExtension($ArtifactPath)-ine'.xlam'){throw 'NxHost Navigator Product artifact rejected'}
if(Get-Process -Name EXCEL -ErrorAction SilentlyContinue){throw 'NxHost Navigator suite requires zero pre-existing Excel processes'}

. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class NxNavigatorNative {
    private delegate bool EnumChildProc(IntPtr hwnd, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumChildProc callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumChildProc callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetClassName(IntPtr hwnd, StringBuilder value, int maximum);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] private static extern int GetWindowText(IntPtr hwnd, StringBuilder value, int maximum);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr hwnd);
    public static string[] VisibleChildClasses(long parent) {
        var values=new List<string>();
        EnumChildWindows(new IntPtr(parent),(hwnd,state)=>{
            if(!IsWindowVisible(hwnd)) return true;
            var name=new StringBuilder(256);
            if(GetClassName(hwnd,name,name.Capacity)>0) values.Add(name.ToString());
            return true;
        },IntPtr.Zero);
        return values.ToArray();
    }
    public static int CountVisibleTopLevelWindows(int wantedProcessId, string prefix) {
        int count=0;
        EnumWindows((hwnd,state)=>{
            uint processId;
            GetWindowThreadProcessId(hwnd,out processId);
            if(processId!=(uint)wantedProcessId || !IsWindowVisible(hwnd)) return true;
            var title=new StringBuilder(256);
            GetWindowText(hwnd,title,title.Capacity);
            if(title.ToString().StartsWith(prefix,StringComparison.Ordinal)) count++;
            return true;
        },IntPtr.Zero);
        return count;
    }
}
'@
$NavigatorTitle=([char]0xB0B4).ToString()+([char]0xC5D1).ToString()+([char]0xC140).ToString()+' Navigator'

function Release-ComObject([object]$Value){if($null-ne$Value-and[Runtime.InteropServices.Marshal]::IsComObject($Value)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)}catch{}}}
function Get-Sha256([string]$Path){(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()}
function Write-Utf8Json([string]$Path,[object]$Value){$encoding=New-Object Text.UTF8Encoding($false);[IO.File]::WriteAllText($Path,($Value|ConvertTo-Json -Depth 24),$encoding)}
function Add-Case([string]$Name,[bool]$Passed,[object]$Details){$script:Cases.Add([pscustomobject][ordered]@{ordinal=$script:Cases.Count+1;name=$Name;status=if($Passed){'PASS'}else{'FAIL'};details=$Details})}
function Get-NavigatorUiaSnapshot([int64]$ExcelHwnd){
    $excelRoot=[Windows.Automation.AutomationElement]::FromHandle([IntPtr]$ExcelHwnd)
    if($null-eq$excelRoot){return [pscustomobject][ordered]@{process_id=0;visible_titles=0;edits=0;combos=0;lists=0;bounds=@()}}
    $processId=[int]$excelRoot.Current.ProcessId
    $root=[Windows.Automation.AutomationElement]::RootElement
    $pidCondition=New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty,$processId)
    $visibleTitles=New-Object 'Collections.Generic.List[object]'
    $titleContainers=New-Object 'Collections.Generic.List[Windows.Automation.AutomationElement]'
    $nativeClasses=New-Object 'Collections.Generic.List[string]'
    $titleCondition=New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty,$NavigatorTitle)
    $titleConditions=[Windows.Automation.Condition[]]@($pidCondition,$titleCondition)
    $exactTitleCondition=[Windows.Automation.AndCondition]::new($titleConditions)
    foreach($element in @($root.FindAll([Windows.Automation.TreeScope]::Descendants,$exactTitleCondition))){
        try{$bounds=$element.Current.BoundingRectangle;if(-not$element.Current.IsOffscreen-and$bounds.Width-gt0-and$bounds.Height-gt0){$titleContainers.Add($element);$nativeHandle=[int64]$element.Current.NativeWindowHandle;if($nativeHandle-ne0){foreach($className in [NxNavigatorNative]::VisibleChildClasses($nativeHandle)){$nativeClasses.Add($className)}};$visibleTitles.Add([pscustomobject][ordered]@{name=[string]$element.Current.Name;hwnd=$nativeHandle;left=[int][Math]::Round($bounds.Left);top=[int][Math]::Round($bounds.Top);width=[int][Math]::Round($bounds.Width);height=[int][Math]::Round($bounds.Height)})}}catch{}
    }
    $counts=@{}
    foreach($entry in @(@('edits',[Windows.Automation.ControlType]::Edit),@('combos',[Windows.Automation.ControlType]::ComboBox),@('lists',[Windows.Automation.ControlType]::List))){
        $typeCondition=New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ControlTypeProperty,$entry[1])
        $count=0
        foreach($container in @($titleContainers.ToArray())){
            foreach($element in @($container.FindAll([Windows.Automation.TreeScope]::Descendants,$typeCondition))){try{$bounds=$element.Current.BoundingRectangle;if(-not$element.Current.IsOffscreen-and$bounds.Width-gt0-and$bounds.Height-gt0){$count++}}catch{}}
        }
        $counts[$entry[0]]=$count
    }
    $nativeEdits=@($nativeClasses|Where-Object{$_.IndexOf('EDIT',[StringComparison]::OrdinalIgnoreCase)-ge0}).Count
    $nativeCombos=@($nativeClasses|Where-Object{$_.IndexOf('COMBOBOX',[StringComparison]::OrdinalIgnoreCase)-ge0}).Count
    $nativeLists=@($nativeClasses|Where-Object{$_.IndexOf('LISTBOX',[StringComparison]::OrdinalIgnoreCase)-ge0}).Count
    [pscustomobject][ordered]@{process_id=$processId;visible_titles=$visibleTitles.Count;edits=[Math]::Max([int]$counts.edits,$nativeEdits);combos=[Math]::Max([int]$counts.combos,$nativeCombos);lists=[Math]::Max([int]$counts.lists,$nativeLists);uia_counts=[ordered]@{edits=[int]$counts.edits;combos=[int]$counts.combos;lists=[int]$counts.lists};native_classes=@($nativeClasses.ToArray()|Sort-Object -Unique);bounds=@($visibleTitles.ToArray())}
}
function Wait-NavigatorState([int64]$ExcelHwnd,[bool]$Visible,[int]$TimeoutMs=10000){
    $timer=[Diagnostics.Stopwatch]::StartNew();$snapshot=$null
    while($timer.ElapsedMilliseconds-lt$TimeoutMs){
        $snapshot=Get-NavigatorUiaSnapshot $ExcelHwnd
        if($Visible){if($snapshot.visible_titles-eq1-and$snapshot.edits-ge1-and$snapshot.combos-ge1-and$snapshot.lists-ge1){return $snapshot}}
        elseif($snapshot.visible_titles-eq0){return $snapshot}
        Start-Sleep -Milliseconds 100
    }
    return $snapshot
}
function Wait-FocusWindowCount([int]$ProcessId,[int]$Expected,[int]$TimeoutMs=5000){
    $timer=[Diagnostics.Stopwatch]::StartNew();$count=-1
    while($timer.ElapsedMilliseconds-lt$TimeoutMs){
        $count=[NxNavigatorNative]::CountVisibleTopLevelWindows($ProcessId,'NX_FOCUS_DLL_')
        if($count-eq$Expected){return $count}
        Start-Sleep -Milliseconds 100
    }
    return $count
}

$script:Cases=New-Object 'Collections.Generic.List[object]'
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$artifact=Join-Path $EvidenceRoot 'Product.xlam'
Copy-Item -LiteralPath $ArtifactPath -Destination $artifact
if((Get-Sha256 $artifact)-cne(Get-Sha256 $ArtifactPath)){throw 'NxHost Navigator Product artifact copy mismatch'}
$artifactSha=Get-Sha256 $artifact
$binaryEvidence=Join-Path $EvidenceRoot 'binary'
$binaryArguments=@{EvidenceRoot=$binaryEvidence;ArtifactPath=$artifact;RunId=$RunId;EvidenceMode='ReleaseBound'}
if(-not[string]::IsNullOrWhiteSpace($NxHostRoot)){$binaryArguments['NxHostRoot']=[IO.Path]::GetFullPath($NxHostRoot)}
if(-not[string]::IsNullOrWhiteSpace($SourceDigest)){$binaryArguments['SourceDigest']=$SourceDigest}
if(-not[string]::IsNullOrWhiteSpace($SnapshotDigest)){$binaryArguments['SnapshotDigest']=$SnapshotDigest}
$binaryOutput=@(& (Join-Path $PSScriptRoot 'Run-NxHostSuite.ps1') @binaryArguments)
$binaryOutput|Write-Output
$x86Dll=Join-Path $binaryEvidence 'nxhost/x86/NxHost32.dll'
$x64Dll=Join-Path $binaryEvidence 'nxhost/x64/NxHost64.dll'
$registerScript=Join-Path $SourceRoot 'build/Register-NxHost.ps1'
$statePath=Join-Path $env:LOCALAPPDATA 'LHexcel/NxHost/state/r57/registry-state.json'
$originalProfile=[string]$env:LHEXCEL_PROFILE_ROOT
$env:LHEXCEL_PROFILE_ROOT=Join-Path $EvidenceRoot 'profile/LHexcel'
[void](New-Item -ItemType Directory -Path $env:LHEXCEL_PROFILE_ROOT -Force)

$excel=$null;$binding=$null;$books=$null;$product=$null;$book=$null;$comAddins=$null;$hostAddin=$null
$registrationAttempted=$false;$registered=$false;$connected=$false;$failure=$null;$cleanupFailure=$null;$excelCleanup=[ordered]@{exit_mode='NOT_STARTED';failure=$null;owned_pid=0};$registryBefore='';$registryAfter=''
try{
    $registrationAttempted=$true
    $registrationOutput=@(& $registerScript -Action Register -OptIn -X86Dll $x86Dll -X64Dll $x64Dll)
    $registrationOutput|Write-Output
    $registered=$true
    if(-not(Test-Path -LiteralPath $statePath -PathType Leaf)){throw 'NxHost registry state receipt missing'}
    $state=[IO.File]::ReadAllText($statePath)|ConvertFrom-Json
    $registryBefore=[string]$state.snapshot_sha256
    Add-Case 'register_exact_dlls' ($registryBefore-match'^[0-9a-f]{64}$') @{registry_snapshot_before=$registryBefore;nxhost32_sha256=Get-Sha256 $x86Dll;nxhost64_sha256=Get-Sha256 $x64Dll}

    $baseline=Get-ExcelProcessBaseline
    $interactive=Start-ExactInteractiveExcel $baseline 'NxHost Navigator Excel'
    $excel=$interactive.excel;$binding=$interactive.binding;$interactive=$null
    if(-not[bool]$binding.owned){throw('Excel ownership rejected: '+[string]$binding.failure)}
    $excel.Visible=$true;$excel.DisplayAlerts=$false
    $comAddins=$excel.COMAddIns
    $hostAddin=$comAddins.Item('LH.NxHost.Connect')
    $hostAddin.Connect=$true;$connected=[bool]$hostAddin.Connect
    if(-not$connected){throw 'NxHost COM add-in did not connect'}
    $books=$excel.Workbooks
    $product=$books.Open($artifact,$false,$true)
    $book=$books.Add()
    $prefix=$product.Name+'!'

    $focusSheet=$null;$focusCell=$null
    try{
        $focusSheet=$book.Worksheets.Item(1)
        $focusCell=$focusSheet.Range('B2')
        [void]$focusCell.Select()
        $settingsResult=[string]$excel.Run($prefix+'NxFocusTryCommitSettingsValues','criss-cross','wide-stripe','#5FC8D8',34,$true,$false)
        if($settingsResult-cne'PASS'){throw('Focus settings probe failed: '+$settingsResult)}
        [void]$excel.Run($prefix+'NxFocusEnable',255)
        [void]$excel.Run($prefix+'NxFocusControllerRefreshActiveSelection')
        $focusBackend=[string]$excel.Run($prefix+'NxFocusControllerBackend')
        $focusDiagnostic=[string]$excel.Run($prefix+'NxFocusControllerBackendDiagnostic')
        Add-Case 'focus_dll_backend' ($focusBackend-ceq'dll'-and[string]::IsNullOrWhiteSpace($focusDiagnostic)) ([ordered]@{backend=$focusBackend;diagnostic=$focusDiagnostic;settings=$settingsResult})
        $focusVisible=Wait-FocusWindowCount ([int]$binding.pid) 8
        Add-Case 'focus_dll_windows_visible' ($focusVisible-eq8) ([ordered]@{visible_windows=$focusVisible;expected=8})
        [void]$excel.Run($prefix+'NxFocusDisable')
        $focusStopped=Wait-FocusWindowCount ([int]$binding.pid) 0
        Add-Case 'focus_dll_stop_zero_windows' ($focusStopped-eq0) ([ordered]@{visible_windows=$focusStopped;expected=0})
    }finally{
        Release-ComObject $focusCell
        Release-ComObject $focusSheet
    }

    $show1=[bool]$excel.Run($prefix+'NxHostShowNavigator')
    $visible1=Wait-NavigatorState ([int64]$excel.Hwnd) $true
    $show1Details=[ordered]@{show_result=$show1;uia=$visible1}
    Add-Case 'show_navigator' ($show1-and$visible1.visible_titles-eq1-and$visible1.edits-ge1-and$visible1.combos-ge1-and$visible1.lists-ge1) $show1Details

    $show2=[bool]$excel.Run($prefix+'NxHostShowNavigator')
    $visible2=Wait-NavigatorState ([int64]$excel.Hwnd) $true
    Add-Case 'reuse_single_navigator' ($show2-and$visible2.visible_titles-eq1) ([ordered]@{show_result=$show2;uia=$visible2})

    $hide=[bool]$excel.Run($prefix+'NxHostHideNavigator')
    $hidden=Wait-NavigatorState ([int64]$excel.Hwnd) $false
    Add-Case 'hide_navigator' ($hide-and$hidden.visible_titles-eq0) ([ordered]@{hide_result=$hide;uia=$hidden})

    $show3=[bool]$excel.Run($prefix+'NxHostShowNavigator')
    $visible3=Wait-NavigatorState ([int64]$excel.Hwnd) $true
    Add-Case 'show_after_hide' ($show3-and$visible3.visible_titles-eq1) ([ordered]@{show_result=$show3;uia=$visible3})
}catch{$failure=$_.Exception.Message}finally{
    try{if($null-ne$product){try{[void]$excel.Run($product.Name+'!NxFocusDisable')}catch{};try{[void]$excel.Run($product.Name+'!NxHostHideNavigator')}catch{};try{[void]$excel.Run($product.Name+'!NxHostBridgeReset')}catch{}}}catch{}
    try{if($null-ne$hostAddin-and[bool]$hostAddin.Connect){$hostAddin.Connect=$false}}catch{}
    try{if($null-ne$book){$book.Close($false)}}catch{}
    try{if($null-ne$product){$product.Close($false)}}catch{}
    Release-ComObject $book;Release-ComObject $product;Release-ComObject $books;Release-ComObject $hostAddin;Release-ComObject $comAddins
    try{if($null-ne$excel){$excel.Quit()}}catch{}
    Release-ComObject $excel
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null-ne$binding-and$null-ne$binding.process){$result=Stop-ExactProcessAfterGrace $binding.process 'NxHost Navigator Excel' 15000 -Detailed;$excelCleanup=[ordered]@{exit_mode=$result.exit_mode;failure=$result.failure;owned_pid=$binding.pid};try{$binding.process.Dispose()}catch{}}
    $cleanupRequired=$registered
    if($registrationAttempted-and(Test-Path -LiteralPath $statePath -PathType Leaf)){
        try{
            $recoveryState=[IO.File]::ReadAllText($statePath)|ConvertFrom-Json
            if([string]::IsNullOrWhiteSpace($registryBefore)){$registryBefore=[string]$recoveryState.snapshot_sha256}
            if([string]$recoveryState.status -eq 'ACTIVE'){$cleanupRequired=$true}
            elseif([string]$recoveryState.status -eq 'RESTORED'){$registryAfter=[string]$recoveryState.restored_snapshot_sha256}
        }catch{if([string]::IsNullOrWhiteSpace($cleanupFailure)){$cleanupFailure='NxHost recovery-state inspection failed: '+$_.Exception.Message}}
    }
    if($cleanupRequired){
        try{$unregistrationOutput=@(& $registerScript -Action Unregister -X86Dll $x86Dll -X64Dll $x64Dll);$unregistrationOutput|Write-Output;$state=[IO.File]::ReadAllText($statePath)|ConvertFrom-Json;$registryAfter=[string]$state.restored_snapshot_sha256}catch{$cleanupFailure=$_.Exception.Message}
    }
    $env:LHEXCEL_PROFILE_ROOT=$originalProfile
}

$remainingExcel=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
$cleanupPassed=($remainingExcel.Count-eq0-and$excelCleanup.exit_mode-eq'NATURAL'-and[string]::IsNullOrWhiteSpace([string]$excelCleanup.failure)-and[string]::IsNullOrWhiteSpace($cleanupFailure)-and$registryBefore-match'^[0-9a-f]{64}$'-and$registryAfter-ceq$registryBefore)
Add-Case 'cleanup_and_registry_restore' $cleanupPassed @{excel=$excelCleanup;remaining_excel=@($remainingExcel|ForEach-Object{[int]$_.Id});registry_snapshot_before=$registryBefore;registry_snapshot_after=$registryAfter;failure=$cleanupFailure}
$passed=@($script:Cases|Where-Object{$_.status-eq'PASS'}).Count
$status=if($null-eq$failure-and$script:Cases.Count-eq9-and$passed-eq9){'PASS'}else{'DIAGNOSTIC'}
$receipt=[ordered]@{
    schema_version=1;suite='NxHostNavigator';status=$status;run_id=$RunId
    artifact_sha256=$artifactSha;source_tree_sha256=$SourceDigest;source_snapshot_sha256=$SnapshotDigest
    nxhost32_sha256=Get-Sha256 $x86Dll;nxhost64_sha256=Get-Sha256 $x64Dll
    cases=@($script:Cases.ToArray());cleanup=[ordered]@{excel=$excelCleanup;registry_snapshot_before=$registryBefore;registry_snapshot_after=$registryAfter;remaining_excel=@($remainingExcel|ForEach-Object{[int]$_.Id})}
    failure=$failure;cleanup_failure=$cleanupFailure;completed_utc=[DateTime]::UtcNow.ToString('o')
}
Write-Utf8Json (Join-Path $EvidenceRoot 'NxHostNavigator.json') $receipt
foreach($process in $remainingExcel){try{$process.Dispose()}catch{}}
if($status-ne'PASS'){Write-Output("DIAGNOSTIC|NxHostNavigator|{0}/9"-f$passed);exit 1}
Write-Output 'PASS|NxHostNavigator|9/9'
exit 0
