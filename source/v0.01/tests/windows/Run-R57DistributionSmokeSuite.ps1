param(
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$ArtifactPath,
    [ValidatePattern('^v0\.01_r\d+$')][string]$ExpectedVersion = 'v0.01_r62',
    [switch]$VerifyNoDllFocus,
    [switch]$VerifyProductAbout,
    [ValidateRange(0,300)][int]$InteractiveProbeSeconds=0,
    [ValidateSet('','internal-xlam','enhanced-dll')][string]$ExpectedProfile='',
    [ValidateSet('active','expired','invalid')][string]$ExpectedPolicy='active'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
$ArtifactPath = [IO.Path]::GetFullPath($ArtifactPath)
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing distribution evidence root is rejected' }
if (-not (Test-Path -LiteralPath $ArtifactPath -PathType Leaf)) { throw 'Distribution artifact missing' }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Distribution smoke requires zero pre-existing Excel processes' }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}
function Get-Sha256([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Write-Json([string]$Path,[object]$Value) {
    $encoding = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 12), $encoding)
}

$excel=$null;$binding=$null;$books=$null;$hostBook=$null;$product=$null;$failure='';$cleanup=$null
$version='';$buildId='';$policyCanOpen=$false;$expiryText=''
$focus=[ordered]@{status='NOT_RUN';backend='';diagnostic='';row_painted=$false;column_painted=$false;shifted=$false;residual_rules=-1}
$beforeHash=Get-Sha256 $ArtifactPath
$previousProfile=$env:LHEXCEL_PROFILE_ROOT
$env:LHEXCEL_PROFILE_ROOT=Join-Path $EvidenceRoot 'profile'
$expectedBuild=$ExpectedVersion.Split('_')[-1]
try {
    $baseline=Get-ExcelProcessBaseline
    $startupFixture=Join-Path $EvidenceRoot 'startup-fixture.xlsx'
    Copy-Item -LiteralPath (Join-Path $SourceRoot 'tests/fixtures/r60-document-navigator/simple.xlsx') -Destination $startupFixture
    $interactive=Start-ExactInteractiveExcel $baseline 'distribution smoke Excel' -TimeoutMs 120000 -StartupWorkbook $startupFixture -ShowWindow
    $excel=$interactive.excel;$binding=$interactive.binding;$interactive=$null
    if(-not[bool]$binding.owned){throw('Excel ownership rejected: '+[string]$binding.failure)}
    $excel.Visible=$true;$excel.DisplayAlerts=$false
    $startupBook=$excel.Workbooks.Item('startup-fixture.xlsx')
    try {
        if([IO.Path]::GetFullPath($startupBook.FullName) -ne [IO.Path]::GetFullPath($startupFixture)){throw 'Startup fixture identity mismatch'}
        $startupBook.Close($false)
    } finally {Release-ComObject $startupBook}
    $books=$excel.Workbooks;$hostBook=$books.Add();$product=$books.Open($ArtifactPath,$false,$true)
    $macro="'"+$product.Name.Replace("'","''")+"'!"
    if($InteractiveProbeSeconds -gt 0){
        $hostBook.Activate()
        $observations=New-Object 'Collections.Generic.List[object]'
        $probeDeadline=[DateTime]::UtcNow.AddSeconds($InteractiveProbeSeconds)
        while([DateTime]::UtcNow -lt $probeDeadline){
            $observations.Add([pscustomobject]@{
                utc=[DateTime]::UtcNow.ToString('o');ready=[bool]$excel.Ready;events=[bool]$excel.EnableEvents
                selection=[string]$excel.Selection.Address();screen_updating=[bool]$excel.ScreenUpdating
                focus_backend=[string]$excel.Run($macro+'NxFocusControllerBackend')
                focus_diagnostic=[string]$excel.Run($macro+'NxFocusControllerBackendDiagnostic')
                picture_availability=[string]$excel.Run($macro+'NxRouteAvailabilitySummary','feature:NX-HANGUL-PICTURE-SEND')
            })
            Write-Json (Join-Path $EvidenceRoot 'interactive-observations.json') $observations.ToArray()
            Start-Sleep -Milliseconds 2000
        }
    }
    $version=[string]$excel.Run($macro+'NxProductVersionText')
    $buildId=[string]$excel.Run($macro+'NxProductBuildIdText')
    if($VerifyProductAbout){
        $aboutMenu=[xml]([string]$excel.Run($macro+'NxRibbonManagementMenuXml'))
        if($aboutMenu.DocumentElement.LastChild.GetAttribute('tag') -ne 'nx1|entry|NX-MGMT-FILE-INFO'){throw 'About must be last menu item'}
        if(-not [bool]$excel.Run($macro+'NxDistributionEntryAllowed','entry','NX-MGMT-FILE-INFO')){throw 'About route unavailable'}
        $expiryText=[string]$excel.Run($macro+'NxDistributionValidUntilText')
        if($ExpectedPolicy -eq 'active' -and $expiryText -ne '2027-12-31'){throw ('About expiry mismatch: '+$expiryText)}
    }
    if($ExpectedProfile){
        $standalone=[bool]$excel.Run($macro+'NxDistributionIsStandaloneXlam')
        $expectStandalone=($ExpectedProfile -eq 'internal-xlam' -and $ExpectedPolicy -ne 'invalid')
        if($standalone -ne $expectStandalone){throw 'Standalone installation profile mismatch'}
        $menu=[string]$excel.Run($macro+'NxRibbonManagementMenuXml')
        $expectInstall=($expectStandalone -and $ExpectedPolicy -eq 'active')
        if($menu.Contains('tag="nx1|entry|NX-MGMT-INSTALL"') -ne $expectInstall -or $menu.Contains('tag="nx1|entry|NX-MGMT-REMOVE"') -ne $expectStandalone){throw 'Installation menu profile mismatch'}
        foreach($entry in @('NX-MGMT-INSTALL','NX-MGMT-REMOVE')){
            $expectAllowed=if($entry -eq 'NX-MGMT-INSTALL'){$expectInstall}else{$expectStandalone}
            if([bool]$excel.Run($macro+'NxDistributionEntryAllowed','entry',$entry) -ne $expectAllowed){throw 'Installation route guard mismatch'}
        }
    }
    if($ExpectedPolicy -eq 'active'){$policyCanOpen=[bool]$excel.Run($macro+'LHEDistributionPolicyCanOpen')}
    else{$policyCanOpen=[bool]$excel.Run($macro+'NxDistributionCanExecute')}
    if([int]($ExpectedVersion -replace '^v0\.01_r','') -ge 76){
        $actualPolicy=[string]$excel.Run($macro+'NxDistributionState')
        if($actualPolicy -ne $ExpectedPolicy){throw ('Distribution state mismatch: '+$actualPolicy)}
        if($policyCanOpen -ne ($ExpectedPolicy -eq 'active')){throw 'Distribution execution gate mismatch'}
        if(-not [bool]$excel.Run($macro+'NxDistributionEntryAllowed','entry','NX-ENTRY-MANAGEMENT')){throw 'Management icon unavailable'}
        if($ExpectedPolicy -ne 'active'){
            if([bool]$excel.Run($macro+'NxDistributionEntryAllowed','entry','NX-ENTRY-ALL')){throw 'All functions still enabled'}
            $menu=[string]$excel.Run($macro+'NxRibbonManagementMenuXml')
            if($menu.Contains('NX-MGMT-SHORTCUTS') -or -not $menu.Contains('NX-MGMT-DISTRIBUTION-STATUS')){throw 'Expired management menu mismatch'}
            if(-not ([string]$excel.Run($macro+'NxDistributionMessage')).Contains('AI')){throw 'Modification warning missing'}
        }
    }
    if($VerifyNoDllFocus){
    $compatible=[bool]$excel.Run($macro+'NxHostBridgeCompatible')
    if($compatible){throw 'Compatible DLL active; no-DLL focus verification cannot be claimed'}
    $sheet=$hostBook.Worksheets.Item(1);$range=$null;$sample=$null
    try{
        $hostBook.Activate();$sheet.Activate()
        $range=$sheet.Range('B2:C3');$range.Value2=1;$range.Select()
        # A non-formula rule must survive focus activation and cleanup.
        $userRange=$sheet.Range('J1:J3')
        $userRule=$userRange.FormatConditions.AddColorScale(2)
        Release-ComObject $userRule
        $userRule=$userRange.FormatConditions.AddDatabar()
        Release-ComObject $userRule
        $userRule=$userRange.FormatConditions.AddIconSetCondition()
        Release-ComObject $userRule;Release-ComObject $userRange
        $userRule=$null;$userRange=$null
        # Exercise the public ribbon route, not the backend-only entry point.
        $excel.EnableEvents=$false
        [void]$excel.Run($macro+'NxRibbonExecuteTag','nx1|feature|NX-DATA-FOCUS-CELL')
        $focus.events_recovered=[bool]$excel.EnableEvents
        if(-not $focus.events_recovered){throw 'Focus startup did not recover selection events'}
        $focus.backend=[string]$excel.Run($macro+'NxFocusControllerBackend')
        $focus.diagnostic=[string]$excel.Run($macro+'NxFocusControllerBackendDiagnostic')
        $sample=$sheet.Range('A2:A3');$focus.row_painted=($sample.DisplayFormat.Interior.Color -ne 16777215);Release-ComObject $sample;$sample=$null
        $sample=$sheet.Range('B1:C1');$focus.column_painted=($sample.DisplayFormat.Interior.Color -ne 16777215);Release-ComObject $sample;$sample=$null
        Release-ComObject $range;$range=$sheet.Range('F6:G7');$range.Select()
        # Selection must update through Excel events without manually calling
        # the controller, otherwise a broken event route would pass this test.
        Start-Sleep -Milliseconds 150
        $sample=$sheet.Range('A6:A7');$focus.shifted=($sample.DisplayFormat.Interior.Color -ne 16777215);Release-ComObject $sample;$sample=$null
        [void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)
        $focus.residual_rules=[int]$excel.Run($macro+'NxFocusManagedRuleCount',$sheet)
        $userRange=$sheet.Range('J1:J3')
        if([int]$userRange.FormatConditions.Count -ne 3){throw 'User conditional formats were not preserved'}
        Release-ComObject $userRange;$userRange=$null
        if($focus.backend -ne 'cf' -or -not $focus.row_painted -or -not $focus.column_painted -or -not $focus.shifted -or $focus.residual_rules -ne 0){throw 'Protected no-DLL focus failed'}
        $focus.status='PASS'
    }finally{
        try{[void]$excel.Run($macro+'NxFocusControllerSetEnabled',$false)}catch{}
        Release-ComObject $sample;Release-ComObject $range;Release-ComObject $sheet
        $sample=$null;$range=$null;$sheet=$null
    }
    }
} catch { $failure=$_.Exception.Message } finally {
    if($null-ne$product){try{$product.Close($false)}catch{}}
    if($null-ne$hostBook){try{$hostBook.Close($false)}catch{}}
    Release-ComObject $product;Release-ComObject $hostBook;Release-ComObject $books
    if($null-ne$excel){try{$excel.Quit()}catch{};Release-ComObject $excel}
    $product=$null;$hostBook=$null;$books=$null;$excel=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null-ne$binding-and$null-ne$binding.process){$cleanup=Stop-ExactProcessAfterGrace $binding.process 'distribution smoke Excel' 15000 -Detailed;try{$binding.process.Dispose()}catch{}}
    $env:LHEXCEL_PROFILE_ROOT=$previousProfile
}

$wait=[Diagnostics.Stopwatch]::StartNew()
do {
    $remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
    if($remaining.Count-eq0){break}
    foreach($process in $remaining){try{$process.Dispose()}catch{}}
    Start-Sleep -Milliseconds 100
} while($wait.ElapsedMilliseconds-lt2000)
$afterHash=Get-Sha256 $ArtifactPath
$passed=([string]::IsNullOrEmpty($failure)-and$version.EndsWith($ExpectedVersion,[StringComparison]::Ordinal)-and$buildId.StartsWith($expectedBuild,[StringComparison]::Ordinal)-and($policyCanOpen -eq ($ExpectedPolicy -eq 'active'))-and$null-ne$cleanup-and[string]$cleanup.exit_mode-ceq'NATURAL'-and$remaining.Count-eq0-and$beforeHash-ceq$afterHash)
$receipt=[pscustomobject][ordered]@{
    schema_version=1;suite='R57DistributionSmoke';status=if($passed){'PASS'}else{'DIAGNOSTIC'}
    artifact_sha256=$afterHash;artifact_sha256_before=$beforeHash;expected_version=$ExpectedVersion;version=$version;build_id=$buildId
    policy_can_open=$policyCanOpen;expected_policy=$ExpectedPolicy;focus_without_dll=$focus;cleanup=$cleanup;remaining_excel=@($remaining|ForEach-Object{[int]$_.Id});failure=$failure
    about_verified=[bool]$VerifyProductAbout;about_expiry_text=$expiryText
}
$path=Join-Path $EvidenceRoot 'R57DistributionSmoke.json';Write-Json $path $receipt
foreach($process in $remaining){try{$process.Dispose()}catch{}}
if($passed){Write-Output 'PASS|R57DistributionSmoke|5/5';Write-Output $path;exit 0}
Write-Output 'DIAGNOSTIC|R57DistributionSmoke';Write-Output $path;exit 24
