param(
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [Parameter(Mandatory=$true)][string]$ArtifactPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
$ArtifactPath = [IO.Path]::GetFullPath($ArtifactPath)
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing r52 distribution evidence root is rejected' }
if (-not (Test-Path -LiteralPath $ArtifactPath -PathType Leaf)) { throw 'r52 distribution artifact missing' }
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'r52 distribution smoke requires zero pre-existing Excel processes' }
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

$excel=$null;$binding=$null;$books=$null;$product=$null;$failure='';$cleanup=$null
$version='';$profile='';$buildId='';$policyCanOpen=$false
try {
    $baseline=Get-ExcelProcessBaseline
    $excel=New-Object -ComObject Excel.Application
    $binding=Get-ExactExcelProcessOwnership $excel $baseline 'r52 distribution smoke Excel'
    if(-not[bool]$binding.owned){throw('Excel ownership rejected: '+[string]$binding.failure)}
    $excel.Visible=$false;$excel.DisplayAlerts=$false
    $books=$excel.Workbooks;$product=$books.Open($ArtifactPath,$false,$true)
    $macro="'"+$product.Name.Replace("'","''")+"'!"
    $version=[string]$excel.Run($macro+'NxProductVersionText')
    $profile=[string]$excel.Run($macro+'NxProductBuildProfileText')
    $buildId=[string]$excel.Run($macro+'NxProductBuildIdText')
    $policyCanOpen=[bool]$excel.Run($macro+'LHEDistributionPolicyCanOpen')
} catch { $failure=$_.Exception.Message } finally {
    if($null-ne$product){try{$product.Close($false)}catch{}}
    Release-ComObject $product;Release-ComObject $books
    if($null-ne$excel){try{$excel.Quit()}catch{};Release-ComObject $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null-ne$binding-and$null-ne$binding.process){$cleanup=Stop-ExactProcessAfterGrace $binding.process 'r52 distribution smoke Excel' 15000 -Detailed;try{$binding.process.Dispose()}catch{}}
}

$wait=[Diagnostics.Stopwatch]::StartNew()
do {
    $remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
    if($remaining.Count-eq0){break}
    foreach($process in $remaining){try{$process.Dispose()}catch{}}
    Start-Sleep -Milliseconds 100
} while($wait.ElapsedMilliseconds-lt2000)
$passed=([string]::IsNullOrEmpty($failure)-and$version.EndsWith('v0.01_r52',[StringComparison]::Ordinal)-and$profile.Contains('enhanced-dll')-and$buildId.StartsWith('r52',[StringComparison]::Ordinal)-and$policyCanOpen-and$null-ne$cleanup-and[string]$cleanup.exit_mode-ceq'NATURAL'-and$remaining.Count-eq0)
$receipt=[pscustomobject][ordered]@{
    schema_version=1;suite='R52DistributionSmoke';status=if($passed){'PASS'}else{'DIAGNOSTIC'}
    artifact_sha256=Get-Sha256 $ArtifactPath;version=$version;profile=$profile;build_id=$buildId
    policy_can_open=$policyCanOpen;cleanup=$cleanup;remaining_excel=@($remaining|ForEach-Object{[int]$_.Id});failure=$failure
}
$path=Join-Path $EvidenceRoot 'R52DistributionSmoke.json';Write-Json $path $receipt
foreach($process in $remaining){try{$process.Dispose()}catch{}}
if($passed){Write-Output 'PASS|R52DistributionSmoke|5/5';Write-Output $path;exit 0}
Write-Output 'DIAGNOSTIC|R52DistributionSmoke';Write-Output $path;exit 24
