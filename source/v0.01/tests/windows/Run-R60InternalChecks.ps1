param([Parameter(Mandatory=$true)][string]$EvidenceRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Close user Excel before checks'}
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Existing evidence rejected'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
. (Join-Path $PSScriptRoot 'Excel-ProcessLifecycle.ps1')
function Release-TestCom($value){if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}}
$excel=$null;$binding=$null;$books=$null;$book=$null;$product=$null;$project=$null;$components=$null;$component=$null;$module=$null
$failure=$null;$cleanup=$null;$policy=$null;$preview=$null
$artifact=Join-Path $PSScriptRoot 'Internal.xlam'
$beforeHash=(Get-FileHash -LiteralPath $artifact -Algorithm SHA256).Hash
$priorProfile=$env:LHEXCEL_PROFILE_ROOT
$env:LHEXCEL_PROFILE_ROOT=Join-Path $EvidenceRoot 'runtime-profile'
try{
    $started=Start-ExactInteractiveExcel (Get-ExcelProcessBaseline) 'R60 internal checks'
    $excel=$started.excel;$binding=$started.binding;$started=$null
    if(-not $binding.owned){throw 'Ownership rejected'}
    $books=$excel.Workbooks;$book=$books.Add()
    $project=$book.VBProject;$components=$project.VBComponents
    foreach($name in @('policy.bas','R60PolicyProbe.bas')){
        $component=$components.Add(1);$module=$component.CodeModule
        $module.AddFromString([IO.File]::ReadAllText((Join-Path $PSScriptRoot $name),[Text.Encoding]::UTF8))
        Release-TestCom $module;Release-TestCom $component;$module=$null;$component=$null
    }
    $macro="'"+$book.Name+"'!"
    $policy=[string]$excel.Run($macro+'PolicyProbe',(Join-Path $EvidenceRoot 'policy-profile'))
    if(-not $policy.StartsWith('PASS|8|')){throw 'Policy checks failed'}
    $product=$books.Open($artifact,$false,$true)
    $book.Activate()
    $preview=[string]$excel.Run($macro+'PreviewProbe',$product.Name)
    if(-not $preview.StartsWith('PASS|12|')){throw 'Preview checks failed'}
}catch{$failure=$_.Exception.ToString()}
finally{
    Release-TestCom $module;Release-TestCom $component;Release-TestCom $components;Release-TestCom $project
    if($null -ne $product){$product.Close($false)}
    if($null -ne $book){$book.Close($false)}
    Release-TestCom $product;Release-TestCom $book;Release-TestCom $books
    if($null -ne $excel){if($null -ne $binding -and $binding.owned){$excel.Quit()};Release-TestCom $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null -ne $binding -and $binding.owned){$cleanup=Stop-ExactProcessAfterGrace $binding.process 'R60 internal checks' 15000 -Detailed;$binding.process.Dispose()}
    $env:LHEXCEL_PROFILE_ROOT=$priorProfile
}
Start-Sleep -Milliseconds 1000
$remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
$afterHash=(Get-FileHash -LiteralPath $artifact -Algorithm SHA256).Hash
$receipt=[pscustomobject]@{failure=$failure;policy=$policy;preview=$preview;artifact_unchanged=($beforeHash -eq $afterHash);artifact_sha256=$afterHash;cleanup=$cleanup;remaining_excel=$remaining}
[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'internal-checks.json'),($receipt|ConvertTo-Json -Depth 6))
if($null -ne $failure -or $null -eq $cleanup -or $cleanup.exit_mode -ne 'NATURAL' -or $remaining.Count -or $beforeHash -ne $afterHash){throw 'Internal checks failed; see receipt'}
Write-Output 'PASS|R60_INTERNAL_CHECKS|policy 8; preview 12'
