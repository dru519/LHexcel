param([Parameter(Mandatory=$true)][string]$ProductXlam,[Parameter(Mandatory=$true)][string]$EvidenceRoot)
$ErrorActionPreference='Stop'
$SourceRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close user Excel first'}
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Evidence exists'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$hash=(Get-FileHash -LiteralPath $ProductXlam).Hash
$copy=Join-Path $EvidenceRoot 'Product.xlam';Copy-Item -LiteralPath $ProductXlam -Destination $copy
$oldProfile=$env:LHEXCEL_PROFILE_ROOT;$env:LHEXCEL_PROFILE_ROOT=Join-Path $EvidenceRoot 'profile'
[void](New-Item -ItemType Directory -Path $env:LHEXCEL_PROFILE_ROOT -Force)
$fileRoute='NOT_RUN';$edge='NOT_RUN';$cancel='NOT_RUN'
function Release-Com($o){if($null-ne$o-and[Runtime.InteropServices.Marshal]::IsComObject($o)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($o)}catch{}}}
$excel=$null;$binding=$null;$addin=$null;$book=$null;$failure='';$cases=New-Object 'Collections.Generic.List[object]'
try{
    $opened=Start-ExactInteractiveExcel (Get-ExcelProcessBaseline) 'data compare test'
    $excel=$opened.excel;$binding=$opened.binding;$opened=$null
    $excel.Visible=$true;$excel.DisplayAlerts=$false
    $books=$excel.Workbooks;$book=$books.Add();$addin=$books.Open($copy,$false,$true)
    $macro="'"+$addin.Name+"'!"
    $project=$addin.VBProject;$components=$project.VBComponents
    $module=$components.Add(2);$module.Name='CNxDataCompareCancelProbe'
    $code=[IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/product/CNxDataCompareCancelProbe.cls'),[Text.Encoding]::UTF8)
    $module.CodeModule.AddFromString($code.Substring($code.IndexOf('Option Explicit')))
    Release-Com $module
    $module=$components.Add(1);$module.Name='T_DataCompare'
    $code=[IO.File]::ReadAllText((Join-Path $SourceRoot 'tests/vba/product/T_DataCompare.bas'),[Text.Encoding]::UTF8)
    $module.CodeModule.AddFromString($code.Substring($code.IndexOf('Option Explicit')))
    Release-Com $module;Release-Com $components;Release-Com $project
    if($env:LHEXCEL_COMPARE_UI-eq'1'){
        $book.Activate()
        $excel.Run($macro+'NxTestDataCompareUiData')
        Write-Output 'UI_READY: inspect owned test workbook, then create ui-done.txt in evidence directory'
        $deadline=[DateTime]::UtcNow.AddMinutes(15)
        while(-not(Test-Path -LiteralPath (Join-Path $EvidenceRoot 'ui-done.txt'))){
            if([DateTime]::UtcNow-gt$deadline){throw 'UI review timed out'}
            Start-Sleep -Milliseconds 500
        }
    }
    $fileRoute=[string]$excel.Run($macro+'NxTestFileCompareRoute')
    Write-Output $fileRoute
    if(-not$fileRoute.StartsWith('PASS|')){throw $fileRoute}
    $edge=[string]$excel.Run($macro+'NxTestDataCompareEdges')
    Write-Output $edge
    if(-not$edge.StartsWith('PASS|')){throw $edge}
    $cancel=[string]$excel.Run($macro+'NxTestDataCompareCancel')
    Write-Output $cancel
    if(-not$cancel.StartsWith('PASS|')){throw $cancel}
    foreach($cells in @(90000,250000,500000,1000000)){
        foreach($pattern in @('same','sparse','all')){
            Write-Output ("START|"+$cells+"|"+$pattern)
            $before=[Diagnostics.Process]::GetProcessById($binding.process.Id).PrivateMemorySize64
            $result=[string]$excel.Run($macro+'NxTestDataCompare',$cells,$pattern)
            $after=[Diagnostics.Process]::GetProcessById($binding.process.Id).PrivateMemorySize64
            $cases.Add(@{result=$result;private_bytes_before=$before;private_bytes_after=$after})
            Write-Output $result
            $cases.ToArray()|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $EvidenceRoot 'progress.json') -Encoding UTF8
            if(-not$result.StartsWith('PASS|')){throw $result}
        }
    }
}catch{$failure=$_.Exception.Message}finally{
    if($addin){try{$addin.Close($false)}catch{}}
    if($book){try{$book.Close($false)}catch{}}
    Release-Com $addin;Release-Com $book;Release-Com $books
    if($excel){try{$excel.Quit()}catch{};Release-Com $excel}
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($binding){$cleanup=Stop-ExactProcessAfterGrace $binding.process 'data compare test' 15000 -Detailed;$binding.process.Dispose()}
    $env:LHEXCEL_PROFILE_ROOT=$oldProfile
}
$unchanged=($hash-eq(Get-FileHash -LiteralPath $ProductXlam).Hash)
$passed=($failure-eq''-and$cleanup.exit_mode-eq'NATURAL'-and$cases.Count-eq12-and$unchanged)
$receipt=@{status=$(if($passed){'PASS'}else{'FAIL'});failure=$failure;file_route=$fileRoute;edges=$edge;cancel=$cancel;cases=$cases.ToArray();cleanup=$cleanup;sha256=$hash;unchanged=$unchanged}
$receipt|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $EvidenceRoot 'data-compare.json') -Encoding UTF8
Write-Output ($receipt|ConvertTo-Json -Depth 5)
if(-not$passed){exit 1}
