param([Parameter(Mandatory=$true)][string]$EvidenceRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Close user Excel before inventory'}
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Existing evidence rejected'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
. (Join-Path $PSScriptRoot 'Excel-ProcessLifecycle.ps1')
function Release-Com($value){
    if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}
$rows=New-Object 'Collections.Generic.List[object]'
foreach($mode in @('automation','interactive')){
    if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Excel appeared before next inventory mode'}
    $excel=$null;$binding=$null;$addins=$null;$cleanup=$null;$failure=$null
    $items=New-Object 'Collections.Generic.List[object]'
    try{
        $baseline=Get-ExcelProcessBaseline
        if($mode -eq 'interactive'){
            $started=Start-ExactInteractiveExcel $baseline 'R60 COM inventory'
            $excel=$started.excel;$binding=$started.binding;$started=$null
        }else{
            $excel=New-Object -ComObject Excel.Application
            $binding=Get-ExactExcelProcessOwnership $excel $baseline 'R60 COM inventory'
        }
        if(-not $binding.owned){throw 'Excel ownership rejected'}
        $addins=$excel.COMAddIns
        for($i=1;$i -le $addins.Count;$i++){
            $item=$null
            try{
                $item=$addins.Item($i)
                $items.Add([pscustomobject]@{progid=[string]$item.ProgId;connected=[bool]$item.Connect})
            }finally{Release-Com $item;$item=$null}
        }
    }catch{$failure=$_.Exception.Message}
    finally{
        Release-Com $addins;$addins=$null
        if($null -ne $excel){
            if($null -ne $binding -and $binding.owned){try{$excel.Quit()}catch{$failure=$_.Exception.Message}}
            Release-Com $excel;$excel=$null
        }
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
        if($null -ne $binding -and $binding.owned){
            $cleanup=Stop-ExactProcessAfterGrace $binding.process 'R60 COM inventory' 15000 -Detailed
            $binding.process.Dispose()
        }
    }
    $settled=[Diagnostics.Stopwatch]::StartNew()
    do{
        $remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id)
        if($remaining.Count -eq 0){break}
        Start-Sleep -Milliseconds 100
    }while($settled.ElapsedMilliseconds -lt 2000)
    $status=if($null -eq $failure -and $null -ne $cleanup -and $cleanup.exit_mode -eq 'NATURAL' -and $remaining.Count -eq 0){'PASS'}else{'FAIL'}
    $rows.Add([pscustomobject]@{mode=$mode;status=$status;addins=@($items.ToArray());failure=$failure;cleanup=$cleanup;remaining_excel=$remaining})
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'com-inventory.json'),(ConvertTo-Json -InputObject @($rows.ToArray()) -Depth 8),(New-Object Text.UTF8Encoding($false)))
    if($status -ne 'PASS'){throw ('Inventory failed: '+$mode+' '+$failure)}
}
Write-Output 'PASS|COM_INVENTORY|2 modes; observation only, not CTP acceptance'
