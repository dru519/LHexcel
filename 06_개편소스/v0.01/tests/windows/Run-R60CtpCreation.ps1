param([Parameter(Mandatory=$true)][string]$EvidenceRoot)
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Close user Excel before CTP creation'}
. (Join-Path $PSScriptRoot 'Excel-ProcessLifecycle.ps1')
function Release-CtpCom($value){
    if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}
function ConvertTo-CtpFailureDetail($record,[string]$callerPhase){
    $exceptions=New-Object 'Collections.Generic.List[object]'
    $exception=$record.Exception;$depth=0
    while($null -ne $exception -and $depth -lt 8){
        $exceptions.Add([pscustomobject]@{depth=$depth;type=$exception.GetType().FullName;
            hresult=('0x{0:X8}' -f [int]$exception.HResult);message=$exception.Message})
        $exception=$exception.InnerException;$depth++
    }
    [pscustomobject]@{caller_phase=$callerPhase;exceptions=@($exceptions.ToArray())}
}
$excel=$null;$binding=$null;$books=$null;$book=$null;$addins=$null;$addin=$null;$automation=$null
$failure=$null;$cleanup=$null
$phase='start_excel';$failureDetail=$null
$cleanupFailures=New-Object 'Collections.Generic.List[string]'
$widths=New-Object 'Collections.Generic.List[object]'
try{
    $started=Start-ExactInteractiveExcel (Get-ExcelProcessBaseline) 'R60 CTP creation'
    $excel=$started.excel;$binding=$started.binding;$started=$null
    if(-not $binding.owned){throw 'Excel ownership rejected'}
    $phase='create_owned_workbook'
    $books=$excel.Workbooks
    $book=$books.Add()
    $book.Activate()
    $phase='discover_addin'
    $addins=$excel.COMAddIns
    $addin=$addins.Item('LH.MyExcel.DocNavCtpProbe.Connect')
    $phase='verify_connection'
    if(-not $addin.Connect){throw 'Probe not connected; no forced connection attempted'}
    $phase='get_automation_object'
    $automation=$addin.Object
    if($null -eq $automation){throw 'CTP automation object unavailable'}
    $phase='ensure_ctp'
    if(-not $automation.EnsureActivePane()){throw 'CTP creation failed'}
    $phase='width_readback'
    foreach($dip in @(216,196,32)){
        if(-not $automation.SetActiveWidthDip($dip)){throw 'Width request rejected'}
        Start-Sleep -Milliseconds 500
        $widths.Add([pscustomobject]@{requested_dip=$dip;office_points=[int]$automation.ActiveOfficeWidthPoints;office_dip=[int]$automation.ActiveOfficeWidthDip;client_px=[int]$automation.ActiveClientWidthPixels;dpi=[int]$automation.ActiveDpi;hwnd=[long]$automation.ActivePaneHwnd;dock=[string]$automation.ActiveDockPosition;clipped=[int]$automation.ActiveClippedControlCount})
    }
}catch{$failure=$_.Exception.Message;$failureDetail=ConvertTo-CtpFailureDetail $_ $phase}
finally{
    foreach($name in @('automation','addin','addins')){
        try{Release-CtpCom (Get-Variable -Name $name -ValueOnly)}catch{$cleanupFailures.Add('release_'+$name+': '+$_.Exception.Message)}
        Set-Variable -Name $name -Value $null
    }
    if($null -ne $book){try{$book.Close($false)}catch{$cleanupFailures.Add('workbook_close: '+$_.Exception.Message)}}
    foreach($name in @('book','books')){
        try{Release-CtpCom (Get-Variable -Name $name -ValueOnly)}catch{$cleanupFailures.Add('release_'+$name+': '+$_.Exception.Message)}
        Set-Variable -Name $name -Value $null
    }
    if($null -ne $excel){
        if($null -ne $binding -and $binding.owned){try{$excel.Quit()}catch{$cleanupFailures.Add('excel_quit: '+$_.Exception.Message)}}
        try{Release-CtpCom $excel}catch{$cleanupFailures.Add('release_excel: '+$_.Exception.Message)}
        $excel=$null
    }
    try{[GC]::Collect();[GC]::WaitForPendingFinalizers()}catch{$cleanupFailures.Add('gc: '+$_.Exception.Message)}
    if($null -ne $binding -and $binding.owned){
        try{$cleanup=Stop-ExactProcessAfterGrace $binding.process 'R60 CTP creation' 15000 -Detailed}catch{$cleanupFailures.Add('process_stop: '+$_.Exception.Message)}
        try{$binding.process.Dispose()}catch{$cleanupFailures.Add('process_dispose: '+$_.Exception.Message)}
    }
    Start-Sleep -Milliseconds 1000
    $receipt=[pscustomobject]@{failure=$failure;failure_detail=$failureDetail;cleanup_failures=@($cleanupFailures.ToArray());widths=@($widths.ToArray());cleanup=$cleanup;remaining_excel=@(Get-Process EXCEL -ErrorAction SilentlyContinue|ForEach-Object Id);evidence_class='native creation and Office width readback; not independent visual acceptance'}
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'ctp-creation.json'),($receipt|ConvertTo-Json -Depth 6))
}
if($null -ne $failure -or $cleanupFailures.Count -gt 0 -or $null -eq $cleanup -or $cleanup.exit_mode -ne 'NATURAL' -or $receipt.remaining_excel.Count){throw ('CTP creation failed: '+$failure)}
Write-Output 'PASS|CTP_CREATION|width readback only'
