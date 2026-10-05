param([Parameter(Mandatory=$true)][string]$InputFile,[Parameter(Mandatory=$true)][string]$EvidenceRoot)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../../build/Excel-ProcessLifecycle.ps1')
function Release-ComObject([object]$value) {
    if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Fresh evidence required'}
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Existing Excel preserved'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$report=[ordered]@{status='RUNNING';input_sha256=(Get-FileHash -LiteralPath $InputFile).Hash;product_loaded=$false;cases=@()}
$excel=$null;$session=$null;$books=$null;$book=$null;$sheet=$null;$cells=$null
try {
    $session=Start-ExactInteractiveExcel (Get-ExcelProcessBaseline) 'r101-compare-reopen'
    $excel=$session.excel
    $excel.Visible=$true;$excel.DisplayAlerts=$false
    $books=$excel.Workbooks
    for($iteration=1;$iteration -le 3;$iteration++) {
        [IO.File]::AppendAllText((Join-Path $EvidenceRoot 'stages.log'),"open $iteration`r`n")
        $book=$books.Open($InputFile,0,$true)
        $sheet=$book.Worksheets.Item((-join @([char]0xC140,[char]0xCC28,[char]0xC774)))
        $cells=$sheet.Range('A1:E4');$values=$cells.Value2
        if($values[2,2] -ne 'B2' -or $values[3,2] -ne 'D4' -or [string]$values[2,5] -ne '42' -or [string]$values[3,5] -ne '3'){throw 'Unexpected report values'}
        Release-ComObject $cells;$cells=$null;Release-ComObject $sheet;$sheet=$null
        $book.Close($false);Release-ComObject $book;$book=$null
        $report.cases+=@{iteration=$iteration;result='PASS'}
        Write-Output "PASS|reopen|$iteration"
    }
    $report.status='PASS'
} catch {
    $report.status='FAIL';$report.failure=$_.Exception.ToString();$report.location=$_.InvocationInfo.PositionMessage
} finally {
    Release-ComObject $cells;Release-ComObject $sheet
    if($null -ne $book){try{$book.Close($false)}catch{};Release-ComObject $book}
    Release-ComObject $books
    if($null -ne $excel){try{$excel.Quit()}catch{};Release-ComObject $excel}
    $sessionExcel=$null;$excel=$null;$books=$null
    if($null -ne $session){$session.excel=$null}
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null -ne $session){$report.cleanup=Stop-ExactProcessAfterGrace $session.binding.process 'r101-compare-reopen' 15000 -Detailed}
    if($report.cleanup.exit_mode -ne 'NATURAL'){$report.status='FAIL'}
    $report.input_sha256_after=(Get-FileHash -LiteralPath $InputFile).Hash
    if($report.input_sha256 -ne $report.input_sha256_after){$report.status='FAIL'}
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'reopen.json'),($report|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
}
if($report.status -ne 'PASS'){throw 'Reopen verification failed'}
