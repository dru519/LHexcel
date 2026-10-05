param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RunId = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }

$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$ProductXlam = [IO.Path]::GetFullPath($ProductXlam)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if (-not (Test-Path -LiteralPath $ProductXlam -PathType Leaf)) { throw 'Product.xlam not found' }
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing r57 usability evidence root is rejected' }
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) { throw 'r57 usability suite requires zero pre-existing Excel processes' }
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

function Convert-CodePoints([int[]]$Points) { return -join ($Points | ForEach-Object { [char]$_ }) }
function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}
function Write-Json([string]$Path, [object]$Value) {
    $encoding = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, (($Value | ConvertTo-Json -Depth 16) + "`n"), $encoding)
}
function Wait-ExcelProcessSnapshot([int]$TimeoutMs = 10000) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $processes = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
        if ($processes.Count -eq 0) { return @() }
        foreach ($candidate in $processes) { try { $candidate.Dispose() } catch {} }
        if ($watch.ElapsedMilliseconds -lt $TimeoutMs) { Start-Sleep -Milliseconds 100 }
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    return @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
}

$newSheetMode = Convert-CodePoints @(0xC0C8,0x0020,0xC2DC,0xD2B8,0x0020,0xB9CC,0xB4E4,0xAE30)
$sourceMode = Convert-CodePoints @(0xC6D0,0xBCF8,0x0020,0xBCC0,0xACBD)
$numberPreset = Convert-CodePoints @(0xC22B,0xC790,0x0020,0xBCC0,0xD658)
$sheetIndexName = Convert-CodePoints @(0xB0B4,0xC5D1,0xC140,0x005F,0xC2DC,0xD2B8,0xBAA9,0xB85D)
$numberHeader = Convert-CodePoints @(0xBC88,0xD638)
$sheetNameHeader = Convert-CodePoints @(0xC2DC,0xD2B8,0xBA85)
$visibilityHeader = Convert-CodePoints @(0xD45C,0xC2DC,0x0020,0xC0C1,0xD0DC)
$visibleLabel = Convert-CodePoints @(0xD45C,0xC2DC)
$hiddenLabel = Convert-CodePoints @(0xC228,0xAE40)
$veryHiddenLabel = Convert-CodePoints @(0xC644,0xC804,0x0020,0xC228,0xAE40)
$countsText = Convert-CodePoints @(0xB300,0xBD84,0xB958,0x0020,0x0031,0x0032,0xAC1C,0x0020,0x00B7,0x0020,0xAE30,0xB2A5,0x0020,0x0034,0x0036,0xAC1C)
$commandsText = Convert-CodePoints @(0xBA54,0xB274,0x00B7,0xBA85,0xB839,0x0020,0x0031,0x0034,0x0037,0xAC1C)
$productTitle = (Convert-CodePoints @(0xB0B4,0xC5D1,0xC140)) + ' v0.01_r57'

$cases = New-Object 'Collections.Generic.List[object]'
function Add-Case([string]$Name, [bool]$Passed, [object]$Details) {
    $script:cases.Add([pscustomobject][ordered]@{name=$Name;status=if($Passed){'PASS'}else{'FAIL'};details=$Details})
}

$excel=$null;$binding=$null;$books=$null;$product=$null;$book=$null;$worksheets=$null;$sourceSheet=$null
$newSource=$null;$newResultSheet=$null;$sourceTarget=$null;$hiddenSheet=$null;$veryHiddenSheet=$null;$indexSheet=$null
$failure=$null;$cleanup=[ordered]@{exit_mode='NOT_STARTED';failure=$null;owned_pid=0}
try {
    $baseline = Get-ExcelProcessBaseline
    $excel = New-Object -ComObject Excel.Application
    $binding = Get-ExactExcelProcessOwnership $excel $baseline 'r57 usability Excel'
    if (-not [bool]$binding.owned) { throw ('Excel ownership rejected: ' + [string]$binding.failure) }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $books = $excel.Workbooks
    $product = $books.Open($ProductXlam, $false, $true)
    $macro = "'" + ([string]$product.Name).Replace("'", "''") + "'!"
    $book = $books.Add()
    $worksheets = $book.Worksheets
    $sourceSheet = $worksheets.Item(1)
    [void]$book.Activate()
    [void]$sourceSheet.Activate()

    $version = [string]$excel.Run($macro + 'NxProductVersionText')
    $counts = [string]$excel.Run($macro + 'NxProductCountsText')
    $commands = [string]$excel.Run($macro + 'NxProductCommandCountText')
    Add-Case 'version_and_counts' ($version -ceq $productTitle -and $counts -ceq $countsText -and $commands -ceq $commandsText) `
        ([ordered]@{version=$version;counts=$counts;commands=$commands})

    $newSource = $sourceSheet.Range('A1:B1')
    $newSource.NumberFormat = '@'
    $newSource.Cells.Item(1,1).Value2 = '1,234'
    $newSource.Cells.Item(1,2).Value2 = '2,500'
    [void]$newSource.Select()
    $beforeNewSheetCount = [int]$worksheets.Count
    [void]$excel.Run($macro + 'NxDataRunNormalizationForNativeTest', $newSource, $numberPreset, 'yyyy-mm-dd', $newSheetMode)
    $afterNewSheetCount = [int]$worksheets.Count
    $newResultSheet = $worksheets.Item($afterNewSheetCount)
    $newSheetPassed = ($afterNewSheetCount -eq ($beforeNewSheetCount + 1) -and
        [string]$newSource.Cells.Item(1,1).Value2 -ceq '1,234' -and
        [double]$newResultSheet.Cells.Item(1,1).Value2 -eq 1234 -and
        [double]$newResultSheet.Cells.Item(1,2).Value2 -eq 2500)
    Add-Case 'normalization_new_sheet_preserves_source' $newSheetPassed `
        ([ordered]@{before_sheets=$beforeNewSheetCount;after_sheets=$afterNewSheetCount;source=[string]$newSource.Cells.Item(1,1).Value2;result=[double]$newResultSheet.Cells.Item(1,1).Value2})

    [void]$sourceSheet.Activate()
    $sourceTarget = $sourceSheet.Range('A2:B2')
    $sourceTarget.NumberFormat = '@'
    $sourceTarget.Cells.Item(1,1).Value2 = '3,000'
    $sourceTarget.Cells.Item(1,2).Value2 = '4,500'
    [void]$sourceTarget.Select()
    $beforeSourceSheetCount = [int]$worksheets.Count
    [void]$excel.Run($macro + 'NxDataRunNormalizationForNativeTest', $sourceTarget, $numberPreset, 'yyyy-mm-dd', $sourceMode)
    $sourceModePassed = ([int]$worksheets.Count -eq $beforeSourceSheetCount -and
        [double]$sourceTarget.Cells.Item(1,1).Value2 -eq 3000 -and
        [double]$sourceTarget.Cells.Item(1,2).Value2 -eq 4500)
    Add-Case 'normalization_source_mode_is_typed' $sourceModePassed `
        ([ordered]@{sheet_count=[int]$worksheets.Count;first=[double]$sourceTarget.Cells.Item(1,1).Value2;second=[double]$sourceTarget.Cells.Item(1,2).Value2})

    $hiddenSheet = $worksheets.Add()
    $hiddenSheet.Name = 'HiddenCase'
    $hiddenSheet.Visible = 0
    $veryHiddenSheet = $worksheets.Add()
    $veryHiddenSheet.Name = 'VeryHiddenCase'
    $veryHiddenSheet.Visible = 2
    [void]$sourceSheet.Activate()
    [void]$book.Activate()
    [void]$excel.Run($macro + 'NxCmdInfo', 'RB_SHOW_WORKSHEET_LIST')
    $indexSheet = $worksheets.Item($worksheets.Count)
    $headersPassed = ([string]$indexSheet.Name -ceq $sheetIndexName -and
        [string]$indexSheet.Cells.Item(1,1).Value2 -ceq $numberHeader -and
        [string]$indexSheet.Cells.Item(1,2).Value2 -ceq $sheetNameHeader -and
        [string]$indexSheet.Cells.Item(1,3).Value2 -ceq $visibilityHeader)
    Add-Case 'sheet_index_korean_headers' $headersPassed `
        ([ordered]@{sheet=[string]$indexSheet.Name;headers=@([string]$indexSheet.Cells.Item(1,1).Value2,[string]$indexSheet.Cells.Item(1,2).Value2,[string]$indexSheet.Cells.Item(1,3).Value2)})

    $listedNames = New-Object 'Collections.Generic.List[string]'
    $listedStates = @{}
    for ($row = 2; $row -le [int]$indexSheet.UsedRange.Rows.Count; $row++) {
        $listedName = [string]$indexSheet.Cells.Item($row,2).Value2
        $listedState = [string]$indexSheet.Cells.Item($row,3).Value2
        $listedNames.Add($listedName)
        $listedStates[$listedName] = $listedState
    }
    $visibilityPassed = ($listedStates['HiddenCase'] -ceq $hiddenLabel -and
        $listedStates['VeryHiddenCase'] -ceq $veryHiddenLabel -and
        @($listedStates.Values | Where-Object { $_ -ceq $visibleLabel }).Count -ge 1)
    Add-Case 'sheet_index_visibility_labels' $visibilityPassed ([ordered]@{states=$listedStates})
    Add-Case 'sheet_index_excludes_itself' (-not $listedNames.Contains($sheetIndexName)) ([ordered]@{names=$listedNames.ToArray()})
} catch {
    $failure = $_.Exception.Message
} finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $product) { try { $product.Close($false) } catch {} }
    foreach ($item in @($indexSheet,$veryHiddenSheet,$hiddenSheet,$sourceTarget,$newResultSheet,$newSource,$sourceSheet,$worksheets,$book,$product,$books)) { Release-ComObject $item }
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if ($null -ne $binding -and $null -ne $binding.process) {
        $stop = Stop-ExactProcessAfterGrace $binding.process 'r57 usability Excel' 15000 -Detailed
        $cleanup = [ordered]@{exit_mode=[string]$stop.exit_mode;failure=$stop.failure;owned_pid=[int]$binding.pid}
        try { $binding.process.Dispose() } catch {}
    }
}

$remaining = @(Wait-ExcelProcessSnapshot 10000)
$cleanupPassed = ($remaining.Count -eq 0 -and [string]$cleanup.exit_mode -ceq 'NATURAL' -and [string]::IsNullOrWhiteSpace([string]$cleanup.failure))
Add-Case 'cleanup_zero_excel' $cleanupPassed ([ordered]@{cleanup=$cleanup;remaining=@($remaining | ForEach-Object { [int]$_.Id })})
$passed = @($cases | Where-Object status -eq 'PASS').Count
$status = if ([string]::IsNullOrWhiteSpace([string]$failure) -and $cases.Count -eq 7 -and $passed -eq 7) { 'PASS' } else { 'DIAGNOSTIC' }
$receipt = [ordered]@{
    schema_version=1;suite='R57Usability';status=$status;run_id=$RunId
    artifact=$ProductXlam;artifact_sha256=(Get-FileHash -LiteralPath $ProductXlam -Algorithm SHA256).Hash.ToLowerInvariant()
    cases=$cases.ToArray();cleanup=$cleanup;failure=$failure;completed_utc=[DateTime]::UtcNow.ToString('o')
}
$receiptPath = Join-Path $EvidenceRoot 'R57Usability.json'
Write-Json $receiptPath $receipt
foreach ($process in $remaining) { try { $process.Dispose() } catch {} }
if ($status -eq 'PASS') { Write-Output ('PASS|R57Usability|' + $passed + '/' + $cases.Count); Write-Output $receiptPath; exit 0 }
Write-Output ('DIAGNOSTIC|R57Usability|' + $passed + '/' + $cases.Count); Write-Output $receiptPath; exit 24
