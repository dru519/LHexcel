param([Parameter(Mandatory=$true)][string]$Path)

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}
. (Join-Path $PSScriptRoot 'Excel-ProcessLifecycle.ps1')
function Assert-ValidXlamArtifact([string]$Artifact) {
    if (-not (Test-Path -LiteralPath $Artifact -PathType Leaf)) { throw 'Blank XLAM artifact missing' }
    $info = Get-Item -LiteralPath $Artifact -ErrorAction Stop
    if ($info.Length -lt 4) { throw 'Blank XLAM artifact empty' }
    $bytes = [IO.File]::ReadAllBytes($Artifact)
    if ($bytes[0] -ne 80 -or $bytes[1] -ne 75) { throw 'Blank XLAM artifact invalid OOXML ZIP signature' }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = $null
    try { $zip = [IO.Compression.ZipFile]::OpenRead($Artifact); if ($zip.Entries.Count -eq 0) { throw 'Blank XLAM artifact invalid OOXML ZIP' } }
    catch { throw ('Blank XLAM artifact invalid OOXML ZIP: ' + $_.Exception.Message) }
    finally { if ($null -ne $zip) { $zip.Dispose() } }
    $hash = (Get-FileHash -LiteralPath $Artifact -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($hash)) { throw 'Blank XLAM artifact hash missing' }
}

function Assert-ValidXlamArtifactWithRetry([string]$Artifact, [int]$TimeoutMs) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $lastFailure = 'validation not started'
    while ($watch.ElapsedMilliseconds -lt $TimeoutMs) {
        try {
            Assert-ValidXlamArtifact $Artifact
            return
        } catch {
            $lastFailure = $_.Exception.Message
        }
        Start-Sleep -Milliseconds 100
    }
    throw ('Blank XLAM validation timeout: path=' + $Artifact + '; timeout_ms=' + $TimeoutMs + '; cause=' + $lastFailure)
}

if ([string]::IsNullOrWhiteSpace($Path)) { throw 'XLAM destination is required' }
$fullPath = [IO.Path]::GetFullPath($Path)
if ([IO.Path]::GetExtension($fullPath) -ine '.xlam') { throw 'Destination must have .xlam extension' }
if (Test-Path -LiteralPath $fullPath) { throw 'Existing XLAM destination is rejected' }
# A fresh temporary destination is required so a stale Excel artifact cannot be promoted.
$parent = Split-Path -Parent $fullPath
if (-not (Test-Path -LiteralPath $parent -PathType Container)) { [void](New-Item -ItemType Directory -Path $parent -Force) }

$excel = $null
$workbooks = $null
$book = $null
$excelProcess = $null
$excelBinding = $null
$excelOwned = $false
$cleanupFailures = @()
try {
    $excelBaseline = Get-ExcelProcessBaseline
    $excel = New-Object -ComObject Excel.Application
    $excelBinding = Get-ExactExcelProcessOwnership $excel $excelBaseline 'blank XLAM Excel'
    $excelOwned = [bool]$excelBinding.owned
    if (-not $excelOwned) { throw ('blank XLAM Excel ownership rejected: ' + [string]$excelBinding.failure) }
    $excel.Visible = $false
    $excelProcess = $excelBinding.process
    $workbooks = $excel.Workbooks
    $book = $workbooks.Add()
    $FileFormat = 55
    $book.SaveAs($fullPath, $FileFormat)
} finally {
    try { if ($book) { $book.Close($false) } } catch { $cleanupFailures += $_.Exception.Message }
    if ($excelOwned -and $null -ne $excel) {
        try { [void](Assert-ExactExcelProcessOwnership $excelBinding) } catch { $cleanupFailures += $_.Exception.Message; $excelOwned = $false }
    }
    try { if ($excelOwned -and $null -ne $excel) { $excel.Quit() } } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $book } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $workbooks } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $excel } catch { $cleanupFailures += $_.Exception.Message }
    $book = $null
    $workbooks = $null
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($excelOwned -and $null -ne $excelProcess) {
        $excelCleanup=Stop-ExactProcessAfterGrace $excelProcess 'blank XLAM Excel'
        if(-not [string]::IsNullOrWhiteSpace($excelCleanup)){$cleanupFailures += $excelCleanup}
    }
    if ($null -ne $excelProcess) { try{$excelProcess.Dispose()}catch{$cleanupFailures += $_.Exception.Message}; $excelProcess=$null }
}
if($cleanupFailures.Count -ne 0){throw ('Blank XLAM cleanup failed: '+($cleanupFailures -join ' | '))}
Wait-ExactArtifactReadable $fullPath 10000
Assert-ValidXlamArtifactWithRetry $fullPath 10000

$fullPath
