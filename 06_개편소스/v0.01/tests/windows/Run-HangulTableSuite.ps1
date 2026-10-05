param(
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [string]$RunId = '',
    [string]$ArtifactPath = '',
    [string]$SourceDigest = '',
    [string]$SnapshotDigest = '',
    [switch]$UseRangePicker
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }

$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$DataRoot = [IO.Path]::GetFullPath($DataRoot)
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [Guid]::NewGuid().ToString().ToLowerInvariant() }
if ($RunId -notmatch '^[0-9a-f-]{36}$') { throw 'RunId must be a lowercase GUID' }
if (-not [string]::IsNullOrWhiteSpace($SourceDigest) -and $SourceDigest -notmatch '^[0-9a-f]{64}$') { throw 'SourceDigest must be a lowercase SHA-256' }
if (-not [string]::IsNullOrWhiteSpace($SnapshotDigest) -and $SnapshotDigest -notmatch '^[0-9a-f]{64}$') { throw 'SnapshotDigest must be a lowercase SHA-256' }
[void](New-Item -ItemType Directory -Path $EvidenceRoot -Force)
$originalProfileOverride=[string]$env:LHEXCEL_PROFILE_ROOT
$isolatedProfileRoot=Join-Path $EvidenceRoot 'profile/LHexcel'
[void](New-Item -ItemType Directory -Path $isolatedProfileRoot -Force)
$env:LHEXCEL_PROFILE_ROOT=$isolatedProfileRoot
$isolatedHwpxRoot=[IO.Path]::GetFullPath((Join-Path $isolatedProfileRoot 'Temp/Hwpx'))
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}
function Get-Sha256([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Test-PathInside([string]$Path,[string]$Boundary) {
    if([string]::IsNullOrWhiteSpace($Path)-or[string]::IsNullOrWhiteSpace($Boundary)){return $false}
    try{
        $full=[IO.Path]::GetFullPath($Path)
        $root=[IO.Path]::GetFullPath($Boundary).TrimEnd('\')
        return $full.StartsWith($root+'\',[StringComparison]::OrdinalIgnoreCase)
    }catch{return $false}
}
function Write-Json([string]$Path, [object]$Value) {
    $utf8 = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 24), $utf8)
}
function Get-HangulProcessSnapshot {
    $rows = New-Object 'Collections.Generic.List[object]'
    foreach ($name in @('Hwp','HwpViewer','hword','HwpFrame')) {
        foreach ($process in @(Get-Process -Name $name -ErrorAction SilentlyContinue)) {
            try {
                [void]$process.Handle
                $rows.Add([pscustomobject][ordered]@{
                    name=[string]$process.ProcessName
                    pid=[int]$process.Id
                    started_utc=$process.StartTime.ToUniversalTime().ToString('o')
                    executable=[IO.Path]::GetFullPath($process.MainModule.FileName)
                    main_window_handle=[int64]$process.MainWindowHandle
                    main_window_title=[string]$process.MainWindowTitle
                    responding=[bool]$process.Responding
                })
            } finally { try { $process.Dispose() } catch {} }
        }
    }
    @($rows.ToArray())
}
function Wait-HancomDocumentWindow([int[]]$BaselinePids,[string]$ExpectedPath,[int]$TimeoutMs=30000) {
    $needle=[IO.Path]::GetFileNameWithoutExtension($ExpectedPath)
    $timer=[Diagnostics.Stopwatch]::StartNew()
    while($timer.ElapsedMilliseconds -lt $TimeoutMs) {
        foreach($row in @(Get-HangulProcessSnapshot | Where-Object { $_.pid -notin $BaselinePids -and $_.main_window_handle -ne 0 })) {
            if(([string]$row.main_window_title).IndexOf($needle,[StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
            try {
                $element=[Windows.Automation.AutomationElement]::FromHandle([IntPtr]([int64]$row.main_window_handle))
                if($null -eq $element) { continue }
                $bounds=$element.Current.BoundingRectangle
                if($bounds.Width -le 0 -or $bounds.Height -le 0 -or $element.Current.IsOffscreen) { continue }
                return [pscustomobject][ordered]@{
                    pid=[int]$row.pid
                    process_name=[string]$row.name
                    executable=[string]$row.executable
                    expected_path=[IO.Path]::GetFullPath($ExpectedPath)
                    expected_file=[IO.Path]::GetFileName($ExpectedPath)
                    window_title=[string]$row.main_window_title
                    automation_name=[string]$element.Current.Name
                    bounds=[ordered]@{left=[int][Math]::Round($bounds.Left);top=[int][Math]::Round($bounds.Top);width=[int][Math]::Round($bounds.Width);height=[int][Math]::Round($bounds.Height)}
                    observed_after_ms=[int]$timer.ElapsedMilliseconds
                }
            } catch {}
        }
        Start-Sleep -Milliseconds 100
    }
    return $null
}
function Request-HancomWindowClose([int64]$WindowHandle) {
    if($WindowHandle -eq 0){return $false}
    try{
        $element=[Windows.Automation.AutomationElement]::FromHandle([IntPtr]$WindowHandle)
        if($null-eq$element){return $false}
        $pattern=$null
        if($element.TryGetCurrentPattern([Windows.Automation.WindowPattern]::Pattern,[ref]$pattern)){
            ([Windows.Automation.WindowPattern]$pattern).Close()
            return $true
        }
    }catch{}
    return $false
}
function Add-HangulCase([string]$Name,[bool]$Passed,[object]$Details) {
    $script:Cases.Add([pscustomobject][ordered]@{ordinal=$script:Cases.Count+1;name=$Name;status=if($Passed){'PASS'}else{'FAIL'};details=$Details})
}
function Read-ZipText([object]$Archive,[string]$Name) {
    $entry=$Archive.GetEntry($Name); if($null -eq $entry){return ''}
    $stream=$entry.Open(); try { $reader=New-Object IO.StreamReader($stream,[Text.Encoding]::UTF8); return $reader.ReadToEnd() } finally { try{$reader.Dispose()}catch{};try{$stream.Dispose()}catch{} }
}
function Test-HwpxPackage([string]$Path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive=[IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $names=@($archive.Entries|ForEach-Object FullName)
        $header=Read-ZipText $archive 'Contents/header.xml'
        $section=Read-ZipText $archive 'Contents/section0.xml'
        $mimetype=Read-ZipText $archive 'mimetype'
        return [pscustomobject]@{
            entries=$names.Count
            mimetype=($mimetype -eq 'application/hwp+zip')
            font=($header -match 'face="' -and $header -match 'type="HFT"')
            font_size=($header -match 'height="1300"')
            title_shading=($header -match 'borderFill')
            table_layout=($section -match 'cellSz' -and $section -match 'rowCnt' -and $section -match 'colCnt')
            has_external_relationship=($names -match 'externalLinks|vbaProject|oleObject').Count -gt 0
            section_bytes=$section.Length
        }
    } finally { $archive.Dispose() }
}

function Test-HwpxAlternateSettings([string]$Path,[string]$ExpectedFont) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive=[IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $header=Read-ZipText $archive 'Contents/header.xml'
        $fontNeedle='face=' + [char]34 + $ExpectedFont + [char]34
        return [pscustomobject]@{
            font=($header.Contains($fontNeedle))
            font_size=($header -match 'height="1100"')
            simple_style=(-not $header.Contains('#EEF3F8') -and -not $header.Contains('DOUBLE_SLIM'))
        }
    } finally { $archive.Dispose() }
}

$script:Cases=New-Object 'Collections.Generic.List[object]'
$artifact=Join-Path $EvidenceRoot 'Product.xlam'
$receiptPath=Join-Path $EvidenceRoot 'HangulTable.json'
$hancomReceiptPath=Join-Path $EvidenceRoot 'HancomTableOpen.json'
$manifest=Join-Path $SourceRoot 'build/manifests/Product.json'
if(Test-Path -LiteralPath $artifact){throw 'Existing Hangul Product.xlam evidence artifact is rejected'}
if(Test-Path -LiteralPath $receiptPath){throw 'Existing HangulTable receipt is rejected'}
if(Test-Path -LiteralPath $hancomReceiptPath){throw 'Existing HancomTableOpen receipt is rejected'}
$preexistingHangul=@(Get-HangulProcessSnapshot); if($preexistingHangul.Count -ne 0){throw 'Hangul native suite requires zero pre-existing Hangul processes'}
if([string]::IsNullOrWhiteSpace($ArtifactPath)){
    & (Join-Path $SourceRoot 'build/Build-Xlam.ps1') -DataRoot $DataRoot -OutputPath $artifact -ManifestPath $manifest -RibbonPath (Join-Path $SourceRoot 'src/ribbon/customUI14.xml')
    if($LASTEXITCODE -ne 0 -or -not(Test-Path -LiteralPath $artifact -PathType Leaf)){throw 'Current Product.xlam build failed'}
    Add-HangulCase 'product_build' $true @{artifact=$artifact;source='built'}
}else{
    $ArtifactPath=[IO.Path]::GetFullPath($ArtifactPath)
    if(-not(Test-Path -LiteralPath $ArtifactPath -PathType Leaf)-or[IO.Path]::GetExtension($ArtifactPath)-ine'.xlam'){throw 'Hangul supplied Product artifact rejected'}
    Copy-Item -LiteralPath $ArtifactPath -Destination $artifact
    if((Get-Sha256 $artifact)-cne(Get-Sha256 $ArtifactPath)){throw 'Hangul supplied Product artifact copy mismatch'}
    Add-HangulCase 'product_build' $true @{artifact=$artifact;source='supplied-exact-artifact'}
}
$artifactSha=Get-Sha256 $artifact; $manifestSha=Get-Sha256 $manifest
$excelBaseline=Get-ExcelProcessBaseline; if(@($excelBaseline.process_ids).Count -ne 0){throw 'Hangul native suite requires zero pre-existing Excel processes'}
$excel=$null;$binding=$null;$books=$null;$addin=$null;$workbook=$null;$sheet=$null;$selection=$null;$macroPrefix=$null
$transferResult=$false;$lastError='';$outputPath='';$primaryOutputPath='';$primaryHwpxSha='';$hwpxInfo=$null;$alternateInfo=$null;$hancomOpenProof=$null;$createdHangulPids=@();$hangulCleanup=New-Object 'Collections.Generic.List[object]';$excelCleanup=[pscustomobject][ordered]@{exit_mode='NOT_STARTED';failure=$null;owned_pid=0};$failure=$null
try {
    $excel=New-Object -ComObject Excel.Application
    $binding=Get-ExactExcelProcessOwnership $excel $excelBaseline 'HangulTable Excel'
    if(-not [bool]$binding.owned){throw ('Excel ownership rejected: '+[string]$binding.failure)}
    $excel.Visible=$true
    $excel.DisplayAlerts=$false
    $books=$excel.Workbooks
    $addin=$books.Open($artifact,$false,$true)
    $workbook=$books.Add()
    $sheet=$workbook.Worksheets.Item(1)
    $sheet.Name="HangulTransfer"
    $sheet.Cells.Item(1,1).Value2=[string][char]0xD488+[char]0xBAA9
    $sheet.Cells.Item(1,2).Value2=[string][char]0xC218+[char]0xB7C9
    $sheet.Cells.Item(1,3).Value2=[string][char]0xB2E8+[char]0xAC00
    $sheet.Cells.Item(2,1).Value2=[string][char]0xC0AC+[char]0xACFC
    $sheet.Cells.Item(2,2).Value2=[double]2
    $sheet.Cells.Item(2,3).Value2=[double]1500
    $sheet.Cells.Item(3,1).Value2=[string][char]0xBC30
    $sheet.Cells.Item(3,2).Value2=[double]3
    $sheet.Cells.Item(3,3).Value2=[double]2000
    [void]$sheet.Activate()
    $selection=$sheet.Range($sheet.Cells.Item(1,1),$sheet.Cells.Item(3,3))
    [void]$selection.Select()
    $transferMacro="{0}!NxHangulTransferCurrentSelection" -f $addin.Name
    $errorMacro="{0}!NxHangulLastError" -f $addin.Name
    $outputMacro="{0}!NxHangulLastOutputPath" -f $addin.Name
    if($UseRangePicker){
        . (Join-Path $SourceRoot 'tests/windows/Test-R72HangulPicker.ps1')
        $pickerMacro="{0}!NxHangulSendSelection" -f $addin.Name
        $cancel=Invoke-R72HangulPicker $excel ([int]$binding.pid) $pickerMacro 'cancel' ''
        $cancelPath=[string]$excel.Run($outputMacro)
        Add-HangulCase 'picker_cancel_no_output' ([string]::IsNullOrEmpty($cancelPath)) $cancel
        $newSelection=$sheet.Range('A1:B2')
        [void]$newSelection.Select()
        Release-ComObject $newSelection
        $pick=Invoke-R72HangulPicker $excel ([int]$binding.pid) $pickerMacro 'accept' '$A$1:$B$2'
        Add-HangulCase 'picker_selected_range' ($pick.status -eq 'PASS') $pick
        $transferResult=-not [string]::IsNullOrEmpty([string]$excel.Run($outputMacro))
    }else{
        $transferResult=[bool]$excel.Run($transferMacro)
    }
    $lastError=[string]$excel.Run($errorMacro)
    $outputPath=[string]$excel.Run($outputMacro)
    $outputInsideIsolatedRoot=Test-PathInside $outputPath $isolatedHwpxRoot
    Add-HangulCase "selection_transfer" ($transferResult-and$outputInsideIsolatedRoot) @{last_error=$lastError;output_path=$outputPath;isolated_hwpx_root=$isolatedHwpxRoot;output_inside_isolated_root=$outputInsideIsolatedRoot}
    if(-not $transferResult -or -not$outputInsideIsolatedRoot -or -not(Test-Path -LiteralPath $outputPath -PathType Leaf)){throw ("HWPX transfer failed or escaped isolated output root: "+$lastError)}
    $primaryOutputPath=$outputPath
    $primaryHwpxSha=Get-Sha256 $primaryOutputPath
    $hancomOpenProof=Wait-HancomDocumentWindow @($preexistingHangul|ForEach-Object{[int]$_.pid}) $primaryOutputPath 30000
    $createdHangulPids=@((Get-HangulProcessSnapshot)|ForEach-Object{[int]$_.pid}|Sort-Object -Unique)
    Add-HangulCase "hwpx_output" $true @{path=$primaryOutputPath;sha256=$primaryHwpxSha}
    $hwpxInfo=Test-HwpxPackage $outputPath
    if($UseRangePicker){
        $archive=[IO.Compression.ZipFile]::OpenRead($outputPath)
        try{$section=Read-ZipText $archive 'Contents/section0.xml'}finally{$archive.Dispose()}
        Add-HangulCase 'picker_output_dimensions' ($section -match 'rowCnt="2"' -and $section -match 'colCnt="2"') @{expected_rows=2;expected_columns=2}
    }
    $packagePassed=$hwpxInfo.mimetype -and $hwpxInfo.font -and $hwpxInfo.font_size -and $hwpxInfo.title_shading -and $hwpxInfo.table_layout -and (-not $hwpxInfo.has_external_relationship)
    Add-HangulCase "hwpx_settings_and_structure" $packagePassed $hwpxInfo
    Add-HangulCase "hancom_opened" ($null -ne $hancomOpenProof) $hancomOpenProof
    $alternateFont=[string][char]0xD568+[char]0xCD08+[char]0xB86C+[char]0xB3CB+[char]0xC6C0
    $alternateStyle=[string][char]0xAC04+[char]0xACB0+[char]0xD615
    $settingsMacro="{0}!NxHangulTryApplySettingsValues" -f $addin.Name
    $settingsResult=[string]$excel.Run($settingsMacro,$alternateFont,11,$true,$true,$alternateStyle)
    if($settingsResult -ne 'PASS'){throw ('Alternate HWPX settings failed: '+$settingsResult)}
    $transferResult=[bool]$excel.Run($transferMacro)
    $lastError=[string]$excel.Run($errorMacro)
    $outputPath=[string]$excel.Run($outputMacro)
    $alternateInsideIsolatedRoot=Test-PathInside $outputPath $isolatedHwpxRoot
    if(-not $transferResult -or -not$alternateInsideIsolatedRoot -or -not(Test-Path -LiteralPath $outputPath -PathType Leaf)){throw ("Alternate HWPX transfer failed or escaped isolated output root: "+$lastError)}
    $alternateInfo=Test-HwpxAlternateSettings $outputPath $alternateFont
    Add-HangulCase "alternate_font_size_and_table_style" ($alternateInsideIsolatedRoot -and $alternateInfo.font -and $alternateInfo.font_size -and $alternateInfo.simple_style) @{isolated_hwpx_root=$isolatedHwpxRoot;output_inside_isolated_root=$alternateInsideIsolatedRoot;package=$alternateInfo}
    $createdHangulPids=@($createdHangulPids + @((Get-HangulProcessSnapshot)|ForEach-Object{[int]$_.pid}) | Sort-Object -Unique)
    Add-HangulCase "clipboard_released" ([int]$excel.CutCopyMode -eq 0) @{cut_copy_mode=[int]$excel.CutCopyMode}
}catch{$failure=$_.Exception.Message}finally{
    $ownedHangul=@(Get-HangulProcessSnapshot|Where-Object{$_.pid -in $createdHangulPids})
    foreach($processInfo in $ownedHangul){$process=Get-Process -Id $processInfo.pid -ErrorAction SilentlyContinue;if($null -ne $process){$uiaCloseRequested=Request-HancomWindowClose ([int64]$processInfo.main_window_handle);Start-Sleep -Milliseconds 250;$processCloseRequested=$false;try{if(-not$process.HasExited){$processCloseRequested=[bool]$process.CloseMainWindow()}}catch{};$result=Stop-ExactProcessAfterGrace $process ("HangulTable "+$process.ProcessName) 30000 -Detailed;$hangulCleanup.Add([pscustomobject][ordered]@{pid=$process.Id;uia_close_requested=$uiaCloseRequested;process_close_requested=$processCloseRequested;exit_mode=$result.exit_mode;failure=$result.failure});try{$process.Dispose()}catch{}}}
    try{if($null -ne $selection){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($selection)}}catch{}
    try{if($null -ne $sheet){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet)}}catch{}
    try{if($null -ne $workbook){$workbook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)}}catch{}
    try{if($null -ne $addin){$addin.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addin)}}catch{}
    try{Release-ComObject $books}catch{}
    $selection=$null;$sheet=$null;$workbook=$null;$addin=$null;$books=$null
    try{if($null -ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}}catch{}
    $excel=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect();[GC]::WaitForPendingFinalizers()
    if($null -ne $binding -and $null -ne $binding.process){$result=Stop-ExactProcessAfterGrace $binding.process "HangulTable Excel" 15000 -Detailed;$excelCleanup=[pscustomobject][ordered]@{exit_mode=$result.exit_mode;failure=$result.failure;owned_pid=$binding.pid};try{$binding.process.Dispose()}catch{}}
}
$env:LHEXCEL_PROFILE_ROOT=$originalProfileOverride
$remainingHangul=@(Get-HangulProcessSnapshot)
$failedCleanup=@($hangulCleanup | Where-Object { $_.exit_mode -ne "NATURAL" -or $_.failure })
if($remainingHangul.Count -ne 0 -and $null -eq $failure){$failure="Hangul process residue remained"}
if($failedCleanup.Count -ne 0 -and $null -eq $failure){$failure="Hangul cleanup failed"}
if(($excelCleanup.exit_mode -ne 'NATURAL' -or -not[string]::IsNullOrWhiteSpace([string]$excelCleanup.failure))-and$null-eq$failure){$failure='HangulTable Excel did not exit naturally'}
$passed=@($script:Cases | Where-Object { $_.status -eq "PASS" }).Count
$expectedCases=if($UseRangePicker){10}else{7}
$status=if($null -eq $failure -and $script:Cases.Count -eq $expectedCases -and $passed -eq $expectedCases){"PASS"}else{"DIAGNOSTIC"}
$primaryHashUnchanged=(-not [string]::IsNullOrWhiteSpace($primaryOutputPath) -and (Test-Path -LiteralPath $primaryOutputPath -PathType Leaf) -and (Get-Sha256 $primaryOutputPath) -ceq $primaryHwpxSha)
$naturalHancomClose=($createdHangulPids.Count -gt 0 -and $hangulCleanup.Count -gt 0 -and @($hangulCleanup|Where-Object{$_.exit_mode -ne 'NATURAL' -or $_.failure}).Count -eq 0 -and $remainingHangul.Count -eq 0)
$hancomStatus=if($null -ne $hancomOpenProof -and $primaryHashUnchanged -and $naturalHancomClose){'PASS'}else{'DIAGNOSTIC'}
$receipt=[ordered]@{
    schema_version=2
    suite="HangulTable"
    run_id=$RunId
    status=$status
    artifact_sha256=$artifactSha
    source_tree_sha256=$SourceDigest
    source_snapshot_sha256=$SnapshotDigest
    manifest_sha256=$manifestSha
    transport="Excel selection to embedded HWPX with v3.4 table settings"
    isolated_hwpx_root=$isolatedHwpxRoot
    created_hangul_pids=@($createdHangulPids)
    hwpx_path=$outputPath
    hwpx_info=$hwpxInfo
    alternate_hwpx_info=$alternateInfo
    cases=@($script:Cases.ToArray())
    cleanup=[ordered]@{excel=$excelCleanup;hangul=@($hangulCleanup.ToArray());remaining_hangul=@($remainingHangul)}
    failure=$failure
    completed_utc=[DateTime]::UtcNow.ToString("o")
}
Write-Json $receiptPath $receipt
Write-Json $hancomReceiptPath ([ordered]@{
    schema_version=1
    suite='HancomTableOpen'
    status=$hancomStatus
    run_id=$RunId
    artifact_sha256=$artifactSha
    source_tree_sha256=$SourceDigest
    source_snapshot_sha256=$SnapshotDigest
    hwpx_path=$primaryOutputPath
    hwpx_sha256=$primaryHwpxSha
    hwpx_hash_unchanged=$primaryHashUnchanged
    window=$hancomOpenProof
    cleanup=[ordered]@{natural_close=$naturalHancomClose;processes=@($hangulCleanup.ToArray());remaining_hangul=@($remainingHangul)}
    completed_utc=[DateTime]::UtcNow.ToString('o')
})
if($status -ne "PASS" -or $hancomStatus -ne 'PASS') { Write-Output ("DIAGNOSTIC|HangulTable|{0}/{1}" -f $passed,$expectedCases); exit 1 }
Write-Output ("PASS|HangulTable|{0}/{1}" -f $passed,$expectedCases)
exit 0
