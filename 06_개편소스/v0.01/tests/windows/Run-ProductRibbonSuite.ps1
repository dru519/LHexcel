param(
    [string]$Suite = 'ProductRibbon',
    [ValidateSet('Green')][string]$Mode = 'Green',
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [string]$EvidenceRoot = '',
    [string]$RunId = '',
    [string]$SourceDigest = '',
    [string]$SnapshotDigest = '',
    [string]$ArtifactPath = '',
    [string]$StartupArtifactPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }
Add-Type -AssemblyName Microsoft.VisualBasic
$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$failurePhase = 'preflight'
$terminalCode = 0
$failureMessage = $null
$failureDetail = $null
$probeStage = 'not-started'
$cleanupFailures = @()
$cleanupExitMode = 'NOT_STARTED'

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

function Get-HexDigest([string]$Path, [string]$Algorithm = 'SHA256') {
    (Get-FileHash -LiteralPath $Path -Algorithm $Algorithm -ErrorAction Stop).Hash.ToLowerInvariant()
}

function Get-TextSha256([string]$Text) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try { ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Write-Json([string]$Path, [object]$Value) {
    $utf8 = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 30), $utf8)
}

function Write-ProductRibbonOwnership([string]$Evidence, [object]$Binding, [string]$Artifact, [string]$State) {
    if ($null -eq $Binding -or -not [bool]$Binding.owned) { throw 'ProductRibbon ownership receipt binding rejected' }
    Write-Json (Join-Path $Evidence 'ProductRibbon.Ownership.json') ([ordered]@{
        schema_version=1
        suite='ProductRibbon'
        run_id=$RunId
        state=$State
        pid=[int]$Binding.pid
        started_utc=[string]$Binding.started_utc
        executable=[IO.Path]::GetFullPath([string]$Binding.executable)
        artifact=[IO.Path]::GetFullPath($Artifact)
    })
}

function Get-ProductManifest {
    $path = Join-Path $SourceRoot 'build/manifests/Product.json'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'Product manifest missing' }
    $manifest = Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
    $manifestProperties = @($manifest.PSObject.Properties.Name)
    if ([int]$manifest.schema_version -ne 3 -or $manifestProperties -contains 'runtime') { throw 'Runtime-free Product manifest contract rejected' }
    $featureContract = Get-Content -LiteralPath (Join-Path $SourceRoot 'contracts/feature-contract.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $expectedFeatureCount = @($featureContract.lineage.release_feature_ids).Count
    if ([int]$manifest.release_scope.feature_count -ne $expectedFeatureCount -or @($manifest.release_scope.included_feature_ids).Count -ne $expectedFeatureCount) { throw 'Product release feature inventory rejected' }
    $excluded = @($manifest.release_scope.excluded_feature_ids)
    if ($excluded.Count -ne 0) { throw 'Product release scope must not carry retired feature placeholders' }
    $hangulModules = @($manifest.modules | Where-Object { [string]$_.path -match '^src/vba/features/hangul/' })
    if ($hangulModules.Count -lt 9 -or @($manifest.modules | Where-Object { [string]$_.path -match '^src/vba/features/hwpx/' }).Count -ne 0 -or @($hangulModules | Where-Object { [string]$_.path -match 'NxHwpxController|NxHwpxSerializer|NxHwpxEmbeddedResources|NxHangulSettings' }).Count -lt 4) { throw 'Product embedded HWPX Hangul source inventory rejected' }
    return $manifest
}

function Get-PeArchitecture([string]$Executable) {
    $stream = [IO.File]::Open($Executable, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    $reader = New-Object IO.BinaryReader($stream)
    try {
        if ($reader.ReadUInt16() -ne 0x5A4D) { throw 'Excel executable DOS signature rejected' }
        $stream.Position = 0x3C
        $peOffset = $reader.ReadInt32()
        $stream.Position = $peOffset
        if ($reader.ReadUInt32() -ne 0x00004550) { throw 'Excel executable PE signature rejected' }
        $machine = $reader.ReadUInt16()
        if ($machine -eq 0x8664) { return 64 }
        if ($machine -eq 0x014C) { return 32 }
        throw ('Excel executable PE machine rejected: ' + $machine)
    } finally { $reader.Dispose(); $stream.Dispose() }
}

function Get-OfficeProductReleaseIds {
    $values = New-Object 'Collections.Generic.List[string]'
    foreach ($view in @(
        [Microsoft.Win32.RegistryView]::Registry64,
        [Microsoft.Win32.RegistryView]::Registry32
    )) {
        $base = $null
        $configuration = $null
        try {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                $view
            )
            $configuration = $base.OpenSubKey('SOFTWARE\Microsoft\Office\ClickToRun\Configuration', $false)
            if ($null -ne $configuration) {
                $rawReleaseIds = $configuration.GetValue('ProductReleaseIds', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                foreach ($entry in @($rawReleaseIds)) {
                    foreach ($token in ([string]$entry -split '[,;]')) {
                        $trimmed = $token.Trim()
                        if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                            [void]$values.Add($trimmed)
                        }
                    }
                }
            }
        } finally {
            if ($null -ne $configuration) { $configuration.Dispose() }
            if ($null -ne $base) { $base.Dispose() }
        }
    }
    return @($values | Sort-Object -Unique)
}

function Get-RegisteredExcelExecutable {
    $candidates = New-Object 'Collections.Generic.List[string]'
    foreach ($view in @(
        [Microsoft.Win32.RegistryView]::Registry64,
        [Microsoft.Win32.RegistryView]::Registry32
    )) {
        $base = $null
        $appPath = $null
        try {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                $view
            )
            $appPath = $base.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', $false)
            if ($null -ne $appPath) {
                $raw = [string]$appPath.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                if (-not [string]::IsNullOrWhiteSpace($raw)) { [void]$candidates.Add($raw) }
            }
        } finally {
            if ($null -ne $appPath) { $appPath.Dispose() }
            if ($null -ne $base) { $base.Dispose() }
        }
    }
    foreach ($candidate in @($candidates | Sort-Object -Unique)) {
        try {
            $full = [IO.Path]::GetFullPath(([string]$candidate).Trim().Trim([char]34))
            if ((Test-Path -LiteralPath $full -PathType Leaf) -and [IO.Path]::GetFileName($full) -ceq 'EXCEL.EXE') { return $full }
        } catch { }
    }
    throw 'Registered EXCEL.EXE path unavailable'
}

function Start-InteractiveProductExcel([object]$Baseline) {
    $excelPath = Get-RegisteredExcelExecutable
    $launch = $null
    $candidate = $null
    $candidateBinding = $null
    $transferred = $false
    try {
        $launch = Start-Process -FilePath $excelPath -ArgumentList @('/x') -WindowStyle Normal -PassThru
        if ($null -eq $launch) { throw 'Interactive Excel process launch failed' }
        for ($attempt = 0; $attempt -lt 150; $attempt++) {
            try {
                $candidate = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
                $candidateBinding = Get-ExactExcelProcessOwnership $candidate $Baseline 'ProductRibbon interactive Excel'
                if ([bool]$candidateBinding.owned -and [int]$candidateBinding.pid -eq $launch.Id -and [bool]$candidate.UserControl) {
                    $result = [pscustomobject]@{ excel=$candidate; binding=$candidateBinding }
                    $candidate = $null
                    $candidateBinding = $null
                    $transferred = $true
                    return $result
                }
            } catch { }
            finally {
                if (-not $transferred) {
                    if ($null -ne $candidateBinding -and $null -ne $candidateBinding.process) { try { $candidateBinding.process.Dispose() } catch { } }
                    Release-ComObject $candidate
                    $candidateBinding = $null
                    $candidate = $null
                }
            }
            Start-Sleep -Milliseconds 200
        }
        throw 'Interactive Excel COM binding timed out'
    } finally {
        $failedLaunchCleanupError = $null
        if (-not $transferred -and $null -ne $launch) {
            try {
                if (-not $launch.HasExited) {
                    $failedLaunchCleanup = Stop-ExactProcessAfterGrace -Process $launch -Label 'ProductRibbon interactive launch' -GraceMs 10000 -Detailed
                    if ($failedLaunchCleanup.failure) { $failedLaunchCleanupError = [string]$failedLaunchCleanup.failure }
                }
            } catch { $failedLaunchCleanupError = $_.Exception.Message }
        }
        if ($null -ne $launch) { try { $launch.Dispose() } catch { } }
        if (-not [string]::IsNullOrWhiteSpace($failedLaunchCleanupError)) { throw $failedLaunchCleanupError }
    }
}

function Resolve-StartupProductProject([object]$Excel, [string]$ExpectedPath) {
    $expected = [IO.Path]::GetFullPath($ExpectedPath)
    $match = $null
    $vbe = $null
    $projects = $null
    try {
        $vbe = $Excel.VBE
        $projects = $vbe.VBProjects
        for ($index = 1; $index -le [int]$projects.Count; $index++) {
            $candidate = $null
            try {
                $candidate = $projects.Item($index)
                $candidateFile = [string]$candidate.FileName
                if ([string]::IsNullOrWhiteSpace($candidateFile)) { continue }
                $candidatePath = [IO.Path]::GetFullPath($candidateFile)
                if ([string]::Equals($candidatePath, $expected, [StringComparison]::OrdinalIgnoreCase)) {
                    if ($null -ne $match) {
                        Release-ComObject $match
                        $match = $null
                        throw 'ProductRibbon startup XLAM was not auto-loaded exactly once'
                    }
                    $match = $candidate
                    $candidate = $null
                }
            } finally { Release-ComObject $candidate }
        }
    } finally {
        Release-ComObject $projects
        Release-ComObject $vbe
    }
    if ($null -eq $match) { throw 'ProductRibbon startup XLAM was not auto-loaded' }
    return $match
}

function Resolve-ProductUiHostWorkbook([object]$Books) {
    $hostWorkbook = $null
    for ($index = 1; $index -le [int]$Books.Count; $index++) {
        $candidate = $null
        try {
            $candidate = $Books.Item($index)
            if (-not [bool]$candidate.IsAddin) {
                if ($null -ne $hostWorkbook) {
                    Release-ComObject $hostWorkbook
                    $hostWorkbook = $null
                    throw 'Interactive Excel startup host workbook count rejected'
                }
                $hostWorkbook = $candidate
                $candidate = $null
            }
        } finally { Release-ComObject $candidate }
    }
    if ($null -eq $hostWorkbook) { $hostWorkbook = $Books.Add() }
    return $hostWorkbook
}

function Get-OfficeIdentity([object]$Excel, [object]$Book, [object]$ExcelBinding) {
    $releaseIds = @(Get-OfficeProductReleaseIds)
    $officeVersion = if (@($releaseIds | Where-Object { $_ -match '2024' }).Count) {
        'Office 2024'
    } elseif (@($releaseIds | Where-Object { $_ -match '^(O365|M365)' }).Count) {
        'Microsoft 365'
    } else {
        $measured = if ($releaseIds.Count -eq 0) { 'NONE' } else { $releaseIds -join ',' }
        throw ('Office 2024 or Microsoft 365 ProductReleaseIds evidence required; measured_product_release_ids=' + $measured)
    }
    [void](Assert-ExactExcelProcessOwnership $ExcelBinding)
    $bitness = Get-PeArchitecture ([string]$ExcelBinding.executable)
    if ($bitness -ne 64) { throw 'Excel x64 evidence required' }
    return [ordered]@{
        method = 'exact-owned Excel process PE + Excel COM + ClickToRun registry'
        office_version = $officeVersion
        compatibility_baseline = 'Office 2024'
        architecture = 'x64'
        excel_version = [string]$Excel.Version
        excel_build = [string]$Excel.Build
        product_release_ids = $releaseIds
        macro_bitness = $bitness
    }
}

function Invoke-PackageProbe([string]$Artifact) {
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Artifact)
    try {
        $parts = @($archive.Entries | Where-Object { $_.FullName -eq 'customUI/customUI14.xml' })
        $relations = @($archive.Entries | Where-Object { $_.FullName -eq '_rels/.rels' })
        if ($parts.Count -ne 1 -or $relations.Count -ne 1) { throw 'Custom UI v2 package part count rejected' }
        $reader = New-Object IO.StreamReader($relations[0].Open(), [Text.Encoding]::UTF8)
        try { [xml]$rels = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $matches = @($rels.Relationships.Relationship | Where-Object { $_.Type -eq 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility' })
        if ($matches.Count -ne 1) { throw 'Custom UI v2 relationship rejected' }
        return [ordered]@{ method='OOXML ZIP/XML reopen'; custom_ui_parts=$parts.Count; relationship_type=[string]$matches[0].Type }
    } finally { $archive.Dispose() }
}

function Invoke-RibbonCallbackProbe([object]$Excel, [object]$Project, [string]$RuntimeArtifact, [string]$InspectionArtifact, [object]$RibbonUiResults, [bool]$RequireSavedCompileState) {
    $vbe = $null; $bars = $null; $control = $null; $components = $null; $compileTarget = $null; $registryModule = $null; $references = $null
    try {
        if ($null -eq $Project -or -not [string]::Equals([IO.Path]::GetFullPath([string]$Project.FileName), [IO.Path]::GetFullPath($RuntimeArtifact), [StringComparison]::OrdinalIgnoreCase)) { throw 'Product ribbon callback artifact identity rejected' }
        if ($null -eq $RibbonUiResults -or @($RibbonUiResults.cases).Count -ne 3) { throw 'Actual Product ribbon UI results missing' }
        $vbe = $Excel.VBE
        $bars = $vbe.CommandBars
        $control = Find-VbeCompileControl $bars
        if ($null -eq $control -or [int]$control.Id -ne 578) { throw 'VBE Compile command unavailable' }
        Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
        $archive = [IO.Compression.ZipFile]::OpenRead($InspectionArtifact)
        try {
            $embeddedEntries = @($archive.Entries | Where-Object { $_.FullName -ceq 'customUI/customUI14.xml' })
            if ($embeddedEntries.Count -ne 1) { throw 'Embedded Product ribbon XML part count rejected' }
            $entryStream = $embeddedEntries[0].Open(); $memory = New-Object IO.MemoryStream
            try { $entryStream.CopyTo($memory); $embeddedBytes = $memory.ToArray() } finally { $memory.Dispose(); $entryStream.Dispose() }
        } finally { $archive.Dispose() }
        $sourceXmlPath = Join-Path $SourceRoot 'src/ribbon/customUI14.xml'
        $sourceBytes = [IO.File]::ReadAllBytes($sourceXmlPath)
        $sha = [Security.Cryptography.SHA256]::Create()
        try {
            $embeddedXmlSha256 = ([BitConverter]::ToString($sha.ComputeHash($embeddedBytes))).Replace('-','').ToLowerInvariant()
            $sourceXmlSha256 = ([BitConverter]::ToString($sha.ComputeHash($sourceBytes))).Replace('-','').ToLowerInvariant()
        } finally { $sha.Dispose() }
        if ($embeddedXmlSha256 -ne $sourceXmlSha256) { throw 'Embedded Product ribbon XML SHA differs from source XML' }
        [xml]$embeddedXml = [Text.Encoding]::UTF8.GetString($embeddedBytes)
        [xml]$sourceXml = [Text.Encoding]::UTF8.GetString($sourceBytes)
        $productTabs = @($embeddedXml.SelectNodes('//*[local-name()="tab" and @id="NX-TAB"]'))
        $homeTabs = @($embeddedXml.SelectNodes('//*[local-name()="tab" and @idMso="TabHome"]'))
        if ($productTabs.Count -ne 1 -or $homeTabs.Count -ne 1) { throw 'Embedded dual Ribbon tab identity rejected' }
        $embeddedGroups = @($productTabs[0].SelectNodes('./*[local-name()="group"]'))
        $embeddedButtons = @($productTabs[0].SelectNodes('.//*[local-name()="button"]'))
        $expectedButtonCount = @($sourceXml.SelectNodes('//*[local-name()="tab" and @id="NX-TAB"]//*[local-name()="button"]')).Count
        $homeGroups = @($homeTabs[0].SelectNodes('./*[local-name()="group"]'))
        if ($embeddedGroups.Count -ne 12 -or $homeGroups.Count -ne 1) { throw 'Embedded Product ribbon group count rejected' }
        if ($embeddedButtons.Count -ne $expectedButtonCount) { throw 'Embedded Product ribbon button count rejected' }
        $callbacks = @(
            @($embeddedXml.SelectNodes('//*[@onLoad]')) | ForEach-Object { [string]$_.onLoad }
            @($embeddedXml.SelectNodes('//*[@onAction]')) | ForEach-Object { [string]$_.onAction }
            @($embeddedXml.SelectNodes('//*[@getContent]')) | ForEach-Object { [string]$_.getContent }
            @($embeddedXml.SelectNodes('//*[@getPressed]')) | ForEach-Object { [string]$_.getPressed }
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique
        $expectedCallbacks = @('NxRibbonExecute','NxRibbonGetAllFunctionsContent','NxRibbonGetFavoritesContent','NxRibbonGetFocusContent','NxRibbonGetManagementContent','NxRibbonOnLoad') | Sort-Object -Unique
        if (($callbacks -join [char]31) -cne ($expectedCallbacks -join [char]31)) { throw 'Ribbon callback declarations rejected' }
        $featureContract = Get-Content -LiteralPath (Join-Path $SourceRoot 'contracts/feature-contract.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $registeredFeatures = @($featureContract.lineage.release_feature_ids | ForEach-Object { [string]$_ } | Sort-Object)
        $expectedFeatureCount = @($featureContract.lineage.release_feature_ids).Count
        if ($registeredFeatures.Count -ne $expectedFeatureCount -or @($registeredFeatures | Sort-Object -Unique).Count -ne $expectedFeatureCount -or $registeredFeatures -notcontains 'NX-HANGUL-TABLE-SEND') { throw 'Registered Product feature inventory rejected' }
        $source = ''
        $components = $Project.VBComponents
        $componentCount = [int]$components.Count
        for ($componentIndex = 1; $componentIndex -le $componentCount; $componentIndex++) {
            $component = $null
            $codeModule = $null
            try {
                $component = $components.Item($componentIndex)
                $codeModule = $component.CodeModule
                if ($codeModule.CountOfLines -gt 0) { $source += $codeModule.Lines(1, $codeModule.CountOfLines) + "`n" }
            } finally { Release-ComObject $codeModule; Release-ComObject $component }
        }
        $gradePolicyPatterns = @(
            '(?m)^[ \t]*RegisterDataFeature\s+registry,\s*"NX-DATA-UNIQUE-COUNT",\s*NxExecutionFast,\s*NxCapabilityRead\s+Or\s+NxCapabilityWorksheetCreate,',
            '(?m)^[ \t]*RegisterDataFeature\s+registry,\s*"NX-DATA-DUPLICATE-LIST",\s*NxExecutionFast,\s*NxCapabilityRead\s+Or\s+NxCapabilityWorksheetCreate,',
            '(?ms)^[ \t]*RegisterFileFeature\s+registry,\s*"NX-FILE-FOLDER-CREATE",\s*NxCapabilityRead\s+Or\s+NxCapabilityDirectoryCreate,\s*50\s+.*?^Private Sub RegisterFileFeature\b(?:(?!^End Sub).)*?definition\.Configure\s+featureId,\s*"NX-CAT-FILE",(?:(?!^End Sub).)*?"disable_feature",\s*"disable_feature",\s*NxExecutionPlanned,',
            '(?ms)\A(?=.*?^[ \t]*If\s+mRequest\.OutputMode\s*=\s*"\uC6D0\uBCF8 \uBCC0\uACBD"\s+Then\s+plan\.AddSideEffect\s+NxCreateFeatureCellEffect\("NX-DATA-NORMALIZE",\s*context\.SelectionTargetIdentity,\s*_\s*NxCellValueMask\s+Or\s+NxCellNumberFormatMask,\s*mRequest\.ActualCellCount\)\s+ElseIf\s+mRequest\.OutputMode\s*=\s*"\uC624\uB978\uCABD \uC0C8 \uC5F4 \uC0BD\uC785"\s+Then\s+plan\.AddSideEffect\s+NxCreateCommandOwnedEffect\("NX-DATA-NORMALIZE",\s*NxColumnInsert,\s*context\.WorkbookIdentity,\s*_\s*"delete only columns inserted by this conversion",\s*mRequest\.ActualCellCount\)\s+Else\s+plan\.AddSideEffect\s+NxCreateCommandOwnedEffect\("NX-DATA-NORMALIZE",\s*NxWorksheetCreate,\s*context\.WorkbookIdentity,\s*_\s*"delete only the command-created normalization result sheet",\s*mRequest\.ActualCellCount\)\s+End\s+If).*?^[ \t]*If\s+mRequest\.IsAnalysis\s+Then\s+mAnalysisValues\s*=\s*NxDataAnalysisValues\(mRequest\)\s+mExpectedOutputCells\s*=\s*NxDataAnalysisOutputCells\(mFeatureId,\s*mAnalysisValues\)\s+plan\.AddSideEffect\s+NxCreateCommandOwnedEffect\(mFeatureId,\s*NxWorksheetCreate,\s*context\.WorkbookIdentity,\s*_\s*"remove only the command-created analysis workbook or sheet",\s*mExpectedOutputCells\)\s+Else\s+Set\s+mGroups\s*=\s*NxDataGroups\(mRequest\)\s+End\s+If\s+If\s+Not\s+mRequest\.IsAnalysis\s+And\s+mFeatureId\s*=\s*"NX-DATA-DUPLICATE-LIST"\s+Then\s+mExpectedOutputCells\s*=\s*NxDataDuplicateOutputCellCount\(mGroups\)\s+plan\.AddSideEffect\s+NxCreateCommandOwnedEffect\(mFeatureId,\s*NxWorksheetCreate,\s*context\.WorkbookIdentity,\s*_\s*"delete only the command-created duplicate result sheet",\s*mExpectedOutputCells\)\s+End\s+If',
            '(?m)^[ \t]*Case\s+NxCellMutation\s+actualCellCount\s*=\s*MaximumCount\(actualCellCount,\s*context\.ActualCellCount\)\s+Case\s+NxFileCreate,\s*NxFileReplace\s+fileMutation\s*=\s*True\s+If\s+effect\.Kind\s*=\s*NxFileReplace\s+Then\s+existingContentOverwrite\s*=\s*True\s+Case\s+NxBinaryPackageWrite\s+fileMutation\s*=\s*True\s+Case\s+NxDirectoryBatchCreate\s+fileMutation\s*=\s*True\s+Case\s+NxWorksheetCreate,\s*NxColumnInsert\s+resultSheetCount\s*=\s*resultSheetCount\s*\+\s*1\s+actualCellCount\s*=\s*MaximumCount\(actualCellCount,\s*effect\.MaximumCells\)',
            '(?m)^[ \t]*Set\s+decision\s*=\s*NxExecutionGradePolicy\.ResolveDecision\(definition,\s*actualCellCount,\s*fileMutation,\s*_\s*externalDocumentMutation,\s*manualRecoveryMutation,\s*\(command\.FeatureId\s*=\s*"NX-AI-IMAGE"\),\s*riskyFormula,\s*resultSheetCount,\s*existingContentOverwrite,\s*fileMutation\)',
            '(?m)^[ \t]*If\s+resultSheetCount\s*=\s*1\s+Then\s+resolved\s*=\s*RaiseGrade\(resolved,\s*NxExecutionGuarded\)',
            '(?m)^[ \t]*If\s+resultSheetCount\s*>\s*1\s+Or\s+existingContentOverwrite\s+Or\s+fileCreation\s+Then\s+resolved\s*=\s*RaiseGrade\(resolved,\s*NxExecutionPlanned\)'
        )
        foreach ($pattern in $gradePolicyPatterns) {
            if ($source -notmatch $pattern) { throw ('Compiled Product grade policy source rejected: ' + $pattern) }
        }
        foreach ($callback in $callbacks) {
            if ($source -notmatch ('(?im)^Public Sub ' + [regex]::Escape($callback) + '\s*\(')) { throw ('Ribbon callback source missing: ' + $callback) }
        }
        $compileTarget = $components.Item('NxGeneratedProductRegistry')
        if ($null -eq $compileTarget) { throw 'Product compile target module missing' }
        $registryModule = $compileTarget.CodeModule
        $registrySource = [string]$registryModule.Lines(1, [int]$registryModule.CountOfLines)
        $artifactFeatures = @([regex]::Matches($registrySource, "(?m)^' Feature: (NX-[A-Z0-9-]+)\r?$") | ForEach-Object { $_.Groups[1].Value } | Sort-Object)
        if ($artifactFeatures.Count -ne $expectedFeatureCount -or ($artifactFeatures -join [char]31) -cne ($registeredFeatures -join [char]31)) { throw 'Actual Product feature registry inventory mismatch' }
        [void]$compileTarget.Activate()
        $activeProjectFile = [string]$vbe.ActiveVBProject.FileName
        if (-not [string]::Equals([IO.Path]::GetFullPath($activeProjectFile), [IO.Path]::GetFullPath($RuntimeArtifact), [StringComparison]::OrdinalIgnoreCase)) { throw 'Product compile project identity mismatch' }
        $compileSaved = -not [bool]$control.Enabled
        if ($RequireSavedCompileState -and -not $compileSaved) { throw 'Product saved compile state is not disabled on fresh read-only reopen' }
        $references = $Project.References
        $referenceCount = [int]$references.Count
        $brokenReferenceCount = 0
        for ($referenceIndex = 1; $referenceIndex -le $referenceCount; $referenceIndex++) {
            $reference = $null
            try { $reference = $references.Item($referenceIndex); if ([bool]$reference.IsBroken) { $brokenReferenceCount++ } } finally { Release-ComObject $reference }
        }
        if ($brokenReferenceCount -ne 0) { throw 'Product reopened project has broken references' }
        $script:productCompileContext = [ordered]@{
            target_module='NxGeneratedProductRegistry'
            active_project_artifact_name=[IO.Path]::GetFileName($activeProjectFile)
            compile_control_id=578
            compile_saved=$compileSaved
            compile_state_requirement=if($RequireSavedCompileState){'saved-readonly-reopen'}else{'runtime-execution-attested'}
            reference_count=$referenceCount
            broken_reference_count=$brokenReferenceCount
        }
        return [ordered]@{
            method=if($RequireSavedCompileState){'deployed Product.xlam embedded XML + saved VBE compile + actual ribbon UI callbacks'}else{'startup-loaded Product.xlam embedded XML + runtime VBE + actual ribbon UI callbacks'}
            callbacks=$callbacks
            compile_control_id=578
            compile_saved=$compileSaved
            compile_state_requirement=if($RequireSavedCompileState){'saved-readonly-reopen'}else{'runtime-execution-attested'}
            active_project_artifact_name=[IO.Path]::GetFileName($activeProjectFile)
            reference_count=$referenceCount
            broken_reference_count=0
            group_count=$embeddedGroups.Count
            button_count=$embeddedButtons.Count
            feature_count=$registeredFeatures.Count
            embedded_xml_sha256=$embeddedXmlSha256
            source_xml_sha256=$sourceXmlSha256
            macro_execution='actual-ribbon-ui'
            grade_policy_contract_attested=$true
            grade_policy_pattern_count=$gradePolicyPatterns.Count
        }
    } finally { Release-ComObject $references; Release-ComObject $registryModule; Release-ComObject $compileTarget; Release-ComObject $components; Release-ComObject $control; Release-ComObject $bars; Release-ComObject $vbe }
}

function Initialize-WindowApi {
    if ('NxProductWindowApi' -as [type]) { return }
    Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
public sealed class NxOwnedWindowSnapshot {
  public long Handle;
  public string ClassName;
  public string Title;
}
public struct NxClientRect {
  public int Left;
  public int Top;
  public int Right;
  public int Bottom;
}
public struct NxClientPoint {
  public int X;
  public int Y;
}
public static class NxProductWindowApi {
  public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr parameter);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr parameter);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int maximum);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int maximum);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int command);
  [DllImport("user32.dll")] public static extern void SwitchToThisWindow(IntPtr hwnd, bool altTab);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint fromThread, uint toThread, bool attach);
  [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
  [DllImport("user32.dll", SetLastError=true)] public static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);
  [DllImport("user32.dll", SetLastError=true)] public static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll", SetLastError=true)] public static extern bool GetClientRect(IntPtr hwnd, out NxClientRect rectangle);
  [DllImport("user32.dll", SetLastError=true)] public static extern bool ClientToScreen(IntPtr hwnd, ref NxClientPoint point);
  public static bool ActivateForegroundWindow(IntPtr hwnd) {
    ShowWindow(hwnd, 9);
    SwitchToThisWindow(hwnd, true);
    for (var attempt = 0; attempt < 20; attempt++) {
      if (GetForegroundWindow() == hwnd) return true;
      SetForegroundWindow(hwnd);
      Thread.Sleep(50);
    }
    var foreground = GetForegroundWindow();
    uint targetProcessId;
    uint targetThread = GetWindowThreadProcessId(hwnd, out targetProcessId);
    uint foregroundProcessId;
    uint foregroundThread = GetWindowThreadProcessId(foreground, out foregroundProcessId);
    uint currentThread = GetCurrentThreadId();
    bool attachedForeground = foregroundThread != 0 && currentThread != foregroundThread && AttachThreadInput(currentThread, foregroundThread, true);
    bool attachedTarget = targetThread != 0 && currentThread != targetThread && AttachThreadInput(currentThread, targetThread, true);
    bool activated = false;
    try {
      ShowWindow(hwnd, 9);
      BringWindowToTop(hwnd);
      SetForegroundWindow(hwnd);
      SetFocus(hwnd);
      for (var attempt = 0; attempt < 20; attempt++) {
        if (GetForegroundWindow() == hwnd) { activated = true; break; }
        Thread.Sleep(50);
      }
    } finally {
      if (attachedTarget) AttachThreadInput(currentThread, targetThread, false);
      if (attachedForeground) AttachThreadInput(currentThread, foregroundThread, false);
    }
    return activated;
  }
  public static IntPtr FindVisibleTopLevelWindow(string title, uint wantedProcessId) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((hwnd, parameter) => {
      uint processId;
      GetWindowThreadProcessId(hwnd, out processId);
      if (processId != wantedProcessId || !IsWindowVisible(hwnd)) return true;
      var text = new StringBuilder(2048);
      GetWindowText(hwnd, text, text.Capacity);
      if (!String.Equals(text.ToString(), title, StringComparison.Ordinal)) return true;
      found = hwnd;
      return false;
    }, IntPtr.Zero);
    return found;
  }
  public static NxOwnedWindowSnapshot[] SnapshotVisibleTopLevelWindows(uint wantedProcessId) {
    var rows = new List<NxOwnedWindowSnapshot>();
    EnumWindows((hwnd, parameter) => {
      uint processId;
      GetWindowThreadProcessId(hwnd, out processId);
      if (processId != wantedProcessId || !IsWindowVisible(hwnd)) return true;
      var title = new StringBuilder(2048);
      var className = new StringBuilder(512);
      GetWindowText(hwnd, title, title.Capacity);
      GetClassName(hwnd, className, className.Capacity);
      rows.Add(new NxOwnedWindowSnapshot {
        Handle = hwnd.ToInt64(),
        ClassName = className.ToString(),
        Title = title.ToString()
      });
      return true;
    }, IntPtr.Zero);
    return rows.ToArray();
  }
  public static NxOwnedWindowSnapshot[] SnapshotChildWindows(long parentHandle) {
    var rows = new List<NxOwnedWindowSnapshot>();
    EnumChildWindows(new IntPtr(parentHandle), (hwnd, parameter) => {
      var title = new StringBuilder(2048);
      var className = new StringBuilder(512);
      GetWindowText(hwnd, title, title.Capacity);
      GetClassName(hwnd, className, className.Capacity);
      if (title.Length > 0) {
        rows.Add(new NxOwnedWindowSnapshot {
          Handle = hwnd.ToInt64(),
          ClassName = className.ToString(),
          Title = title.ToString()
        });
      }
      return true;
    }, IntPtr.Zero);
    return rows.ToArray();
  }
}
"@
}

function ConvertFrom-CodePoints([int[]]$CodePoints) {
    return -join @($CodePoints | ForEach-Object { [char]$_ })
}

function Find-UiElementWithin([object]$Root, [object]$Condition, [int]$TimeoutMs) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    do {
        $element = $Root.FindFirst([Windows.Automation.TreeScope]::Descendants, $Condition)
        if ($null -ne $element) { return $element }
        Start-Sleep -Milliseconds 100
    } while ($timer.ElapsedMilliseconds -lt $TimeoutMs)
    return $null
}

function Find-UiElementByName([object]$Root, [string]$Name, [int]$TimeoutMs) {
    $condition = [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::NameProperty,
        $Name
    )
    return Find-UiElementWithin $Root $condition $TimeoutMs
}

function Find-UiElementByNameAndProcessId([object]$Root, [string]$Name, [int]$ProcessId, [int]$TimeoutMs) {
    $conditions = [Windows.Automation.Condition[]]@(
        [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::NameProperty,
            $Name
        ),
        [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::ProcessIdProperty,
            $ProcessId
        )
    )
    $condition = [Windows.Automation.AndCondition]::new($conditions)
    return Find-UiElementWithin $Root $condition $TimeoutMs
}

function Get-UiaDescendantsWithRetry(
    [object]$Root,
    [object]$Condition,
    [object]$Scope = [Windows.Automation.TreeScope]::Descendants
) {
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        try {
            return @($Root.FindAll($Scope, $Condition))
        } catch {
            if ($attempt -eq 10) { throw }
            Start-Sleep -Milliseconds 100
        }
    }
}

function Get-UiElementDiagnosticSnapshot([object]$Root, [string]$Name, [int]$ProcessId) {
    $conditions = [Windows.Automation.Condition[]]@(
        [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::NameProperty,
            $Name
        ),
        [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::ProcessIdProperty,
            $ProcessId
        )
    )
    $condition = [Windows.Automation.AndCondition]::new($conditions)
    $rows = @()
    foreach ($element in @(Get-UiaDescendantsWithRetry $Root $condition)) {
        $bounds = $element.Current.BoundingRectangle
        $ancestors = @()
        $current = $element
        for ($depth = 0; $depth -lt 8 -and $null -ne $current; $depth++) {
            try { $current = [Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($current) }
            catch { $current = $null }
            if ($null -ne $current) {
                $ancestors += [ordered]@{
                    name=[string]$current.Current.Name
                    automation_id=[string]$current.Current.AutomationId
                    control_type=[string]$current.Current.ControlType.ProgrammaticName
                    class_name=[string]$current.Current.ClassName
                }
            }
        }
        $rows += [pscustomobject][ordered]@{
            name=[string]$element.Current.Name
            automation_id=[string]$element.Current.AutomationId
            control_type=[string]$element.Current.ControlType.ProgrammaticName
            class_name=[string]$element.Current.ClassName
            enabled=[bool]$element.Current.IsEnabled
            offscreen=[bool]$element.Current.IsOffscreen
            bounds=[ordered]@{left=$bounds.Left;top=$bounds.Top;width=$bounds.Width;height=$bounds.Height}
            ancestors=$ancestors
        }
    }
    return $rows
}

function Find-OwnedTopLevelWindowElement([string]$Caption, [int]$ProcessId, [int]$TimeoutMs) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    do {
        $handle = [NxProductWindowApi]::FindVisibleTopLevelWindow($Caption, [uint32]$ProcessId)
        if ($handle -ne [IntPtr]::Zero) {
            $element = [Windows.Automation.AutomationElement]::FromHandle($handle)
            if ($null -ne $element) { return $element }
        }
        Start-Sleep -Milliseconds 100
    } while ($timer.ElapsedMilliseconds -lt $TimeoutMs)
    return $null
}

function Get-OwnedVisibleWindowSnapshot([int]$ProcessId) {
    $result = @()
    foreach ($row in @([NxProductWindowApi]::SnapshotVisibleTopLevelWindows([uint32]$ProcessId))) {
        $titleText = [string]$row.Title
        $result += [pscustomobject][ordered]@{
            native_handle=[Int64]$row.Handle
            window_class=[string]$row.ClassName
            title_length=$titleText.Length
            title_sha256=Get-TextSha256 $titleText
            title_code_points=@($titleText.ToCharArray() | ForEach-Object { [int][char]$_ })
            child_windows=@([NxProductWindowApi]::SnapshotChildWindows([Int64]$row.Handle) | ForEach-Object {
                $childText = [string]$_.Title
                [pscustomobject][ordered]@{
                    native_handle=[Int64]$_.Handle
                    window_class=[string]$_.ClassName
                    text_length=$childText.Length
                    text_sha256=Get-TextSha256 $childText
                    text_code_points=@($childText.ToCharArray() | ForEach-Object { [int][char]$_ })
                }
            })
        }
    }
    return $result
}

function Wait-OwnedTopLevelWindowGone([Int64]$Handle, [int]$TimeoutMs) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    do {
        $window = [IntPtr]$Handle
        if (-not [NxProductWindowApi]::IsWindow($window) -or -not [NxProductWindowApi]::IsWindowVisible($window)) { return $true }
        Start-Sleep -Milliseconds 100
    } while ($timer.ElapsedMilliseconds -lt $TimeoutMs)
    return $false
}

function Close-OwnedTopLevelWindow([Int64]$Handle, [string]$Label) {
    $handlePointer = [IntPtr]$Handle
    if ($Handle -eq 0 -or -not [NxProductWindowApi]::IsWindow($handlePointer)) {
        throw ('Owned top-level window is unavailable for close: ' + $Label)
    }
    if (-not [NxProductWindowApi]::PostMessage($handlePointer, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)) {
        throw ('Owned top-level window WM_CLOSE post failed: ' + $Label)
    }
}

function Invoke-UiElement([object]$Element, [string]$Label) {
    if ($null -eq $Element) { throw ('UI element is unavailable: ' + $Label) }
    try {
        $pattern = [Windows.Automation.InvokePattern]$Element.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern)
        $pattern.Invoke()
    } catch { throw ('InvokePattern failed for ' + $Label + ': ' + $_.Exception.Message) }
}

function Expand-UiElement([object]$Element, [string]$Label) {
    if ($null -eq $Element -or -not [bool]$Element.Current.IsEnabled) { throw ('UI menu is unavailable or disabled: ' + $Label) }
    try {
        $pattern = [Windows.Automation.ExpandCollapsePattern]$Element.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern)
        $pattern.Expand()
    } catch {
        try {
            $fallback = [Windows.Automation.InvokePattern]$Element.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern)
            $fallback.Invoke()
        } catch { throw ('Ribbon menu expansion failed for ' + $Label + ': ' + $_.Exception.Message) }
    }
}

function Find-ExpandableUiElementByName([object]$Root, [string]$Name, [int]$TimeoutMs) {
    $condition = [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::NameProperty,
        $Name
    )
    $timer = [Diagnostics.Stopwatch]::StartNew()
    do {
        $elements = @(Get-UiaDescendantsWithRetry $Root $condition)
        foreach ($element in $elements) {
            $pattern = $null
            if ([bool]$element.Current.IsEnabled -and $element.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)) {
                return $element
            }
        }
        Start-Sleep -Milliseconds 100
    } while ($timer.ElapsedMilliseconds -lt $TimeoutMs)
    return $null
}

function Find-VisibleUiElementByNameAndProcessId([object]$Root, [string]$Name, [int]$ProcessId, [int]$TimeoutMs) {
    $conditions = [Windows.Automation.Condition[]]@(
        [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::NameProperty,
            $Name
        ),
        [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::ProcessIdProperty,
            $ProcessId
        )
    )
    $condition = [Windows.Automation.AndCondition]::new($conditions)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    do {
        foreach ($element in @(Get-UiaDescendantsWithRetry $Root $condition)) {
            $bounds = $element.Current.BoundingRectangle
            if (-not [bool]$element.Current.IsOffscreen -and $bounds.Width -gt 0 -and $bounds.Height -gt 0) {
                return $element
            }
        }
        Start-Sleep -Milliseconds 100
    } while ($timer.ElapsedMilliseconds -lt $TimeoutMs)
    return $null
}

function Find-CollapsedExpandableUiElementByNameAndProcessId([object]$Root, [string]$Name, [int]$ProcessId, [int]$TimeoutMs) {
    $conditions = [Windows.Automation.Condition[]]@(
        [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::NameProperty,
            $Name
        ),
        [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::ProcessIdProperty,
            $ProcessId
        )
    )
    $condition = [Windows.Automation.AndCondition]::new($conditions)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    do {
        foreach ($element in @(Get-UiaDescendantsWithRetry $Root $condition)) {
            $bounds = $element.Current.BoundingRectangle
            $pattern = $null
            if (
                [bool]$element.Current.IsEnabled -and
                -not [bool]$element.Current.IsOffscreen -and
                $bounds.Width -gt 0 -and
                $bounds.Height -gt 0 -and
                $element.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern) -and
                $pattern.Current.ExpandCollapseState -eq [Windows.Automation.ExpandCollapseState]::Collapsed
            ) {
                return $element
            }
        }
        Start-Sleep -Milliseconds 100
    } while ($timer.ElapsedMilliseconds -lt $TimeoutMs)
    return $null
}

function Open-ProductRibbonMenuPath(
    [object]$Root,
    [object]$Desktop,
    [string]$MenuName,
    [string]$TargetName,
    [int]$ProcessId
) {
    for ($depth = 0; $depth -lt 3; $depth++) {
        $target = Find-VisibleUiElementByNameAndProcessId $Desktop $TargetName $ProcessId 250
        if ($null -ne $target) { return $target }
        $menu = Find-CollapsedExpandableUiElementByNameAndProcessId $Root $MenuName $ProcessId 3000
        if ($null -eq $menu) { break }
        Expand-UiElement $menu $MenuName
        Start-Sleep -Milliseconds 300
    }
    return Find-VisibleUiElementByNameAndProcessId $Desktop $TargetName $ProcessId 1000
}

function Get-WorkbookFingerprint([object]$Book) {
    $worksheets = $null
    try {
        $worksheets = $Book.Worksheets
        $worksheetCount = [int]$worksheets.Count
        $rows = @('workbook=' + [string]$Book.Name, 'worksheet_count=' + $worksheetCount)
        for ($sheetIndex = 1; $sheetIndex -le $worksheetCount; $sheetIndex++) {
            $sheet = $null; $cells = $null
            try {
                $sheet = $worksheets.Item($sheetIndex)
                $rows += ('sheet=' + [string]$sheet.Name)
                $cells = $sheet.Cells
                for ($row = 1; $row -le 4; $row++) {
                    for ($column = 1; $column -le 2; $column++) {
                        $cell = $null
                        try {
                            $cell = $cells.Item($row,$column)
                            $rows += ($row.ToString() + ',' + $column.ToString() + '|value=' + [Convert]::ToString($cell.Value2, [Globalization.CultureInfo]::InvariantCulture) + '|formula=' + [string]$cell.Formula)
                        } finally { Release-ComObject $cell }
                    }
                }
            } finally { Release-ComObject $cells; Release-ComObject $sheet }
        }
        return Get-TextSha256 (($rows -join "`n") + "`n")
    } finally { Release-ComObject $worksheets }
}

function Select-ProductRibbonTab([object]$Root) {
    $tabLabel = ConvertFrom-CodePoints @(0xB0B4,0xC5D1,0xC140)
    $tab = Find-UiElementByName $Root $tabLabel 15000
    if ($null -eq $tab -or $tab.Current.ControlType -ne [Windows.Automation.ControlType]::TabItem) { throw 'LHexcel Ribbon tab identity rejected' }
    try { $tab.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select() } catch { throw 'LHexcel Ribbon tab selection failed' }
    Start-Sleep -Milliseconds 400
}

function Select-HomeRibbonTab([object]$Root) {
    $tabLabel = ConvertFrom-CodePoints @(0xD648)
    $condition = [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::NameProperty,
        $tabLabel
    )
    $matches = @(Get-UiaDescendantsWithRetry $Root $condition)
    $tabs = @($matches | Where-Object { $_.Current.ControlType -eq [Windows.Automation.ControlType]::TabItem })
    if ($tabs.Count -ne 1) { throw 'Excel Home Ribbon tab identity rejected' }
    try { $tabs[0].GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select() } catch { throw 'Excel Home Ribbon tab selection failed' }
    Start-Sleep -Milliseconds 400
}

function Get-VisibleRibbonElementByName([object]$Root, [string]$Name) {
    $condition = [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::NameProperty,
        $Name
    )
    $matches = @(Get-UiaDescendantsWithRetry $Root $condition)
    $visible = @($matches | Where-Object {
        $bounds = $_.Current.BoundingRectangle
        $_.Current.ControlType -ne [Windows.Automation.ControlType]::TabItem -and $bounds.Width -gt 0 -and $bounds.Height -gt 0
    })
    if ($visible.Count -lt 1) { throw ('Visible Ribbon element missing: ' + $Name) }
    return $visible[0]
}

function Test-VisibleUiLabels([object]$Root, [int]$ProcessId, [string[]]$ExpectedLabels) {
    foreach ($label in $ExpectedLabels) {
        $condition = [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::NameProperty,
            $label
        )
        $matches = @(Get-UiaDescendantsWithRetry $Root $condition | Where-Object {
            $bounds = $_.Current.BoundingRectangle
            [int]$_.Current.ProcessId -eq $ProcessId -and
                $_.Current.ControlType -eq [Windows.Automation.ControlType]::MenuItem -and
                $bounds.Width -gt 0 -and $bounds.Height -gt 0
        })
        if ($matches.Count -eq 0) { return $false }
    }
    return $true
}

function Get-UiParentRibbonGroup([object]$Element) {
    $current = $Element
    for ($depth = 0; $depth -lt 5 -and $null -ne $current; $depth++) {
        try { $current = [Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($current) }
        catch { return $null }
        if ($null -ne $current -and $current.Current.ControlType -eq [Windows.Automation.ControlType]::Group) {
            return $current
        }
    }
    return $null
}

function Find-ProductHomeRibbonGroup(
    [object]$Root,
    [object]$Desktop,
    [int]$ProcessId,
    [string]$MenuName,
    [string[]]$ExpectedLabels
) {
    $timer = [Diagnostics.Stopwatch]::StartNew()
    do {
        $condition = [Windows.Automation.PropertyCondition]::new(
            [Windows.Automation.AutomationElement]::NameProperty,
            $MenuName
        )
        $candidates = @(Get-UiaDescendantsWithRetry $Root $condition | Where-Object {
            $bounds = $_.Current.BoundingRectangle
            $_.Current.ControlType -ne [Windows.Automation.ControlType]::TabItem -and
                $bounds.Width -gt 0 -and $bounds.Height -gt 0
        })
        foreach ($candidate in $candidates) {
            try {
                $pattern = $candidate.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern)
                $pattern.Expand()
            } catch { continue }
            $childrenTimer = [Diagnostics.Stopwatch]::StartNew()
            do {
                if (Test-VisibleUiLabels $Desktop $ProcessId $ExpectedLabels) {
                    $group = Get-UiParentRibbonGroup $candidate
                    [Windows.Forms.SendKeys]::SendWait('{ESC}')
                    Start-Sleep -Milliseconds 150
                    if ($null -eq $group) { throw 'Candidate Home Ribbon group identity unavailable' }
                    return $group
                }
                Start-Sleep -Milliseconds 100
            } while ($childrenTimer.ElapsedMilliseconds -lt 2000)
            [Windows.Forms.SendKeys]::SendWait('{ESC}')
            Start-Sleep -Milliseconds 150
        }
        Start-Sleep -Milliseconds 100
    } while ($timer.ElapsedMilliseconds -lt 15000)
    throw 'Candidate Home Ribbon management menu unavailable'
}

function Wait-UiElementGone([object]$Root, [string]$Name, [int]$ProcessId, [int]$TimeoutMs) {
    $conditions = [Windows.Automation.Condition[]]@(
        [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty, $Name),
        [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty, $ProcessId)
    )
    $condition = [Windows.Automation.AndCondition]::new($conditions)
    $timer = [Diagnostics.Stopwatch]::StartNew()
    do {
        if ($null -eq $Root.FindFirst([Windows.Automation.TreeScope]::Descendants, $condition)) { return $true }
        Start-Sleep -Milliseconds 100
    } while ($timer.ElapsedMilliseconds -lt $TimeoutMs)
    return $false
}

function Get-ActualProductRibbonXml([string]$Artifact) {
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $archive = [IO.Compression.ZipFile]::OpenRead($Artifact)
    try {
        $entries = @($archive.Entries | Where-Object { $_.FullName -ceq 'customUI/customUI14.xml' })
        if ($entries.Count -ne 1) { throw 'Actual Product ribbon XML part identity rejected' }
        $reader = New-Object IO.StreamReader($entries[0].Open(), [Text.Encoding]::UTF8)
        try { [xml]$document = $reader.ReadToEnd() } finally { $reader.Dispose() }
        return ,$document
    } finally { $archive.Dispose() }
}

function Get-ActualProductRibbonLabels([xml]$Ribbon, [string]$FeatureId, [string]$MenuId) {
    $tabs = @($Ribbon.SelectNodes('//*[local-name()="tab" and @id="NX-TAB"]'))
    if ($tabs.Count -ne 1) { throw 'Actual Product ribbon tab identity rejected' }
    $menus = @($tabs[0].SelectNodes('.//*[local-name()="menu"]') | Where-Object { $_.GetAttribute('id') -ceq $MenuId })
    if ($menus.Count -ne 1) { throw ('Actual Product ribbon menu identity rejected: ' + $MenuId) }
    $featureTag = 'nx1|feature|' + $FeatureId
    $buttons = @($menus[0].SelectNodes('.//*[local-name()="button"]') | Where-Object { $_.GetAttribute('tag') -ceq $featureTag })
    if ($buttons.Count -ne 1) { throw ('Actual Product ribbon feature identity rejected: ' + $FeatureId) }
    $ribbonLabel = [string]$buttons[0].GetAttribute('label')
    $menuLabel = [string]$menus[0].GetAttribute('label')
    if ([string]::IsNullOrWhiteSpace($ribbonLabel) -or [string]::IsNullOrWhiteSpace($menuLabel)) {
        throw ('Actual Product ribbon label metadata missing: ' + $FeatureId)
    }
    return [pscustomobject]@{ribbon=$ribbonLabel;menu=$menuLabel}
}

function Select-WorkbookRange([object]$Book, [string]$Address) {
    $worksheets = $null; $sheet = $null; $range = $null
    try {
        [void]$Book.Activate()
        $worksheets = $Book.Worksheets
        $sheet = $worksheets.Item(1)
        [void]$sheet.Activate()
        $range = $sheet.Range($Address)
        [void]$range.Select()
    } finally {
        Release-ComObject $range
        Release-ComObject $sheet
        Release-ComObject $worksheets
    }
}

function Invoke-ActualProductRibbonCases([object]$Excel, [object]$UiHostBook, [object]$ExcelBinding, [string]$Artifact) {
    $script:probeStage = 'actual-ribbon:setup'
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    Add-Type -AssemblyName System.Windows.Forms
    Initialize-WindowApi
    [void](Assert-ExactExcelProcessOwnership $ExcelBinding)
    $artifactShaBefore = Get-HexDigest $Artifact
    $scratchBook = $null; $scratchWorksheets = $null; $sheet = $null; $scratchRange = $null
    try {
        if ($null -eq $UiHostBook -or [bool]$UiHostBook.IsAddin) {
            throw 'Actual Product ribbon host workbook rejected'
        }
        # Reuse the suite-owned visible workbook. Adding a second workbook creates a
        # second Excel SDI window and makes process-only UIA menu lookup ambiguous.
        $scratchBook = $UiHostBook
        try {
            $scratchWorksheets = $scratchBook.Worksheets
            $sheet = $scratchWorksheets.Item(1)
            $scratchRange = $sheet.Range('A1:B4')
            $scratchValues = New-Object 'object[,]' 4,2
            $scratchValues[0,0] = 'Key'; $scratchValues[0,1] = 'Value'
            $scratchValues[1,0] = 'A'; $scratchValues[1,1] = [double]10
            $scratchValues[2,0] = 'B'; $scratchValues[2,1] = [double]20
            $scratchValues[3,0] = 'C'; $scratchValues[3,1] = [double]30
            $scratchRange.Value2 = $scratchValues
            [void]$scratchRange.Select()
        } finally {
            Release-ComObject $scratchRange; $scratchRange = $null
            Release-ComObject $sheet; $sheet = $null
            Release-ComObject $scratchWorksheets; $scratchWorksheets = $null
        }
        $excelRoot = [Windows.Automation.AutomationElement]::FromHandle([IntPtr][Int64]$Excel.Hwnd)
        $desktop = [Windows.Automation.AutomationElement]::RootElement
        $ribbonXml = Get-ActualProductRibbonXml $Artifact
        $formCaptionPrefix = ConvertFrom-CodePoints @(0xB0B4,0xC5D1,0xC140,0x20,0x2D,0x20)
        $cases = @(
            [pscustomobject]@{feature_id='NX-DATA-UNIQUE-COUNT';menu_id='NX-DATA-MENU';grade='Fast';resolved=0;form_name='FNxDataAnalyze'},
            [pscustomobject]@{feature_id='NX-DATA-DUPLICATE-LIST';menu_id='NX-DATA-MENU';grade='Guarded';resolved=1;form_name='FNxDataAnalyze'},
            [pscustomobject]@{feature_id='NX-FILE-FOLDER-CREATE';menu_id='NX-FILE-MENU';grade='Planned';resolved=2;form_name='FNxFolderCreate'}
        )
        $results = @()
        foreach ($case in $cases) {
            $caseId = [string]$case.feature_id
            $script:probeStage = 'actual-ribbon:' + $caseId + ':activate'
            Select-WorkbookRange $scratchBook 'A1:B4'
            $script:probeStage = 'actual-ribbon:' + $caseId + ':fingerprint-before'
            $before = Get-WorkbookFingerprint $scratchBook
            $script:probeStage = 'actual-ribbon:' + $caseId + ':select-tab'
            Select-ProductRibbonTab $excelRoot
            $labels = Get-ActualProductRibbonLabels $ribbonXml $caseId ([string]$case.menu_id)
            $ribbonName = [string]$labels.ribbon
            if ($null -ne $case.menu_id) {
                $script:probeStage = 'actual-ribbon:' + $caseId + ':expand-menu'
                $menuName = [string]$labels.menu
                $ribbonElement = Open-ProductRibbonMenuPath $excelRoot $desktop $menuName $ribbonName ([int]$ExcelBinding.pid)
            } else {
                $ribbonElement = Find-VisibleUiElementByNameAndProcessId $desktop $ribbonName ([int]$ExcelBinding.pid) 10000
            }
            $script:probeStage = 'actual-ribbon:' + $caseId + ':find-control'
            if ($null -eq $ribbonElement) {
                $script:windowDiagnosticContext = [ordered]@{
                    target_name=$ribbonName
                    target_candidates=@(Get-UiElementDiagnosticSnapshot $desktop $ribbonName ([int]$ExcelBinding.pid))
                    menu_name=[string]$labels.menu
                    menu_candidates=@(Get-UiElementDiagnosticSnapshot $desktop ([string]$labels.menu) ([int]$ExcelBinding.pid))
                }
                throw ('Actual Product ribbon control ownership rejected: ' + $case.feature_id)
            }
            Invoke-UiElement $ribbonElement $case.feature_id
            $script:probeStage = 'actual-ribbon:' + $caseId + ':find-form'
            $formCaption = $formCaptionPrefix + $ribbonName
            $form = Find-OwnedTopLevelWindowElement $formCaption ([int]$ExcelBinding.pid) 15000
            if ($null -eq $form) {
                $uiaCandidate = Find-UiElementByNameAndProcessId $desktop $formCaption ([int]$ExcelBinding.pid) 1000
                $uiaCandidateContext = if ($null -eq $uiaCandidate) {
                    $null
                } else {
                    [ordered]@{
                        control_type=[string]$uiaCandidate.Current.ControlType.ProgrammaticName
                        native_handle=[Int64]$uiaCandidate.Current.NativeWindowHandle
                        process_id=[int]$uiaCandidate.Current.ProcessId
                        name_sha256=Get-TextSha256 ([string]$uiaCandidate.Current.Name)
                        name_code_points=@(([string]$uiaCandidate.Current.Name).ToCharArray() | ForEach-Object { [int][char]$_ })
                    }
                }
                $script:windowDiagnosticContext = [ordered]@{
                    expected_process_id=[int]$ExcelBinding.pid
                    expected_caption_sha256=Get-TextSha256 $formCaption
                    expected_caption_code_points=@($formCaption.ToCharArray() | ForEach-Object { [int][char]$_ })
                    top_level_windows=@(Get-OwnedVisibleWindowSnapshot ([int]$ExcelBinding.pid))
                    uia_candidate=$uiaCandidateContext
                }
                throw ([string]$case.form_name + ' owned top-level window unavailable: ' + $case.feature_id)
            }
            $formHandle = [Int64]$form.Current.NativeWindowHandle
            if ($form.Current.Name -cne $formCaption -or $form.Current.ProcessId -ne [int]$ExcelBinding.pid -or $form.Current.NativeWindowHandle -eq 0) {
                throw ([string]$case.form_name + ' owned top-level window identity rejected: ' + $case.feature_id)
            }
            $formControlType = [string]$form.Current.ControlType.ProgrammaticName
            $script:probeStage = 'actual-ribbon:' + $caseId + ':route-verified'
            $script:probeStage = 'actual-ribbon:' + $caseId + ':close'
            Add-Type -AssemblyName System.Windows.Forms
            $formClosed = $false
            for ($cancelAttempt = 0; $cancelAttempt -lt 3 -and -not $formClosed; $cancelAttempt++) {
                if (-not [NxProductWindowApi]::ActivateForegroundWindow([IntPtr]$formHandle)) {
                    throw ([string]$case.form_name + ' foreground activation failed: ' + $case.feature_id)
                }
                $formForeground = $false
                for ($foregroundAttempt = 0; $foregroundAttempt -lt 50; $foregroundAttempt++) {
                    if ([NxProductWindowApi]::GetForegroundWindow() -eq [IntPtr]$formHandle) { $formForeground = $true; break }
                    Start-Sleep -Milliseconds 100
                }
                if (-not $formForeground) { throw ([string]$case.form_name + ' foreground ownership rejected: ' + $case.feature_id) }
                [Windows.Forms.SendKeys]::SendWait('{ESC}')
                $formClosed = Wait-OwnedTopLevelWindowGone $formHandle 3000
            }
            if (-not $formClosed) {
                if (-not [NxProductWindowApi]::ActivateForegroundWindow([IntPtr]$formHandle)) {
                    throw ([string]$case.form_name + ' close fallback activation failed: ' + $case.feature_id)
                }
                Start-Sleep -Milliseconds 200
                [Windows.Forms.SendKeys]::SendWait('%{F4}')
                $formClosed = Wait-OwnedTopLevelWindowGone $formHandle 5000
            }
            if (-not $formClosed) { throw ([string]$case.form_name + ' did not close: ' + $case.feature_id) }
            $script:probeStage = 'actual-ribbon:' + $caseId + ':fingerprint-after'
            $after = Get-WorkbookFingerprint $scratchBook
            $artifactShaAfterCase = Get-HexDigest $Artifact
            if ($after -ne $before) { throw ('Scratch workbook changed during route/close: ' + $case.feature_id) }
            if ($artifactShaAfterCase -ne $artifactShaBefore) { throw ('Product.xlam SHA changed during actual ribbon UI case: ' + $case.feature_id) }
            $results += [pscustomobject][ordered]@{
                feature_id=$case.feature_id;callback='NxRibbonExecute';grade=$case.grade;resolved=[int]$case.resolved
                workbook_before_sha256=$before;workbook_after_sha256=$after;artifact_sha256=$artifactShaAfterCase
                route_verified=$true;form_opened=$true;form=[string]$case.form_name;form_control_type=$formControlType
                form_native_handle=$formHandle;scratch_range='A1:B4';ribbon_allowed=$true
                ribbon_caption_sha256=Get-TextSha256 $ribbonName
                route_proof='actual-ribbon-control-to-owned-form'
                close_input='owned-form-escape-or-alt-f4-cancel'
            }
        }
        if (@($results | Where-Object { $_.resolved -eq 0 }).Count -ne 1 -or @($results | Where-Object { $_.resolved -eq 1 }).Count -ne 1 -or @($results | Where-Object { $_.resolved -eq 2 }).Count -ne 1) { throw 'Actual Product ribbon grade inventory rejected' }
        return [ordered]@{method='actual Product.xlam Ribbon -> feature-specific owned form route/cancel';macro_execution='actual-ribbon-ui';artifact_sha_before=$artifactShaBefore;artifact_sha_after=Get-HexDigest $Artifact;cases=$results}
    } finally {
        Release-ComObject $scratchRange; Release-ComObject $sheet; Release-ComObject $scratchWorksheets
        # The outer suite owns and closes UiHostBook after callback attestation.
        $scratchBook = $null
    }
}

function Invoke-UiAutomationProbe([object]$Excel) {
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
    Add-Type -AssemblyName System.Windows.Forms
    Initialize-WindowApi
    $handle = [IntPtr][int64]$Excel.Hwnd
    $dpi = [int][NxProductWindowApi]::GetDpiForWindow($handle)
    if ($dpi -ne 192) { throw ('ENVIRONMENT_MISMATCH: current Excel scale must be 200 percent; measured DPI=' + $dpi) }
    $foregroundVerified = [NxProductWindowApi]::ActivateForegroundWindow($handle)
    if (-not $foregroundVerified) { throw 'Excel foreground verification failed before Ribbon UI search' }
    $root = [Windows.Automation.AutomationElement]::FromHandle($handle)
    $clientRect = New-Object NxClientRect
    if (-not [NxProductWindowApi]::GetClientRect($handle, [ref]$clientRect)) { throw 'Excel client bounds unavailable' }
    $clientOrigin = New-Object NxClientPoint
    if (-not [NxProductWindowApi]::ClientToScreen($handle, [ref]$clientOrigin)) { throw 'Excel client origin unavailable' }
    $clientWidth = [int]($clientRect.Right - $clientRect.Left)
    $clientHeight = [int]($clientRect.Bottom - $clientRect.Top)
    if ($clientWidth -le 0 -or $clientHeight -le 0) { throw 'Excel client bounds are empty' }
    Select-HomeRibbonTab $root
    $desktop = [Windows.Automation.AutomationElement]::RootElement
    [uint32]$excelProcessId = 0
    [void][NxProductWindowApi]::GetWindowThreadProcessId($handle, [ref]$excelProcessId)
    if ($excelProcessId -eq 0) { throw 'Excel process identity unavailable for Home Ribbon verification' }
    $catalogSource = Get-Content -LiteralPath (Join-Path $SourceRoot 'src/vba/ui/NxRibbonMenuCatalog.bas') -Raw -Encoding UTF8
    $managementLabels = @([regex]::Matches($catalogSource, 'NxManagementButton\("[^"]+", "([^"]+)"') | ForEach-Object { $_.Groups[1].Value })
    if ($managementLabels.Count -ne 6) { throw 'Candidate Home Ribbon management inventory rejected' }
    $homeMenuLabel = ConvertFrom-CodePoints @(0xB0B4,0xC5D1,0xC140)
    $homeGroup = Find-ProductHomeRibbonGroup $root $desktop ([int]$excelProcessId) $homeMenuLabel $managementLabels
    $homeControls = @()
    $homeGroupBounds = $homeGroup.Current.BoundingRectangle
    $homeControls += [ordered]@{label=$homeMenuLabel;control_type=[string]$homeGroup.Current.ControlType.ProgrammaticName;left=$homeGroupBounds.Left;top=$homeGroupBounds.Top;width=$homeGroupBounds.Width;height=$homeGroupBounds.Height}
    foreach ($label in @(
        (ConvertFrom-CodePoints @(0xC990,0xACA8,0xCC3E,0xAE30)),
        (ConvertFrom-CodePoints @(0xD3EC,0xCEE4,0xC2A4,0xC140)),
        (ConvertFrom-CodePoints @(0xC804,0xCCB4,0xAE30,0xB2A5))
    )) {
        $element = Get-VisibleRibbonElementByName $homeGroup $label
        $bounds = $element.Current.BoundingRectangle
        $homeControls += [ordered]@{label=$label;control_type=[string]$element.Current.ControlType.ProgrammaticName;left=$bounds.Left;top=$bounds.Top;width=$bounds.Width;height=$bounds.Height}
    }
    Select-ProductRibbonTab $root
    # Windows PowerShell 5.1 treats UTF-8 without BOM as the active ANSI code
    # page. Keep this launcher ASCII-only and construct localized UI labels at
    # runtime so Korean Office installations parse the probe deterministically.
    $expected = @(
        (ConvertFrom-CodePoints @(0xB0B4,0xC5D1,0xC140)),
        (ConvertFrom-CodePoints @(0xC800,0xC7A5)),
        (ConvertFrom-CodePoints @(0xC778,0xC1C4)),
        (ConvertFrom-CodePoints @(0xBCF5,0xBD99)),
        (ConvertFrom-CodePoints @(0xC0BD,0xC785)),
        (ConvertFrom-CodePoints @(0xD30C,0xC77C,0xAD00,0xB9AC)),
        (ConvertFrom-CodePoints @(0xB370,0xC774,0xD130)),
        (ConvertFrom-CodePoints @(0x20)),
        (ConvertFrom-CodePoints @(0x20)),
        (ConvertFrom-CodePoints @(0xC2A4,0xD0C0,0xC77C)),
        (ConvertFrom-CodePoints @(0xCD94,0xAC00,0xAE30,0xB2A5)),
        (ConvertFrom-CodePoints @(0xC815,0xBCF4,0xC9C4,0xB2E8))
    )
    $groups = @()
    foreach ($label in $expected) {
        if ([string]::IsNullOrWhiteSpace($label)) { continue }
        $condition = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty, $label)
        $element = Find-UiElementWithin $root $condition 10000
        if ($null -eq $element) { throw ('Ribbon group missing: ' + $label) }
        $bounds = $element.Current.BoundingRectangle
        if ($bounds.Width -le 0 -or $bounds.Height -le 0) { throw ('Ribbon group has empty bounds: ' + $label) }
        $groups += [ordered]@{ label=$label; control_type=[string]$element.Current.ControlType.ProgrammaticName; left=$bounds.Left; top=$bounds.Top; width=$bounds.Width; height=$bounds.Height }
    }
    [Windows.Forms.SendKeys]::SendWait('%')
    $keytips = @()
    for ($attempt = 0; $attempt -lt 20 -and $keytips.Count -lt 1; $attempt++) {
        Start-Sleep -Milliseconds 250
        $all = @(Get-UiaDescendantsWithRetry $root ([Windows.Automation.Condition]::TrueCondition))
        $keytips = @($all | ForEach-Object { [string]$_.Current.AccessKey } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    }
    [Windows.Forms.SendKeys]::SendWait('{ESC}')
    if ($keytips.Count -lt 1) { throw 'Ribbon KeyTip accessibility evidence missing' }
    $windowBounds = $root.Current.BoundingRectangle
    $script:uiProbeContext = [ordered]@{
        windows_scale_percent=200; host_dpi=$dpi; scale_changed_by_test=$false
        excel_window_bounds=[ordered]@{left=$windowBounds.Left;top=$windowBounds.Top;width=$windowBounds.Width;height=$windowBounds.Height}
        excel_client_bounds=[ordered]@{left=$clientOrigin.X;top=$clientOrigin.Y;width=$clientWidth;height=$clientHeight}
    }
    return [ordered]@{
        method='UIAutomationClient + Win32 current 200 percent environment; no scale or window mutation'
        windows_scale_percent=200;host_dpi=$dpi;scale_changed_by_test=$false
        foreground_verified=$foregroundVerified;keytip_accessible=$true;keytips=$keytips
        excel_window_bounds=[ordered]@{left=$windowBounds.Left;top=$windowBounds.Top;width=$windowBounds.Width;height=$windowBounds.Height}
        excel_client_bounds=[ordered]@{left=$clientOrigin.X;top=$clientOrigin.Y;width=$clientWidth;height=$clientHeight}
        home_controls=$homeControls;product_groups=$groups
    }
}

function Invoke-GradeProbe([object]$RibbonUiResults, [object]$CallbackResults) {
    if ($null -eq $RibbonUiResults -or @($RibbonUiResults.cases).Count -ne 3) { throw 'Actual Product ribbon grade results missing' }
    if ($null -eq $CallbackResults -or $CallbackResults.grade_policy_contract_attested -ne $true -or [int]$CallbackResults.grade_policy_pattern_count -ne 8) {
        throw 'Compiled Product grade policy attestation missing'
    }
    $fast = @($RibbonUiResults.cases | Where-Object { $_.grade -ceq 'Fast' -and [int]$_.resolved -eq 0 })
    $guarded = @($RibbonUiResults.cases | Where-Object { $_.grade -ceq 'Guarded' -and [int]$_.resolved -eq 1 })
    $planned = @($RibbonUiResults.cases | Where-Object { $_.grade -ceq 'Planned' -and [int]$_.resolved -eq 2 })
    if ($fast.Count -ne 1 -or $guarded.Count -ne 1 -or $planned.Count -ne 1) { throw 'Fast/Guarded/Planned actual ribbon UI policy probe failed' }
    foreach ($case in @($fast[0], $guarded[0], $planned[0])) {
        if ($case.route_verified -ne $true -or $case.form_opened -ne $true -or $case.route_proof -cne 'actual-ribbon-control-to-owned-form') {
            throw ('Actual Product ribbon form route evidence rejected: ' + [string]$case.feature_id)
        }
    }
    return [ordered]@{
        method='actual Product.xlam Ribbon form route + compiled Product.xlam grade policy attestation'
        Fast=$fast[0]
        Guarded=$guarded[0]
        Planned=$planned[0]
        policy_attested_from_compiled_product=$true
    }
}

function Write-ProbeEvidence([string]$Path, [string]$Name, [object]$Measured, [string]$Artifact) {
    if ($null -eq $Measured -or [string]::IsNullOrWhiteSpace([string]$Measured.method)) { throw ($Name + ' raw evidence missing') }
    Write-Json $Path ([ordered]@{
        schema_version=1; suite='ProductRibbon'; status='PASS'; evidence=$Name; method=$Measured.method; run_id=$RunId
        artifact_sha256=Get-HexDigest $Artifact
        source_tree_sha256=$SourceDigest
        source_snapshot_sha256=$SnapshotDigest
        measured=$Measured
    })
}

function Test-ListColumnWidths([string]$Actual,[string]$Expected) {
    if ($Actual -ceq $Expected) { return $true }
    $actualParts=$Actual.Split(';'); $expectedParts=$Expected.Split(';')
    if ($actualParts.Count -ne $expectedParts.Count) { return $false }
    for ($index=0; $index -lt $actualParts.Count; $index++) {
        if ($actualParts[$index] -notmatch '^\s*(\d+(?:\.\d+)?)\s+pt\s*$') { return $false }
        $actualWidth=[double]::Parse($Matches[1],[Globalization.CultureInfo]::InvariantCulture)
        if ($expectedParts[$index] -notmatch '^\s*(\d+(?:\.\d+)?)\s+pt\s*$') { return $false }
        $expectedWidth=[double]::Parse($Matches[1],[Globalization.CultureInfo]::InvariantCulture)
        # MSForms persists widths at twip precision: observed 176 -> 175.95 pt.
        if ([Math]::Abs($actualWidth-$expectedWidth) -gt 0.051) { return $false }
        if ($expectedWidth -eq 0 -and $actualWidth -ne 0) { return $false }
    }
    return $true
}
function Invoke-ProductUserFormAttestation([object]$Books, [string]$Artifact, [string]$ManifestPath, [object]$ExistingProject = $null) {
    # Build-Xlam has already saved the artifact. Reopen that closed artifact
    # read-only, then inspect the persisted Designer.Controls surface.
    $reopened = $null
    $project = $null
    $releaseProject = $false
    if ($null -ne $ExistingProject) {
        if (-not [string]::Equals([IO.Path]::GetFullPath([string]$ExistingProject.FileName), [IO.Path]::GetFullPath($Artifact), [StringComparison]::OrdinalIgnoreCase)) { throw 'Product UserForm startup project identity rejected' }
        $project = $ExistingProject
    } else {
        if ($null -eq $Books) { throw 'Product UserForm attestation workbook unavailable' }
        $reopened = $Books.Open($Artifact, $false, $true)
    }
    try {
        if ($null -ne $reopened) {
            if (-not [bool]$reopened.IsAddin -or [IO.Path]::GetFullPath([string]$reopened.FullName) -ne [IO.Path]::GetFullPath($Artifact)) { throw 'Product UserForm reopen identity rejected' }
            $project = $reopened.VBProject
            $releaseProject = $true
        }
        $components = $null; $manifest = $null; $forms = @(); $bindings = @()
        $expectedControlCount = 0; $expectedTextBoxCount = 0; $expectedListBoxCount = 0
        try {
            $manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
            $bindings = @($manifest.form_bindings)
            if ($bindings.Count -eq 0) { throw 'Product UserForm form_bindings count rejected' }
            $components = $project.VBComponents
            foreach ($binding in $bindings) {
                $layoutPath = Join-Path $SourceRoot ([string]$binding.layout); $layout = Get-Content -LiteralPath $layoutPath -Raw -Encoding UTF8 | ConvertFrom-Json
                $component = $null; $designer = $null; $controls = $null; $rows = @()
                try {
                    $component = $components.Item([string]$layout.name); $designer = $component.Designer; $controls = $designer.Controls
                    $expectedControls=@($layout.controls)
                    $expectedControlCount += $expectedControls.Count
                    $expectedTextBoxCount += @($expectedControls | Where-Object {$_.type -eq 'TextBox'}).Count
                    $expectedListBoxCount += @($expectedControls | Where-Object {$_.type -eq 'ListBox'}).Count
                    if([int]$controls.Count -ne $expectedControls.Count){throw ('Product UserForm control count rejected: '+[string]$layout.name)}
                    $actualNames=@(); for ($i = 0; $i -lt [int]$controls.Count; $i++) {
                        $control = $null
                        try {
                            $control = $controls.Item($i); $actualNames += [string]$control.Name
                            $expected=@($expectedControls | Where-Object {$_.name -eq [string]$control.Name}); if($expected.Count -ne 1){throw ('Product UserForm unexpected control: '+[string]$control.Name)}; $type=[string]$expected[0].type
                            $controlTypeMap = @{ILabelControl='Label';IMdcCombo='ComboBox';IMdcText='TextBox';ICommandButton='CommandButton';IMdcCheckBox='CheckBox';IMdcList='ListBox'}
                            $comTypeName = [Microsoft.VisualBasic.Information]::TypeName($control)
                            if (-not $controlTypeMap.ContainsKey($comTypeName)) { throw ('Product UserForm unsupported COM control type: '+[string]$control.Name+' ('+$comTypeName+')') }
                            $actualType = [string]$controlTypeMap[$comTypeName]
                            if($actualType -ne $type){throw ('Product UserForm control type mismatch: '+[string]$control.Name)}
                            $row = [ordered]@{name=[string]$control.Name;type=$type;tab_index=[int]$control.TabIndex;tab_stop=[bool]$control.TabStop;enabled=[bool]$control.Enabled}
                            if ($type -eq 'TextBox') {
                                $expectedValue=if($expected[0].PSObject.Properties.Name -contains 'value'){[string]$expected[0].value}elseif($expected[0].PSObject.Properties.Name -contains 'text'){[string]$expected[0].text}else{''}
                                $expectedLocked=if($expected[0].PSObject.Properties.Name -contains 'locked'){[bool]$expected[0].locked}else{$false}
                                $expectedMultiLine=if($expected[0].PSObject.Properties.Name -contains 'multi_line'){[bool]$expected[0].multi_line}else{$false}
                                $row.value=[string]$control.Value; $row.locked=[bool]$control.Locked; $row.multi_line=[bool]$control.MultiLine
                                if($row.value -ne $expectedValue -or $row.locked -ne $expectedLocked -or $row.multi_line -ne $expectedMultiLine){throw ('Product UserForm TextBox mismatch: '+[string]$control.Name)}
                            }
                            elseif ($type -eq 'ListBox') {
                                $row.integral_height=[bool]$control.IntegralHeight
                                if ($expected[0].PSObject.Properties.Name -contains 'integral_height') {
                                    if ($expected[0].integral_height -isnot [bool] -or $row.integral_height -ne $expected[0].integral_height) { throw ('Product UserForm ListBox IntegralHeight mismatch: '+[string]$control.Name) }
                                    $row.width=[double]$control.Width; $row.height=[double]$control.Height
                                    if ([Math]::Abs($row.width-[double]$expected[0].width) -gt 0.051 -or [Math]::Abs($row.height-[double]$expected[0].height) -gt 0.051) { throw ('Product UserForm fixed ListBox dimensions mismatch: '+($row|ConvertTo-Json -Compress)) }
                                }
                                $items=@(); for($j=0;$j -lt [int]$control.ListCount;$j++){$items += [string]$control.List($j)}
                                $row.column_count=[int]$control.ColumnCount; $row.column_widths=[string]$control.ColumnWidths; $row.multi_select=([int]$control.MultiSelect); $row.items_sha256=(Get-TextSha256 (($items -join "`n") + "`n")); $row.list_index=[int]$control.ListIndex
                                $expectedColumns=if($expected[0].PSObject.Properties.Name -contains 'column_count'){[int]$expected[0].column_count}else{1}
                                $expectedWidths=if($expected[0].PSObject.Properties.Name -contains 'column_widths'){[string]$expected[0].column_widths}else{''}
                                $expectedMulti=if($expected[0].PSObject.Properties.Name -contains 'multi_select' -and [bool]$expected[0].multi_select){1}else{0}
                                $expectedItems=@()
                                if($expected[0].PSObject.Properties.Name -contains 'items'){$expectedItems=@($expected[0].items|ForEach-Object {[string]$_})}
                                $expectedIndex=if($expected[0].PSObject.Properties.Name -contains 'selected_index'){[int]$expected[0].selected_index}else{-1}
                                if($row.column_count -ne $expectedColumns -or -not (Test-ListColumnWidths $row.column_widths $expectedWidths) -or $row.multi_select -ne $expectedMulti){throw ('Product UserForm ListBox mismatch: '+[string]$control.Name+'; actual='+($row|ConvertTo-Json -Compress)+'; expected columns='+$expectedColumns+'; widths='+$expectedWidths+'; multi='+$expectedMulti)}
                                if($expectedItems.Count -gt 0){
                                    if(@($expectedItems|Where-Object {$_.Contains('|') -or $_.Contains([string][char]31)}).Count -ne 0){throw ('Product UserForm ListBox metadata separator rejected: '+[string]$control.Name)}
                                    $expectedListTag='nxlist1|'+[string]$expectedIndex+'|'+($expectedItems -join [char]31)
                                    $actualListTag=[string]$control.Tag
                                    $row.item_count=$expectedItems.Count; $row.selected_index=$expectedIndex; $row.list_tag_sha256=Get-TextSha256 $actualListTag
                                    if($actualListTag -cne $expectedListTag -or $items.Count -ne 0 -or $row.list_index -ne -1){throw ('Product UserForm ListBox metadata mismatch: '+[string]$control.Name)}
                                } elseif($row.items_sha256 -ne (Get-TextSha256 "`n") -or $row.list_index -ne -1){throw ('Product UserForm ListBox mismatch: '+[string]$control.Name)}
                            }
                            elseif ($type -eq 'ComboBox') {
                                $expectedComboItems=@($expected[0].items|ForEach-Object {[string]$_})
                                if(@($expectedComboItems|Where-Object {$_.Contains('|') -or $_.Contains([string][char]31)}).Count -ne 0){throw ('Product UserForm ComboBox metadata separator rejected: '+[string]$control.Name)}
                                $expectedComboTag='nxcombo1|'+[string]([int]$expected[0].selected_index)+'|'+($expectedComboItems -join [char]31)
                                $actualComboTag=[string]$control.Tag
                                $row.style=[int]$control.Style; $row.match_required=[bool]$control.MatchRequired; $row.item_count=$expectedComboItems.Count; $row.selected_index=[int]$expected[0].selected_index; $row.combo_tag_sha256=Get-TextSha256 $actualComboTag
                                if($row.style -ne 2 -or -not $row.match_required -or $actualComboTag -cne $expectedComboTag){throw ('Product UserForm ComboBox metadata mismatch: '+[string]$control.Name)}
                            }
                            if($row.tab_stop -ne [bool]$expected[0].tab_stop -or $row.enabled -ne [bool]$expected[0].enabled){throw ('Product UserForm tab/enabled mismatch: '+[string]$control.Name)}
                            $rows += [pscustomobject]$row
                        } finally { Release-ComObject $control }
                    }
                    if(@($actualNames|Sort-Object -Unique).Count -ne $expectedControls.Count){throw ('Product UserForm control names rejected: '+[string]$layout.name)}
                    $tabs=@($rows | Where-Object {$_.tab_stop} | ForEach-Object {[int]$_.tab_index}); if($tabs.Count -ne (@($tabs|Sort-Object -Unique).Count)){throw ('Product UserForm tab order rejected: '+[string]$layout.name)}
                    $expectedTabOrder=@($expectedControls | Where-Object {[bool]$_.tab_stop} | Sort-Object {[int]$_.tab_index} | ForEach-Object {[string]$_.name})
                    $actualTabOrder=@($rows | Where-Object {$_.tab_stop} | Sort-Object {[int]$_.tab_index} | ForEach-Object {[string]$_.name})
                    if(($actualTabOrder -join [char]31) -cne ($expectedTabOrder -join [char]31)){throw ('Product UserForm tab order rejected: '+[string]$layout.name)}
                    if($null -ne $layout.default_control){$defaultReadback=$null; try{$defaultReadback=$controls.Item([string]$layout.default_control); $defaultExpected=@($expectedControls|Where-Object {$_.name -eq [string]$layout.default_control}); if($defaultExpected.Count -ne 1 -or [string]$defaultExpected[0].type -ne 'CommandButton' -or -not [bool]$defaultReadback.Default){throw ('Product UserForm default control rejected: '+[string]$layout.name)}} finally {Release-ComObject $defaultReadback}}
                    if($null -ne $layout.cancel_control){$cancelReadback=$null; try{$cancelReadback=$controls.Item([string]$layout.cancel_control); $cancelExpected=@($expectedControls|Where-Object {$_.name -eq [string]$layout.cancel_control}); if($cancelExpected.Count -ne 1 -or [string]$cancelExpected[0].type -ne 'CommandButton' -or -not [bool]$cancelReadback.Cancel){throw ('Product UserForm cancel control rejected: '+[string]$layout.name)}} finally {Release-ComObject $cancelReadback}}
                    $forms += [pscustomobject][ordered]@{name=[string]$layout.name;layout=[string]$binding.layout;code=[string]$binding.code;layout_sha256=Get-HexDigest $layoutPath;code_sha256=Get-HexDigest (Join-Path $SourceRoot ([string]$binding.code));default_control=if($null -eq $layout.default_control){$null}else{[string]$layout.default_control};cancel_control=if($null -eq $layout.cancel_control){$null}else{[string]$layout.cancel_control};controls=@($rows)}
                } finally { Release-ComObject $controls; Release-ComObject $designer; Release-ComObject $component }
            }
        } finally {
            Release-ComObject $components
            if ($releaseProject) { Release-ComObject $project }
        }
        $all=@($forms | ForEach-Object {@($_.controls)}); $textboxes=@($all|Where-Object {$_.type -eq 'TextBox'}).Count; $listboxes=@($all|Where-Object {$_.type -eq 'ListBox'}).Count
        if ($forms.Count -ne $bindings.Count -or $all.Count -ne $expectedControlCount -or $textboxes -ne $expectedTextBoxCount -or $listboxes -ne $expectedListBoxCount) { throw 'Product UserForm aggregate counts rejected' }
        $formBindings = @($bindings|ForEach-Object {[ordered]@{layout=[string]$_.layout;code=[string]$_.code}})
        if ($null -ne $ExistingProject) {
            return [ordered]@{method='startup-loaded Product.xlam VBProject Designer.Controls readback';form_bindings=$formBindings;forms=@($forms);total_forms=$forms.Count;total_controls=$all.Count;textboxes=$textboxes;listboxes=$listboxes}
        }
        return [ordered]@{method='built Product.xlam close/reopen + VBProject Designer.Controls readback';form_bindings=$formBindings;forms=@($forms);total_forms=$forms.Count;total_controls=$all.Count;textboxes=$textboxes;listboxes=$listboxes}
    } finally {
        if ($null -ne $reopened) { try {$reopened.Close($false)} catch {}; Release-ComObject $reopened }
    }
}

$excel = $null; $interactiveExcel = $null; $books = $null; $book = $null; $productProject = $null; $uiHostBook = $null; $excelProcess = $null; $excelBinding = $null; $excelOwned = $false; $evidence = $null; $artifact = $null; $inspectionArtifact = $null; $artifactShaBefore = $null; $artifactShaAfter = $null; $probeContext = $null; $productCompileContext = $null; $userFormProbe = $null; $uiProbeContext = $null; $windowDiagnosticContext = $null; $startupMode = $false
try {
    $failurePhase = 'preflight'
    if (-not [string]::IsNullOrWhiteSpace($ArtifactPath) -and -not [string]::IsNullOrWhiteSpace($StartupArtifactPath)) {
        throw 'ArtifactPath and StartupArtifactPath are mutually exclusive'
    }
    $manifest = Get-ProductManifest
    $evidence = if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) { Join-Path $DataRoot 'runtime/evidence/vba' } else { $EvidenceRoot }
    New-Item -ItemType Directory -Force -Path $evidence | Out-Null
    $startupMode = -not [string]::IsNullOrWhiteSpace($StartupArtifactPath)
    $artifact = if ($startupMode) { [IO.Path]::GetFullPath($StartupArtifactPath) } else { Join-Path $evidence 'Product.xlam' }
    $failurePhase = 'build'
    if ($startupMode) {
        if (-not (Test-Path -LiteralPath $artifact -PathType Leaf) -or [IO.Path]::GetExtension($artifact) -ine '.xlam') { throw 'ProductRibbon startup Product artifact rejected' }
    } elseif ([string]::IsNullOrWhiteSpace($ArtifactPath)) {
        & (Join-Path $SourceRoot 'build/Build-Xlam.ps1') -DataRoot $DataRoot -OutputPath $artifact -ManifestPath (Join-Path $SourceRoot 'build/manifests/Product.json') -RibbonPath (Join-Path $SourceRoot 'src/ribbon/customUI14.xml')
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $artifact -PathType Leaf)) { throw 'Build-Xlam failed' }
    } else {
        $ArtifactPath = [IO.Path]::GetFullPath($ArtifactPath)
        if (-not (Test-Path -LiteralPath $ArtifactPath -PathType Leaf) -or [IO.Path]::GetExtension($ArtifactPath) -ine '.xlam') { throw 'ProductRibbon supplied Product artifact rejected' }
        Copy-Item -LiteralPath $ArtifactPath -Destination $artifact -Force
        if ((Get-HexDigest $artifact) -cne (Get-HexDigest $ArtifactPath)) { throw 'ProductRibbon supplied Product artifact copy mismatch' }
    }
    $inspectionArtifact = $artifact
    if ($startupMode) {
        $inspectionArtifact = Join-Path $evidence 'Product.xlam'
        Copy-Item -LiteralPath $artifact -Destination $inspectionArtifact -Force
        if ((Get-HexDigest $inspectionArtifact) -cne (Get-HexDigest $artifact)) { throw 'ProductRibbon startup inspection artifact copy mismatch' }
    }
    $artifactShaBefore = Get-HexDigest $inspectionArtifact

    $failurePhase = 'preflight'
    $excelBaseline = Get-ExcelProcessBaseline
    $interactiveExcel = Start-InteractiveProductExcel $excelBaseline
    $excel = $interactiveExcel.excel
    $excelBinding = $interactiveExcel.binding
    $interactiveExcel = $null
    $excelOwned = [bool]$excelBinding.owned
    if (-not $excelOwned) { throw ('ProductRibbon Excel ownership rejected: ' + [string]$excelBinding.failure) }
    $excelProcess = $excelBinding.process
    if (-not [bool]$excel.UserControl) { throw 'ProductRibbon Excel is not user controlled' }
    if ([string]$excel.Version -notmatch '^16\.') { throw 'Office 2024 or Microsoft 365 Excel 16.x family required' }
    $vbePreflight = $null
    try { $vbePreflight = $excel.VBE } catch { throw ('Product VBOM access blocked by Trust Center or AccessVBOM policy: ' + $_.Exception.Message) } finally { Release-ComObject $vbePreflight }
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.WindowState = -4143
    $books = $excel.Workbooks
    if ($startupMode) {
        $uiHostBook = Resolve-ProductUiHostWorkbook $books
        $uiHostBook.Activate()
        $initialPid = [int]$excelBinding.pid
        $refreshedBinding = Get-ExactExcelProcessOwnership $excel $excelBaseline 'ProductRibbon active startup host'
        if (-not [bool]$refreshedBinding.owned -or [int]$refreshedBinding.pid -ne $initialPid) {
            if ($null -ne $refreshedBinding.process) { $refreshedBinding.process.Dispose() }
            throw 'ProductRibbon startup Excel ownership changed after host activation'
        }
        if ($null -ne $excelProcess) { $excelProcess.Dispose() }
        $excelBinding = $refreshedBinding
        $excelProcess = $excelBinding.process
        Write-ProductRibbonOwnership $evidence $excelBinding $artifact 'startup-host-bound'
        $productProject = Resolve-StartupProductProject $excel $artifact
    } else {
        if ([int]$books.Count -eq 0) { $uiHostBook = $books.Add() }
        elseif ([int]$books.Count -eq 1) { $uiHostBook = $books.Item(1) }
        else { throw 'Interactive Excel startup workbook count rejected' }
        $book = $books.Open($artifact, $false, $true)
        if ($null -eq $book -or -not [bool]$book.IsAddin -or -not [bool]$book.ReadOnly) { throw 'ProductRibbon XLAM did not open as an add-in workbook' }
        $productProject = $book.VBProject
        Write-ProductRibbonOwnership $evidence $excelBinding $artifact 'direct-open-bound'
    }
    $uiHostBook.Activate()
    $expectedArtifactName = [IO.Path]::GetFileName($artifact)
    $loadedProjectPath = [IO.Path]::GetFullPath([string]$productProject.FileName)
    if (-not [string]::Equals($loadedProjectPath, [IO.Path]::GetFullPath($artifact), [StringComparison]::OrdinalIgnoreCase)) { throw 'ProductRibbon opened workbook identity mismatch' }
    if (-not $startupMode -and [string]$book.Name -cne $expectedArtifactName) { throw 'ProductRibbon XLAM did not open as the expected add-in workbook' }
    $probeContext = [ordered]@{
        artifact_name=[IO.Path]::GetFileName($artifact)
        artifact_sha256=$artifactShaBefore
        book_name=if($startupMode){$expectedArtifactName}else{[string]$book.Name}
        is_addin=if($startupMode){$true}else{[bool]$book.IsAddin}
        read_only=if($startupMode){$null}else{[bool]$book.ReadOnly}
        load_surface=if($startupMode){'VBE.VBProjects'}else{'Workbooks.Open'}
        loaded_project_path=$loadedProjectPath
        workbook_count=[int]$books.Count
        user_control=[bool]$excel.UserControl
        automation_security=[int]$excel.AutomationSecurity
        macro_execution=if($startupMode){'actual-ribbon-ui-startup'}else{'actual-ribbon-ui'}
        startup_mode=$startupMode
    }

    $failurePhase = 'probe'
    $script:probeStage = 'package'
    $packageProbe = Invoke-PackageProbe $inspectionArtifact
    $script:probeStage = 'office-identity'
    $identityProbe = Get-OfficeIdentity $excel $book $excelBinding
    $script:probeStage = 'ui:current-200'
    $uiHostBook.Activate()
    $current200 = Invoke-UiAutomationProbe $excel
    $ribbonUiProbe = Invoke-ActualProductRibbonCases $excel $uiHostBook $excelBinding $inspectionArtifact
    $script:probeStage = 'ribbon-callback'
    $callbackProbe = Invoke-RibbonCallbackProbe $excel $productProject $artifact $inspectionArtifact $ribbonUiProbe (-not $startupMode)
    $gradeProbe = Invoke-GradeProbe $ribbonUiProbe $callbackProbe
    $script:probeStage = 'userform-attestation'
    if ($startupMode) {
        $userFormProbe = Invoke-ProductUserFormAttestation $books $artifact (Join-Path $SourceRoot 'build/manifests/Product.json') $productProject
    } else {
        Release-ComObject $productProject; $productProject=$null
        $book.Close($false); Release-ComObject $book; $book=$null
        $userFormProbe = Invoke-ProductUserFormAttestation $books $artifact (Join-Path $SourceRoot 'build/manifests/Product.json')
    }
    $artifactShaAfter = Get-HexDigest $inspectionArtifact
    if ($artifactShaAfter -ne $artifactShaBefore) { throw 'Product.xlam artifact SHA changed during ProductRibbon verification' }

    $failurePhase = 'evidence'
    Write-ProbeEvidence (Join-Path $evidence 'ProductRibbon.Package.json') 'Package' $packageProbe $inspectionArtifact
    Write-ProbeEvidence (Join-Path $evidence 'ProductRibbon.RibbonXml.json') 'RibbonXml' $callbackProbe $inspectionArtifact
    Write-ProbeEvidence (Join-Path $evidence 'ProductRibbon.OfficeIdentity.json') 'OfficeIdentity' $identityProbe $inspectionArtifact
    Write-ProbeEvidence (Join-Path $evidence 'ProductRibbon.Ui.Current200.json') 'Ui.Current200' $current200 $inspectionArtifact
    $fastGradeCase = $gradeProbe.Fast
    Write-ProbeEvidence (Join-Path $evidence 'ProductRibbon.Fast.json') 'Fast' ([ordered]@{
        method=$gradeProbe.method;execution_grade='Fast';resolved=[int]$fastGradeCase.resolved
        route_feature_id=[string]$fastGradeCase.feature_id;route_verified=[bool]$fastGradeCase.route_verified
        form_opened=[bool]$fastGradeCase.form_opened;route_proof=[string]$fastGradeCase.route_proof
        policy_attested_from_compiled_product=[bool]$gradeProbe.policy_attested_from_compiled_product
    }) $inspectionArtifact
    $guardedGradeCase = $gradeProbe.Guarded
    Write-ProbeEvidence (Join-Path $evidence 'ProductRibbon.Guarded.json') 'Guarded' ([ordered]@{
        method=$gradeProbe.method;execution_grade='Guarded';resolved=[int]$guardedGradeCase.resolved
        route_feature_id=[string]$guardedGradeCase.feature_id;route_verified=[bool]$guardedGradeCase.route_verified
        form_opened=[bool]$guardedGradeCase.form_opened;route_proof=[string]$guardedGradeCase.route_proof
        policy_attested_from_compiled_product=[bool]$gradeProbe.policy_attested_from_compiled_product
    }) $inspectionArtifact
    $plannedGradeCase = $gradeProbe.Planned
    Write-ProbeEvidence (Join-Path $evidence 'ProductRibbon.Planned.json') 'Planned' ([ordered]@{
        method=$gradeProbe.method;execution_grade='Planned';resolved=[int]$plannedGradeCase.resolved
        route_feature_id=[string]$plannedGradeCase.feature_id;route_verified=[bool]$plannedGradeCase.route_verified
        form_opened=[bool]$plannedGradeCase.form_opened;route_proof=[string]$plannedGradeCase.route_proof
        policy_attested_from_compiled_product=[bool]$gradeProbe.policy_attested_from_compiled_product
    }) $inspectionArtifact
    Write-ProbeEvidence (Join-Path $evidence 'ProductRibbon.UserForms.json') 'UserForms' $userFormProbe $inspectionArtifact
    $terminalCode = 0
} catch {
    $failureMessage = $_.Exception.Message
    $stackFunctions = @($_.ScriptStackTrace -split '\r?\n' | ForEach-Object {
        $frame = [regex]::Match([string]$_, '^(?:[^:]+:\s*|at\s+)([^,]+),')
        if ($frame.Success) { $frame.Groups[1].Value.Trim() }
    })
    $failureDetail = [ordered]@{
        stage = $script:probeStage
        exception_type = $_.Exception.GetType().FullName
        script_line = [int]$_.InvocationInfo.ScriptLineNumber
        source_line = ([string]$_.InvocationInfo.Line).Trim()
        stack_functions = $stackFunctions
    }
    $terminalCode = switch ($failurePhase) {
        'preflight' { 10 }; 'build' { 22 }; 'probe' { 23 }; 'evidence' { 24 }; default { 24 }
    }
    [Console]::Error.WriteLine(('ProductRibbon ' + $failurePhase + ' failure at ' + $script:probeStage + ' line ' + $failureDetail.script_line + ': ' + $failureMessage))
} finally {
    if ($null -ne $productProject) {
        try { Release-ComObject $productProject } catch { $cleanupFailures += $_.Exception.Message }
    }
    if ($null -ne $book) {
        try { $book.Close($false) } catch { $cleanupFailures += $_.Exception.Message }
        try { Release-ComObject $book } catch { $cleanupFailures += $_.Exception.Message }
    }
    if ($excelOwned -and $null -ne $excel) {
        try { [void](Assert-ExactExcelProcessOwnership $excelBinding) } catch { $cleanupFailures += $_.Exception.Message; $excelOwned = $false }
    }
    if ($null -ne $uiHostBook) {
        try { $uiHostBook.Close($false) } catch { $cleanupFailures += $_.Exception.Message }
        try { Release-ComObject $uiHostBook } catch { $cleanupFailures += $_.Exception.Message }
    }
    if ($null -ne $books) {
        try { Release-ComObject $books } catch { $cleanupFailures += $_.Exception.Message }
    }
    if ($null -ne $excel) {
        try { if ($excelOwned -and $null -ne $excel) { $excel.Quit() } } catch { $cleanupFailures += $_.Exception.Message }
        try { Release-ComObject $excel } catch { $cleanupFailures += $_.Exception.Message }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($excelOwned -and $null -ne $excelProcess) {
        try {
            $cleanupResult = Stop-ExactProcessAfterGrace -Process $excelProcess -Label 'ProductRibbon Excel' -GraceMs 10000 -Detailed
            $cleanupExitMode = [string]$cleanupResult.exit_mode
            if ($cleanupResult.failure) { $cleanupFailures += [string]$cleanupResult.failure }
        } catch { $cleanupFailures += $_.Exception.Message }
    }
    if ($null -ne $excelProcess) { try { $excelProcess.Dispose() } catch { $cleanupFailures += $_.Exception.Message } }
    if ($cleanupFailures.Count -gt 0) {
        $failurePhase = 'cleanup'
        $terminalCode = 25
        if ([string]::IsNullOrWhiteSpace($failureMessage)) { $failureMessage = 'Owned Excel cleanup failed: ' + ($cleanupFailures -join ' | ') }
    }
    if ($terminalCode -eq 0) {
        Write-Json (Join-Path $evidence 'ProductRibbon.json') ([ordered]@{
            schema_version=1; suite='ProductRibbon'; status='PASS'; mode='Green'; run_id=$RunId
            source_digest=$SourceDigest; snapshot_digest=$SnapshotDigest; artifact_sha256=$artifactShaAfter
            environment=[ordered]@{excel_bitness='64-bit';office=[string]$identityProbe.office_version;'compatibility_baseline'='Office 2024'};cleanup=[ordered]@{status='PASS';exit_mode=$cleanupExitMode;failures=@()}
            run=[ordered]@{total=3;passed=3;failed=0;skipped=0}
        })
    } elseif ($null -ne $evidence -and (Test-Path -LiteralPath $evidence -PathType Container)) {
        $cleanupStatus = if ($cleanupFailures.Count -eq 0 -and $cleanupExitMode -in @('NATURAL','FORCED_CONTAINED')) { 'PASS' } elseif ($cleanupFailures.Count -eq 0) { 'NOT_REACHED' } else { 'FAIL' }
        Write-Json (Join-Path $evidence 'ProductRibbon.json') ([ordered]@{
            schema_version=1; suite='ProductRibbon'; status='INCOMPLETE'; mode='Green'; run_id=$RunId
            source_digest=$SourceDigest; snapshot_digest=$SnapshotDigest; artifact_sha256=if($null -ne $inspectionArtifact -and (Test-Path -LiteralPath $inspectionArtifact -PathType Leaf)){Get-HexDigest $inspectionArtifact}else{$null}
            failure_phase=$failurePhase; exit_code=$terminalCode; failure=$failureMessage; error_context=$failureDetail; probe_context=$probeContext; compile_context=$productCompileContext; dpi_context=$uiProbeContext; window_context=$windowDiagnosticContext
            cleanup=[ordered]@{status=$cleanupStatus;exit_mode=$cleanupExitMode;failures=$cleanupFailures}
            run=[ordered]@{total=3;passed=0;failed=1;skipped=0}
        })
    }
    exit $terminalCode
}
