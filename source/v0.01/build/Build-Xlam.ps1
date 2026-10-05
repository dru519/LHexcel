param(
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [string]$ManifestPath = '',
    [string]$RibbonPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Excel-ProcessLifecycle.ps1')
if ([string]::IsNullOrWhiteSpace($ManifestPath)) { $ManifestPath = Join-Path $PSScriptRoot 'manifests/Product.json' }
if ([string]::IsNullOrWhiteSpace($RibbonPath)) { $RibbonPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'src/ribbon/customUI14.xml' }

function Assert-NoReparsePath([string]$Path, [string]$ExpectedType, [string]$Boundary) {
    $boundaryFull = [IO.Path]::GetFullPath($Boundary).TrimEnd('\')
    $pathFull = [IO.Path]::GetFullPath($Path)
    if ($pathFull -ne $boundaryFull -and -not $pathFull.StartsWith($boundaryFull + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw ('Product reparse check escaped trust boundary: ' + $Path)
    }
    $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    $cursor = $item
    while ($null -ne $cursor) {
        if (($cursor.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('Product reparse path rejected: ' + $cursor.FullName) }
        if ([IO.Path]::GetFullPath($cursor.FullName).TrimEnd('\') -eq $boundaryFull) { break }
        $parentPath = Split-Path -Parent $cursor.FullName
        if ([string]::IsNullOrEmpty($parentPath) -or $parentPath -eq $cursor.FullName) { throw 'Product reparse boundary was not reached' }
        $cursor = Get-Item -LiteralPath $parentPath -Force -ErrorAction Stop
    }
    if ($ExpectedType -eq 'Leaf' -and $item.PSIsContainer) { throw ('Product file path rejected: ' + $Path) }
    if ($ExpectedType -eq 'Container' -and -not $item.PSIsContainer) { throw ('Product directory path rejected: ' + $Path) }
    return $item
}

function Get-VerifiedHash([string]$Path, [string]$Boundary) {
    [void](Assert-NoReparsePath $Path 'Leaf' $Boundary)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
}

function Get-VerifiedHashWithRetry([string]$Path, [string]$Boundary, [int]$TimeoutMs = 10000) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $lastFailure = 'hash not started'
    while ($watch.ElapsedMilliseconds -lt $TimeoutMs) {
        try { return (Get-VerifiedHash $Path $Boundary) }
        catch { $lastFailure = $_.Exception.Message }
        Start-Sleep -Milliseconds 100
    }
    throw ('Shared artifact hash timeout: path=' + $Path + '; timeout_ms=' + $TimeoutMs + '; cause=' + $lastFailure)
}

$sourceRoot = [IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$root = [IO.Path]::GetFullPath($DataRoot)
$destination = [IO.Path]::GetFullPath($OutputPath)
if ([IO.Path]::GetExtension($destination) -ine '.xlam') { throw 'Product output must have .xlam extension' }
if (Test-Path -LiteralPath $destination -ErrorAction Stop) { throw 'Existing product XLAM destination is rejected' }
$parent = Split-Path -Parent $destination
if (-not (Test-Path -LiteralPath $parent -PathType Container -ErrorAction Stop)) { [void](New-Item -ItemType Directory -Path $parent -Force -ErrorAction Stop) }
[void](Assert-NoReparsePath $sourceRoot 'Container' $sourceRoot)
[void](Assert-NoReparsePath $root 'Container' $root)
[void](Assert-NoReparsePath $parent 'Container' $parent)

$manifestTool = Join-Path $sourceRoot 'tools/build_product_manifest.py'
[void](Assert-NoReparsePath $manifestTool 'Leaf' $sourceRoot)
[void](Assert-NoReparsePath $ManifestPath 'Leaf' $sourceRoot)
[void](Assert-NoReparsePath $RibbonPath 'Leaf' $sourceRoot)
$python = (Get-Command python -ErrorAction Stop).Source
$symbolCatalogTool = Join-Path $sourceRoot 'tools/generate_symbol_catalog.py'
[void](Assert-NoReparsePath $symbolCatalogTool 'Leaf' $sourceRoot)
& $python $symbolCatalogTool --check
if ($LASTEXITCODE -ne 0) { throw 'Generated symbol catalog is stale' }
& $python $manifestTool --check
if ($LASTEXITCODE -ne 0) { throw 'Product manifest is stale' }

$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$manifestProperties = @($manifest.PSObject.Properties.Name)
if ([int]$manifest.schema_version -ne 3 -or [int]$manifest.release_scope.feature_count -ne 49 -or $manifestProperties -contains 'runtime') { throw 'Product release scope rejected' }
$included = @($manifest.release_scope.included_feature_ids)
$excluded = @($manifest.release_scope.excluded_feature_ids)
if ($included.Count -ne 49 -or @($included | Sort-Object -Unique).Count -ne 49 -or $included -notcontains 'NX-HANGUL-TABLE-SEND' -or $included -notcontains 'NX-HANGUL-PICTURE-SEND' -or $included -notcontains 'NX-FILE-FILE-COMPARE' -or $included -notcontains 'NX-FILE-SHEET-COMPARE' -or @($included | Where-Object { [string]$_ -like 'NX-HWPX-*' -or [string]$_ -like 'NX-DATA-ROSTER-*' }).Count -ne 0) { throw 'Product included feature scope rejected' }
if ($excluded.Count -ne 0) { throw 'Product excluded feature scope rejected' }
$removedModules = @($manifest.modules | Where-Object { [string]$_.path -match '^src/vba/features/hwpx/' -or [string]$_.path -match '(?:^|/)CNxRoster|(?:^|/)NxRoster|(?:^|/)FNxRoster' })
$hangulModules = @($manifest.modules | Where-Object { [string]$_.path -match '^src/vba/features/hangul/' })
if ($removedModules.Count -ne 0 -or $hangulModules.Count -lt 9 -or @($hangulModules | Where-Object { [string]$_.path -match 'NxHwpxController|NxHwpxSerializer|NxHwpxEmbeddedResources|NxHangulSettings' }).Count -lt 4) { throw 'Product Hangul HWPX source inventory rejected' }
if ($manifest.ribbon.path -ne 'src/ribbon/customUI14.xml') { throw 'Product ribbon manifest path rejected' }
$ribbonHash = Get-VerifiedHash $RibbonPath $sourceRoot
if ($ribbonHash -ne ([string]$manifest.ribbon.sha256).ToLowerInvariant()) { throw 'Product ribbon SHA mismatch' }

# Excel COM cannot reliably create and reopen a workbook on Parallels shared
# folders. Build and validate on local NTFS, then copy one hash-bound candidate
# to the shared destination directory for the final same-directory rename.
$localStageBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
[void](Assert-NoReparsePath $localStageBase 'Container' $localStageBase)
$publishRoot = Join-Path $localStageBase ('naeexcel-product-publish-' + [guid]::NewGuid().ToString('N'))
$publishStage = Join-Path $publishRoot ([IO.Path]::GetFileName($destination))
$publishCandidate = Join-Path $parent ('product-candidate-' + [guid]::NewGuid().ToString('N') + '.xlam')
if (Test-Path -LiteralPath $publishCandidate -ErrorAction Stop) { throw 'Existing Product publish candidate rejected' }
$published = $false
$publishPhase = 'initialize'
try {
    $publishPhase = 'create-stage-directory'
    [void](New-Item -ItemType Directory -Path $publishRoot -Force -ErrorAction Stop)
    [void](Assert-NoReparsePath $publishRoot 'Container' $localStageBase)
    $publishPhase = 'create-blank-xlam'
    & (Join-Path $PSScriptRoot 'New-BlankXlam.ps1') -Path $publishStage
    if ($LASTEXITCODE -ne 0) { throw 'Blank product XLAM creation failed' }
    $publishPhase = 'stabilize-blank-xlam'
    Wait-ExactArtifactReadable $publishStage 10000
    $publishPhase = 'import-product-vba'
    & (Join-Path $PSScriptRoot 'Import-ProductVba.ps1') -DataRoot $root -XlamPath $publishStage -ManifestPath $ManifestPath
    if ($LASTEXITCODE -ne 0) { throw 'Product VBA import failed' }
    $publishPhase = 'stabilize-imported-xlam'
    Wait-ExactArtifactReadable $publishStage 10000
    $publishPhase = 'inject-ribbon'
    & $python (Join-Path $PSScriptRoot 'inject_ribbon.py') --xlam $publishStage --ribbon $RibbonPath --ribbon-sha $manifest.ribbon.sha256
    if ($LASTEXITCODE -ne 0) { throw 'Product ribbon injection failed' }
    $publishPhase = 'stabilize-ribbon-xlam'
    Wait-ExactArtifactReadable $publishStage 10000
    $publishPhase = 'hash-ribbon-xlam'
    $stageDigest = Get-VerifiedHash $publishStage $publishRoot
    $publishPhase = 'copy-publish-candidate'
    Copy-Item -LiteralPath $publishStage -Destination $publishCandidate -ErrorAction Stop
    Wait-ExactArtifactReadable $publishCandidate 10000
    $candidateDigest = Get-VerifiedHashWithRetry $publishCandidate $parent 10000
    if ($candidateDigest -ne $stageDigest) { throw 'Product publish candidate integrity mismatch' }
    $publishPhase = 'publish-final-xlam'
    Move-Item -LiteralPath $publishCandidate -Destination $destination -ErrorAction Stop
    $published = $true
    $publishPhase = 'verify-final-xlam'
    Wait-ExactArtifactReadable $destination 10000
    if ((Get-VerifiedHashWithRetry $destination $parent 10000) -ne $stageDigest) { throw 'Published Product XLAM integrity mismatch' }
} catch {
    $publishError = $_.Exception.Message
    if ($published -and (Test-Path -LiteralPath $destination -PathType Leaf -ErrorAction SilentlyContinue)) {
        Remove-Item -LiteralPath $destination -Force -ErrorAction Stop
    }
    throw ('Publish transaction failed; phase=' + $publishPhase + '; cleaning final outputs: ' + $publishError)
} finally {
    if (Test-Path -LiteralPath $publishCandidate -PathType Leaf -ErrorAction SilentlyContinue) { Remove-Item -LiteralPath $publishCandidate -Force -ErrorAction Stop }
    if (Test-Path -LiteralPath $publishRoot -ErrorAction SilentlyContinue) {
        [void](Assert-NoReparsePath $publishRoot 'Container' $localStageBase)
        Remove-Item -LiteralPath $publishRoot -Force -Recurse -ErrorAction Stop
    }
}
Write-Output $destination
