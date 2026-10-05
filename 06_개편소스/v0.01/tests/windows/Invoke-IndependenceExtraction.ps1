param(
  [switch]$FixtureMode,
  [switch]$FrozenMode,
  [Parameter(Mandatory=$true)][string]$Candidate,
  [Parameter(Mandatory=$true)][string]$FixtureRoot,
  [string]$ReferenceCandidate = '',
  [string]$ReferenceRoot = '',
  [string]$ReferenceVbaRoot = '',
  [ValidateSet('','original','v3x')][string]$ReferenceIdentity = '',
  [string]$FreezeManifest = '',
  [string]$OutputRoot = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ToolVersion = 'G012-independence-extractor/1.1'
if (($FixtureMode -and $FrozenMode) -or (-not $FixtureMode -and -not $FrozenMode)) { throw 'FixtureMode and FrozenMode are mutually exclusive' }
if ($FixtureMode -and ((-not [string]::IsNullOrWhiteSpace($ReferenceCandidate)) -or (-not [string]::IsNullOrWhiteSpace($ReferenceRoot)) -or (-not [string]::IsNullOrWhiteSpace($ReferenceVbaRoot)) -or (-not [string]::IsNullOrWhiteSpace($ReferenceIdentity)) -or (-not [string]::IsNullOrWhiteSpace($FreezeManifest)))) { throw 'Reference arguments are forbidden in fixture mode' }
if ($FrozenMode -and [string]::IsNullOrWhiteSpace($FreezeManifest)) { throw 'FrozenMode requires FreezeManifest' }
if (-not (Test-Path -LiteralPath $Candidate -PathType Leaf)) { throw 'Candidate package is missing' }
if (-not (Test-Path -LiteralPath $FixtureRoot -PathType Container)) { throw 'FixtureRoot is missing' }
if ($FrozenMode -and ([string]::IsNullOrWhiteSpace($ReferenceCandidate) -or [string]::IsNullOrWhiteSpace($ReferenceRoot) -or [string]::IsNullOrWhiteSpace($ReferenceIdentity))) { throw 'FrozenMode requires ReferenceCandidate, ReferenceRoot, and ReferenceIdentity' }

function Get-Sha256([string]$Path) { return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Get-Canonical([string]$Path) { return [IO.Path]::GetFullPath((Resolve-Path -LiteralPath $Path).Path) }
function Get-Utf8SortKey([string]$Value) { return ([BitConverter]::ToString([Text.Encoding]::UTF8.GetBytes($Value))).Replace('-','') }
function Get-ManifestValue($Object, [string[]]$Names) {
  foreach ($name in $Names) {
    $property = $Object.PSObject.Properties[$name]
    if ($null -ne $property -and $null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return [string]$property.Value }
  }
  return ''
}
function Add-FileRecord([System.Collections.ArrayList]$Records, [string]$Subject, [string]$Path, [string]$Kind) {
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw ('Missing extracted artifact: ' + $Path) }
  $item = Get-Item -LiteralPath $Path
  [void]$Records.Add([ordered]@{ subject=$Subject; kind=$Kind; path=(Get-Canonical $Path); sha256=(Get-Sha256 $Path); size=$item.Length; access_utc=[DateTime]::UtcNow.ToString('o'); tool_version=$ToolVersion })
}
function Export-ReferenceRoot([string]$Source, [string]$Destination, [System.Collections.ArrayList]$Records) {
  $files = @(Get-ChildItem -LiteralPath $Source -File -Recurse | Sort-Object @{ Expression = { Get-Utf8SortKey ([string]$_.FullName) } })
  if ($files.Count -eq 0) { throw 'ReferenceRoot contains no artifacts' }
  foreach ($file in $files) {
    $relative = $file.FullName.Substring((Get-Canonical $Source).Length).TrimStart([char[]]"\/")
    $target = Join-Path $Destination $relative
    New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
    Copy-Item -LiteralPath $file.FullName -Destination $target -Force
    Add-FileRecord $Records 'reference-root' $target ([IO.Path]::GetExtension($target).TrimStart('.').ToLowerInvariant())
    $Records[$Records.Count - 1]['inventory_path'] = $relative.Replace('\','/')
  }
}
function Export-ReferenceVbaRoot([string]$Source, [string]$Destination, [System.Collections.ArrayList]$Records) {
  $allowedExtensions = @('.bas','.cls','.frm')
  $files = @(Get-ChildItem -LiteralPath $Source -File -Recurse | Where-Object { ($allowedExtensions -contains $_.Extension.ToLowerInvariant()) -or ($_.Name -like 'VBA__*.txt') } | Sort-Object @{ Expression = { Get-Utf8SortKey ([string]$_.FullName) } })
  if ($files.Count -eq 0) { throw 'ReferenceVbaRoot contains no supported VBA source artifacts' }
  $sourceRoot = Get-Canonical $Source
  foreach ($file in $files) {
    $relative = $file.FullName.Substring($sourceRoot.Length).TrimStart([char[]]"\/")
    $target = Join-Path $Destination $relative
    New-Item -ItemType Directory -Path (Split-Path -Parent $target) -Force | Out-Null
    Copy-Item -LiteralPath $file.FullName -Destination $target -Force
    Add-FileRecord $Records 'reference-vba' $target 'archived-source'
  }
}
function Export-ReferenceCandidate([string]$Source, [string]$Destination, [string]$Subject, [System.Collections.ArrayList]$Records) {
  $excel = $null; $book = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open((Get-Canonical $Source), 0, $true)
    $components = $book.VBProject.VBComponents
    if ($null -eq $components) { throw 'Reference workbook VBComponents are unavailable after load' }
    $componentCount = [int]$components.Count
    if ($componentCount -eq 0) { throw 'Reference workbook has no VBProject components' }
    for ($componentIndex = 1; $componentIndex -le $componentCount; $componentIndex++) {
      $component = $components.Item($componentIndex)
      $name = [string]$component.Name
      if ([string]::IsNullOrWhiteSpace($name)) { throw 'VBA component name is missing' }
      $type = switch ([int]$component.Type) { 1 {'standard'} 2 {'class'} 3 {'form'} 100 {'document'} default { 'unknown' } }
      if ($type -eq 'unknown') { throw ('Unsupported VBA component type: ' + $name) }
      $before = @(Get-ChildItem -LiteralPath $Destination -File -Recurse -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
      $extension = switch ($type) { 'standard' { '.bas' } 'class' { '.cls' } 'form' { '.frm' } 'document' { '.cls' } }
      $component.Export((Join-Path $Destination ($name + $extension)))
      $after = @(Get-ChildItem -LiteralPath $Destination -File -Recurse | Where-Object { $before -notcontains $_.FullName })
      if ($after.Count -eq 0) { throw ('VBA component export missing: ' + $name) }
      foreach ($file in $after) { Add-FileRecord $Records $Subject $file.FullName $type }
      if ($type -eq 'form' -and -not (Get-ChildItem -LiteralPath $Destination -Filter ($name + '.frx') -File -Recurse)) { throw ('UserForm FRX export missing: ' + $name) }
      try { [Runtime.InteropServices.Marshal]::ReleaseComObject($component) | Out-Null } catch {}
    }
  } catch { throw ('Reference extraction INCOMPLETE: ' + $_.Exception.Message) }
  finally {
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    foreach ($obj in @($book,$excel)) { if ($null -ne $obj) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($obj) | Out-Null } catch {} } }
  }
}

$root = if ([string]::IsNullOrWhiteSpace($OutputRoot)) { Join-Path $FixtureRoot 'independence-extraction' } else { $OutputRoot }
New-Item -ItemType Directory -Path $root -Force | Out-Null
$records = New-Object System.Collections.ArrayList
$extractionErrors = New-Object System.Collections.ArrayList
$sourceCommit = ''
$sourceTree = ''
$python = Get-Command python -ErrorAction SilentlyContinue
if ($null -eq $python) { throw 'python is required for package extraction' }
$extractor = Join-Path $PSScriptRoot '../../tools/extract_package.py'
$report = Join-Path $root 'candidate-package.json'
$candidateSha = Get-Sha256 $Candidate
$candidateVbaDir = Join-Path $root 'candidate-vba'
$referenceReport = Join-Path $root 'reference-package.json'
$referenceVbaDir = Join-Path $root 'reference-vba'
$referenceCandidateSha = ''
$referenceVbaMode = ''
if ($FrozenMode) {
  if (-not (Test-Path -LiteralPath $FreezeManifest -PathType Leaf)) { throw 'FreezeManifest is missing' }
  $freeze = Get-Content -LiteralPath $FreezeManifest -Raw | ConvertFrom-Json
  $expectedSha = Get-ManifestValue $freeze @('candidate_sha256','artifact_sha256','CandidateSha256')
  $candidateDescriptor = $freeze.PSObject.Properties['candidate']
  if ([string]::IsNullOrWhiteSpace($expectedSha) -and $null -ne $candidateDescriptor -and $null -ne $candidateDescriptor.Value) { $expectedSha = Get-ManifestValue $candidateDescriptor.Value @('sha256','artifact_sha256') }
  if ([string]::IsNullOrWhiteSpace($expectedSha)) { throw 'FreezeManifest candidate artifact SHA is missing' }
  if ($expectedSha -notmatch '^[0-9a-fA-F]{64}$') { throw 'FreezeManifest candidate artifact SHA is invalid' }
  if ($candidateSha -ne $expectedSha.ToLowerInvariant()) { throw 'Candidate artifact SHA does not match FreezeManifest' }
  $sourceCommit = Get-ManifestValue $freeze @('frozen_source_commit','approved_commit','source_commit','commit','commit_sha')
  $sourceTree = Get-ManifestValue $freeze @('source_tree_digest','source_tree_sha256','source_tree','tree','tree_sha','tree_sha256','tree_identity','source_tree_identity')
  if (($sourceCommit -notmatch '^[0-9a-fA-F]{40}$') -or ($sourceTree -notmatch '^[0-9a-fA-F]{64}$')) { throw 'FreezeManifest source commit/tree identity is invalid' }
}
& $python.Source $extractor $Candidate --json $report
if ($LASTEXITCODE -ne 0) { throw ('Package extraction returned ' + $LASTEXITCODE) }
Add-FileRecord $records 'candidate-artifact' $Candidate 'package'
Add-FileRecord $records 'candidate-package' $report 'report'
if ($FrozenMode) {
  New-Item -ItemType Directory -Path $candidateVbaDir -Force | Out-Null
  try { Export-ReferenceCandidate $Candidate $candidateVbaDir 'candidate-vba' $records } catch { [void]$extractionErrors.Add($_.Exception.Message) }
  $referenceDir = Join-Path $root 'reference'
  New-Item -ItemType Directory -Path $referenceDir -Force | Out-Null
  if (-not [string]::IsNullOrWhiteSpace($ReferenceCandidate)) {
    if (-not (Test-Path -LiteralPath $ReferenceCandidate -PathType Leaf)) { throw 'ReferenceCandidate is missing' }
    New-Item -ItemType Directory -Path $referenceVbaDir -Force | Out-Null
    $referenceCandidateSha = Get-Sha256 $ReferenceCandidate
    & $python.Source $extractor $ReferenceCandidate --json $referenceReport
    if ($LASTEXITCODE -ne 0) { [void]$extractionErrors.Add('Reference package extraction returned ' + $LASTEXITCODE) }
    elseif (Test-Path -LiteralPath $referenceReport -PathType Leaf) { Add-FileRecord $records 'reference-package' $referenceReport 'report' }
    else { [void]$extractionErrors.Add('Reference package report is missing') }
  }
  if (-not [string]::IsNullOrWhiteSpace($ReferenceRoot)) {
    if (-not (Test-Path -LiteralPath $ReferenceRoot -PathType Container)) { throw 'ReferenceRoot is missing' }
    try { Export-ReferenceRoot $ReferenceRoot (Join-Path $referenceDir 'root') $records } catch { [void]$extractionErrors.Add($_.Exception.Message) }
  }
  if (-not [string]::IsNullOrWhiteSpace($ReferenceVbaRoot)) {
    if (-not (Test-Path -LiteralPath $ReferenceVbaRoot -PathType Container)) { throw 'ReferenceVbaRoot is missing' }
    $referenceRootCanonical = (Get-Canonical $ReferenceRoot).TrimEnd([char[]]"\/")
    $referenceVbaCanonical = Get-Canonical $ReferenceVbaRoot
    $referencePrefix = $referenceRootCanonical + [IO.Path]::DirectorySeparatorChar
    if (($referenceVbaCanonical -ne $referenceRootCanonical) -and (-not $referenceVbaCanonical.StartsWith($referencePrefix, [StringComparison]::OrdinalIgnoreCase))) { throw 'ReferenceVbaRoot must be contained within ReferenceRoot' }
    $referenceVbaMode = 'archived-source'
    try { Export-ReferenceVbaRoot $ReferenceVbaRoot $referenceVbaDir $records } catch { [void]$extractionErrors.Add($_.Exception.Message) }
  } elseif (-not [string]::IsNullOrWhiteSpace($ReferenceCandidate)) {
    $referenceVbaMode = 'excel-com'
    try { Export-ReferenceCandidate $ReferenceCandidate $referenceVbaDir 'reference-vba' $records } catch { [void]$extractionErrors.Add($_.Exception.Message) }
  }
}
$counts = [ordered]@{ candidate_package=@($records | Where-Object {$_.subject -eq 'candidate-package'}).Count; candidate_vba=@($records | Where-Object {$_.subject -eq 'candidate-vba'}).Count; reference_package=@($records | Where-Object {$_.subject -eq 'reference-package'}).Count; reference_vba=@($records | Where-Object {$_.subject -eq 'reference-vba'}).Count; reference_root=@($records | Where-Object {$_.subject -eq 'reference-root'}).Count }
$requiredCounts = if ($FrozenMode) { @('candidate_package','candidate_vba','reference_package','reference_vba','reference_root') } else { @('candidate_package') }
$status = 'PASS'
foreach ($required in $requiredCounts) { if ([int]$counts[$required] -eq 0) { $status = 'INCOMPLETE' } }
if ($extractionErrors.Count -gt 0) { $status = 'INCOMPLETE' }
$inventoryRows = @($records | Where-Object {$_.subject -eq 'reference-root'} | ForEach-Object { [ordered]@{ path=[string]$_.inventory_path; sha256=$_.sha256; size=[Int64]$_.size } } | Sort-Object @{ Expression = { Get-Utf8SortKey ([string]$_.path) } })
$inventoryCanonical = ''
foreach ($row in $inventoryRows) { $inventoryCanonical += [string]$row.path + "`n" + [string]$row.sha256 + "`n" + [string][Int64]$row.size + "`n" }
$inventorySha = if ($inventoryRows.Count -gt 0) { $bytes = [Text.Encoding]::UTF8.GetBytes($inventoryCanonical); $hash = [Security.Cryptography.SHA256]::Create(); try { ([BitConverter]::ToString($hash.ComputeHash($bytes))).Replace('-','').ToLowerInvariant() } finally { $hash.Dispose() } } else { '' }
$manifest = [ordered]@{ mode=if ($FixtureMode) { 'fixture' } else { 'frozen' }; candidate=(Get-Canonical $Candidate); candidate_sha256=$candidateSha; report=(Get-Canonical $report); reference_access=$FrozenMode; reference_identity=$ReferenceIdentity; reference_candidate_sha256=$referenceCandidateSha; reference_vba_mode=$referenceVbaMode; source_commit=$sourceCommit; source_tree=$sourceTree; reference_root_inventory=[ordered]@{ records=$inventoryRows; inventory_sha256=$inventorySha; root_sha256=$inventorySha }; status=$status; extraction_errors=@($extractionErrors); subjects=$records; complete_counts=$counts; tool_version=$ToolVersion; access_utc=[DateTime]::UtcNow.ToString('o') }
$manifest | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $root 'access-manifest.json') -Encoding UTF8
if ($status -ne 'PASS') { throw ('Independence extraction INCOMPLETE: ' + (@($extractionErrors) -join '; ')) }
Write-Host ('PASS independence extraction: ' + $report)
