param([string]$OutputRoot = (Join-Path $PSScriptRoot 'out/NxHost'), [ValidateSet('Debug','Release')][string]$Configuration = 'Release', [switch]$ProtectedLoader)
$ErrorActionPreference = 'Stop'; Set-StrictMode -Version Latest
$source = Split-Path $PSScriptRoot -Parent
$hostRoot = Join-Path $source 'src/dotnet/NxHost'
$out = [IO.Path]::GetFullPath($OutputRoot)
$csc32 = Join-Path $env:windir 'Microsoft.NET/Framework/v4.0.30319/csc.exe'
$csc64 = Join-Path $env:windir 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
if (!(Test-Path $csc32) -or !(Test-Path $csc64)) { throw 'installed .NET Framework 4.8 csc.exe not found' }
$framework32 = Join-Path $env:windir 'Microsoft.NET\Framework\v4.0.30319'
$framework64 = Join-Path $env:windir 'Microsoft.NET\Framework64\v4.0.30319'
$gacRoots = @(
  (Join-Path $env:windir 'assembly'),
  (Join-Path $env:windir 'Microsoft.NET\assembly')
)
$gac = @()
foreach ($assemblyName in @('Extensibility.dll', 'Office.dll', 'Microsoft.Office.Interop.Excel.dll')) {
  $match = $null
  foreach ($gacRoot in $gacRoots) {
    if (Test-Path -LiteralPath $gacRoot -PathType Container) {
      $match = Get-ChildItem -LiteralPath $gacRoot -Recurse -Filter $assemblyName -ErrorAction SilentlyContinue |
        Sort-Object FullName | Select-Object -First 1 -ExpandProperty FullName
      if ($match) { break }
    }
  }
  if (-not $match) { throw ('required Office interop assembly not found: ' + $assemblyName) }
  $gac += $match
}
foreach($arch in @(@{Name='x86'; Csc=$csc32; Dll='NxHost32.dll'}, @{Name='x64'; Csc=$csc64; Dll='NxHost64.dll'})) {
  $dir = Join-Path $out $arch.Name; New-Item -ItemType Directory -Path $dir -Force | Out-Null
  $platform = if ($arch.Name -eq 'x86') { 'x86' } else { 'x64' }
  $frameworkRoot = if ($arch.Name -eq 'x86') { $framework32 } else { $framework64 }
  $args = @('/nologo','/target:library',('/platform:' + $platform),'/optimize+',('/out:' + (Join-Path $dir $arch.Dll)))
  foreach($referenceName in @('System.dll','System.Drawing.dll','System.Windows.Forms.dll','System.Xml.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll','System.Web.Extensions.dll')) {
    $referencePath = Join-Path $frameworkRoot $referenceName
    if (-not (Test-Path -LiteralPath $referencePath -PathType Leaf)) { throw ('framework reference missing: ' + $referencePath) }
    $args += '/reference:' + $referencePath
  }
  foreach($referencePath in $gac){$args += '/reference:' + $referencePath}
  $references = @($args | Where-Object {$_ -like '/reference:*'})
  if ($ProtectedLoader) {
    $corePath=Join-Path $dir ($arch.Dll.Replace('NxHost','NxCore'))
    $args=@($args | Where-Object {$_ -notlike '/out:*'}) + ('/out:'+$corePath)
  }
  foreach($asset in (Get-ChildItem (Join-Path $hostRoot 'Assets') -Filter '*.png' | Sort-Object Name)) {
    $args += '/resource:' + $asset.FullName + ',NxHost.Brand.' + $asset.Name
  }
  $args += (Get-ChildItem $hostRoot -Recurse -Filter *.cs | Sort-Object FullName | ForEach-Object FullName)
  & $arch.Csc @args; if($LASTEXITCODE -ne 0){throw "NxHost $($arch.Name) compilation failed"}
  if ($ProtectedLoader) {
    $metadata=[Reflection.Assembly]::ReflectionOnlyLoadFrom($corePath)
    $generated=Join-Path $out ('generated/'+$arch.Name)
    & python -X utf8 (Join-Path $source 'tools/generate_protected_loader.py') --core $corePath --mvid $metadata.ManifestModule.ModuleVersionId.ToString('D') --out $generated
    if($LASTEXITCODE -ne 0){throw 'Protected loader generation failed'}
    $loaderArgs=@('/nologo','/target:library',('/platform:'+$platform),'/optimize+',('/out:'+(Join-Path $dir $arch.Dll)))+$references
    $loaderArgs+=(Get-ChildItem (Join-Path $source 'src/dotnet/NxLoader') -Filter *.cs | ForEach-Object FullName)
    $loaderArgs+=(Get-ChildItem $generated -Filter *.cs | ForEach-Object FullName)
    & $arch.Csc @loaderArgs
    if($LASTEXITCODE -ne 0){throw 'Protected loader compilation failed'}
  }
}
Write-Output ('PASS|NxHostBuild|'+$out)
# enhanced-dll emits out/NxHost/x86/NxHost32.dll and out/NxHost/x64/NxHost64.dll
# internal-xlam profile intentionally emits no host sidecar or registration.
# PlatformTarget is x86/x64 for deterministic framework csc builds.
