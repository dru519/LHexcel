param([Parameter(Mandatory=$true)][string]$EvidenceRoot,[string]$ServiceSource='')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Fresh evidence required' }
New-Item -ItemType Directory -Path $EvidenceRoot | Out-Null
$source=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$gac=@()
foreach($name in @('Extensibility.dll','Office.dll','Microsoft.Office.Interop.Excel.dll')) {
    $found=@(foreach($root in @("$env:windir/assembly","$env:windir/Microsoft.NET/assembly")) {
        Get-ChildItem -LiteralPath $root -Recurse -Filter $name -ErrorAction SilentlyContinue
    }) | Select-Object -First 1
    if($null -eq $found) { throw ('Reference missing: '+$name) }
    $gac+=$found.FullName
}
if(!$ServiceSource) { $ServiceSource=Join-Path $source 'src/dotnet/NxHost/WorkbookCompareService.cs' }
Get-FileHash -LiteralPath $ServiceSource | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'service-hash.json') -Encoding UTF8
foreach($arch in @('x86','x64')) {
    $framework=if($arch -eq 'x86') { "$env:windir/Microsoft.NET/Framework/v4.0.30319" } else { "$env:windir/Microsoft.NET/Framework64/v4.0.30319" }
    $exe=Join-Path $EvidenceRoot ("probe-"+$arch+".exe")
    $arguments=@('/nologo','/target:exe',('/platform:'+$arch),'/optimize+',('/out:'+$exe))
    foreach($name in @('System.dll','System.Drawing.dll','System.Windows.Forms.dll','System.Xml.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll','System.Web.Extensions.dll')) { $arguments+=('/reference:'+(Join-Path $framework $name)) }
    foreach($path in $gac) { $arguments+=('/reference:'+$path) }
    $arguments+=@(Get-ChildItem -LiteralPath (Join-Path $source 'src/dotnet/NxHost') -Filter '*.cs' -Recurse | Where-Object Name -ne 'WorkbookCompareService.cs' | ForEach-Object FullName)
    $arguments+=$ServiceSource
    $arguments+=(Join-Path $PSScriptRoot 'R66CompareProbe.cs')
    & (Join-Path $framework 'csc.exe') @arguments
    if($LASTEXITCODE -ne 0) { throw 'Probe compile failed' }
    & $exe (Join-Path $EvidenceRoot ($arch+'.json')) | Tee-Object -FilePath (Join-Path $EvidenceRoot ($arch+'.txt'))
    if($LASTEXITCODE -ne 0) { throw 'Probe failed' }
}
