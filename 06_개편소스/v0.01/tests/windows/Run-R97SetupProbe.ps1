param([Parameter(Mandatory=$true)][string]$PackageZip,[Parameter(Mandatory=$true)][string]$PackageVersion,[Parameter(Mandatory=$true)][string]$EvidenceRoot)
$ErrorActionPreference='Stop'
$source=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Fresh evidence required'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$framework=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319'
$binary=Join-Path $EvidenceRoot 'SetupProbe.exe'
$arguments=@('/nologo','/target:exe','/platform:x64','/main:NxSetupR97Tests',('/out:'+$binary))
foreach($name in @('System.dll','System.Core.dll','System.Xml.dll','System.Xml.Linq.dll','System.Drawing.dll','System.Windows.Forms.dll','System.IO.Compression.dll')){$arguments+='/reference:'+(Join-Path $framework $name)}
$arguments+=Get-ChildItem (Join-Path $source 'src/dotnet/NxSetup') -Filter '*.cs' | Sort-Object Name | ForEach-Object FullName
$arguments+=Join-Path $source 'tests/dotnet/NxSetupR97Tests.cs'
& (Join-Path $framework 'csc.exe') @arguments
if($LASTEXITCODE -ne 0){throw 'Setup probe compile failed'}
& $binary $PackageZip $PackageVersion | Tee-Object -FilePath (Join-Path $EvidenceRoot 'results.txt')
if($LASTEXITCODE -ne 0){throw 'Setup probe failed'}
