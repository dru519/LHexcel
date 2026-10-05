param([Parameter(Mandatory=$true)][string]$EvidenceRoot)
$ErrorActionPreference='Stop'
$source=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Fresh evidence required'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$framework=Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319'
$binary=Join-Path $EvidenceRoot 'ResultsProbe.exe'
$args=@('/nologo','/target:exe','/platform:x64',('/out:'+$binary))
foreach($name in @('System.dll','System.Drawing.dll','System.Windows.Forms.dll','System.Xml.dll','System.IO.Compression.dll','System.IO.Compression.FileSystem.dll','System.Web.Extensions.dll')){
    $args+='/reference:'+(Join-Path $framework $name)
}
foreach($name in @('Extensibility.dll','Office.dll','Microsoft.Office.Interop.Excel.dll')){
    $found=$null
    foreach($root in @((Join-Path $env:WINDIR 'assembly'),(Join-Path $env:WINDIR 'Microsoft.NET/assembly'))){
        $found=Get-ChildItem -LiteralPath $root -Recurse -Filter $name | Sort-Object FullName | Select-Object -First 1 -ExpandProperty FullName
        if($found){break}
    }
    if(-not $found){throw ('Missing interop '+$name)}
    $args+='/reference:'+$found
}
$args+=Get-ChildItem (Join-Path $source 'src/dotnet/NxHost') -Recurse -Filter '*.cs' | Sort-Object FullName | ForEach-Object FullName
$args+=Join-Path $PSScriptRoot 'R67ResultsProbe.cs'
& (Join-Path $framework 'csc.exe') @args
if($LASTEXITCODE -ne 0){throw 'Results probe compile failed'}
& $binary | Tee-Object -FilePath (Join-Path $EvidenceRoot 'results.txt')
if($LASTEXITCODE -ne 0){throw 'Results probe failed'}
