param(
    [Parameter(Mandatory=$true)][string]$PackageZip,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [ValidatePattern('^v0\.01_r[0-9]+$')][string]$Version = 'v0.01_r105'
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'Artifact-Antivirus.ps1')
if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) {
    throw 'Use Windows PowerShell 5.1 (powershell.exe) for installer build verification'
}
$PackageZip = [IO.Path]::GetFullPath($PackageZip)
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath $PackageZip -PathType Leaf)) { throw 'Package ZIP is required' }
if (Test-Path -LiteralPath $OutputPath) { throw 'Refusing to overwrite installer' }
$framework = Join-Path $env:windir 'Microsoft.NET/Framework/v4.0.30319'
$compiler = Join-Path $framework 'csc.exe'
if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) { throw 'Installed Framework compiler is missing' }
$compilerSignature = Get-AuthenticodeSignature -LiteralPath $compiler
if ($compilerSignature.Status -ne 'Valid' -or $null -eq $compilerSignature.SignerCertificate -or
    $compilerSignature.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation(,|$)') {
    throw 'Installed Framework compiler must have a valid Microsoft publisher signature'
}
$source = Join-Path (Split-Path $PSScriptRoot -Parent) 'src/dotnet/NxSetup'
$stage = Join-Path ([IO.Path]::GetTempPath()) ('nx-setup-build-' + [Guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $stage)
$encoding = New-Object Text.UTF8Encoding($false)
[IO.File]::WriteAllText((Join-Path $stage 'version.txt'), $Version, $encoding)
$hash = (Get-FileHash -LiteralPath $PackageZip -Algorithm SHA256).Hash.ToLowerInvariant()
[IO.File]::WriteAllText((Join-Path $stage 'package.sha256'), $hash, $encoding)
$candidate = Join-Path $stage 'NxSetup.exe'
$arguments = @('/nologo','/target:winexe','/platform:x86','/optimize+',
    ('/out:' + $candidate), ('/win32manifest:' + (Join-Path $source 'app.manifest')),
    ('/win32icon:' + (Join-Path $source 'Assets/lhexcel.ico')),
    ('/resource:' + (Join-Path $source 'Assets/lhexcel.ico') + ',NxSetup.Brand.Icon'),
    ('/resource:' + (Join-Path $source 'Assets/brand-48.png') + ',NxSetup.Brand.Mark'),
    ('/resource:' + $PackageZip + ',NxSetup.Package.zip'),
    ('/resource:' + (Join-Path $stage 'version.txt') + ',NxSetup.Version'),
    ('/resource:' + (Join-Path $stage 'package.sha256') + ',NxSetup.Package.sha256'))
foreach ($name in @('System.dll','System.Core.dll','System.Xml.dll','System.Xml.Linq.dll','System.Drawing.dll','System.Windows.Forms.dll','System.IO.Compression.dll')) {
    $arguments += '/reference:' + (Join-Path $framework $name)
}
$sources = @(Get-ChildItem -LiteralPath $source -Filter '*.cs' | Sort-Object Name)
$sourceHashes = @(@($sources) + @(Get-Item -LiteralPath (Join-Path $source 'app.manifest')) + @(Get-ChildItem -LiteralPath (Join-Path $source 'Assets') -File) | ForEach-Object {
    [pscustomobject]@{ name=$_.Name; sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant() }
})
$arguments += @($sources | ForEach-Object FullName)
& $compiler @arguments
if ($LASTEXITCODE -ne 0) { throw ('Setup compile failed; preserved at ' + $stage) }
$preflight = Invoke-NxArtifactAntivirus -FilePath $candidate -ReportPath (Join-Path $stage 'antivirus-before-execution.json')
$check = Start-Process -FilePath $candidate -ArgumentList '--verify-only' -PassThru -Wait -WindowStyle Hidden
if ($check.ExitCode -ne 0) { throw ('Embedded package verification failed; preserved at ' + $stage) }
$postflight = Invoke-NxArtifactAntivirus -FilePath $candidate -ReportPath (Join-Path $stage 'antivirus-before-publication.json')
if ($preflight.sha256_after -ne $postflight.sha256_after) { throw 'Installer changed during verification' }
[void](New-Item -ItemType Directory -Path (Split-Path $OutputPath -Parent) -Force)
[IO.File]::Copy($candidate, $OutputPath, $false)
$publishedHash = (Get-FileHash -LiteralPath $OutputPath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($publishedHash -ne $postflight.sha256_after) { throw 'Published installer differs from scanned bytes' }
$receipt = [pscustomobject]@{ status='COMPILED_EMBEDDED_PACKAGE_VERIFIED'; version=$Version; signed=$false;
    installer=$OutputPath; installer_sha256=$publishedHash;
    package_sha256=$hash; installed_acceptance='NOT_RUN'; antivirus='LOCAL_DEFENDER_PASS';
    security_scope='Current local definitions only; not a publisher identity or a safety guarantee';
    compiler_signature=[string]$compilerSignature.Status; compiler_sha256=(Get-FileHash -LiteralPath $compiler -Algorithm SHA256).Hash.ToLowerInvariant();
    source_hashes=$sourceHashes; evidence_directory=$stage }
$receiptJson = $receipt | ConvertTo-Json -Depth 5
[IO.File]::WriteAllText((Join-Path $stage 'build-receipt.json'), $receiptJson, $encoding)
$receiptJson
