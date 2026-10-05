[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PackageZip,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^v0\.01_r[0-9]+$')]
    [string]$Version
)

$ErrorActionPreference = 'Stop'
$sourceRoot = Split-Path -Parent $PSScriptRoot
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
$packagePath = (Resolve-Path -LiteralPath $PackageZip).Path
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('nx-r73-setup-' + [Guid]::NewGuid().ToString('N'))
$nativeRoot = Join-Path $testRoot 'native'
$packageRoot = Join-Path $testRoot 'package'

if (-not (Test-Path -LiteralPath $compiler -PathType Leaf)) {
    throw "Required .NET Framework compiler not found: $compiler"
}

New-Item -ItemType Directory -Path $nativeRoot, $packageRoot | Out-Null
Copy-Item -LiteralPath (Join-Path $sourceRoot 'src\dotnet\NxSetup\Package.cs') -Destination $nativeRoot
Copy-Item -LiteralPath (Join-Path $sourceRoot 'src\dotnet\NxSetup\NativeInstaller.cs') -Destination $nativeRoot
Copy-Item -LiteralPath (Join-Path $sourceRoot 'tests\dotnet\NxSetupNativeTests.cs') -Destination $nativeRoot
Copy-Item -LiteralPath (Join-Path $sourceRoot 'src\dotnet\NxSetup\Package.cs') -Destination $packageRoot
Copy-Item -LiteralPath (Join-Path $sourceRoot 'tests\dotnet\NxSetupPackageTests.cs') -Destination $packageRoot

Push-Location $nativeRoot
try {
    & $compiler /nologo /target:exe /platform:x86 /out:NativeTests.exe `
        /r:System.Xml.Linq.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll `
        Package.cs NativeInstaller.cs NxSetupNativeTests.cs
    if ($LASTEXITCODE -ne 0) { throw "Native test compilation failed: $LASTEXITCODE" }
    & (Join-Path $nativeRoot 'NativeTests.exe') $packagePath $Version
    if ($LASTEXITCODE -ne 0) { throw "Native tests failed: $LASTEXITCODE" }
}
finally { Pop-Location }

Push-Location $packageRoot
try {
    & $compiler /nologo /target:exe /platform:x86 /out:PackageTests.exe `
        /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll `
        Package.cs NxSetupPackageTests.cs
    if ($LASTEXITCODE -ne 0) { throw "Package test compilation failed: $LASTEXITCODE" }
    & (Join-Path $packageRoot 'PackageTests.exe')
    if ($LASTEXITCODE -ne 0) { throw "Package tests failed: $LASTEXITCODE" }
}
finally { Pop-Location }

Write-Output "PASS|R73SetupTests|X86_PROCESS|package=$packagePath|version=$Version|artifacts=$testRoot"
