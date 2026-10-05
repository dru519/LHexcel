param([Parameter(Mandatory=$true)][string]$SourceRoot,
      [Parameter(Mandatory=$true)][string]$HostRoot,
      [Parameter(Mandatory=$true)][string]$EvidenceRoot)
$ErrorActionPreference='Stop'
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Fresh loader evidence required'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$compiler=Join-Path $env:windir 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
$probe=Join-Path $EvidenceRoot 'Probe.exe'
& $compiler /nologo /target:exe /platform:x64 ("/out:"+$probe) (Join-Path $SourceRoot 'tests/dotnet/ProtectedLoaderProbe.cs')
if($LASTEXITCODE -ne 0){throw 'Loader probe compile failed'}
$results=@()
foreach($mode in @('normal','missing','changed','binding_collision')) {
    $folder=Join-Path $EvidenceRoot $mode
    [void](New-Item -ItemType Directory -Path $folder)
    $loader=Join-Path $folder 'NxHost64.dll'
    $core=Join-Path $folder 'NxCore64.dll'
    Copy-Item -LiteralPath (Join-Path $HostRoot 'x64/NxHost64.dll') -Destination $loader
    if($mode -ne 'missing'){Copy-Item -LiteralPath (Join-Path $HostRoot 'x64/NxCore64.dll') -Destination $core}
    if($mode -eq 'changed'){
        $bytes=[IO.File]::ReadAllBytes($core)
        $bytes[$bytes.Length-1]=$bytes[$bytes.Length-1] -bxor 1
        [IO.File]::WriteAllBytes($core,$bytes)
    }
    if($mode -eq 'binding_collision'){
        $other=Join-Path $folder 'other'
        [void](New-Item -ItemType Directory -Path $other)
        Copy-Item -LiteralPath $core -Destination (Join-Path $other 'NxCore64.dll')
    }
    $line=& $probe $mode $loader
    $results+=@{case=$mode;exit_code=$LASTEXITCODE;result=[string]$line}
    Write-Output $line
}
@{evidence_class='windows_netframework_x64_loader_no_excel';cases=$results} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'report.json') -Encoding UTF8
if(@($results|Where-Object {$_.exit_code -ne 0}).Count){throw 'Loader probe failed'}
