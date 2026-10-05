param([Parameter(Mandatory=$true)][string]$EvidenceRoot)
$ErrorActionPreference='Stop'
# PowerShell can be ARM64 while Excel is emulated x64. Probe explicit processes.
$compiler=Join-Path $env:windir 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
foreach($architecture in @('x86','x64')){
    $probe=Join-Path $EvidenceRoot ('NavigatorComProbe-'+$architecture+'.exe')
    & $compiler /nologo /target:exe ("/platform:"+$architecture) ("/out:"+$probe) (Join-Path $PSScriptRoot 'NxNavigatorComProbe.cs')
    if($LASTEXITCODE -ne 0){throw 'COM probe compilation failed'}
    $output=@(& $probe)
    $exitCode=$LASTEXITCODE
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot ('navigator-com-'+$architecture+'.txt')),($output -join "`n"),(New-Object Text.UTF8Encoding($false)))
    if($exitCode -ne 0){throw ('COM probe failed: '+$architecture+' '+($output -join '|'))}
}
