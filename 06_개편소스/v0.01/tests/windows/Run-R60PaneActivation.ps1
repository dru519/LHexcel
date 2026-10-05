param([Parameter(Mandatory=$true)][string]$EvidenceRoot)
$ErrorActionPreference='Stop'
$pane=$null;$failure=$null
try{$pane=New-Object -ComObject LH.MyExcel.DocNavCtpProbe.Pane}
catch{$failure=$_.Exception.ToString()}
finally{
    if($null -ne $pane -and [Runtime.InteropServices.Marshal]::IsComObject($pane)){
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($pane)
    }
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'pane-activation.json'),([pscustomobject]@{failure=$failure;host_bits=([IntPtr]::Size*8)}|ConvertTo-Json))
}
# Exit the host to unload the managed assembly before the parent restores registration.
if($null -ne $failure){exit 1}
