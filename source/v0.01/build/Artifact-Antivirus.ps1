# Build-time local evidence only; this is not a publisher signature or a safety guarantee.
function Get-NxAntivirusVerdict {
    param([Parameter(Mandatory=$true)]$Evidence)
    if ($Evidence.error) { return 'SCAN_ERROR' }
    if ($Evidence.scanner_trusted -ne $true) { return 'SCANNER_NOT_TRUSTED' }
    if ($Evidence.antivirus_before -ne $true -or $Evidence.antivirus_after -ne $true -or
        $Evidence.realtime_before -ne $true -or $Evidence.realtime_after -ne $true) { return 'PROTECTION_NOT_ACTIVE' }
    foreach ($age in @($Evidence.signature_age_hours_before, $Evidence.signature_age_hours_after)) {
        if ($null -eq $age -or $age -lt 0 -or $age -gt 48) { return 'SIGNATURES_NOT_CURRENT' }
    }
    if ($null -eq $Evidence.scan_exit_code -or $Evidence.scan_exit_code -ne 0) { return 'SCAN_NOT_CLEAN' }
    # Fail closed on unrecognized or localized output, not just a successful process exit.
    if ($Evidence.scan_output -notmatch '(?im)\bfound no threats\.\s*$') { return 'CLEAN_RESULT_NOT_CONFIRMED' }
    if ($Evidence.sha256_before -notmatch '^[0-9a-fA-F]{64}$' -or
        $Evidence.sha256_after -ne $Evidence.sha256_before) { return 'ARTIFACT_CHANGED_OR_MISSING' }
    if (@($Evidence.new_threats).Count -gt 0) { return 'THREAT_RECORDED' }
    return 'LOCAL_DEFENDER_PASS'
}

function Invoke-NxArtifactAntivirus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [Parameter(Mandatory=$true)][string]$ReportPath
    )
    $ErrorActionPreference = 'Stop'
    Set-StrictMode -Version Latest
    $FilePath = [IO.Path]::GetFullPath($FilePath)
    $ReportPath = [IO.Path]::GetFullPath($ReportPath)
    if (Test-Path -LiteralPath $ReportPath) { throw 'Existing antivirus evidence must not be overwritten' }
    $started = [DateTime]::UtcNow
    $evidence = [ordered]@{
        status='HOLD'; scope='LOCAL_DEFENDER_SCAN_ONLY'; file=$FilePath; started_utc=$started.ToString('o');
        completed_utc=$null; error=$null; scanner=$null; scanner_trusted=$false; scanner_sha256=$null;
        antivirus_before=$false; antivirus_after=$false; realtime_before=$false; realtime_after=$false;
        signature_age_hours_before=$null; signature_age_hours_after=$null;
        definitions_before=$null; definitions_after=$null; engine=$null;
        scan_exit_code=$null; scan_output=''; sha256_before=$null; sha256_after=$null;
        new_threats=@(); artifact_signature=$null;
        scan_mode='CUSTOM_DIAGNOSTIC_EXCLUSIONS_IGNORED'; protection_settings_changed=$false
    }
    try {
        if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) {
            throw 'Use Windows PowerShell 5.1 (powershell.exe) for Defender evidence'
        }
        if (-not (Test-Path -LiteralPath $FilePath -PathType Leaf)) { throw 'Scan target is missing' }
        Import-Module (Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/Modules/Defender') -ErrorAction Stop
        $before = Get-MpComputerStatus -ErrorAction Stop
        $evidence.antivirus_before = $before.AntivirusEnabled
        $evidence.realtime_before = $before.RealTimeProtectionEnabled
        $evidence.signature_age_hours_before = ($started - $before.AntivirusSignatureLastUpdated.ToUniversalTime()).TotalHours
        $evidence.definitions_before = $before.AntivirusSignatureVersion
        $evidence.engine = $before.AMEngineVersion
        if (-not $evidence.antivirus_before -or -not $evidence.realtime_before) { throw 'Defender protection is not active' }
        if ($evidence.signature_age_hours_before -lt 0 -or $evidence.signature_age_hours_before -gt 48) {
            throw 'Defender definitions must be no older than 48 hours; update through the approved policy'
        }
        $platform = Join-Path $env:ProgramData 'Microsoft/Windows Defender/Platform'
        $versions = @(Get-ChildItem -LiteralPath $platform -Directory -ErrorAction Stop |
            Where-Object { $_.Name -match '^\d+\.\d+\.\d+\.\d+-\d+$' } |
            Sort-Object { [version]($_.Name.Split('-')[0]) } -Descending)
        if ($versions.Count -eq 0) { throw 'Installed Defender platform was not found' }
        $scanner = Join-Path $versions[0].FullName 'MpCmdRun.exe'
        $signature = Get-AuthenticodeSignature -LiteralPath $scanner
        $evidence.scanner = $scanner
        $evidence.scanner_trusted = ($signature.Status -eq 'Valid' -and
            $null -ne $signature.SignerCertificate -and $signature.SignerCertificate.Subject -match 'O=Microsoft Corporation(,|$)')
        if (-not $evidence.scanner_trusted) { throw 'Defender scanner publisher signature is not valid' }
        $evidence.scanner_sha256 = (Get-FileHash -LiteralPath $scanner -Algorithm SHA256).Hash.ToLowerInvariant()
        $evidence.sha256_before = (Get-FileHash -LiteralPath $FilePath -Algorithm SHA256).Hash.ToLowerInvariant()
        $evidence.artifact_signature = [string](Get-AuthenticodeSignature -LiteralPath $FilePath).Status
        # This switch applies to this one custom scan, NOT to real-time protection or
        # security preferences. It inspects excluded paths and archives, reports threats
        # to stdout, and does not remediate them. Never execute a rejected artifact.
        $output = & $scanner -Scan -ScanType 3 -File $FilePath -DisableRemediation 2>&1
        $evidence.scan_exit_code = $LASTEXITCODE
        $evidence.scan_output = ($output | Out-String).Trim()
        $evidence.sha256_after = (Get-FileHash -LiteralPath $FilePath -Algorithm SHA256).Hash.ToLowerInvariant()
        $after = Get-MpComputerStatus -ErrorAction Stop
        $evidence.antivirus_after = $after.AntivirusEnabled
        $evidence.realtime_after = $after.RealTimeProtectionEnabled
        $evidence.signature_age_hours_after = ([DateTime]::UtcNow - $after.AntivirusSignatureLastUpdated.ToUniversalTime()).TotalHours
        $evidence.definitions_after = $after.AntivirusSignatureVersion
        # Diagnostic scan detections are in stdout; include any concurrent real-time detection too.
        $evidence.new_threats = @(Get-MpThreatDetection -ErrorAction Stop | Where-Object {
            ($_.InitialDetectionTime.ToUniversalTime() -ge $started -or
             $_.LastThreatStatusChangeTime.ToUniversalTime() -ge $started) -and
            (($_.Resources -join "`n").IndexOf($FilePath, [StringComparison]::OrdinalIgnoreCase) -ge 0)
        } | Select-Object ThreatID, InitialDetectionTime, LastThreatStatusChangeTime, ActionSuccess, Resources)
    } catch {
        $evidence.error = $_.Exception.Message
    }
    $evidence.completed_utc = [DateTime]::UtcNow.ToString('o')
    $evidence.status = Get-NxAntivirusVerdict -Evidence ([pscustomobject]$evidence)
    [void](New-Item -ItemType Directory -Path (Split-Path $ReportPath -Parent) -Force)
    $json = $evidence | ConvertTo-Json -Depth 6
    $stream = [IO.File]::Open($ReportPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $bytes = (New-Object Text.UTF8Encoding($false)).GetBytes($json)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally { $stream.Dispose() }
    if ($evidence.status -ne 'LOCAL_DEFENDER_PASS') {
        throw ('Antivirus gate HOLD: ' + $evidence.status + '; evidence: ' + $ReportPath)
    }
    return [pscustomobject]$evidence
}
