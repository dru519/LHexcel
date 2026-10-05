param(
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [string]$EvidenceRoot = '',
    [string]$RunId = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }

$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$DataRoot = [IO.Path]::GetFullPath($DataRoot)
if ($DataRoot -ne $SourceRoot) { throw 'DataRoot must be the current v0.01 source root' }
if ([string]::IsNullOrWhiteSpace($RunId)) { $RunId = [guid]::NewGuid().ToString() }
if ($RunId -notmatch '^[0-9a-fA-F-]{36}$') { throw 'RunId must be a GUID' }

$verifyRoot = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'LHExcel\Verify'))
$EvidenceRoot = if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) {
    Join-Path $verifyRoot ('r57-xlstart-' + $RunId)
} else {
    [IO.Path]::GetFullPath($EvidenceRoot)
}
if (-not $EvidenceRoot.StartsWith($verifyRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'EvidenceRoot must stay under LOCALAPPDATA LHExcel Verify'
}
if (Test-Path -LiteralPath $EvidenceRoot) {
    if (-not (Test-Path -LiteralPath $EvidenceRoot -PathType Container)) { throw 'EvidenceRoot must be a directory' }
    if (@(Get-ChildItem -LiteralPath $EvidenceRoot -Force).Count -ne 0) { throw 'Existing non-empty EvidenceRoot is rejected' }
} else {
    [void](New-Item -ItemType Directory -Path $EvidenceRoot -Force)
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')

function Assert-ChildPath([string]$Path, [string]$Root, [switch]$AllowRoot) {
    $full = [IO.Path]::GetFullPath($Path)
    $boundary = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    if (($AllowRoot -and $full -eq $boundary) -or $full.StartsWith($boundary + '\', [StringComparison]::OrdinalIgnoreCase)) {
        return $full
    }
    throw ('Path escaped boundary: ' + $full)
}

function Get-Sha256([string]$Path) {
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
}

function Get-TextSha256([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').ToLowerInvariant()
    } finally { $sha.Dispose() }
}

function Get-StableDigest([object]$Value) {
    Get-TextSha256 ($Value | ConvertTo-Json -Depth 30 -Compress)
}

function Write-Receipt([string]$Name, [object]$Value) {
    $path = Assert-ChildPath (Join-Path $EvidenceRoot $Name) $EvidenceRoot
    $utf8 = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($path, ($Value | ConvertTo-Json -Depth 30), $utf8)
}

function Get-ExcelProcessSnapshot {
    $rows = @()
    foreach ($candidate in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
        try {
            [void]$candidate.Handle
            $rows += [pscustomobject][ordered]@{
                pid = [int]$candidate.Id
                started_utc = $candidate.StartTime.ToUniversalTime().ToString('o')
                executable = [IO.Path]::GetFullPath($candidate.MainModule.FileName)
            }
        } catch {
            $rows += [pscustomobject][ordered]@{pid=[int]$candidate.Id;started_utc=$null;executable=$null}
        } finally {
            try { $candidate.Dispose() } catch { }
        }
    }
    return @($rows | Sort-Object pid)
}

function Wait-ExcelProcessSnapshot([object]$Expected, [int]$TimeoutMs = 15000) {
    $expectedDigest = Get-StableDigest @($Expected)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $stableReads = 0
    $latest = @()
    do {
        $latest = @(Get-ExcelProcessSnapshot)
        if ((Get-StableDigest @($latest)) -ceq $expectedDigest) {
            $stableReads++
            if ($stableReads -ge 3) { return @($latest) }
        } else {
            $stableReads = 0
        }
        Start-Sleep -Milliseconds 200
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    return @($latest)
}

function Stop-RecordedRibbonExcel([string]$ReceiptPath, [string]$ExpectedRunId, [string]$ExpectedArtifact, [object]$Baseline) {
    if (-not (Test-Path -LiteralPath $ReceiptPath -PathType Leaf)) {
        return [pscustomobject][ordered]@{status='NOT_AVAILABLE';exit_mode='NOT_RECORDED';pid=$null;failure=$null}
    }
    $process = $null
    try {
        $receipt = Get-Content -LiteralPath $ReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([string]$receipt.run_id -cne $ExpectedRunId) { throw 'Ribbon ownership receipt run ID mismatch' }
        if (-not [string]::Equals([IO.Path]::GetFullPath([string]$receipt.artifact), [IO.Path]::GetFullPath($ExpectedArtifact), [StringComparison]::OrdinalIgnoreCase)) { throw 'Ribbon ownership receipt artifact mismatch' }
        $processId = [int]$receipt.pid
        if (@($Baseline | Where-Object { [int]$_.pid -eq $processId }).Count -ne 0) { throw 'Ribbon ownership receipt points to a baseline Excel process' }
        $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
        if ($null -eq $process) {
            return [pscustomobject][ordered]@{status='PASS';exit_mode='ALREADY_EXITED';pid=$processId;failure=$null}
        }
        [void]$process.Handle
        $startedUtc = $process.StartTime.ToUniversalTime().ToString('o')
        $executable = [IO.Path]::GetFullPath($process.MainModule.FileName)
        if ($startedUtc -cne [string]$receipt.started_utc -or -not [string]::Equals($executable, [IO.Path]::GetFullPath([string]$receipt.executable), [StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($executable) -cne 'EXCEL.EXE') {
            throw 'Ribbon ownership receipt no longer matches PID, start time, and EXCEL.EXE'
        }
        $cleanup = Stop-ExactProcessAfterGrace -Process $process -Label 'InternalXlamStartup recorded ribbon Excel' -GraceMs 10000 -Detailed
        if ($cleanup.failure) { throw [string]$cleanup.failure }
        return [pscustomobject][ordered]@{status='PASS';exit_mode=[string]$cleanup.exit_mode;pid=$processId;failure=$null}
    } catch {
        return [pscustomobject][ordered]@{status='FAIL';exit_mode='CONTAINMENT_FAILED';pid=if($null -eq $process){$null}else{[int]$process.Id};failure=$_.Exception.Message}
    } finally {
        if ($null -ne $process) { try { $process.Dispose() } catch { } }
    }
}

function Get-XlstartSnapshot([string]$Root) {
    if (-not (Test-Path -LiteralPath $Root)) { return @() }
    $rootItem = Get-Item -LiteralPath $Root -Force
    if (-not $rootItem.PSIsContainer) { throw 'XLSTART path is not a directory' }
    if (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'XLSTART root reparse point rejected' }
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $rows = @()
    foreach ($item in @(Get-ChildItem -LiteralPath $rootFull -Force -Recurse)) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw ('XLSTART reparse point rejected: ' + $item.FullName)
        }
        $full = Assert-ChildPath $item.FullName $rootFull
        $relative = $full.Substring($rootFull.Length).TrimStart('\')
        $rows += [pscustomobject][ordered]@{
            path = $relative
            kind = if ($item.PSIsContainer) { 'directory' } else { 'file' }
            length = if ($item.PSIsContainer) { 0 } else { [Int64]$item.Length }
            sha256 = if ($item.PSIsContainer) { $null } else { Get-Sha256 $full }
            attributes = [int]$item.Attributes
            last_write_utc = $item.LastWriteTimeUtc.ToString('o')
        }
    }
    return @($rows | Sort-Object path)
}

function Get-FileTreeDigest([string]$Root) {
    if (-not (Test-Path -LiteralPath $Root -PathType Container)) { return Get-StableDigest @() }
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $rows = @()
    foreach ($item in @(Get-ChildItem -LiteralPath $rootFull -Force -Recurse)) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw ('Protected tree reparse point rejected: ' + $item.FullName)
        }
        if (-not $item.PSIsContainer) {
            $full = Assert-ChildPath $item.FullName $rootFull
            $rows += [pscustomobject][ordered]@{
                path = $full.Substring($rootFull.Length).TrimStart('\')
                length = [Int64]$item.Length
                sha256 = Get-Sha256 $full
            }
        }
    }
    return Get-StableDigest @($rows | Sort-Object path)
}

function Convert-RegistryValue([object]$Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -is [byte[]]) { return [Convert]::ToBase64String($Value) }
    if ($Value -is [string[]]) { return @($Value) }
    return $Value
}

function Get-RegistryTreeRows(
    [Microsoft.Win32.RegistryHive]$Hive,
    [Microsoft.Win32.RegistryView]$View,
    [string]$SubKey,
    [string]$ValuePattern,
    [bool]$Recursive
) {
    $base = $null
    $queue = New-Object 'Collections.Generic.Queue[string]'
    $rows = @()
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
        $queue.Enqueue($SubKey)
        while ($queue.Count -gt 0) {
            $path = $queue.Dequeue()
            $key = $null
            try {
                $key = $base.OpenSubKey($path, $false)
                if ($null -eq $key) {
                    if ($path -eq $SubKey) {
                        $rows += [pscustomobject][ordered]@{
                            hive=[string]$Hive;view=[string]$View;path=$path;exists=$false
                            name=$null;kind=$null;value=$null
                        }
                    }
                    continue
                }
                $rows += [pscustomobject][ordered]@{
                    hive=[string]$Hive;view=[string]$View;path=$path;exists=$true
                    name=$null;kind=$null;value=$null
                }
                foreach ($name in @($key.GetValueNames() | Sort-Object)) {
                    if (-not [string]::IsNullOrWhiteSpace($ValuePattern) -and $name -notmatch $ValuePattern) { continue }
                    $rows += [pscustomobject][ordered]@{
                        hive=[string]$Hive;view=[string]$View;path=$path;exists=$true
                        name=[string]$name;kind=[string]$key.GetValueKind($name)
                        value=Convert-RegistryValue $key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                    }
                }
                if ($Recursive) {
                    foreach ($child in @($key.GetSubKeyNames() | Sort-Object)) {
                        $queue.Enqueue($path + '\' + $child)
                    }
                }
            } finally {
                if ($null -ne $key) { $key.Dispose() }
            }
        }
    } finally {
        if ($null -ne $base) { $base.Dispose() }
    }
    return @($rows)
}

function Get-RegistryBaseline {
    $rows = @()
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
        $rows += @(Get-RegistryTreeRows ([Microsoft.Win32.RegistryHive]::CurrentUser) $view 'Software\Microsoft\Office\16.0\Excel\Options' '^(OPEN[0-9]*|AltStartupPath)$' $false)
        foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
            foreach ($path in @(
                'Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect',
                'Software\Microsoft\Office\Excel\Addins\LHExcel.FocusEngine',
                'Software\Classes\LH.NxHost.BridgeService',
                'Software\Classes\LHExcel.FocusEngine'
            )) {
                $rows += @(Get-RegistryTreeRows $hive $view $path '' $true)
            }
        }
    }
    return @($rows | Sort-Object hive,view,path,name)
}

function Test-NxHostRegistrationAbsent([object[]]$RegistrySnapshot) {
    return @($RegistrySnapshot | Where-Object {
        [bool]$_.exists -and ([string]$_.path -match 'LH\.NxHost\.(Connect|BridgeService)')
    }).Count -eq 0
}

function Get-RegisteredExcelExecutable {
    $candidates = New-Object 'Collections.Generic.List[string]'
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
        $base = $null
        $key = $null
        try {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
            $key = $base.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', $false)
            if ($null -ne $key) {
                $raw = [string]$key.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                if (-not [string]::IsNullOrWhiteSpace($raw)) { [void]$candidates.Add($raw) }
            }
        } finally {
            if ($null -ne $key) { $key.Dispose() }
            if ($null -ne $base) { $base.Dispose() }
        }
    }
    foreach ($candidate in @($candidates | Sort-Object -Unique)) {
        $full = [IO.Path]::GetFullPath($candidate.Trim().Trim([char]34))
        if ((Test-Path -LiteralPath $full -PathType Leaf) -and [IO.Path]::GetFileName($full) -ceq 'EXCEL.EXE') {
            return $full
        }
    }
    throw 'Registered EXCEL.EXE path unavailable'
}

function Start-InteractiveStartupExcel([object]$Baseline, [string]$Label) {
    $excelPath = Get-RegisteredExcelExecutable
    $launch = $null
    $candidate = $null
    $binding = $null
    $transferred = $false
    try {
        $launch = Start-Process -FilePath $excelPath -ArgumentList @('/x') -WindowStyle Normal -PassThru
        if ($null -eq $launch) { throw ($Label + ' launch failed') }
        for ($attempt = 0; $attempt -lt 150; $attempt++) {
            try {
                $candidate = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
                $binding = Get-ExactExcelProcessOwnership $candidate $Baseline $Label
                if ([bool]$binding.owned -and [int]$binding.pid -eq [int]$launch.Id -and [bool]$candidate.UserControl) {
                    $result = [pscustomobject]@{excel=$candidate;binding=$binding}
                    $candidate = $null
                    $binding = $null
                    $transferred = $true
                    return $result
                }
            } catch { }
            finally {
                if (-not $transferred) {
                    if ($null -ne $binding -and $null -ne $binding.process) { try { $binding.process.Dispose() } catch { } }
                    Release-ComObject $candidate
                    $binding = $null
                    $candidate = $null
                }
            }
            Start-Sleep -Milliseconds 200
        }
        throw ($Label + ' COM binding timed out')
    } finally {
        if (-not $transferred -and $null -ne $launch) {
            try {
                if (-not $launch.HasExited) {
                    $cleanup = Stop-ExactProcessAfterGrace -Process $launch -Label $Label -GraceMs 10000 -Detailed
                    if ($cleanup.failure) { throw [string]$cleanup.failure }
                }
            } finally { $launch.Dispose() }
        } elseif ($null -ne $launch) {
            $launch.Dispose()
        }
    }
}

function Resolve-StartupAddinProject([object]$Excel, [string]$ExpectedPath) {
    $expected = [IO.Path]::GetFullPath($ExpectedPath)
    $match = $null
    $vbe = $null
    $projects = $null
    try {
        $vbe = $Excel.VBE
        $projects = $vbe.VBProjects
        for ($index = 1; $index -le [int]$projects.Count; $index++) {
            $candidate = $null
            try {
                $candidate = $projects.Item($index)
                $candidateFile = [string]$candidate.FileName
                if ([string]::IsNullOrWhiteSpace($candidateFile)) { continue }
                if ([string]::Equals([IO.Path]::GetFullPath($candidateFile), $expected, [StringComparison]::OrdinalIgnoreCase)) {
                    if ($null -ne $match) {
                        Release-ComObject $match
                        $match = $null
                        throw 'Startup Product XLAM loaded more than once'
                    }
                    $match = $candidate
                    $candidate = $null
                }
            } finally { Release-ComObject $candidate }
        }
    } finally {
        Release-ComObject $projects
        Release-ComObject $vbe
    }
    if ($null -eq $match) { throw 'Startup Product XLAM was not auto-loaded' }
    return $match
}

function Resolve-HostWorkbook([object]$Books) {
    $hostWorkbook = $null
    for ($index = 1; $index -le [int]$Books.Count; $index++) {
        $candidate = $null
        try {
            $candidate = $Books.Item($index)
            if (-not [bool]$candidate.IsAddin) {
                if ($null -ne $hostWorkbook) {
                    Release-ComObject $hostWorkbook
                    $hostWorkbook = $null
                    throw 'Startup host workbook count rejected'
                }
                $hostWorkbook = $candidate
                $candidate = $null
            }
        } finally { Release-ComObject $candidate }
    }
    if ($null -eq $hostWorkbook) { $hostWorkbook = $Books.Add() }
    return $hostWorkbook
}

function Get-RuleFingerprint([object]$Rule) {
    $row = [ordered]@{
        type=[int]$Rule.Type
        formula1=[string]$Rule.Formula1
        stop_if_true=[bool]$Rule.StopIfTrue
        interior_color=[Int64]$Rule.Interior.Color
    }
    return Get-StableDigest $row
}

function Get-ForbiddenLoadedModules([object]$Binding) {
    $rows = @()
    [void](Assert-ExactExcelProcessOwnership $Binding)
    $Binding.process.Refresh()
    foreach ($module in @($Binding.process.Modules)) {
        $name = [string]$module.ModuleName
        $path = [string]$module.FileName
        if ($name -match '(?i)(NxHost|FocusEngine)' -or $path -match '(?i)(NxHost|FocusEngine)') {
            $rows += [pscustomobject][ordered]@{name=$name;path=$path}
        }
    }
    return @($rows)
}

function Invoke-FocusStartupPass([string]$Label, [string]$TargetPath, [string]$ProfileRoot, [bool]$CommitSettings) {
    $excel = $null
    $binding = $null
    $books = $null
    $addinProject = $null
    $workbook = $null
    $worksheets = $null
    $sheet = $null
    $userRange = $null
    $userRule = $null
    $macro = $null
    $result = $null
    $failure = $null
    $cleanupFailures = New-Object 'Collections.Generic.List[string]'
    $exitMode = 'NOT_STARTED'
    $targetShaBefore = Get-Sha256 $TargetPath
    try {
        $baseline = Get-ExcelProcessBaseline
        if (@($baseline.process_ids).Count -ne 0) { throw ($Label + ' requires zero Excel processes') }
        $started = Start-InteractiveStartupExcel $baseline $Label
        $excel = $started.excel
        $binding = $started.binding
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $books = $excel.Workbooks
        $workbook = Resolve-HostWorkbook $books
        [void]$workbook.Activate()
        $initialPid = [int]$binding.pid
        $refreshedBinding = Get-ExactExcelProcessOwnership $excel $baseline ($Label + ' active host')
        if (-not [bool]$refreshedBinding.owned -or [int]$refreshedBinding.pid -ne $initialPid) {
            if ($null -ne $refreshedBinding.process) { $refreshedBinding.process.Dispose() }
            throw ($Label + ' Excel ownership changed after host activation')
        }
        if ($null -ne $binding.process) { $binding.process.Dispose() }
        $binding = $refreshedBinding
        $addinProject = Resolve-StartupAddinProject $excel $TargetPath
        $worksheets = $workbook.Worksheets
        $sheet = $worksheets.Item(1)
        $userRange = $sheet.Range('A1:Z100')
        $userRule = $userRange.FormatConditions.Add(2, $null, '=MOD(ROW()+COLUMN(),3)=0')
        $userRule.StopIfTrue = $false
        $userRule.Interior.Color = 65535
        $userFingerprintBefore = Get-RuleFingerprint $userRule
        $workbook.Saved = $true
        [void]$sheet.Activate()
        [void]$sheet.Range('B2').Select()
        $macro = "'" + ([IO.Path]::GetFileName($TargetPath)).Replace("'", "''") + "'!"

        if ($CommitSettings) {
            $commit = [string]$excel.Run($macro + 'NxFocusTryCommitSettingsValues', 'vertical', 'wide-stripe', '#78C6A3', 47, $false, $true)
            if ($commit -cne 'PASS') { throw ($Label + ' focus settings commit failed: ' + $commit) }
        } else {
            $colorsBeforeEnable = [string]$excel.Run($macro + 'NxFocusLoadColors')
            if ([string]($colorsBeforeEnable.Split(',')[0]) -cne '#78C6A3') { throw ($Label + ' persisted focus color readback failed') }
            [void]$excel.Run($macro + 'NxFocusControllerEnable')
        }
        Start-Sleep -Milliseconds 250

        $backend = [string]$excel.Run($macro + 'NxFocusControllerBackend')
        $managedRules = [int]$excel.Run($macro + 'NxFocusManagedRuleCount', $sheet)
        $settingsPath = [IO.Path]::GetFullPath([string]$excel.Run($macro + 'NxFocusSettingsPath'))
        if (-not $settingsPath.StartsWith([IO.Path]::GetFullPath($ProfileRoot).TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
            throw ($Label + ' focus settings escaped isolated profile')
        }
        if (-not (Test-Path -LiteralPath $settingsPath -PathType Leaf)) { throw ($Label + ' focus settings file missing') }
        $settingsSha = Get-Sha256 $settingsPath
        $colors = [string]$excel.Run($macro + 'NxFocusLoadColors')
        $userFingerprintAfter = Get-RuleFingerprint $userRule
        $forbiddenModules = @(Get-ForbiddenLoadedModules $binding)
        if ($backend -cne 'cf') { throw ($Label + ' focus backend is not cf: ' + $backend) }
        if ($managedRules -ne 2) { throw ($Label + ' managed focus rule count rejected: ' + $managedRules) }
        if ([string]($colors.Split(',')[0]) -cne '#78C6A3') { throw ($Label + ' focus color readback rejected') }
        if ($userFingerprintAfter -cne $userFingerprintBefore) { throw ($Label + ' user conditional format changed') }
        if (-not [bool]$workbook.Saved) { throw ($Label + ' workbook Saved state changed') }
        if ($forbiddenModules.Count -ne 0) { throw ($Label + ' forbidden focus DLL module loaded') }

        [void]$excel.Run($macro + 'NxFocusDisable')
        Start-Sleep -Milliseconds 150
        $rulesAfterDisable = [int]$excel.Run($macro + 'NxFocusManagedRuleCount', $sheet)
        if ($rulesAfterDisable -ne 0) { throw ($Label + ' managed focus rules remained after disable') }
        if ((Get-RuleFingerprint $userRule) -cne $userFingerprintBefore) { throw ($Label + ' user conditional format changed during disable') }

        $result = [pscustomobject][ordered]@{
            label=$Label
            startup_artifact=[IO.Path]::GetFullPath([string]$addinProject.FileName)
            startup_artifact_sha256=$targetShaBefore
            backend=$backend
            managed_rules=$managedRules
            rules_after_disable=$rulesAfterDisable
            settings_path=$settingsPath
            settings_sha256=$settingsSha
            selected_color=[string]($colors.Split(',')[0])
            user_rule_preserved=($userFingerprintAfter -ceq $userFingerprintBefore)
            saved_state_preserved=[bool]$workbook.Saved
            forbidden_modules=$forbiddenModules
            pid=[int]$binding.pid
        }
    } catch {
        $failure = $_.Exception.Message + ' | ' + $_.InvocationInfo.PositionMessage
    } finally {
        $ownershipVerified = $false
        if ($null -ne $excel -and -not [string]::IsNullOrWhiteSpace([string]$macro)) {
            try { [void]$excel.Run($macro + 'NxFocusDisable') } catch { }
        }
        if ($null -ne $excel -and $null -ne $binding -and [bool]$binding.owned) {
            try { [void](Assert-ExactExcelProcessOwnership $binding); $ownershipVerified = $true } catch { [void]$cleanupFailures.Add($_.Exception.Message) }
        }
        if ($null -ne $workbook) { try { $workbook.Close($false) } catch { [void]$cleanupFailures.Add($_.Exception.Message) } }
        foreach ($item in @($userRule,$userRange,$sheet,$worksheets,$workbook,$addinProject,$books)) {
            try { Release-ComObject $item } catch { [void]$cleanupFailures.Add($_.Exception.Message) }
        }
        if ($ownershipVerified) {
            try { $excel.Quit() } catch { [void]$cleanupFailures.Add($_.Exception.Message) }
        }
        try { Release-ComObject $excel } catch { [void]$cleanupFailures.Add($_.Exception.Message) }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        if ($null -ne $binding -and $null -ne $binding.process) {
            try {
                $cleanup = Stop-ExactProcessAfterGrace -Process $binding.process -Label $Label -GraceMs 15000 -Detailed
                $exitMode = [string]$cleanup.exit_mode
                if ($cleanup.failure) { [void]$cleanupFailures.Add([string]$cleanup.failure) }
            } catch { [void]$cleanupFailures.Add($_.Exception.Message) }
            try { $binding.process.Dispose() } catch { [void]$cleanupFailures.Add($_.Exception.Message) }
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($failure)) { throw $failure }
    if ($cleanupFailures.Count -ne 0) { throw ($Label + ' cleanup failed: ' + ($cleanupFailures -join ' | ')) }
    if ($exitMode -cne 'NATURAL') { throw ($Label + ' did not exit naturally: ' + $exitMode) }
    $result | Add-Member -NotePropertyName cleanup_exit_mode -NotePropertyValue $exitMode
    return $result
}

$productName = [string]([char]0xB0B4) + [char]0xC5D1 + [char]0xC140
$targetName = $productName + ' v0.01_r57.xlam'
$xlstart = [IO.Path]::GetFullPath((Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'))
$targetPath = [IO.Path]::GetFullPath((Join-Path $xlstart $targetName))
$stagePath = $targetPath + '.candidate'
$backupRoot = [IO.Path]::GetFullPath((Join-Path $EvidenceRoot 'xlstart-backup'))
$profileRoot = [IO.Path]::GetFullPath((Join-Path $EvidenceRoot 'profile\LHexcel'))
$buildRoot = [IO.Path]::GetFullPath((Join-Path $EvidenceRoot 'build'))
$candidatePath = [IO.Path]::GetFullPath((Join-Path $buildRoot 'Product.xlam'))
$projectRoot = [IO.Path]::GetFullPath((Join-Path $DataRoot '../..'))
$completedRoot = Join-Path $projectRoot '01_완성본(고도화)'
$distributionRoot = Join-Path $projectRoot '02_배포본(고도화)'

$script:transactionStarted = $false
$script:installedByRun = $false
$script:collisionManifest = New-Object 'Collections.Generic.List[object]'
$xlstartExistedBefore = Test-Path -LiteralPath $xlstart

function Install-XlstartCandidate([string]$Candidate, [string]$Target, [string]$Backup) {
    $root = [IO.Path]::GetFullPath((Split-Path $Target -Parent))
    [void](Assert-ChildPath $Target $root)
    [void](Assert-ChildPath $Backup $EvidenceRoot)
    [void](New-Item -ItemType Directory -Path $Backup -Force)
    $script:transactionStarted = $true
    foreach ($item in @(Get-ChildItem -LiteralPath $root -Force)) {
        $isProductCollision = $item.Name -match '(?i)^(Product|LHExcel).*\.xlam$'
        $isKoreanCollision = $item.Name.StartsWith($productName, [StringComparison]::OrdinalIgnoreCase) -and $item.Extension -ieq '.xlam'
        if (-not ($isProductCollision -or $isKoreanCollision)) { continue }
        if ($item.PSIsContainer -or ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw ('Unsafe XLSTART collision rejected: ' + $item.FullName)
        }
        $backupPath = Assert-ChildPath (Join-Path $Backup $item.Name) $Backup
        if (Test-Path -LiteralPath $backupPath) { throw ('Existing XLSTART backup collision: ' + $backupPath) }
        $row = [pscustomobject][ordered]@{
            original=[IO.Path]::GetFullPath($item.FullName)
            backup=$backupPath
            sha256=Get-Sha256 $item.FullName
            attributes=[int]$item.Attributes
            last_write_utc=$item.LastWriteTimeUtc.ToString('o')
            moved=$false
        }
        [void]$script:collisionManifest.Add($row)
        Move-Item -LiteralPath $item.FullName -Destination $backupPath
        $row.moved = $true
    }
    if (Test-Path -LiteralPath $Target) { throw 'XLSTART target still exists after collision isolation' }
    if (Test-Path -LiteralPath $stagePath) { throw 'XLSTART stage already exists' }
    Copy-Item -LiteralPath $Candidate -Destination $stagePath
    if ((Get-Sha256 $Candidate) -cne (Get-Sha256 $stagePath)) { throw 'XLSTART candidate hash mismatch' }
    Move-Item -LiteralPath $stagePath -Destination $Target
    $script:installedByRun = $true
    if ((Get-Sha256 $Candidate) -cne (Get-Sha256 $Target)) { throw 'XLSTART installed artifact hash mismatch' }
}

function Restore-XlstartTransaction([object[]]$BeforeSnapshot) {
    $failures = New-Object 'Collections.Generic.List[string]'
    try {
        if ($script:installedByRun -and (Test-Path -LiteralPath $targetPath -PathType Leaf)) {
            try {
                [IO.File]::SetAttributes($targetPath, [IO.FileAttributes]::Normal)
                Remove-Item -LiteralPath $targetPath -Force
            } catch { [void]$failures.Add($_.Exception.Message) }
        }
        if ($script:transactionStarted -and (Test-Path -LiteralPath $stagePath -PathType Leaf)) {
            try {
                [IO.File]::SetAttributes($stagePath, [IO.FileAttributes]::Normal)
                Remove-Item -LiteralPath $stagePath -Force
            } catch { [void]$failures.Add($_.Exception.Message) }
        }
        foreach ($row in $script:collisionManifest.ToArray()) {
            if (-not [bool]$row.moved) { continue }
            try {
                $original = Assert-ChildPath ([string]$row.original) $xlstart
                $backup = Assert-ChildPath ([string]$row.backup) $backupRoot
                if (-not (Test-Path -LiteralPath $backup -PathType Leaf)) { throw ('XLSTART backup missing: ' + $backup) }
                if (Test-Path -LiteralPath $original) { throw ('XLSTART restore collision: ' + $original) }
                if ((Get-Sha256 $backup) -cne [string]$row.sha256) { throw ('XLSTART backup hash changed: ' + $backup) }
                [IO.File]::SetAttributes($backup, [IO.FileAttributes]::Normal)
                Move-Item -LiteralPath $backup -Destination $original
                (Get-Item -LiteralPath $original -Force).LastWriteTimeUtc = [DateTime]::Parse([string]$row.last_write_utc).ToUniversalTime()
                [IO.File]::SetAttributes($original, [IO.FileAttributes][int]$row.attributes)
                if ((Get-Sha256 $original) -cne [string]$row.sha256) { throw ('Restored XLSTART hash mismatch: ' + $original) }
            } catch { [void]$failures.Add($_.Exception.Message) }
        }
        if (-not $xlstartExistedBefore -and (Test-Path -LiteralPath $xlstart -PathType Container) -and @(Get-ChildItem -LiteralPath $xlstart -Force).Count -eq 0) {
            try { Remove-Item -LiteralPath $xlstart -Force } catch { [void]$failures.Add($_.Exception.Message) }
        }
    } catch {
        [void]$failures.Add($_.Exception.Message)
    }
    $after = @()
    try { $after = @(Get-XlstartSnapshot $xlstart) } catch { [void]$failures.Add($_.Exception.Message) }
    $restored = (Get-StableDigest @($BeforeSnapshot)) -ceq (Get-StableDigest @($after))
    if (-not $restored) { [void]$failures.Add('XLSTART post-run snapshot differs from baseline') }
    return [pscustomobject][ordered]@{
        restored=($restored -and $failures.Count -eq 0)
        before_digest=Get-StableDigest @($BeforeSnapshot)
        after_digest=Get-StableDigest @($after)
        failures=@($failures)
    }
}

$originalProfilePresent = Test-Path Env:LHEXCEL_PROFILE_ROOT
$originalProfile = if ($originalProfilePresent) { [string]$env:LHEXCEL_PROFILE_ROOT } else { $null }
$phase = 'preflight'
$status = 'DIAGNOSTIC'
$failure = $null
$baselineExcel = @()
$baselineXlstart = @()
$baselineRegistry = @()
$baselineRegistryDigest = $null
$completedDigestBefore = $null
$distributionDigestBefore = $null
$candidateHash = $null
$installedHashBeforeCleanup = $null
$ribbonReceipt = $null
$ribbonContainment = $null
$focusFirst = $null
$focusRestart = $null
$cleanupReceipt = $null

try {
    $baselineExcel = @(Get-ExcelProcessSnapshot)
    $baselineXlstart = @(Get-XlstartSnapshot $xlstart)
    $baselineRegistry = @(Get-RegistryBaseline)
    $baselineRegistryDigest = Get-StableDigest @($baselineRegistry)
    $completedDigestBefore = Get-FileTreeDigest $completedRoot
    $distributionDigestBefore = Get-FileTreeDigest $distributionRoot
    Write-Receipt 'InternalXlamStartup.Baseline.json' ([ordered]@{
        schema_version=1;run_id=$RunId;excel_processes=$baselineExcel
        xlstart_path=$xlstart;xlstart_existed=$xlstartExistedBefore
        xlstart_snapshot=$baselineXlstart;xlstart_digest=Get-StableDigest @($baselineXlstart)
        registry_snapshot=$baselineRegistry;registry_digest=$baselineRegistryDigest
        nxhost_registration_absent=Test-NxHostRegistrationAbsent $baselineRegistry
        completed_tree_digest=$completedDigestBefore;distribution_tree_digest=$distributionDigestBefore
    })
    if ($baselineExcel.Count -ne 0) { $status = 'HOLD'; throw 'Pre-existing Excel process detected' }
    if (-not (Test-NxHostRegistrationAbsent $baselineRegistry)) { $status = 'HOLD'; throw 'Pre-existing NxHost registration detected' }
    if (-not $xlstartExistedBefore) { [void](New-Item -ItemType Directory -Path $xlstart -Force) }
    [void](Get-XlstartSnapshot $xlstart)
    [void](New-Item -ItemType Directory -Path $profileRoot -Force)
    [void](New-Item -ItemType Directory -Path $buildRoot -Force)
    $env:LHEXCEL_PROFILE_ROOT = $profileRoot

    $phase = 'build'
    & (Join-Path $SourceRoot 'build/Build-Xlam.ps1') -DataRoot $DataRoot -OutputPath $candidatePath -ManifestPath (Join-Path $SourceRoot 'build/manifests/Product.json') -RibbonPath (Join-Path $SourceRoot 'src/ribbon/customUI14.xml')
    if (-not (Test-Path -LiteralPath $candidatePath -PathType Leaf)) { throw 'Current Product.xlam build failed' }
    if (@(Get-ExcelProcessSnapshot).Count -ne 0) { throw 'Product build left an Excel process behind' }
    $candidateHash = Get-Sha256 $candidatePath

    $phase = 'install'
    Install-XlstartCandidate $candidatePath $targetPath $backupRoot
    $installedHash = Get-Sha256 $targetPath
    Write-Receipt 'InternalXlamStartup.Install.json' ([ordered]@{
        schema_version=1;run_id=$RunId;candidate=$candidatePath;candidate_sha256=$candidateHash
        target=$targetPath;installed_sha256=$installedHash
        collision_backups=$script:collisionManifest.ToArray()
        hash_match=($candidateHash -ceq $installedHash)
    })

    $phase = 'ribbon'
    $ribbonEvidence = Join-Path $EvidenceRoot 'ribbon'
    $powershell = (Get-Command powershell.exe -ErrorAction Stop).Source
    $ribbonRunId = [guid]::NewGuid().ToString()
    $ribbonArgs = @(
        '-NoProfile','-ExecutionPolicy','Bypass','-File',(Join-Path $SourceRoot 'tests/windows/Run-ProductRibbonSuite.ps1'),
        '-DataRoot',$DataRoot,'-EvidenceRoot',$ribbonEvidence,'-StartupArtifactPath',$targetPath,
        '-RunId',$ribbonRunId
    )
    $savedErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $ribbonOutput = @(& $powershell @ribbonArgs 2>&1 | ForEach-Object { [string]$_ })
        $ribbonExit = $LASTEXITCODE
    } finally {
        $ErrorActionPreference = $savedErrorActionPreference
    }
    $ribbonOutput | ForEach-Object { Write-Output $_ }
    $ribbonContainment = Stop-RecordedRibbonExcel (Join-Path $ribbonEvidence 'ProductRibbon.Ownership.json') $ribbonRunId $targetPath $baselineExcel
    if ([string]$ribbonContainment.status -ceq 'FAIL') { throw ('Installed Product ribbon containment failed: ' + [string]$ribbonContainment.failure) }
    $ribbonReceiptPath = Join-Path $ribbonEvidence 'ProductRibbon.json'
    if ($ribbonExit -ne 0 -or -not (Test-Path -LiteralPath $ribbonReceiptPath -PathType Leaf)) {
        throw ('Installed Product ribbon acceptance failed with exit code ' + $ribbonExit)
    }
    $ribbonReceipt = Get-Content -LiteralPath $ribbonReceiptPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([string]$ribbonReceipt.status -cne 'PASS') { throw 'Installed Product ribbon receipt is not PASS' }
    Write-Receipt 'InternalXlamStartup.Ribbon.json' ([ordered]@{
        schema_version=1;run_id=$RunId;status=[string]$ribbonReceipt.status
        run=$ribbonReceipt.run;artifact_sha256=[string]$ribbonReceipt.artifact_sha256
        child_exit_code=$ribbonExit;child_output=$ribbonOutput;containment=$ribbonContainment
    })
    if (@(Wait-ExcelProcessSnapshot $baselineExcel).Count -ne 0) { throw 'Ribbon acceptance left an Excel process behind' }

    $phase = 'focus-first'
    $focusFirst = Invoke-FocusStartupPass 'InternalXlamStartup focus first' $targetPath $profileRoot $true
    Write-Receipt 'InternalXlamStartup.Focus.json' ([ordered]@{schema_version=1;run_id=$RunId;status='PASS';result=$focusFirst})
    if (@(Wait-ExcelProcessSnapshot $baselineExcel).Count -ne 0) { throw 'First focus pass left an Excel process behind' }

    $phase = 'focus-restart'
    $focusRestart = Invoke-FocusStartupPass 'InternalXlamStartup focus restart' $targetPath $profileRoot $false
    if ([string]$focusRestart.settings_sha256 -cne [string]$focusFirst.settings_sha256) { throw 'Focus settings digest changed across restart' }
    Write-Receipt 'InternalXlamStartup.Restart.json' ([ordered]@{schema_version=1;run_id=$RunId;status='PASS';result=$focusRestart;settings_digest_match=$true})
    if (@(Wait-ExcelProcessSnapshot $baselineExcel).Count -ne 0) { throw 'Restart focus pass left an Excel process behind' }
    if ((Get-Sha256 $targetPath) -cne $candidateHash) { throw 'Installed Product.xlam changed during acceptance' }
    $status = 'PASS'
} catch {
    if ($status -cne 'HOLD') { $status = 'DIAGNOSTIC' }
    $failure = $_.Exception.Message + ' | ' + $_.InvocationInfo.PositionMessage
} finally {
    $phaseBeforeCleanup = $phase
    $phase = 'cleanup'
    if (Test-Path -LiteralPath $targetPath -PathType Leaf) {
        try { $installedHashBeforeCleanup = Get-Sha256 $targetPath } catch { }
    }
    $xlstartCleanup = Restore-XlstartTransaction $baselineXlstart
    $registryAfter = @()
    $registryRestored = $false
    $completedRestored = $false
    $distributionRestored = $false
    try {
        $registryAfter = @(Get-RegistryBaseline)
        $registryRestored = (Get-StableDigest @($registryAfter)) -ceq $baselineRegistryDigest
    } catch { $failure = (@($failure,$_.Exception.Message) | Where-Object {$_}) -join ' | ' }
    try { $completedRestored = (Get-FileTreeDigest $completedRoot) -ceq $completedDigestBefore } catch { $failure = (@($failure,$_.Exception.Message) | Where-Object {$_}) -join ' | ' }
    try { $distributionRestored = (Get-FileTreeDigest $distributionRoot) -ceq $distributionDigestBefore } catch { $failure = (@($failure,$_.Exception.Message) | Where-Object {$_}) -join ' | ' }
    $excelAfter = @(Wait-ExcelProcessSnapshot $baselineExcel)
    $excelRestored = (Get-StableDigest @($excelAfter)) -ceq (Get-StableDigest @($baselineExcel))
    $nxHostAbsentAfter = Test-NxHostRegistrationAbsent $registryAfter
    $artifactUnchanged = ($null -eq $installedHashBeforeCleanup -or $null -eq $candidateHash -or $installedHashBeforeCleanup -ceq $candidateHash)
    $cleanupPass = ([bool]$xlstartCleanup.restored -and $registryRestored -and $completedRestored -and $distributionRestored -and $excelRestored)
    $cleanupReceipt = [ordered]@{
        status=if($cleanupPass){'PASS'}else{'RECOVERY_REQUIRED'}
        xlstart_restored=[bool]$xlstartCleanup.restored
        xlstart_before_digest=[string]$xlstartCleanup.before_digest
        xlstart_after_digest=[string]$xlstartCleanup.after_digest
        xlstart_failures=@($xlstartCleanup.failures)
        registry_restored=$registryRestored
        registry_before_digest=$baselineRegistryDigest
        registry_after_digest=Get-StableDigest @($registryAfter)
        completed_tree_restored=$completedRestored
        distribution_tree_restored=$distributionRestored
        excel_restored=$excelRestored
        excel_residue=$excelAfter.Count
        nxhost_registration_absent=$nxHostAbsentAfter
        installed_artifact_unchanged_before_removal=$artifactUnchanged
    }
    if (-not $cleanupPass) { $status = 'RECOVERY_REQUIRED' }
    Write-Receipt 'InternalXlamStartup.Cleanup.json' ([ordered]@{schema_version=1;run_id=$RunId;cleanup=$cleanupReceipt})
    if ($originalProfilePresent) { $env:LHEXCEL_PROFILE_ROOT = $originalProfile } else { Remove-Item Env:LHEXCEL_PROFILE_ROOT -ErrorAction SilentlyContinue }
    Write-Receipt 'InternalXlamStartup.json' ([ordered]@{
        schema_version=1;suite='InternalXlamStartup';status=$status;run_id=$RunId
        phase=$phaseBeforeCleanup;failure=$failure
        candidate_sha256=$candidateHash;installed_sha256_before_cleanup=$installedHashBeforeCleanup
        ribbon=$ribbonReceipt;ribbon_containment=$ribbonContainment;focus_first=$focusFirst;focus_restart=$focusRestart;cleanup=$cleanupReceipt
    })
}

if ($status -ceq 'PASS') {
    Write-Output 'PASS|InternalXlamStartup'
    exit 0
}
if ($status -ceq 'HOLD') {
    [Console]::Error.WriteLine(('HOLD|InternalXlamStartup|' + $failure))
    exit 10
}
if ($status -ceq 'RECOVERY_REQUIRED') {
    [Console]::Error.WriteLine(('RECOVERY_REQUIRED|InternalXlamStartup|' + $failure))
    exit 25
}
[Console]::Error.WriteLine(('DIAGNOSTIC|InternalXlamStartup|' + $failure))
exit 23
