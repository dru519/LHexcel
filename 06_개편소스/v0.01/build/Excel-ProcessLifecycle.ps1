Set-StrictMode -Version Latest

function Wait-ExactArtifactReadable([string]$Path, [int]$TimeoutMs = 10000) {
    if ([string]::IsNullOrWhiteSpace($Path) -or $TimeoutMs -le 0) { throw 'Shared artifact visibility wait input rejected' }
    $fullPath = [IO.Path]::GetFullPath($Path)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $lastFailure = 'not visible'
    $stableReadCount = 0
    while ($watch.ElapsedMilliseconds -lt $TimeoutMs) {
        $stream = $null
        try {
            if (-not (Test-Path -LiteralPath $fullPath -PathType Leaf -ErrorAction Stop)) { throw 'leaf not visible' }
            $item = Get-Item -LiteralPath $fullPath -Force -ErrorAction Stop
            if ($item.PSIsContainer -or $item.Length -le 0) { throw 'leaf is not a non-empty file' }
            $stream = [IO.File]::Open($fullPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
            if ($stream.Length -le 0) { throw 'leaf is not readable' }
            $stableReadCount++
            if ($stableReadCount -ge 3) { return }
            $lastFailure = 'awaiting stable reads: ' + $stableReadCount + '/3'
        } catch {
            $lastFailure = $_.Exception.Message
            $stableReadCount = 0
        } finally {
            if ($null -ne $stream) { $stream.Dispose() }
        }
        Start-Sleep -Milliseconds 100
    }
    throw ('Shared artifact visibility timeout: path=' + $fullPath + '; timeout_ms=' + $TimeoutMs + '; cause=' + $lastFailure)
}

function Initialize-ExactExcelWindowApi {
    if ('NxExactExcelWindowApi' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class NxExactExcelWindowApi {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  delegate bool EnumProc(IntPtr h, IntPtr p);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, IntPtr p);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc callback, IntPtr p);
  [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, System.Text.StringBuilder s, int n);
  [DllImport("oleacc.dll")] static extern int AccessibleObjectFromWindow(IntPtr h, uint id, ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object value);
  public static object NativeDocument(int processId) {
    object result=null;
    EnumWindows((top,p)=> {
      uint owner; GetWindowThreadProcessId(top,out owner); if(owner!=processId)return true;
      EnumChildWindows(top,(child,q)=> {
        var cls=new System.Text.StringBuilder(256);GetClassName(child,cls,256);
        if(cls.ToString()!="EXCEL7")return true;
        Guid iid=new Guid("00020400-0000-0000-C000-000000000046");object value;
        if(AccessibleObjectFromWindow(child,0xFFFFFFF0,ref iid,out value)==0){result=value;return false;}
        return true;
      },IntPtr.Zero);return result==null;
    },IntPtr.Zero);return result;
  }
}
'@
}

function Get-ExcelProcessBaseline {
    $captured = [DateTime]::UtcNow
    $rows = @()
    foreach ($candidate in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
        try {
            [void]$candidate.Handle
            $rows += [pscustomobject]@{
                pid = [int]$candidate.Id
                started_utc = $candidate.StartTime.ToUniversalTime().ToString('o')
                executable = [IO.Path]::GetFullPath($candidate.MainModule.FileName)
            }
        } catch {
            $rows += [pscustomobject]@{pid=[int]$candidate.Id;started_utc=$null;executable=$null}
        } finally {
            try { $candidate.Dispose() } catch { }
        }
    }
    return [pscustomobject]@{
        captured_utc = $captured.ToString('o')
        captured_ticks = $captured.Ticks
        process_ids = @($rows | ForEach-Object { [int]$_.pid })
        processes = @($rows)
    }
}

function Get-ExactExcelProcessOwnership([object]$Excel, [object]$Baseline, [string]$Label) {
    Initialize-ExactExcelWindowApi
    $process = $null
    $binding = [ordered]@{
        owned = $false; failure = $null; label = $Label; process = $null
        pid = 0; hwnd = 0; started_utc = $null; executable = $null
    }
    try {
        if ($null -eq $Excel -or $null -eq $Baseline) { throw ($Label + ' ownership input unavailable') }
        $hwnd = [Int64]$Excel.Hwnd
        if ($hwnd -eq 0) { throw ($Label + ' HWND unavailable') }
        [uint32]$excelProcessId = 0
        [void][NxExactExcelWindowApi]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$excelProcessId)
        if ($excelProcessId -eq 0) { throw ($Label + ' HWND PID unavailable') }
        if (@($Baseline.process_ids) -contains [int]$excelProcessId) { throw ($Label + ' reused a baseline EXCEL PID') }
        $process = Get-Process -Id ([int]$excelProcessId) -ErrorAction Stop
        [void]$process.Handle
        $started = $process.StartTime.ToUniversalTime()
        $executable = [IO.Path]::GetFullPath($process.MainModule.FileName)
        if ($started.Ticks -lt [Int64]$Baseline.captured_ticks) { throw ($Label + ' process StartTime predates baseline') }
        if ([IO.Path]::GetFileName($executable) -cne 'EXCEL.EXE') { throw ($Label + ' executable is not exact EXCEL.EXE') }
        [uint32]$confirmProcessId = 0
        [void][NxExactExcelWindowApi]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$confirmProcessId)
        if ($confirmProcessId -ne $excelProcessId -or [int]$process.Id -ne [int]$excelProcessId) { throw ($Label + ' HWND/PID exact binding changed') }
        $binding.owned = $true
        $binding.process = $process
        $binding.pid = [int]$excelProcessId
        $binding.hwnd = $hwnd
        $binding.started_utc = $started.ToString('o')
        $binding.executable = $executable
        $process = $null
    } catch {
        $binding.failure = $_.Exception.Message
    } finally {
        if ($null -ne $process) { try { $process.Dispose() } catch { } }
    }
    return [pscustomobject]$binding
}

function Get-ExactRegisteredExcelExecutable {
    $candidates = New-Object 'Collections.Generic.List[string]'
    foreach ($view in @(
        [Microsoft.Win32.RegistryView]::Registry64,
        [Microsoft.Win32.RegistryView]::Registry32
    )) {
        $base = $null
        $appPath = $null
        try {
            $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
                [Microsoft.Win32.RegistryHive]::LocalMachine,
                $view
            )
            $appPath = $base.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe', $false)
            if ($null -ne $appPath) {
                $raw = [string]$appPath.GetValue('', $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                if (-not [string]::IsNullOrWhiteSpace($raw)) { [void]$candidates.Add($raw) }
            }
        } finally {
            if ($null -ne $appPath) { $appPath.Dispose() }
            if ($null -ne $base) { $base.Dispose() }
        }
    }
    foreach ($candidate in @($candidates | Sort-Object -Unique)) {
        try {
            $full = [IO.Path]::GetFullPath(([string]$candidate).Trim().Trim([char]34))
            if ((Test-Path -LiteralPath $full -PathType Leaf) -and [IO.Path]::GetFileName($full) -ceq 'EXCEL.EXE') { return $full }
        } catch { }
    }
    throw 'Registered EXCEL.EXE path unavailable'
}

function Start-ExactInteractiveExcel([object]$Baseline, [string]$Label, [int]$TimeoutMs = 30000, [switch]$SafeMode, [string]$StartupWorkbook = '', [switch]$ShowWindow) {
    if ($null -eq $Baseline -or [string]::IsNullOrWhiteSpace($Label) -or $TimeoutMs -le 0) {
        throw 'Interactive Excel launch input rejected'
    }
    $excelPath = Get-ExactRegisteredExcelExecutable
    $launch = $null
    $candidate = $null
    $candidateBinding = $null
    $transferred = $false
    try {
        $launchArguments = @('/x')
        if ($SafeMode) { $launchArguments = @('/safe') }
        if ($StartupWorkbook) {
            $startupPath=[IO.Path]::GetFullPath($StartupWorkbook)
            if (-not (Test-Path -LiteralPath $startupPath -PathType Leaf) -or [IO.Path]::GetExtension($startupPath) -ne '.xlsx' -or $startupPath.Contains('"')) { throw 'Startup fixture must be an existing XLSX' }
            $launchArguments += ('"'+$startupPath+'"')
        }
        $windowStyle = if ($ShowWindow) { 'Normal' } else { 'Hidden' }
        $launch = Start-Process -FilePath $excelPath -ArgumentList $launchArguments -WindowStyle $windowStyle -PassThru
        if ($null -eq $launch) { throw ($Label + ' interactive launch failed') }
        $deadline = [DateTime]::UtcNow.AddMilliseconds($TimeoutMs)
        while ([DateTime]::UtcNow -lt $deadline) {
            try {
                try { $candidate = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') }
                catch {
                    Initialize-ExactExcelWindowApi
                    $nativeDocument=[NxExactExcelWindowApi]::NativeDocument($launch.Id)
                    try {if($null -ne $nativeDocument){$candidate=$nativeDocument.Application}}
                    finally {if($null -ne $nativeDocument -and [Runtime.InteropServices.Marshal]::IsComObject($nativeDocument)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($nativeDocument)}}
                }
                if($null -eq $candidate){throw 'Owned Excel document not ready'}
                $candidateBinding = Get-ExactExcelProcessOwnership $candidate $Baseline $Label
                if ([bool]$candidateBinding.owned -and [int]$candidateBinding.pid -eq $launch.Id -and [bool]$candidate.UserControl) {
                    $result = [pscustomobject]@{excel=$candidate;binding=$candidateBinding}
                    $candidate = $null
                    $candidateBinding = $null
                    $transferred = $true
                    return $result
                }
            } catch { }
            finally {
                if (-not $transferred) {
                    if ($null -ne $candidateBinding -and $null -ne $candidateBinding.process) {
                        try { $candidateBinding.process.Dispose() } catch { }
                    }
                    if ($null -ne $candidate -and [Runtime.InteropServices.Marshal]::IsComObject($candidate)) {
                        # A rejected ROT result may share its RCW with another live fixture.
                        # Release only this acquisition, never invalidate the baseline owner.
                        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($candidate) } catch { }
                    }
                    $candidateBinding = $null
                    $candidate = $null
                }
            }
            Start-Sleep -Milliseconds 100
        }
        throw ($Label + ' interactive COM binding timed out')
    } finally {
        $cleanupFailure = $null
        if (-not $transferred -and $null -ne $launch) {
            try {
                if (-not $launch.HasExited) {
                    # Without a COM binding the workbook inventory is unknown.
                    # Preserve even our launched PID rather than discard user data.
                    $cleanupFailure = $Label + ' unbound Excel preserved for review; pid=' + $launch.Id
                }
            } catch { $cleanupFailure = $_.Exception.Message }
        }
        if ($null -ne $launch) { try { $launch.Dispose() } catch { } }
        if (-not [string]::IsNullOrWhiteSpace($cleanupFailure)) { throw $cleanupFailure }
    }
}

function Assert-ExactExcelProcessOwnership([object]$Binding) {
    Initialize-ExactExcelWindowApi
    if ($null -eq $Binding -or -not [bool]$Binding.owned -or $null -eq $Binding.process) { throw 'Excel process is not owned' }
    [uint32]$hwndProcessId = 0
    [void][NxExactExcelWindowApi]::GetWindowThreadProcessId([IntPtr][Int64]$Binding.hwnd, [ref]$hwndProcessId)
    [void]$Binding.process.Handle
    $started = $Binding.process.StartTime.ToUniversalTime().ToString('o')
    $executable = [IO.Path]::GetFullPath($Binding.process.MainModule.FileName)
    if ([int]$Binding.process.Id -ne [int]$Binding.pid -or
        $started -cne [string]$Binding.started_utc -or $executable -cne [string]$Binding.executable -or
        [IO.Path]::GetFileName($executable) -cne 'EXCEL.EXE' -or
        ($hwndProcessId -ne 0 -and $hwndProcessId -ne [int]$Binding.pid)) {
        throw 'Excel ownership exact binding no longer matches PID/StartTime/EXCEL.EXE'
    }
    return $true
}

function Find-CommandBarControlById([object]$Controls, [int]$ControlId, [int]$Depth) {
    if ($null -eq $Controls -or $Depth -gt 6) { return $null }
    $count = 0
    try { $count = [int]$Controls.Count } catch { return $null }
    for ($index = 1; $index -le $count; $index++) {
        $candidate = $null; $children = $null; $found = $null
        try {
            $candidate = $Controls.Item($index)
            if ([int]$candidate.Id -eq $ControlId) { $result = $candidate; $candidate = $null; return $result }
            try { $children = $candidate.Controls } catch { $children = $null }
            if ($null -ne $children -and [int]$children.Count -gt 0) {
                $found = Find-CommandBarControlById $children $ControlId ($Depth + 1)
                if ($null -ne $found) { $result = $found; $found = $null; return $result }
            }
        } finally {
            Release-ComObject $found; Release-ComObject $children; Release-ComObject $candidate
        }
    }
    return $null
}

function Find-VbeCompileControl([object]$Bars) {
    $barCount = 0
    try { $barCount = [int]$Bars.Count } catch { return $null }
    for ($barIndex = 1; $barIndex -le $barCount; $barIndex++) {
        $bar = $null; $controls = $null; $found = $null
        try {
            $bar = $Bars.Item($barIndex); $controls = $bar.Controls
            $found = Find-CommandBarControlById $controls 578 0
            if ($null -ne $found) { $result = $found; $found = $null; return $result }
        } finally {
            Release-ComObject $found; Release-ComObject $controls; Release-ComObject $bar
        }
    }
    return $null
}

function Start-ExactOwnedCompileDialogWatcher([int]$ExcelPid, [Int64]$ExcelHwnd, [Int64]$VbeHwnd) {
    Start-Job -ArgumentList $ExcelPid,$ExcelHwnd,$VbeHwnd -ScriptBlock {
        param($excelProcessId,$excelHwnd,$vbeHwnd)
        Add-Type -TypeDefinition @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class NxExactCompileWatcher {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr state);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc cb, IntPtr state);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int max);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int max);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);
  static string Text(IntPtr h) { var b=new StringBuilder(4096); GetWindowText(h,b,b.Capacity); return b.ToString(); }
  static string Body(IntPtr h) { var rows=new List<string>(); var top=Text(h); if(!String.IsNullOrWhiteSpace(top)) rows.Add(top); EnumChildWindows(h,(c,s)=>{var t=Text(c);if(!String.IsNullOrWhiteSpace(t))rows.Add(t);return true;},IntPtr.Zero); return String.Join("\n",rows.ToArray()); }
  public static string[] Poll(uint wanted,long excelRoot,long vbeRoot) { var rows=new List<string>(); EnumWindows((h,s)=>{uint pid;GetWindowThreadProcessId(h,out pid);if(pid==wanted&&IsWindowVisible(h)&&h.ToInt64()!=excelRoot&&h.ToInt64()!=vbeRoot){var c=new StringBuilder(256);GetClassName(h,c,c.Capacity);if(c.ToString()=="#32770")rows.Add(h.ToInt64()+"|"+Body(h));}return true;},IntPtr.Zero);return rows.ToArray(); }
  public static bool Close(long value,uint wanted) { var h=new IntPtr(value);uint pid;GetWindowThreadProcessId(h,out pid);return pid==wanted&&PostMessage(h,0x0010,IntPtr.Zero,IntPtr.Zero); }
}
'@
        $detected = @()
        for ($attempt = 0; $attempt -lt 50; $attempt++) {
            $dialogs = @([NxExactCompileWatcher]::Poll([uint32]$excelProcessId,[Int64]$excelHwnd,[Int64]$vbeHwnd))
            if ($dialogs.Count -gt 0) {
                foreach ($dialog in $dialogs) {
                    $parts = $dialog.Split('|',2)
                    $closed = [NxExactCompileWatcher]::Close([Int64]$parts[0],[uint32]$excelProcessId)
                    $detected += [pscustomobject]@{hwnd=[Int64]$parts[0];text=$parts[1];closed=$closed}
                }
            } elseif ($detected.Count -gt 0) {
                return [pscustomobject]@{status='DIALOG_DETECTED';owner_pid=[int]$excelProcessId;dialog_count=$detected.Count;dialog_text=($detected.text -join "`n---`n");closed=(-not (@($detected | Where-Object {-not $_.closed}).Count))}
            }
            Start-Sleep -Milliseconds 100
        }
        if ($detected.Count -gt 0) { return [pscustomobject]@{status='DIALOG_DETECTED';owner_pid=[int]$excelProcessId;dialog_count=$detected.Count;dialog_text=($detected.text -join "`n---`n");closed=(-not (@($detected | Where-Object {-not $_.closed}).Count))} }
        return [pscustomobject]@{status='NO_DIALOG';owner_pid=[int]$excelProcessId;dialog_count=0;dialog_text=$null;closed=$false}
    }
}

function New-ExactProcessLifecycleResult([string]$ExitMode, [string]$Failure, [bool]$IncludeDetails) {
    if ($IncludeDetails) { return [pscustomobject]@{exit_mode=$ExitMode;failure=$Failure} }
    return $Failure
}

function Stop-ExactProcessAfterGrace([Diagnostics.Process]$Process, [string]$Label, [int]$GraceMs = 10000, [switch]$Detailed) {
    if ($null -eq $Process) { return (New-ExactProcessLifecycleResult 'NATURAL' $null $Detailed.IsPresent) }
    $processId = 0
    try { $processId = [int]$Process.Id } catch {
        Write-Host ('lifecycle_event=FORCED_FAILED exit_mode=FORCED_FAILED reason=identity_unavailable label=' + $Label + ' pid=unknown grace_ms=' + $GraceMs + ' grace_elapsed_ms=0')
        return (New-ExactProcessLifecycleResult 'FORCED_FAILED' ($Label + ' identity unavailable during cleanup') $Detailed.IsPresent)
    }
    $deadline = [Diagnostics.Stopwatch]::StartNew()
    try {
        if ($Process.HasExited) {
            Write-Host ('lifecycle_event=NATURAL exit_mode=NATURAL label=' + $Label + ' pid=' + $processId + ' grace_ms=' + $GraceMs + ' grace_elapsed_ms=' + $deadline.ElapsedMilliseconds)
            return (New-ExactProcessLifecycleResult 'NATURAL' $null $Detailed.IsPresent)
        }
    } catch {
        Write-Host ('lifecycle_event=FORCED_FAILED exit_mode=FORCED_FAILED reason=identity_unavailable label=' + $Label + ' pid=' + $processId + ' grace_ms=' + $GraceMs + ' grace_elapsed_ms=' + $deadline.ElapsedMilliseconds)
        return (New-ExactProcessLifecycleResult 'FORCED_FAILED' ($Label + ' identity unavailable during cleanup') $Detailed.IsPresent)
    }
    while ($deadline.ElapsedMilliseconds -lt $GraceMs) {
        try {
            if ($Process.HasExited) {
                Write-Host ('lifecycle_event=NATURAL exit_mode=NATURAL label=' + $Label + ' pid=' + $processId + ' grace_ms=' + $GraceMs + ' grace_elapsed_ms=' + $deadline.ElapsedMilliseconds)
                return (New-ExactProcessLifecycleResult 'NATURAL' $null $Detailed.IsPresent)
            }
        } catch {
            Write-Host ('lifecycle_event=FORCED_FAILED exit_mode=FORCED_FAILED reason=identity_unavailable label=' + $Label + ' pid=' + $processId + ' grace_ms=' + $GraceMs + ' grace_elapsed_ms=' + $deadline.ElapsedMilliseconds)
            return (New-ExactProcessLifecycleResult 'FORCED_FAILED' ($Label + ' identity unavailable during cleanup') $Detailed.IsPresent)
        }
        Start-Sleep -Milliseconds 200
    }
    try { $Process.Kill() } catch {
        Write-Host ('lifecycle_event=FORCED_FAILED exit_mode=FORCED_FAILED reason=kill_failed label=' + $Label + ' pid=' + $processId + ' grace_ms=' + $GraceMs + ' grace_elapsed_ms=' + $deadline.ElapsedMilliseconds)
        return (New-ExactProcessLifecycleResult 'FORCED_FAILED' ($Label + ' forced cleanup failed: ' + $_.Exception.Message) $Detailed.IsPresent)
    }
    try {
        if (-not $Process.WaitForExit(30000)) {
            Write-Host ('lifecycle_event=FORCED_FAILED exit_mode=FORCED_FAILED reason=survived label=' + $Label + ' pid=' + $processId + ' grace_ms=' + $GraceMs + ' grace_elapsed_ms=' + $deadline.ElapsedMilliseconds)
            return (New-ExactProcessLifecycleResult 'FORCED_FAILED' ($Label + ' survived forced cleanup: ' + $processId) $Detailed.IsPresent)
        }
    } catch {
        Write-Host ('lifecycle_event=FORCED_FAILED exit_mode=FORCED_FAILED reason=wait_failed label=' + $Label + ' pid=' + $processId + ' grace_ms=' + $GraceMs + ' grace_elapsed_ms=' + $deadline.ElapsedMilliseconds)
        return (New-ExactProcessLifecycleResult 'FORCED_FAILED' ($Label + ' forced cleanup wait failed: ' + $_.Exception.Message) $Detailed.IsPresent)
    }
    Write-Host ('lifecycle_event=FORCED_CONTAINED exit_mode=FORCED_CONTAINED label=' + $Label + ' pid=' + $processId + ' grace_ms=' + $GraceMs + ' grace_elapsed_ms=' + $deadline.ElapsedMilliseconds)
    return (New-ExactProcessLifecycleResult 'FORCED_CONTAINED' $null $Detailed.IsPresent)
}
