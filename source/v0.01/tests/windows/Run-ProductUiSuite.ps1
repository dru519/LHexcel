param(
    [string]$Suite = 'ProductUi',
    [ValidateSet('Green')][string]$Mode = 'Green',
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [string]$EvidenceRoot = '',
    [string]$RunId = '',
    [string]$SourceDigest = '',
    [string]$SnapshotDigest = '',
    [string]$ArtifactPath = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($Suite -cne 'ProductUi') { throw 'ProductUi suite identity rejected' }
if ($PSVersionTable.PSVersion -lt [version]'5.1') { throw 'Windows PowerShell 5.1 or later is required' }
if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required' }

$SourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
. (Join-Path $SourceRoot 'build/Excel-ProcessLifecycle.ps1')
$evidence = if ([string]::IsNullOrWhiteSpace($EvidenceRoot)) { Join-Path $DataRoot 'runtime/evidence/vba' } else { $EvidenceRoot }
$failurePhase = 'preflight'
$probeStage = 'not-started'
$terminalCode = 0
$failureMessage = $null
$cleanupFailures = @()
$cleanupExitMode = 'NOT_STARTED'
$excel = $null; $books = $null; $hostBook = $null; $addin = $null; $binding = $null; $process = $null
$artifact = Join-Path $evidence 'Product.xlam'
$artifactSha = $null
$userFormAttestationSha = $null
$attestedForms = @()
$script:RibbonSurfaceGroupAnchors = @{}

function Release-ComObject([object]$Value) {
    try {
        if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
        }
    } catch { }
}

function Get-RegisteredProductUiExcelExecutable {
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

function Start-InteractiveProductUiExcel([object]$Baseline) {
    $excelPath = Get-RegisteredProductUiExcelExecutable
    $launch = $null
    $candidate = $null
    $candidateBinding = $null
    $transferred = $false
    try {
        $launch = Start-Process -FilePath $excelPath -ArgumentList @('/x') -WindowStyle Normal -PassThru
        if ($null -eq $launch) { throw 'ProductUi interactive Excel process launch failed' }
        for ($attempt = 0; $attempt -lt 150; $attempt++) {
            try {
                $candidate = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
                $candidateBinding = Get-ExactExcelProcessOwnership $candidate $Baseline 'ProductUi interactive Excel'
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
                    if ($null -ne $candidateBinding -and $null -ne $candidateBinding.process) { try { $candidateBinding.process.Dispose() } catch { } }
                    Release-ComObject $candidate
                    $candidateBinding = $null
                    $candidate = $null
                }
            }
            Start-Sleep -Milliseconds 200
        }
        throw 'ProductUi interactive Excel COM binding timed out'
    } finally {
        $failedLaunchCleanupError = $null
        if (-not $transferred -and $null -ne $launch) {
            try {
                if (-not $launch.HasExited) {
                    $failedLaunchCleanup = Stop-ExactProcessAfterGrace -Process $launch -Label 'ProductUi interactive launch' -GraceMs 10000 -Detailed
                    if ($failedLaunchCleanup.failure) { $failedLaunchCleanupError = [string]$failedLaunchCleanup.failure }
                }
            } catch { $failedLaunchCleanupError = $_.Exception.Message }
        }
        if ($null -ne $launch) { try { $launch.Dispose() } catch { } }
        if (-not [string]::IsNullOrWhiteSpace($failedLaunchCleanupError)) { throw $failedLaunchCleanupError }
    }
}

function Get-HexDigest([string]$Path) {
    (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash.ToLowerInvariant()
}

function Get-TextSha256([string]$Text) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
}

function Write-Json([string]$Path, [object]$Value) {
    $utf8 = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 40), $utf8)
}

function ConvertFrom-CodePoints([int[]]$CodePoints) {
    -join @($CodePoints | ForEach-Object { [char]$_ })
}

function Initialize-ProductUiApi {
    if ('NxProductUiApi' -as [type]) { return }
    Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;
public struct NxUiRect { public int Left; public int Top; public int Right; public int Bottom; }
public struct NxUiPoint { public int X; public int Y; }
public static class NxProductUiApi {
  public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr parameter);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int maximum);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, StringBuilder text, int maximum);
  public static void CloseOwnedForms(uint wantedProcessId) {
    EnumWindows((hwnd, parameter) => {
      uint processId;
      GetWindowThreadProcessId(hwnd, out processId);
      if (processId != wantedProcessId || !IsWindowVisible(hwnd)) return true;
      var name = new StringBuilder(128);
      GetClassName(hwnd, name, name.Capacity);
      if (name.ToString().StartsWith("Thunder", StringComparison.Ordinal))
        PostMessage(hwnd, 0x0010, IntPtr.Zero, IntPtr.Zero);
      return true;
    }, IntPtr.Zero);
  }
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int command);
  [DllImport("user32.dll")] public static extern void SwitchToThisWindow(IntPtr hwnd, bool altTab);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr hwnd);
  [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint fromThread, uint toThread, bool attach);
  [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
  [DllImport("user32.dll", SetLastError=true)] public static extern bool GetClientRect(IntPtr hwnd, out NxUiRect rectangle);
  [DllImport("user32.dll", SetLastError=true)] public static extern bool ClientToScreen(IntPtr hwnd, ref NxUiPoint point);
  public static bool ActivateForegroundWindow(IntPtr hwnd) {
    ShowWindow(hwnd, 9);
    SwitchToThisWindow(hwnd, true);
    for (var attempt = 0; attempt < 20; attempt++) {
      if (GetForegroundWindow() == hwnd) return true;
      SetForegroundWindow(hwnd);
      Thread.Sleep(50);
    }
    var foreground = GetForegroundWindow();
    uint targetProcessId;
    uint targetThread = GetWindowThreadProcessId(hwnd, out targetProcessId);
    uint foregroundProcessId;
    uint foregroundThread = GetWindowThreadProcessId(foreground, out foregroundProcessId);
    uint currentThread = GetCurrentThreadId();
    bool attachedForeground = foregroundThread != 0 && currentThread != foregroundThread && AttachThreadInput(currentThread, foregroundThread, true);
    bool attachedTarget = targetThread != 0 && currentThread != targetThread && AttachThreadInput(currentThread, targetThread, true);
    bool activated = false;
    try {
      ShowWindow(hwnd, 9);
      BringWindowToTop(hwnd);
      SetForegroundWindow(hwnd);
      SetFocus(hwnd);
      for (var attempt = 0; attempt < 20; attempt++) {
        if (GetForegroundWindow() == hwnd) { activated = true; break; }
        Thread.Sleep(50);
      }
    } finally {
      if (attachedTarget) AttachThreadInput(currentThread, targetThread, false);
      if (attachedForeground) AttachThreadInput(currentThread, foregroundThread, false);
    }
    return activated;
  }
  public static IntPtr FindVisibleTopLevelWindow(string caption, uint wantedProcessId) {
    IntPtr found = IntPtr.Zero;
    EnumWindows((hwnd, parameter) => {
      uint processId;
      GetWindowThreadProcessId(hwnd, out processId);
      if (processId != wantedProcessId || !IsWindowVisible(hwnd)) return true;
      var text = new StringBuilder(2048);
      GetWindowText(hwnd, text, text.Capacity);
      if (!String.Equals(text.ToString(), caption, StringComparison.Ordinal)) return true;
      found = hwnd;
      return false;
    }, IntPtr.Zero);
    return found;
  }
}
"@
}

function Find-UiElementsByName([object]$Root, [string]$Name) {
    $condition = [Windows.Automation.PropertyCondition]::new(
        [Windows.Automation.AutomationElement]::NameProperty,
        $Name
    )
    @($Root.FindAll([Windows.Automation.TreeScope]::Descendants, $condition))
}

function Find-VisibleUiElement([object]$Root, [string]$Name, [int]$ProcessId, [int]$TimeoutMs, [string]$Preference = '') {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $visible = @(Find-UiElementsByName $Root $Name | Where-Object {
            $bounds = $_.Current.BoundingRectangle
            [int]$_.Current.ProcessId -eq $ProcessId -and $bounds.Width -gt 0 -and $bounds.Height -gt 0
        })
        if (-not [string]::IsNullOrWhiteSpace($Preference)) {
            $preferred = @($visible | Where-Object { [string]$_.Current.ControlType.ProgrammaticName -match $Preference })
            if ($preferred.Count -gt 0) { return $preferred[0] }
        }
        if ($visible.Count -gt 0) { return $visible[0] }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    throw ('Visible UI element unavailable: ' + $Name)
}

function Get-UiElementDiagnostic([object]$Element) {
    if ($null -eq $Element) { return $null }
    try {
        $bounds = $Element.Current.BoundingRectangle
        [ordered]@{
            name = [string]$Element.Current.Name
            automation_id = [string]$Element.Current.AutomationId
            class_name = [string]$Element.Current.ClassName
            control_type = [string]$Element.Current.ControlType.ProgrammaticName
            process_id = [int]$Element.Current.ProcessId
            is_offscreen = [bool]$Element.Current.IsOffscreen
            bounds = [ordered]@{left=$bounds.Left;top=$bounds.Top;width=$bounds.Width;height=$bounds.Height}
        }
    } catch {
        [ordered]@{error=$_.Exception.Message}
    }
}

function Get-UiAncestorDiagnostics([object]$Element) {
    $ancestors = @()
    $current = $Element
    for ($depth = 0; $depth -lt 8 -and $null -ne $current; $depth++) {
        try { $current = [Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($current) }
        catch { break }
        if ($null -ne $current) { $ancestors += Get-UiElementDiagnostic $current }
    }
    @($ancestors)
}

function Get-VisibleMenuItemDiagnostics([object]$Desktop, [int]$ProcessId) {
    $items = @($Desktop.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition) | Where-Object {
        $bounds = $_.Current.BoundingRectangle
        [int]$_.Current.ProcessId -eq $ProcessId -and
            $_.Current.ControlType -eq [Windows.Automation.ControlType]::MenuItem -and
            $bounds.Width -gt 0 -and $bounds.Height -gt 0
    })
    @($items | Select-Object -First 30 | ForEach-Object { Get-UiElementDiagnostic $_ })
}

function Select-RibbonTab([object]$ExcelRoot, [string]$Label) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $tabCount = 0
    do {
        $tabs = @(Find-UiElementsByName $ExcelRoot $Label | Where-Object {
            $_.Current.ControlType -eq [Windows.Automation.ControlType]::TabItem
        })
        $tabCount = $tabs.Count
        if ($tabs.Count -eq 1) {
            $tabs[0].GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
            Start-Sleep -Milliseconds 350
            return
        }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt 15000)
    throw ('Ribbon tab identity rejected: ' + $Label + '; count=' + $tabCount)
}

function Expand-RibbonControl([object]$ExcelRoot, [string]$Label) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $script:LastRibbonExpansionDiagnostic = [ordered]@{label=$Label;candidates=@();selected=$null;ancestors=@()}
    do {
        $candidates = @(Find-UiElementsByName $ExcelRoot $Label | Where-Object {
            $bounds = $_.Current.BoundingRectangle
            $bounds.Width -gt 0 -and $bounds.Height -gt 0
        })
        $script:LastRibbonExpansionDiagnostic.candidates = @($candidates | ForEach-Object { Get-UiElementDiagnostic $_ })
        foreach ($candidate in $candidates) {
            try {
                $pattern = $candidate.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern)
                $pattern.Expand()
                $script:LastRibbonExpansionDiagnostic.selected = Get-UiElementDiagnostic $candidate
                $script:LastRibbonExpansionDiagnostic.ancestors = @(Get-UiAncestorDiagnostics $candidate)
                Start-Sleep -Milliseconds 400
                return $candidate
            } catch { }
        }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt 10000)
    throw ('Expandable Ribbon control unavailable: ' + $Label)
}

function Expand-CollapsedRibbonGroupForChild(
    [object]$ExcelRoot,
    [object]$Desktop,
    [int]$ProcessId,
    [string]$GroupLabel,
    [string]$ChildLabel
) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $diagnostic = [ordered]@{group_label=$GroupLabel;child_label=$ChildLabel;candidates=@();attempts=@();scroll_attempts=@()}
    do {
        if (Test-VisibleUiLabels $Desktop $ProcessId @($ChildLabel)) { return }
        $groupChunks = @(Find-UiElementsByName $ExcelRoot $GroupLabel | Where-Object {
            $_.Current.ControlType -eq [Windows.Automation.ControlType]::Group -and
                [string]$_.Current.ClassName -ceq 'NetUIChunk'
        })
        if ($groupChunks.Count -eq 1 -and [bool]$groupChunks[0].Current.IsOffscreen) {
            $leftLabel = ConvertFrom-CodePoints @(0xC5F4,0x20,0xC67C,0xCABD,0xC73C,0xB85C)
            $leftButtons = @(Find-UiElementsByName $ExcelRoot $leftLabel | Where-Object {
                [int]$_.Current.ProcessId -eq $ProcessId -and
                    $_.Current.ControlType -eq [Windows.Automation.ControlType]::Button -and
                    -not [bool]$_.Current.IsOffscreen
            })
            if ($leftButtons.Count -eq 1) {
                $leftButtons[0].GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()
                $diagnostic.scroll_attempts += Get-UiElementDiagnostic $leftButtons[0]
                Start-Sleep -Milliseconds 250
                continue
            }
        }
        $candidates = @(Find-UiElementsByName $ExcelRoot $GroupLabel | Where-Object {
            $bounds = $_.Current.BoundingRectangle
            $bounds.Width -gt 0 -and $bounds.Height -gt 0 -and -not [bool]$_.Current.IsOffscreen
        })
        $diagnostic.candidates = @($candidates | ForEach-Object { Get-UiElementDiagnostic $_ })
        foreach ($candidate in $candidates) {
            $pattern = $null
            if (-not $candidate.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$pattern)) { continue }
            if ($pattern.Current.ExpandCollapseState -ne [Windows.Automation.ExpandCollapseState]::Collapsed) { continue }
            $diagnostic.attempts += Get-UiElementDiagnostic $candidate
            $pattern.Expand()
            Start-Sleep -Milliseconds 350
            if (Test-VisibleUiLabels $Desktop $ProcessId @($ChildLabel)) { return }
        }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt 10000)
    throw ('Collapsed Ribbon group child unavailable: ' + $GroupLabel + ' > ' + $ChildLabel + '; diagnostic=' + ($diagnostic | ConvertTo-Json -Depth 8 -Compress))
}

function Test-VisibleUiLabels([object]$Desktop, [int]$ProcessId, [string[]]$ExpectedLabels) {
    foreach ($label in $ExpectedLabels) {
        $matches = @(Find-UiElementsByName $Desktop $label | Where-Object {
            $bounds = $_.Current.BoundingRectangle
            [int]$_.Current.ProcessId -eq $ProcessId -and
                $_.Current.ControlType -eq [Windows.Automation.ControlType]::MenuItem -and
                $bounds.Width -gt 0 -and $bounds.Height -gt 0
        })
        if ($matches.Count -eq 0) { return $false }
    }
    return $true
}

function Get-VisibleMenuChildDiagnostics([object]$Desktop, [int]$ProcessId, [string]$Label) {
    $rows = @()
    foreach ($menu in @(Find-UiElementsByName $Desktop $Label | Where-Object {
        $bounds = $_.Current.BoundingRectangle
        $_.Current.ControlType -eq [Windows.Automation.ControlType]::Menu -and
            $bounds.Width -gt 0 -and $bounds.Height -gt 0
    })) {
        foreach ($child in @($menu.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition))) {
            $bounds = $child.Current.BoundingRectangle
            if (-not [string]::IsNullOrWhiteSpace([string]$child.Current.Name) -and $bounds.Width -gt 0 -and $bounds.Height -gt 0) {
                $rows += [pscustomobject][ordered]@{
                    name=[string]$child.Current.Name
                    automation_id=[string]$child.Current.AutomationId
                    class_name=[string]$child.Current.ClassName
                    control_type=[string]$child.Current.ControlType.ProgrammaticName
                    process_id=[int]$child.Current.ProcessId
                    enabled=[bool]$child.Current.IsEnabled
                    offscreen=[bool]$child.Current.IsOffscreen
                    bounds=[ordered]@{left=$bounds.Left;top=$bounds.Top;width=$bounds.Width;height=$bounds.Height}
                }
            }
        }
        if ($rows.Count -eq 0) {
            $menuBounds = $menu.Current.BoundingRectangle
            foreach ($child in @($Desktop.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition))) {
                $bounds = $child.Current.BoundingRectangle
                if (
                    -not [string]::IsNullOrWhiteSpace([string]$child.Current.Name) -and
                    $bounds.Width -gt 0 -and $bounds.Height -gt 0 -and
                    $bounds.Left -ge $menuBounds.Left -and $bounds.Right -le $menuBounds.Right -and
                    $bounds.Top -ge $menuBounds.Top -and $bounds.Bottom -le $menuBounds.Bottom
                ) {
                    $rows += [pscustomobject][ordered]@{
                        name=[string]$child.Current.Name
                        automation_id=[string]$child.Current.AutomationId
                        class_name=[string]$child.Current.ClassName
                        control_type=[string]$child.Current.ControlType.ProgrammaticName
                        process_id=[int]$child.Current.ProcessId
                        enabled=[bool]$child.Current.IsEnabled
                        offscreen=[bool]$child.Current.IsOffscreen
                        bounds=[ordered]@{left=$bounds.Left;top=$bounds.Top;width=$bounds.Width;height=$bounds.Height}
                    }
                }
            }
        }
    }
    return $rows
}

function Get-UiGroupAnchorKey([object]$Element) {
    $current = $Element
    for ($depth = 0; $depth -lt 5 -and $null -ne $current; $depth++) {
        try { $current = [Windows.Automation.TreeWalker]::ControlViewWalker.GetParent($current) }
        catch { return '' }
        if ($null -ne $current -and $current.Current.ControlType -eq [Windows.Automation.ControlType]::Group) {
            $bounds = $current.Current.BoundingRectangle
            return ([string]$current.Current.ClassName + '|' + [string]$current.Current.Name + '|' +
                [string]$bounds.Left + '|' + [string]$bounds.Top + '|' + [string]$bounds.Width + '|' + [string]$bounds.Height)
        }
    }
    return ''
}

function Expand-RibbonControlForExpectedChildren(
    [object]$ExcelRoot,
    [object]$Desktop,
    [int]$ProcessId,
    [string]$Label,
    [string[]]$ExpectedLabels,
    [string]$Surface
) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $requiredGroupAnchor = if ($script:RibbonSurfaceGroupAnchors.ContainsKey($Surface)) {
        [string]$script:RibbonSurfaceGroupAnchors[$Surface]
    } else { '' }
    $expandedKeys = @{}
    $script:LastRibbonExpansionDiagnostic = [ordered]@{label=$Label;candidates=@();attempts=@();selected=$null;ancestors=@();required_group_anchor=$requiredGroupAnchor}

    do {
        $candidates = @(Find-UiElementsByName $ExcelRoot $Label | Where-Object {
            $bounds = $_.Current.BoundingRectangle
            $groupAnchor = Get-UiGroupAnchorKey $_
            $bounds.Width -gt 0 -and $bounds.Height -gt 0 -and
                ([string]::IsNullOrWhiteSpace($requiredGroupAnchor) -or $groupAnchor -ceq $requiredGroupAnchor)
        })
        $script:LastRibbonExpansionDiagnostic.candidates = @($candidates | ForEach-Object { Get-UiElementDiagnostic $_ })
        $expandedThisPass = $false
        foreach ($candidate in $candidates) {
            try { $candidateKey = (@($candidate.GetRuntimeId()) -join '.') }
            catch { $candidateKey = '' }
            if ([string]::IsNullOrWhiteSpace($candidateKey)) {
                $candidateKey = ((Get-UiElementDiagnostic $candidate) | ConvertTo-Json -Depth 4 -Compress)
            }
            if ($expandedKeys.ContainsKey($candidateKey)) { continue }
            $expandedKeys[$candidateKey] = $true
            try {
                $pattern = $candidate.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern)
                $pattern.Expand()
            } catch { continue }

            $expandedThisPass = $true
            $script:LastRibbonExpansionDiagnostic.selected = Get-UiElementDiagnostic $candidate
            $script:LastRibbonExpansionDiagnostic.ancestors = @(Get-UiAncestorDiagnostics $candidate)
            $script:LastRibbonExpansionDiagnostic.attempts += [ordered]@{
                candidate=$script:LastRibbonExpansionDiagnostic.selected
                ancestors=$script:LastRibbonExpansionDiagnostic.ancestors
            }
            $childrenWatch = [Diagnostics.Stopwatch]::StartNew()
            do {
                if (Test-VisibleUiLabels $Desktop $ProcessId $ExpectedLabels) {
                    $groupAnchor = Get-UiGroupAnchorKey $candidate
                    if ([string]::IsNullOrWhiteSpace($groupAnchor)) { throw ('Ribbon group identity unavailable: ' + $Label) }
                    if ([string]::IsNullOrWhiteSpace($requiredGroupAnchor)) {
                        $script:RibbonSurfaceGroupAnchors[$Surface] = $groupAnchor
                    }
                    return $candidate
                }
                Start-Sleep -Milliseconds 100
            } while ($childrenWatch.ElapsedMilliseconds -lt 2000)
            break
        }
        if ($expandedThisPass) {
            Start-Sleep -Milliseconds 100
            continue
        }
        [Windows.Forms.SendKeys]::SendWait('{ESC}')
        Start-Sleep -Milliseconds 200
        $expandedKeys = @{}
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt 15000)
    $script:LastRibbonExpansionDiagnostic.menu_children = @(Get-VisibleMenuChildDiagnostics $Desktop $ProcessId $Label)
    throw ('Expandable Ribbon control with exact children unavailable: ' + $Label + '; diagnostic=' + ($script:LastRibbonExpansionDiagnostic | ConvertTo-Json -Depth 12 -Compress))
}

function Invoke-UiElement([object]$Element, [string]$Label) {
    if ($null -eq $Element) { throw ('UI element missing: ' + $Label) }
    try { $Element.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke() }
    catch { throw ('UI invoke failed: ' + $Label) }
}

function Get-XmlLabel([xml]$RibbonXml, [string]$ControlId) {
    $node = $RibbonXml.SelectSingleNode('//*[@id="' + $ControlId + '"]')
    if ($null -eq $node -or [string]::IsNullOrWhiteSpace([string]$node.label)) { throw ('Ribbon XML label missing: ' + $ControlId) }
    [string]$node.label
}

function Get-FeatureLabel([object]$FeatureContract, [string]$FeatureId) {
    $feature = @($FeatureContract.features | Where-Object { [string]$_.id -ceq $FeatureId })
    if ($feature.Count -ne 1) { throw ('Feature label identity rejected: ' + $FeatureId) }
    [string]$feature[0].ribbon.label
}

function Get-FormLayout([object]$ProductManifest, [string]$FormName) {
    $binding = @($ProductManifest.form_bindings | Where-Object { [IO.Path]::GetFileNameWithoutExtension([IO.Path]::GetFileNameWithoutExtension([string]$_.layout)) -ceq $FormName })
    if ($binding.Count -ne 1) { throw ('Product form binding rejected: ' + $FormName) }
    $layoutPath = Join-Path $SourceRoot ([string]$binding[0].layout)
    Get-Content -LiteralPath $layoutPath -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Get-ExcelClientBounds([IntPtr]$Handle) {
    $rect = New-Object NxUiRect
    $origin = New-Object NxUiPoint
    if (-not [NxProductUiApi]::GetClientRect($Handle, [ref]$rect)) { throw 'Excel client rectangle unavailable' }
    if (-not [NxProductUiApi]::ClientToScreen($Handle, [ref]$origin)) { throw 'Excel client origin unavailable' }
    $width = [int]($rect.Right - $rect.Left); $height = [int]($rect.Bottom - $rect.Top)
    if ($width -le 0 -or $height -le 0) { throw 'Excel client rectangle is empty' }
    [ordered]@{left=$origin.X;top=$origin.Y;width=$width;height=$height}
}

function Get-ScratchFingerprint([object]$Book) {
    $worksheets = $null; $sheet = $null; $range = $null
    try {
        $worksheets = $Book.Worksheets; $sheet = $worksheets.Item(1); $range = $sheet.Range('A1:B4')
        $rows = @([string]$worksheets.Count, [string]$range.Address($true,$true,1,$false))
        for ($row = 1; $row -le 4; $row++) {
            for ($column = 1; $column -le 2; $column++) {
                $cell = $null
                try { $cell = $range.Cells.Item($row,$column); $rows += ([string]$cell.Value2 + '|' + [string]$cell.Formula) }
                finally { Release-ComObject $cell }
            }
        }
        Get-TextSha256 (($rows -join "`n") + "`n")
    } finally { Release-ComObject $range; Release-ComObject $sheet; Release-ComObject $worksheets }
}

function Select-ScratchRange([object]$Book) {
    $worksheets = $null; $sheet = $null; $range = $null
    try { [void]$Book.Activate(); $worksheets=$Book.Worksheets; $sheet=$worksheets.Item(1); $range=$sheet.Range('A1:B4'); [void]$range.Select() }
    finally { Release-ComObject $range; Release-ComObject $sheet; Release-ComObject $worksheets }
}

function Find-OwnedTopLevelForm([string]$Caption, [int]$ProcessId, [int]$TimeoutMs) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        $handle = [NxProductUiApi]::FindVisibleTopLevelWindow($Caption, [uint32]$ProcessId)
        if ($handle -ne [IntPtr]::Zero) {
            $form = [Windows.Automation.AutomationElement]::FromHandle($handle)
            if ($null -ne $form -and $form.Current.NativeWindowHandle -ne 0) { return $form }
        }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    throw ('Owned form unavailable: ' + $Caption)
}

function Wait-WindowGone([string]$Caption, [int]$ProcessId, [int]$TimeoutMs) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    do {
        if ([NxProductUiApi]::FindVisibleTopLevelWindow($Caption, [uint32]$ProcessId) -eq [IntPtr]::Zero) { return $true }
        Start-Sleep -Milliseconds 100
    } while ($watch.ElapsedMilliseconds -lt $TimeoutMs)
    return $false
}

function Invoke-RibbonPath([object]$ExcelRoot, [object]$Desktop, [int]$ProcessId, [string]$ProductTabLabel, [xml]$RibbonXml, [string]$ParentControlId, [string]$ActionControlId, [string]$ActionLabelOverride = '', [string]$CollapsedGroupControlId = '', [string]$CollapsedProbeControlId = '') {
    Select-RibbonTab $ExcelRoot $ProductTabLabel
    if (-not [string]::IsNullOrWhiteSpace($CollapsedGroupControlId)) {
        if ([string]::IsNullOrWhiteSpace($CollapsedProbeControlId)) { throw 'Collapsed Ribbon route probe is required' }
        $collapsedGroupLabel = Get-XmlLabel $RibbonXml $CollapsedGroupControlId
        $collapsedProbeLabel = Get-XmlLabel $RibbonXml $CollapsedProbeControlId
        Expand-CollapsedRibbonGroupForChild $ExcelRoot $Desktop $ProcessId $collapsedGroupLabel $collapsedProbeLabel
    }
    $actionLabel = if ([string]::IsNullOrWhiteSpace($ActionLabelOverride)) { Get-XmlLabel $RibbonXml $ActionControlId } else { $ActionLabelOverride }
    if (-not [string]::IsNullOrWhiteSpace($ParentControlId)) {
        $parentLabel = Get-XmlLabel $RibbonXml $ParentControlId
        [void](Expand-RibbonControlForExpectedChildren $ExcelRoot $Desktop $ProcessId $parentLabel @($actionLabel) ('route:' + $ParentControlId + ':' + $ActionControlId))
    }
    $action = Find-VisibleUiElement $Desktop $actionLabel $ProcessId 10000 'MenuItem|Button'
    Invoke-UiElement $action $actionLabel
    Start-Sleep -Milliseconds 350
}

function Invoke-MenuCase([object]$ExcelRoot, [object]$Desktop, [int]$ProcessId, [string]$TabLabel, [string]$Surface, [string]$MenuLabel, [string[]]$ExpectedLabels, [string[]]$RoutePrefixes, [string]$CatalogSha, [string]$ParentGroupLabel = '', [string]$ParentGroupProbeLabel = '') {
    Select-RibbonTab $ExcelRoot $TabLabel
    if (-not [string]::IsNullOrWhiteSpace($ParentGroupLabel)) {
        Expand-CollapsedRibbonGroupForChild $ExcelRoot $Desktop $ProcessId $ParentGroupLabel $ParentGroupProbeLabel
    }
    $menu = Expand-RibbonControlForExpectedChildren $ExcelRoot $Desktop $ProcessId $MenuLabel $ExpectedLabels $Surface
    $bounds = $menu.Current.BoundingRectangle
    $children = @()
    foreach ($label in $ExpectedLabels) {
        try {
            $element = Find-VisibleUiElement $Desktop $label $ProcessId 10000 ''
        } catch {
            $diagnostic = [ordered]@{
                surface=$Surface
                tab_label=$TabLabel
                menu_label=$MenuLabel
                missing_child=$label
                expansion=$script:LastRibbonExpansionDiagnostic
                visible_menu_items=@(Get-VisibleMenuItemDiagnostics $Desktop $ProcessId)
            }
            throw ('Dynamic menu child unavailable; diagnostic=' + ($diagnostic | ConvertTo-Json -Depth 12 -Compress))
        }
        $childBounds = $element.Current.BoundingRectangle
        $children += [ordered]@{label=$label;control_type=[string]$element.Current.ControlType.ProgrammaticName;left=$childBounds.Left;top=$childBounds.Top;width=$childBounds.Width;height=$childBounds.Height}
    }
    [Windows.Forms.SendKeys]::SendWait('{ESC}')
    [ordered]@{
        surface=$Surface;menu_label=$MenuLabel;bounds=[ordered]@{left=$bounds.Left;top=$bounds.Top;width=$bounds.Width;height=$bounds.Height}
        child_count=$children.Count;children=$children;typed_route_prefixes=$RoutePrefixes
        typed_routes_verified=$true;typed_route_source_sha256=$CatalogSha
    }
}

function Invoke-FavoritesMenuCase(
    [object]$ExcelRoot,
    [object]$Desktop,
    [int]$ProcessId,
    [string]$TabLabel,
    [string]$Surface,
    [string]$MenuLabel,
    [string]$FavoriteEditor,
    [string]$FavoriteEmpty,
    [string]$FavoriteAdd,
    [string]$FavoriteRemove,
    [string[]]$AllowedFavoriteLabels,
    [string]$CatalogSha,
    [string]$ParentGroupLabel = '',
    [string]$ParentGroupProbeLabel = ''
) {
    Select-RibbonTab $ExcelRoot $TabLabel
    if (-not [string]::IsNullOrWhiteSpace($ParentGroupLabel)) {
        Expand-CollapsedRibbonGroupForChild $ExcelRoot $Desktop $ProcessId $ParentGroupLabel $ParentGroupProbeLabel
    }
    $menu = Expand-RibbonControlForExpectedChildren $ExcelRoot $Desktop $ProcessId $MenuLabel @($FavoriteEditor,$FavoriteAdd) $Surface
    $bounds = $menu.Current.BoundingRectangle
    $rows = @(Get-VisibleMenuChildDiagnostics $Desktop $ProcessId $MenuLabel | Where-Object {
        [int]$_.process_id -eq $ProcessId -and [string]$_.control_type -eq 'ControlType.MenuItem'
    } | Sort-Object @{Expression={[double]$_.bounds.top}}, @{Expression={[double]$_.bounds.left}})
    $labels = @($rows | ForEach-Object { [string]$_.name })
    $routePrefixes = @('nx1|entry|','fav2|add|')
    if (($labels -join [char]31) -ceq (@($FavoriteEditor,$FavoriteEmpty,$FavoriteAdd) -join [char]31)) {
        $favoriteState = 'empty'
    } else {
        if ($labels.Count -lt 4 -or $labels[0] -cne $FavoriteEditor -or $labels[$labels.Count-2] -cne $FavoriteAdd -or $labels[$labels.Count-1] -cne $FavoriteRemove) {
            throw ('Favorites dynamic menu state rejected: ' + ($labels -join ' | '))
        }
        $favoriteLabels = @($labels[1..($labels.Count-3)])
        if (@($favoriteLabels | Sort-Object -Unique).Count -ne $favoriteLabels.Count) { throw 'Favorites dynamic menu contains duplicate feature labels' }
        foreach ($label in $favoriteLabels) {
            if ($AllowedFavoriteLabels -cnotcontains $label) { throw ('Favorites dynamic menu feature label rejected: ' + $label) }
        }
        $favoriteState = 'populated'
        $routePrefixes += 'fav2|remove|'
    }
    $children = @($rows | ForEach-Object {
        [ordered]@{
            label=[string]$_.name
            control_type=[string]$_.control_type
            left=[double]$_.bounds.left
            top=[double]$_.bounds.top
            width=[double]$_.bounds.width
            height=[double]$_.bounds.height
        }
    })
    [Windows.Forms.SendKeys]::SendWait('{ESC}')
    [ordered]@{
        surface=$Surface;menu_label=$MenuLabel;bounds=[ordered]@{left=$bounds.Left;top=$bounds.Top;width=$bounds.Width;height=$bounds.Height}
        child_count=$children.Count;children=$children;typed_route_prefixes=$routePrefixes
        typed_routes_verified=$true;typed_route_source_sha256=$CatalogSha
    }
}

function Invoke-FormCase([object]$Case, [object]$ExcelRoot, [object]$Desktop, [int]$ProcessId, [string]$ProductTabLabel, [xml]$RibbonXml, [object]$FeatureContract, [object]$ProductManifest, [object]$SurfaceContract, [object]$ScratchBook, [object]$ClientBounds, [object[]]$AttestedForms, [string]$ControlAttestationSha) {
    Select-ScratchRange $ScratchBook
    $before = Get-ScratchFingerprint $ScratchBook
    $layout = Get-FormLayout $ProductManifest ([string]$Case.form)
    # BindSelection may replace a form's design-time title before Show.
    if ($null -ne $Case.PSObject.Properties['runtime_caption']) {
        if ([string]::IsNullOrWhiteSpace([string]$Case.runtime_caption)) { throw 'Runtime form caption override is empty' }
        $layout.caption = [string]$Case.runtime_caption
    }
    $attested = @($AttestedForms | Where-Object { [string]$_.name -ceq [string]$Case.form })
    if ($attested.Count -ne 1) { throw ('ProductRibbon UserForms attestation missing: ' + [string]$Case.form) }
    if ([string]$attested[0].default_control -cne [string]$layout.default_control) { throw ('Attested default control mismatch: ' + [string]$Case.form) }
    $layoutCancelControl = if ($null -eq $layout.cancel_control) { $null } else { [string]$layout.cancel_control }
    $attestedCancelControl = if ($null -eq $attested[0].cancel_control) { $null } else { [string]$attested[0].cancel_control }
    if (($null -eq $layoutCancelControl) -ne ($null -eq $attestedCancelControl) -or ($null -ne $layoutCancelControl -and $layoutCancelControl -cne $attestedCancelControl)) { throw ('Attested cancel control mismatch: ' + [string]$Case.form) }
    $attestedDefault = @($attested[0].controls | Where-Object { [string]$_.name -ceq [string]$layout.default_control -and [string]$_.type -ceq 'CommandButton' })
    if ($attestedDefault.Count -ne 1) { throw ('Attested default CommandButton missing: ' + [string]$Case.form) }
    if ($null -ne $layoutCancelControl) {
        $attestedCancel = @($attested[0].controls | Where-Object { [string]$_.name -ceq $layoutCancelControl -and [string]$_.type -ceq 'CommandButton' })
        if ($attestedCancel.Count -ne 1) { throw ('Attested cancel CommandButton missing: ' + [string]$Case.form) }
    }
    $caseActionLabel = if ($null -eq $Case.PSObject.Properties['action_label']) { '' } else { [string]$Case.action_label }
    $actionOverride = if (-not [string]::IsNullOrWhiteSpace($caseActionLabel)) {
        $caseActionLabel
    } elseif ([string]::IsNullOrWhiteSpace([string]$Case.feature_id)) {
        ''
    } else {
        Get-FeatureLabel $FeatureContract ([string]$Case.feature_id)
    }
    $collapsedGroupId = if ($null -eq $Case.PSObject.Properties['collapsed_group_id']) { '' } else { [string]$Case.collapsed_group_id }
    $collapsedProbeId = if ($null -eq $Case.PSObject.Properties['collapsed_probe_id']) { '' } else { [string]$Case.collapsed_probe_id }
    Invoke-RibbonPath $ExcelRoot $Desktop $ProcessId $ProductTabLabel $RibbonXml ([string]$Case.parent_id) ([string]$Case.action_id) $actionOverride $collapsedGroupId $collapsedProbeId
    $form = Find-OwnedTopLevelForm ([string]$layout.caption) $ProcessId 15000
    $handle = [Int64]$form.Current.NativeWindowHandle
    $bounds = $form.Current.BoundingRectangle
    $surfaceProperty = $SurfaceContract.active_form_surfaces.PSObject.Properties[[string]$Case.form]
    if ($null -eq $surfaceProperty) { throw ('Form surface class missing: ' + [string]$Case.form) }
    $surface = [string]$surfaceProperty.Value
    if ([string]::IsNullOrWhiteSpace($surface)) { throw ('Form surface class empty: ' + [string]$Case.form) }
    $limits = $SurfaceContract.classes.$surface
    $widthRatio = $bounds.Width / [double]$ClientBounds.width
    $heightRatio = $bounds.Height / [double]$ClientBounds.height
    $enforceReference = $true
    if ($null -ne $SurfaceContract.reference.PSObject.Properties['enforce_geometry']) { $enforceReference = [bool]$SurfaceContract.reference.enforce_geometry }
    if ($enforceReference -and ($widthRatio -gt [double]$limits.max_width_ratio -or $heightRatio -gt [double]$limits.max_height_ratio)) {
        $widthRatioText = $widthRatio.ToString('F4', [Globalization.CultureInfo]::InvariantCulture)
        $heightRatioText = $heightRatio.ToString('F4', [Globalization.CultureInfo]::InvariantCulture)
        throw ('Form ratio exceeds surface contract: ' + [string]$Case.form + '; width_ratio=' + $widthRatioText + '; height_ratio=' + $heightRatioText + '; form_bounds=' + $bounds.Width + 'x' + $bounds.Height + '; client_bounds=' + $ClientBounds.width + 'x' + $ClientBounds.height)
    }
    $screen = [Windows.Forms.Screen]::FromHandle([IntPtr]$handle).WorkingArea
    $inside = $bounds.Left -ge $screen.Left -and $bounds.Top -ge $screen.Top -and $bounds.Right -le $screen.Right -and $bounds.Bottom -le $screen.Bottom
    if (-not $inside) { throw ('Form is clipped by the current working area: ' + [string]$Case.form) }
    $forbiddenMatches = @()
    $descendants = @($form.FindAll([Windows.Automation.TreeScope]::Descendants, [Windows.Automation.Condition]::TrueCondition))
    foreach ($element in $descendants) {
        $text = [string]$element.Current.Name
        foreach ($pattern in @($SurfaceContract.forbidden_user_text_patterns)) {
            if (-not [string]::IsNullOrWhiteSpace($text) -and $text.IndexOf([string]$pattern, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                $forbiddenMatches += [ordered]@{form=[string]$Case.form;pattern=[string]$pattern;text_sha256=Get-TextSha256 $text}
            }
        }
    }
    if ($forbiddenMatches.Count -ne 0) { throw ('Internal user text exposed: ' + [string]$Case.form) }
    if (-not [NxProductUiApi]::ActivateForegroundWindow([IntPtr]$handle)) { throw ('Form foreground activation failed: ' + [string]$Case.form) }
    Start-Sleep -Milliseconds 150
    $formClosed = $false
    if ($null -ne $layoutCancelControl) {
        [Windows.Forms.SendKeys]::SendWait('{ESC}')
        $closeMethod = 'escape-cancel'
        $formClosed = Wait-WindowGone ([string]$layout.caption) $ProcessId 3000
    }
    if (-not $formClosed) {
        if (-not [NxProductUiApi]::ActivateForegroundWindow([IntPtr]$handle)) { throw ('Form close fallback activation failed: ' + [string]$Case.form) }
        Start-Sleep -Milliseconds 150
        [Windows.Forms.SendKeys]::SendWait('%{F4}')
        $closeMethod = 'alt-f4'
        $formClosed = Wait-WindowGone ([string]$layout.caption) $ProcessId 5000
    }
    if (-not $formClosed) { throw ('Form did not close safely: ' + [string]$Case.form) }
    $after = Get-ScratchFingerprint $ScratchBook
    if ($after -ne $before) { throw ('Scratch workbook changed while inspecting form: ' + [string]$Case.form) }
    [ordered]@{
        form=[string]$Case.form;feature_id=[string]$Case.feature_id;ui_surface=$surface
        caption_sha256=Get-TextSha256 ([string]$layout.caption);native_handle=$handle
        bounds=[ordered]@{left=$bounds.Left;top=$bounds.Top;width=$bounds.Width;height=$bounds.Height}
        width_ratio=$widthRatio;height_ratio=$heightRatio;inside_work_area=$inside
        default_control=[string]$layout.default_control;cancel_control=$layoutCancelControl
        close_method=$closeMethod
        control_attestation='ProductUi.UserForms.json';control_attestation_sha256=$ControlAttestationSha
        workbook_before_sha256=$before;workbook_after_sha256=$after;forbidden_text_matches=@()
    }
}

try {
    New-Item -ItemType Directory -Force -Path $evidence | Out-Null
    $probeStage = 'preflight:user-documents'
    if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -ne 0) { throw 'ENVIRONMENT_MISMATCH: close all user Excel windows before ProductUi verification' }
    $failurePhase = 'build'
    $probeStage = 'product-ribbon-prerequisite'
    $powershell = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
    $ribbonArguments = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
        (Join-Path $SourceRoot 'tests/windows/Run-ProductRibbonSuite.ps1'),
        '-Suite', 'ProductRibbon', '-Mode', 'Green', '-RunId', $RunId,
        '-DataRoot', $DataRoot, '-EvidenceRoot', $evidence
    )
    if (-not [string]::IsNullOrWhiteSpace($SourceDigest)) {
        $ribbonArguments += @('-SourceDigest', $SourceDigest)
    }
    if (-not [string]::IsNullOrWhiteSpace($SnapshotDigest)) {
        $ribbonArguments += @('-SnapshotDigest', $SnapshotDigest)
    }
    if (-not [string]::IsNullOrWhiteSpace($ArtifactPath)) {
        $ArtifactPath = [IO.Path]::GetFullPath($ArtifactPath)
        if (-not (Test-Path -LiteralPath $ArtifactPath -PathType Leaf) -or [IO.Path]::GetExtension($ArtifactPath) -ine '.xlam') { throw 'ProductUi supplied Product artifact rejected' }
        $ribbonArguments += @('-ArtifactPath', $ArtifactPath)
    }
    & $powershell @ribbonArguments
    if ($LASTEXITCODE -ne 0) { throw 'ProductRibbon prerequisite failed' }
    if (-not (Test-Path -LiteralPath $artifact -PathType Leaf)) { throw 'Product.xlam prerequisite artifact missing' }
    $artifactSha = Get-HexDigest $artifact
    $userFormAttestationPath = Join-Path $evidence 'ProductRibbon.UserForms.json'
    if (-not (Test-Path -LiteralPath $userFormAttestationPath -PathType Leaf)) { throw 'ProductRibbon UserForms attestation missing' }
    $userFormAttestationSha = Get-HexDigest $userFormAttestationPath
    $userFormAttestation = Get-Content -LiteralPath $userFormAttestationPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($userFormAttestation.status -cne 'PASS' -or $userFormAttestation.run_id -cne $RunId -or $userFormAttestation.artifact_sha256 -cne $artifactSha -or $userFormAttestation.source_tree_sha256 -cne $SourceDigest -or $userFormAttestation.source_snapshot_sha256 -cne $SnapshotDigest) { throw 'ProductRibbon UserForms attestation binding rejected' }
    $attestedForms = @($userFormAttestation.measured.forms)
    $productManifest = Get-Content -LiteralPath (Join-Path $SourceRoot 'build/manifests/Product.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $expectedAttestedFormCount = @($productManifest.form_bindings).Count
    if ($expectedAttestedFormCount -le 0 -or $attestedForms.Count -ne $expectedAttestedFormCount) { throw 'ProductRibbon UserForms attestation inventory rejected' }
    $productUiAttestationPath = Join-Path $evidence 'ProductUi.UserForms.json'
    Copy-Item -LiteralPath $userFormAttestationPath -Destination $productUiAttestationPath -Force
    if ((Get-HexDigest $productUiAttestationPath) -cne $userFormAttestationSha) { throw 'ProductUi UserForms attestation copy rejected' }
    $failurePhase = 'probe'
    $probeStage = 'open-owned-excel'
    if (@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -ne 0) { throw 'ENVIRONMENT_MISMATCH: ProductRibbon prerequisite left an Excel process' }
    $baseline = Get-ExcelProcessBaseline
    $interactiveExcel = Start-InteractiveProductUiExcel $baseline
    $excel = $interactiveExcel.excel
    $binding = $interactiveExcel.binding
    $interactiveExcel = $null
    if (-not [bool]$binding.owned) { throw ('ProductUi Excel ownership rejected: ' + [string]$binding.failure) }
    $process = $binding.process
    Write-Json (Join-Path $evidence 'ProductUi.Ownership.json') ([ordered]@{pid=$binding.pid;started_utc=$binding.started_utc;executable=$binding.executable;hwnd=$binding.hwnd;run_id=$RunId})
    if (-not [bool]$excel.UserControl) { throw 'ProductUi Excel is not user controlled' }
    $excel.Visible = $true; $excel.DisplayAlerts = $false; $excel.WindowState = -4143
    $books = $excel.Workbooks
    if ([int]$books.Count -eq 0) { $hostBook = $books.Add() }
    elseif ([int]$books.Count -eq 1) { $hostBook = $books.Item(1) }
    else { throw 'ProductUi interactive Excel startup workbook count rejected' }
    $worksheets = $null; $sheet = $null; $range = $null
    try {
        $worksheets=$hostBook.Worksheets; $sheet=$worksheets.Item(1); $range=$sheet.Range('A1:B4')
        $values=New-Object 'object[,]' 4,2
        $values[0,0]='Key';$values[0,1]='Value';$values[1,0]='A';$values[1,1]=10;$values[2,0]='B';$values[2,1]=20;$values[3,0]='C';$values[3,1]=30
        $range.Value2=$values; [void]$range.Select()
    } finally { Release-ComObject $range; Release-ComObject $sheet; Release-ComObject $worksheets }
    $addin = $books.Open($artifact,$false,$true)
    if ($null -eq $addin -or -not [bool]$addin.IsAddin -or [string]$addin.Name -cne 'Product.xlam') { throw 'ProductUi add-in identity rejected' }
    [void]$hostBook.Activate(); Start-Sleep -Milliseconds 900

    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms
    Initialize-ProductUiApi
    $handle=[IntPtr][Int64]$excel.Hwnd
    $dpi=[int][NxProductUiApi]::GetDpiForWindow($handle)
    if($dpi -ne 192){throw ('ENVIRONMENT_MISMATCH: current Excel scale must be 200 percent; measured DPI='+$dpi)}
    if (-not [NxProductUiApi]::ActivateForegroundWindow($handle)) { throw 'Excel foreground verification failed before ProductUi UI search' }
    $excelRoot=[Windows.Automation.AutomationElement]::FromHandle($handle)
    $desktop=[Windows.Automation.AutomationElement]::RootElement
    $clientBounds=Get-ExcelClientBounds $handle
    [xml]$ribbonXml=Get-Content -LiteralPath (Join-Path $SourceRoot 'src/ribbon/customUI14.xml') -Raw -Encoding UTF8
    $featureContract=Get-Content -LiteralPath (Join-Path $SourceRoot 'contracts/feature-contract.json') -Raw -Encoding UTF8|ConvertFrom-Json
    $surfaceContract=Get-Content -LiteralPath (Join-Path $SourceRoot 'contracts/ui-surface-contract.json') -Raw -Encoding UTF8|ConvertFrom-Json
    $productTabLabel=[string]$ribbonXml.SelectSingleNode('//*[local-name()="tab" and @id="NX-TAB"]').label
    $homeTabLabel=ConvertFrom-CodePoints @(0xD648)
    $catalogPath=Join-Path $SourceRoot 'src/vba/ui/NxRibbonMenuCatalog.bas'
    $catalogSource=Get-Content -LiteralPath $catalogPath -Raw -Encoding UTF8
    foreach($token in @('tag=""nx1|entry|','tag=""fav2|','NxGeneratedNavigationMenuXml')){if($catalogSource -notlike ('*'+$token+'*')){throw ('Typed Ribbon route source missing: '+$token)}}
    $catalogSha=Get-HexDigest $catalogPath
    $navigationPath=Join-Path $SourceRoot 'src/resources/navigation.ko-KR.json'
    $navigation=Get-Content -LiteralPath $navigationPath -Raw -Encoding UTF8|ConvertFrom-Json
    $navigationSha=Get-HexDigest $navigationPath
    $managementLabels=@([regex]::Matches($catalogSource,'NxManagementButton\("[^"]+", "([^"]+)"')|ForEach-Object {$_.Groups[1].Value})
    $categoryLabels=@($navigation.categories|ForEach-Object {[string]$_.label})
    $favoriteEmpty=[regex]::Match($catalogSource,'id=""fav_empty"" label=""([^"]+)""').Groups[1].Value
    $favoriteEditor=[regex]::Match($catalogSource,'id=""fav_editor"" label=""([^"]+)""').Groups[1].Value
    $favoriteAdd=[regex]::Match($catalogSource,'id=""fav_add"" label=""([^"]+)""').Groups[1].Value
    $favoriteRemove=[regex]::Match($catalogSource,'id=""fav_remove"" label=""([^"]+)""').Groups[1].Value
    $allowedFavoriteLabels=@($featureContract.features|Where-Object {[string]$_.owner -cne 'ai' -and [string]$_.id -ne 'NX-DATA-DUPLICATE-LIST'}|ForEach-Object {[string]$_.ribbon.label}|Sort-Object -Unique)
    if($managementLabels.Count -ne 6 -or $categoryLabels.Count -ne 14 -or $allowedFavoriteLabels.Count -lt 1 -or [string]::IsNullOrWhiteSpace($favoriteEditor) -or [string]::IsNullOrWhiteSpace($favoriteEmpty) -or [string]::IsNullOrWhiteSpace($favoriteAdd) -or [string]::IsNullOrWhiteSpace($favoriteRemove)){throw 'Dynamic Ribbon catalog inventory rejected'}

    $probeStage='dynamic-menus'
    $menuCases=@()
    $surface='Product';$tabLabel=$productTabLabel;$prefix='NX-PROD'
    $parentGroupLabel=Get-XmlLabel $ribbonXml 'NX-GRP-MYEXCEL'
    $parentGroupProbeLabel=Get-XmlLabel $ribbonXml 'NX-PROD-FAVORITES'
    $menuCases+=Invoke-MenuCase $excelRoot $desktop ([int]$binding.pid) $tabLabel $surface (Get-XmlLabel $ribbonXml ($prefix+'-MANAGEMENT')) $managementLabels @('nx1|entry|') $catalogSha $parentGroupLabel $parentGroupProbeLabel
    $menuCases+=Invoke-FavoritesMenuCase $excelRoot $desktop ([int]$binding.pid) $tabLabel $surface (Get-XmlLabel $ribbonXml ($prefix+'-FAVORITES')) $favoriteEditor $favoriteEmpty $favoriteAdd $favoriteRemove $allowedFavoriteLabels $catalogSha $parentGroupLabel $parentGroupProbeLabel
    $menuCases+=Invoke-MenuCase $excelRoot $desktop ([int]$binding.pid) $tabLabel $surface (Get-XmlLabel $ribbonXml ($prefix+'-ALL')) $categoryLabels @('nx1|feature|','nx1|command|') $navigationSha $parentGroupLabel $parentGroupProbeLabel
    if($menuCases.Count -ne 3){throw 'Dynamic Product Ribbon three-path inventory rejected'}

    $sampleCases=@(
        [pscustomobject]@{form='FNxFocusSettings';feature_id='';parent_id='NX-PROD-FOCUS';action_id='focus_settings';action_label=(ConvertFrom-CodePoints @(0xD3EC,0xCEE4,0xC2A4,0xC140,0x20,0xC124,0xC815));collapsed_group_id='NX-GRP-MYEXCEL';collapsed_probe_id='NX-PROD-FAVORITES'},
        [pscustomobject]@{form='FNxAi';feature_id='';parent_id='';action_id='NX-UTIL-AI-WORKBENCH'},
        [pscustomobject]@{form='FNxDataAnalyze';feature_id='NX-DATA-UNIQUE-COUNT';parent_id='NX-DATA-MENU';action_id='NX-DATA-MENU-UNIQUE';runtime_caption=(ConvertFrom-CodePoints @(0xB0B4,0xC5D1,0xC140,0x20,0x2D,0x20,0xACE0,0xC720,0xAC12,0x20,0xAC1C,0xC218))},
        [pscustomobject]@{form='FNxDataNormalize';feature_id='NX-DATA-NORMALIZE';parent_id='NX-DATA-MENU';action_id='NX-DATA-MENU-NORMALIZE'},
        [pscustomobject]@{form='FNxDrawTable';feature_id='NX-DRAW-BUSINESS-TABLE';parent_id='';action_id='NX-STYLE-BUSINESS-TABLE-BUTTON'},
        [pscustomobject]@{form='FNxRoleStyle';feature_id='NX-DRAW-ROLE-STYLE';parent_id='';action_id='NX-STYLE-MYEXCEL-FORMAT-BUTTON'},
        [pscustomobject]@{form='FNxFolderCreate';feature_id='NX-FILE-FOLDER-CREATE';parent_id='NX-FILE-MENU';action_id='NX-FILE-MENU-FOLDER-CREATE'},
        [pscustomobject]@{form='FNxFileConsolidate';feature_id='NX-FILE-CONSOLIDATE';parent_id='NX-FILE-MENU';action_id='NX-FILE-MENU-CONSOLIDATE'},
        [pscustomobject]@{form='FNxCalculator';feature_id='NX-UTIL-CALCULATOR';parent_id='';action_id='NX-UTIL-CALCULATOR-BUTTON'},
        [pscustomobject]@{form='FNxSymbols';feature_id='NX-UTIL-SYMBOLS';parent_id='';action_id='NX-UTIL-SYMBOLS-BUTTON'}
    )
    $formCases=@()
    $focusCase=$sampleCases[0]
    $probeStage='form:'+[string]$focusCase.form
    $formCases+=Invoke-FormCase $focusCase $excelRoot $desktop ([int]$binding.pid) $productTabLabel $ribbonXml $featureContract $productManifest $surfaceContract $hostBook $clientBounds $attestedForms $userFormAttestationSha

    $probeStage='direct-no-popup'
    Select-ScratchRange $hostBook
    $directBefore=Get-ScratchFingerprint $hostBook
    Invoke-RibbonPath $excelRoot $desktop ([int]$binding.pid) $productTabLabel $ribbonXml 'NX-COPY-MENU' 'NX-COPY-VISIBLE' (Get-FeatureLabel $featureContract 'NX-DATA-COPY-VISIBLE')
    Start-Sleep -Milliseconds 500
    $windowCondition=[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$binding.pid)
    $popupWindows=@($desktop.FindAll([Windows.Automation.TreeScope]::Children,$windowCondition)|Where-Object {$_.Current.ControlType -eq [Windows.Automation.ControlType]::Window -and [Int64]$_.Current.NativeWindowHandle -ne [Int64]$excel.Hwnd})
    $directAfter=Get-ScratchFingerprint $hostBook
    if($popupWindows.Count -ne 0 -or $directBefore -ne $directAfter){throw 'Direct Ribbon case popup or mutation rejected'}
    $directCase=[ordered]@{feature_id='NX-DATA-COPY-VISIBLE';popup_count=0;workbook_before_sha256=$directBefore;workbook_after_sha256=$directAfter}

    foreach($case in @($sampleCases | Select-Object -Skip 1)){$probeStage='form:'+[string]$case.form;$formCases+=Invoke-FormCase $case $excelRoot $desktop ([int]$binding.pid) $productTabLabel $ribbonXml $featureContract $productManifest $surfaceContract $hostBook $clientBounds $attestedForms $userFormAttestationSha}
    if($formCases.Count -ne 10){throw 'ProductUi ten-form inventory rejected'}
    $sampleNames=@($surfaceContract.native_200_sample_forms|ForEach-Object {[string]$_})
    if(($sampleNames -join [char]31) -cne (@($formCases|ForEach-Object {[string]$_.form}) -join [char]31)){throw 'ProductUi sample order differs from UI surface contract'}
    if((Get-HexDigest $artifact) -ne $artifactSha){throw 'Product.xlam changed during ProductUi verification'}

    $failurePhase='evidence';$probeStage='write-evidence'
    Write-Json (Join-Path $evidence 'ProductUi.Current200.json') ([ordered]@{
        schema_version=1;suite='ProductUi';status='PASS';run_id=$RunId;artifact_sha256=$artifactSha
        source_tree_sha256=$SourceDigest;source_snapshot_sha256=$SnapshotDigest
        windows_scale_percent=200;scale_changed_by_test=$false;excel_client_bounds=$clientBounds
        menu_cases=$menuCases;direct_case=$directCase;form_cases=$formCases;forbidden_text_matches=@()
    })
} catch {
    $failureMessage=$_.Exception.Message
    $terminalCode=switch($failurePhase){'preflight'{10};'build'{22};'probe'{23};'evidence'{24};default{24}}
    [Console]::Error.WriteLine(('ProductUi '+$failurePhase+' failure at '+$probeStage+': '+$failureMessage))
} finally {
    try {
        if($null -ne $binding -and [bool]$binding.owned -and ('NxProductUiApi' -as [type])) {
            [void](Assert-ExactExcelProcessOwnership $binding)
            [NxProductUiApi]::CloseOwnedForms([uint32]$binding.pid)
            Start-Sleep -Milliseconds 350
        }
    } catch { $cleanupFailures += $_.Exception.Message }
    try { if($null -ne $addin){try{$addin.Close($false)}catch{$cleanupFailures+=$_.Exception.Message};Release-ComObject $addin} } catch { $cleanupFailures += $_.Exception.Message }
    try { if($null -ne $hostBook){try{$hostBook.Close($false)}catch{$cleanupFailures+=$_.Exception.Message};Release-ComObject $hostBook} } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $books } catch { $cleanupFailures += $_.Exception.Message }
    try { if($null -ne $binding -and [bool]$binding.owned -and $null -ne $excel){[void](Assert-ExactExcelProcessOwnership $binding)} } catch { $cleanupFailures += $_.Exception.Message }
    try { if($null -ne $excel){try{$excel.Quit()}catch{$cleanupFailures+=$_.Exception.Message};Release-ComObject $excel} } catch { $cleanupFailures += $_.Exception.Message }
    try { [GC]::Collect();[GC]::WaitForPendingFinalizers() } catch { $cleanupFailures += $_.Exception.Message }
    try {
        if($null -ne $process){
            $result=Stop-ExactProcessAfterGrace -Process $process -Label 'ProductUi Excel' -GraceMs 10000 -Detailed
            $cleanupExitMode=[string]$result.exit_mode
            if($result.failure){$cleanupFailures+=[string]$result.failure}
        }
    } catch { $cleanupFailures += $_.Exception.Message }
    try {
        if($cleanupExitMode -cne 'NATURAL' -and $null -ne $binding -and [bool]$binding.owned){
            $remainingProcess=Get-Process -Id ([int]$binding.pid) -ErrorAction SilentlyContinue
            if($null -ne $remainingProcess){
                $remainingStarted=$remainingProcess.StartTime.ToUniversalTime().ToString('o')
                $remainingExecutable=[IO.Path]::GetFullPath($remainingProcess.MainModule.FileName)
                if($remainingStarted -cne [string]$binding.started_utc -or $remainingExecutable -cne [string]$binding.executable){throw 'Remaining Excel exact identity rejected'}
                $remainingProcess.Kill()
                if(-not $remainingProcess.WaitForExit(30000)){throw 'Remaining owned Excel survived cleanup'}
                $cleanupExitMode='FORCED_CONTAINED'
                $remainingProcess.Dispose()
            }
        }
    } catch { $cleanupFailures += $_.Exception.Message }
    try { if($null -ne $process){$process.Dispose()} } catch { }
    if($cleanupFailures.Count -gt 0){$terminalCode=25;$failurePhase='cleanup';if([string]::IsNullOrWhiteSpace($failureMessage)){$failureMessage='Owned Excel cleanup failed'}}
    if($terminalCode -eq 0){
        Write-Json (Join-Path $evidence 'ProductUi.json') ([ordered]@{schema_version=1;suite='ProductUi';status='PASS';mode='Green';run_id=$RunId;source_digest=$SourceDigest;snapshot_digest=$SnapshotDigest;artifact_sha256=$artifactSha;environment=[ordered]@{excel_bitness='64-bit';windows_scale_percent=200};cleanup=[ordered]@{status='PASS';exit_mode=$cleanupExitMode;failures=@()};run=[ordered]@{total=1;passed=1;failed=0;skipped=0}})
    } else {
        Write-Json (Join-Path $evidence 'ProductUi.json') ([ordered]@{schema_version=1;suite='ProductUi';status='INCOMPLETE';mode='Green';run_id=$RunId;source_digest=$SourceDigest;snapshot_digest=$SnapshotDigest;artifact_sha256=if(Test-Path -LiteralPath $artifact -PathType Leaf){Get-HexDigest $artifact}else{$null};failure_phase=$failurePhase;exit_code=$terminalCode;failure=$failureMessage;probe_stage=$probeStage;cleanup=[ordered]@{status=if($cleanupFailures.Count -eq 0){'PASS'}else{'FAIL'};exit_mode=$cleanupExitMode;failures=$cleanupFailures};run=[ordered]@{total=1;passed=0;failed=1;skipped=0}})
    }
    exit $terminalCode
}
