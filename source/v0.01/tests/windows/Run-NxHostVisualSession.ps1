param(
    [Parameter(Mandatory=$true)][string]$ProductXlam,
    [Parameter(Mandatory=$true)][string]$NxHostRoot,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [switch]$Automated,
    [switch]$ExpectFallback,
    [switch]$Enhancements,
    [switch]$DocumentOnly,
    [switch]$ServicesOnly,
    [switch]$SafeMode,
    [switch]$IsolateOtherAddins,
    [string]$PreparedProbeXlam='',
    [string]$PreparedProbeSha256='',
    [switch]$PictureFixture
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
$sourceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$EvidenceRoot = [IO.Path]::GetFullPath($EvidenceRoot)
$ProductXlam = [IO.Path]::GetFullPath($ProductXlam)
if (Test-Path -LiteralPath $EvidenceRoot) { throw 'Fresh evidence root required' }
if (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Pre-existing Excel must be preserved' }
$runtimeRoot = Join-Path $env:LOCALAPPDATA 'LHexcel/NxHost/r101'
if (Test-Path -LiteralPath $runtimeRoot) { throw 'Existing runtime must be preserved' }
$stateRoot = Join-Path $env:LOCALAPPDATA 'LHexcel/NxHost/state/r101'
$statePath = Join-Path $stateRoot 'registry-state.json'
if (Test-Path -LiteralPath $statePath) {
    if ((Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json).status -eq 'ACTIVE') { throw 'Active registration must be preserved' }
}
. (Join-Path $sourceRoot 'build/Excel-ProcessLifecycle.ps1')
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
if (Test-Path -LiteralPath $stateRoot) { Copy-Item -LiteralPath $stateRoot -Destination (Join-Path $EvidenceRoot 'prior-registration-state') -Recurse }
$artifact = Join-Path $EvidenceRoot 'Product.xlam'
Copy-Item -LiteralPath $ProductXlam -Destination $artifact
if($PreparedProbeXlam) {
    if(-not $ServicesOnly -or -not $Enhancements -or $PreparedProbeSha256 -notmatch '^[a-fA-F0-9]{64}$'){throw 'Prepared probe is restricted to verified service diagnostics'}
    if((Get-FileHash -LiteralPath $PreparedProbeXlam).Hash -ne $PreparedProbeSha256){throw 'Prepared probe hash mismatch'}
    Copy-Item -LiteralPath $PreparedProbeXlam -Destination $artifact -Force
}
$x86 = [IO.Path]::GetFullPath((Join-Path $NxHostRoot 'NxHost32.dll'))
$x64 = [IO.Path]::GetFullPath((Join-Path $NxHostRoot 'NxHost64.dll'))
$productHash = (Get-FileHash -LiteralPath $ProductXlam).Hash
$priorProfile = $env:LHEXCEL_PROFILE_ROOT
$priorDiagnostics = $env:LHEXCEL_NXHOST_DIAGNOSTICS
$env:LHEXCEL_NXHOST_DIAGNOSTICS = '1'
$env:LHEXCEL_PROFILE_ROOT = Join-Path $EvidenceRoot 'profile'
[void](New-Item -ItemType Directory -Path $env:LHEXCEL_PROFILE_ROOT)
$excel=$null; $binding=$null; $books=$null; $book=$null; $product=$null; $sheet=$null; $addins=$null; $hostAddin=$null
$registered=$false; $failure=$null; $cleanup=$null; $stopped=$false; $registryBefore=''; $registryAfter=''
$actions = New-Object 'Collections.Generic.List[object]'
$isolatedAddins = New-Object 'Collections.Generic.List[object]'
function Release-Com([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
    }
}
function Release-ComObject([object]$value) { Release-Com $value }
function Write-Receipt([string]$name, [object]$value) {
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot $name), (($value | ConvertTo-Json -Depth 12)+"`n"), (New-Object Text.UTF8Encoding($false)))
}
function Select-FixtureCell([string]$address) {
    $range=$null
    try {
        # Other source books may have been activated by result navigation.
        [void]$book.Activate(); [void]$sheet.Activate()
        $range=$sheet.Range($address); [void]$range.Select()
    } finally { Release-Com $range }
}
try {
    & (Join-Path $sourceRoot 'build/Register-NxHost.ps1') -Action Register -OptIn -X86Dll $x86 -X64Dll $x64
    $registered=$true
    $state=Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
    $registryBefore=[string]$state.snapshot_sha256
    if($Automated){ & (Join-Path $PSScriptRoot 'Test-NxNavigatorComContract.ps1') -EvidenceRoot $EvidenceRoot }
    if($IsolateOtherAddins) {
        # Isolate at startup: disconnecting a machine-installed COM add-in after
        # loading can require elevation and leaves its native modules loaded.
        foreach($diagnosticProgId in @('Detong.KTEHelper','DeTong.KTELoader')) {
        $keyPath='Software\Microsoft\Office\Excel\Addins\'+$diagnosticProgId
        $key=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($keyPath,$true)
        try {
            if($null -eq $key -or @($key.GetValueNames()) -notcontains 'LoadBehavior') {
                throw 'Existing per-user Kutools LoadBehavior required for diagnostic isolation'
            }
            $isolatedAddins.Add([pscustomobject]@{prog_id=$diagnosticProgId;key=$keyPath;had_value=$true;value=$key.GetValue('LoadBehavior');kind=[string]$key.GetValueKind('LoadBehavior')})
            Write-Receipt 'other-addins-before.json' $isolatedAddins.ToArray()
            $key.SetValue('LoadBehavior',0,[Microsoft.Win32.RegistryValueKind]::DWord)
        } finally {if($null -ne $key){$key.Dispose()}}
        }
    }
    $baseline=Get-ExcelProcessBaseline
    $startupFixture=''
    if($ServicesOnly -or $IsolateOtherAddins -or $SafeMode) {
        $startupFixture=Join-Path $EvidenceRoot 'startup-fixture.xlsx'
        Copy-Item -LiteralPath (Join-Path $sourceRoot 'tests/fixtures/r60-document-navigator/simple.xlsx') -Destination $startupFixture
    }
    $interactive=Start-ExactInteractiveExcel $baseline 'DLL visual verification' -TimeoutMs 120000 -SafeMode:$SafeMode -StartupWorkbook $startupFixture
    $excel=$interactive.excel; $binding=$interactive.binding; $interactive=$null
    if (-not $binding.owned) { throw 'Excel ownership rejected' }
    $excel.Visible=$true; $excel.DisplayAlerts=$false; $excel.EnableEvents=$true
    if($startupFixture) {
        $startupBook=$excel.Workbooks.Item('startup-fixture.xlsx')
        try {
            if([IO.Path]::GetFullPath($startupBook.FullName) -ne [IO.Path]::GetFullPath($startupFixture)){throw 'Startup fixture identity mismatch'}
            $startupBook.Close($false)
        } finally {Release-Com $startupBook}
    }
    if($IsolateOtherAddins) {
        $inventory=$excel.COMAddIns
        try {
            for($index=1;$index -le $inventory.Count;$index++) {
                $entry=$inventory.Item($index)
                try {
                    $progId=[string]$entry.ProgId
                    if($progId -like 'LH.NxHost.*' -or -not $entry.Connect){continue}
                    throw ('Unexpected connected add-in in isolated diagnostic: '+$progId)
                } finally {Release-Com $entry}
            }
        } finally {Release-Com $inventory}
    }
    if (-not $SafeMode) {
        $addins=$excel.COMAddIns; $hostAddin=$addins.Item('LH.NxHost.Connect'); $hostAddin.Connect=$true
        if (-not $hostAddin.Connect) { throw 'COM add-in did not connect' }
    } elseif (-not $ServicesOnly) { throw 'Safe mode is scoped to isolated service diagnostics' }
    $books=$excel.Workbooks; $book=$books.Add(); $product=$books.Open($artifact,$false,(-not $Enhancements))
    if ($Enhancements -and -not $PreparedProbeXlam) {
        $cancelClass=$product.VBProject.VBComponents.Add(2);$cancelClass.Name='CNxR65CancelCompare'
        $cancelCode=[IO.File]::ReadAllText((Join-Path $PSScriptRoot '../vba/product/CNxR65CancelCompare.cls'),[Text.Encoding]::UTF8)
        $cancelCode=$cancelCode.Substring($cancelCode.IndexOf('Option Explicit'))
        $cancelClass.CodeModule.AddFromString($cancelCode.Replace("`r`n","`n").Replace("`n","`r`n"))
        $component=$product.VBProject.VBComponents.Add(1); $component.Name='T_R65Enhancement'
        $code=[IO.File]::ReadAllText((Join-Path $PSScriptRoot '../vba/product/T_R65Enhancement.bas'),[Text.Encoding]::UTF8)
        $code=[regex]::Replace($code,'(?m)^Attribute [^\r\n]*\r?\n','')
        $component.CodeModule.AddFromString($code.Replace("`r`n","`n").Replace("`n","`r`n"))
        $component.Activate()
        $control=Find-VbeCompileControl $excel.VBE.CommandBars
        $watcher=Start-ExactOwnedCompileDialogWatcher -ExcelPid ([int]$binding.pid) -ExcelHwnd ([Int64]$binding.hwnd) -VbeHwnd ([Int64]$excel.VBE.MainWindow.HWnd)
        if ($control.Enabled) { $control.Execute() }
        if ($null -eq (Wait-Job $watcher -Timeout 30)) { throw 'Enhancement compile watcher timeout' }
        $observed=@(Receive-Job $watcher -Wait -AutoRemoveJob)
        if ($observed.Count -ne 1 -or $observed[0].status -ne 'NO_DIALOG' -or $control.Enabled) { throw 'Enhancement probe compile failed' }
        $product.Save()
        Release-Com $control; Release-Com $component; Release-Com $cancelClass
        $product.Close($false); Release-Com $product
        $product=$books.Open($artifact,$false,$false)
    }
    $sheets=$book.Worksheets
    try { $sheet=$sheets.Item(1) } finally { Release-Com $sheets }
    $sheet.Name='NxVisualFixture'; [void]$book.Activate(); [void]$sheet.Activate()
    # Keep the owned fixture distinguishable from Excel's disposable blank book.
    $fixtureMarker=$sheet.Range('B2')
    try {$fixtureMarker.Value2='NxVisualFixture'} finally {Release-Com $fixtureMarker}
    Select-FixtureCell 'B2'
    if($PictureFixture) {
        Add-Type -AssemblyName System.Drawing
        $shapes=$book.Worksheets.Item(1).Shapes
        try {
            $names=@()
            foreach($i in @(1,2)) {
                $image=New-Object Drawing.Bitmap 400,100
                $graphics=[Drawing.Graphics]::FromImage($image)
                try { $graphics.Clear($(if($i -eq 1){[Drawing.Color]::Orange}else{[Drawing.Color]::SteelBlue})) }
                finally { $graphics.Dispose() }
                $imagePath=Join-Path $env:LHEXCEL_PROFILE_ROOT ('picture-'+$i+'.png')
                try { $image.Save($imagePath,[Drawing.Imaging.ImageFormat]::Png) } finally { $image.Dispose() }
                $shape=$shapes.AddPicture($imagePath,0,-1,30,(30+($i-1)*70),120,30)
                try { $shape.Name='R67FixturePicture'+$i;$names+=[string]$shape.Name } finally { Release-Com $shape }
            }
            $range=$shapes.Range($names)
            try { $range.Select() } finally { Release-Com $range }
            Write-Receipt 'picture-fixture.json' @{selected_shapes=2;shape_count=$shapes.Count;profile=$env:LHEXCEL_PROFILE_ROOT}
        } finally { Release-Com $shapes }
    }
    $prefix="'"+$product.Name+"'!"
    [void]$excel.Run(($prefix+'NxFocusDisable'))
    if (-not $ServicesOnly -and -not [bool]$excel.Run(($prefix+'NxHostShowNavigator'))) { throw 'Navigator show failed' }
    Write-Receipt 'ready.json' @{status='READY';owned_excel_pid=$binding.pid;hwnd=$excel.Hwnd;product_sha256=$productHash;dll_sha256=(Get-FileHash -LiteralPath $x64).Hash;focus='OFF';ui_verification=if($Automated){'Native UI Automation and fixture-scoped list key message'}else{'External observation required'}}
    Write-Output ('READY|'+$EvidenceRoot)
    if ($Enhancements -or $DocumentOnly) {
        & (Join-Path $PSScriptRoot 'Test-R65EnhancementUi.ps1') -Excel $excel -Sheet $sheet -Prefix $prefix -EvidenceRoot $EvidenceRoot -DocumentOnly:$DocumentOnly -ServicesOnly:$ServicesOnly
        $stopped=$true
    } elseif ($Automated) {
        & (Join-Path $PSScriptRoot 'Test-NxNavigatorUi.ps1') -Excel $excel -Sheet $sheet -Prefix $prefix -EvidenceRoot $EvidenceRoot -ExpectFallback:$ExpectFallback
        $stopped=$true
    }
    # Closed fixture-only commands; no arbitrary code execution and no user workbook access.
    $commands=@('01-protect','02-unprotect','03-hide','04-show','05-stop')
    $next=0; $deadline=[DateTime]::UtcNow.AddMinutes(15)
    while ([DateTime]::UtcNow -lt $deadline -and -not $stopped) {
        $command=$commands[$next]
        if (Test-Path -LiteralPath (Join-Path $EvidenceRoot ($command+'.request'))) {
            switch ($command) {
                '01-protect' { $sheet.Protect(); Select-FixtureCell 'C3' }
                '02-unprotect' { $sheet.Unprotect(); Select-FixtureCell 'D4' }
                '03-hide' { [void]$excel.Run(($prefix+'NxHostHideNavigator')) }
                '04-show' { [void]$excel.Run(($prefix+'NxHostShowNavigator')) }
                '05-stop' {
                    if($PictureFixture) {
                        $shapes=$sheet.Shapes;$dimensions=@()
                        try { for($i=1;$i -le $shapes.Count;$i++) {
                            $shape=$shapes.Item($i)
                            try {$dimensions+=@{name=$shape.Name;width=$shape.Width;height=$shape.Height}}finally{Release-Com $shape}
                        }}finally{Release-Com $shapes}
                        Write-Receipt 'picture-fixture-final.json' @{shapes=$dimensions}
                    }
                    $stopped=$true
                }
            }
            $entry=@{action=$command;utc=[DateTime]::UtcNow.ToString('o');sheet_protected=[bool]$sheet.ProtectContents}
            $actions.Add($entry); Write-Receipt ($command+'.done.json') $entry
            $next++
        }
        Start-Sleep -Milliseconds 100
    }
    if (-not $stopped) { throw 'Visual session timed out' }
} catch { $failure=$_.Exception.ToString() }
finally {
    if ($null -ne $product) { try { [void]$excel.Run(($prefix+'NxHostHideDocumentNavigator')); [void]$excel.Run(($prefix+'NxHostHideNavigator')); [void]$excel.Run(($prefix+'NxFocusDisable')) } catch {} }
    if ($null -ne $book) { try { $book.Close($false) } catch {} }
    if ($null -ne $product) { try { $product.Close($false) } catch {} }
    if ($null -ne $hostAddin) { try { $hostAddin.Connect=$false } catch {} }
    Release-Com $sheet; Release-Com $book; Release-Com $product; Release-Com $books; Release-Com $hostAddin; Release-Com $addins
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-Com $excel }
    $sheet=$null; $book=$null; $product=$null; $books=$null; $hostAddin=$null; $addins=$null; $excel=$null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($null -ne $binding -and $null -ne $binding.process) { $cleanup=Stop-ExactProcessAfterGrace $binding.process 'DLL visual verification' 15000 -Detailed; $binding.process.Dispose() }
    if ($registered) {
        try {
            & (Join-Path $sourceRoot 'build/Register-NxHost.ps1') -Action Unregister -X86Dll $x86 -X64Dll $x64
            $state=Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
            $registryAfter=[string]$state.restored_snapshot_sha256
        } catch { $failure=([string]$failure+'; cleanup: '+$_.Exception.Message) }
    }
    foreach($saved in $isolatedAddins) {
        # Restore only the value affected by this diagnostic's Connect change.
        $key=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($saved.key,$true)
        try {
            if($null -eq $key -and $saved.had_value){throw ('Add-in registration disappeared: '+$saved.prog_id)}
            if($null -ne $key) {
                if($saved.had_value){$key.SetValue('LoadBehavior',$saved.value,([Microsoft.Win32.RegistryValueKind]$saved.kind))}
                else {$key.DeleteValue('LoadBehavior',$false)}
                $hasValue=@($key.GetValueNames()) -contains 'LoadBehavior'
                if($hasValue -ne $saved.had_value -or ($hasValue -and ($key.GetValue('LoadBehavior') -ne $saved.value -or [string]$key.GetValueKind('LoadBehavior') -ne $saved.kind))){throw ('Add-in restore mismatch: '+$saved.prog_id)}
            }
        } catch {$failure=([string]$failure+'; add-in restore: '+$_.Exception.Message)}
        finally {if($null -ne $key){$key.Dispose()}}
    }
    if($IsolateOtherAddins){Write-Receipt 'other-addins-restored.json' $isolatedAddins.ToArray()}
    $env:LHEXCEL_PROFILE_ROOT=$priorProfile
    $env:LHEXCEL_NXHOST_DIAGNOSTICS=$priorDiagnostics
}
$exitedEntries=New-Object 'Collections.Generic.List[int]'
$remaining=@(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object {
    $observed=$_
    try {
        # UIA can retain a handle to a terminated process briefly. Count live processes,
        # and record terminated enumeration entries separately instead of declaring a leak.
        if($observed.HasExited){$exitedEntries.Add($observed.Id)}else{$observed.Id}
    } catch { $observed.Id } finally { $observed.Dispose() }
})
$afterHash=(Get-FileHash -LiteralPath $ProductXlam).Hash
$ok=$null -eq $failure -and $stopped -and $remaining.Count -eq 0 -and $null -ne $cleanup -and $cleanup.exit_mode -eq 'NATURAL' -and $productHash -eq $afterHash -and $registryBefore -ne '' -and $registryBefore -eq $registryAfter
Write-Receipt 'session.json' @{status=if($ok){'PASS'}else{'FAIL'};boundary=if($Automated){'Native Excel CTP UI Automation; exact cases in navigator-ui.json; not physical-input or multi-monitor certification'}else{'fixture and lifecycle only; UI requires separate observations'};actions=$actions.ToArray();failure=$failure;cleanup=$cleanup;remaining_excel=$remaining;exited_process_entries=$exitedEntries.ToArray();product_sha256=$productHash;product_sha256_after=$afterHash;registry_before=$registryBefore;registry_after=$registryAfter}
if (-not $ok) { throw ('Visual session failed: '+$failure) }
Write-Output 'PASS|VisualSession|fixture_and_cleanup'
