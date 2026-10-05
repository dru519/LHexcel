param([Parameter(Mandatory=$true)][string]$SetupExe,[Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [ValidatePattern('^v0\.01_r[0-9]+$')][string]$ExpectedVersion='v0.01_r101',
    [switch]$SeedPreferences,[switch]$VerifyPreferences,[switch]$KeepInstalled,[switch]$AlreadyInstalled)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
$source=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
. (Join-Path $source 'build/Excel-ProcessLifecycle.ps1')
if(Test-Path -LiteralPath $EvidenceRoot) { throw 'Existing evidence is preserved' }
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count) { throw 'Close Excel first' }
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$SetupExe=[IO.Path]::GetFullPath($SetupExe)
$result=[ordered]@{status='RUNNING';installer_sha256=(Get-FileHash -LiteralPath $SetupExe -Algorithm SHA256).Hash;
    install='NOT_RUN';autoload='NOT_RUN';bridge='NOT_RUN';uninstall='NOT_RUN';cleanup='NOT_RUN';failure=$null}
$app=$null;$owner=$null;$installed=$false;$addins=$null;$addin=$null
$previousProfile=$env:LHEXCEL_PROFILE_ROOT
if([string]::IsNullOrWhiteSpace($previousProfile)) {
    # Installed services require local runtime resources even when evidence is on SMB.
    $env:LHEXCEL_PROFILE_ROOT=Join-Path ([IO.Path]::GetTempPath()) ('nx-setup-profile-'+[Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $env:LHEXCEL_PROFILE_ROOT)
}
$result.isolated_profile=$env:LHEXCEL_PROFILE_ROOT
function Invoke-Setup([string]$Action) {
    $p=Start-Process -FilePath $SetupExe -ArgumentList @($Action,'--accept-current-user-changes') -PassThru -Wait -WindowStyle Hidden
    if($p.ExitCode -ne 0) { throw ('Installer failed: '+$Action+'; inspect LOCALAPPDATA/LHexcel/Setup/last-error.txt') }
}
try {
    if($AlreadyInstalled) {
        $receiptPath=Join-Path $env:LOCALAPPDATA ('LHexcel/NxHost/state/'+$ExpectedVersion.Split('_')[1]+'/native-install.xml')
        [xml]$receiptDocument=Get-Content -LiteralPath $receiptPath -Raw
        if($receiptDocument.installation.state -ne 'ACTIVE' -or $receiptDocument.installation.version -ne $ExpectedVersion){throw 'Expected active update receipt'}
        $result.install_origin='verified_update'
    } else { Invoke-Setup '--install' }
    $installed=$true;$result.install='PASS'
    $baseline=Get-ExcelProcessBaseline
    # Reuse the exact /x launch and PID-bound attachment used by the current build.
    # ROT-only lookup can miss a normally started Excel instance until focus changes.
    $interactive=Start-ExactInteractiveExcel $baseline 'r101 installed normal-start Excel'
    $app=$interactive.excel;$owner=$interactive.binding;$interactive=$null
    if(-not $owner.owned) { throw 'Normal-start Excel ownership mismatch' }
    $name=(-join @([char]0xB0B4,[char]0xC5D1,[char]0xC140))+' '+$ExpectedVersion+'.xlam'
    $macro="'"+$name+"'!"
    $version=[string]$app.Run($macro+'NxProductVersionText')
    if(-not $version.EndsWith($ExpectedVersion)) { throw 'Installed automatic XLAM version mismatch' }
    $result.autoload='PASS';$result.version=$version
    if([int]($ExpectedVersion -replace '^v0\.01_r','') -ge 96){
        if([bool]$app.Run($macro+'NxDistributionIsStandaloneXlam')){throw 'Installed package incorrectly permits self installation'}
        $menu=[string]$app.Run($macro+'NxRibbonManagementMenuXml')
        foreach($entry in @('NX-MGMT-INSTALL','NX-MGMT-REMOVE')){
            if($menu.Contains('tag="nx1|entry|'+$entry+'"')){throw 'Installed self-install menu visible'}
            if([bool]$app.Run($macro+'NxDistributionEntryAllowed','entry',$entry)){throw 'Installed self-install route allowed'}
        }
        $result.install_menu_profile='PASS'
        $result.pre_menu_availability=[string]$app.Run($macro+'NxRouteAvailabilitySummary','feature:NX-FILE-MANNER-SAVE')
        $result.pre_menu_ready=[bool]$app.Ready
        $result.pre_menu_workbook_count=[int]$app.Workbooks.Count
        if($result.pre_menu_workbook_count -eq 0){
            $fixtureBooks=$app.Workbooks
            $fixtureBook=$fixtureBooks.Add()
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($fixtureBook)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($fixtureBooks)
            $fixtureBook=$null;$fixtureBooks=$null
        }
        $result.document_availability=[string]$app.Run($macro+'NxRouteAvailabilitySummary','feature:NX-FILE-MANNER-SAVE')
        Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
        $app.Visible=$true
        $uiRoot=[Windows.Automation.AutomationElement]::FromHandle([IntPtr]$app.Hwnd)
        function Find-VisibleNamed([string]$name,[switch]$Expandable){
            $condition=New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::NameProperty,$name)
            $until=[DateTime]::UtcNow.AddSeconds(5)
            do{
                $items=$uiRoot.FindAll([Windows.Automation.TreeScope]::Descendants,$condition)
                foreach($item in $items){
                    $pattern=$null
                    if(-not $item.Current.IsOffscreen -and
                        (-not $Expandable -or $item.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern,[ref]$pattern))){return $item}
                }
                Start-Sleep -Milliseconds 100
            }while([DateTime]::UtcNow -lt $until)
            $visible=@($uiRoot.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition) | Where-Object {-not $_.Current.IsOffscreen} | ForEach-Object {
                [ordered]@{name=$_.Current.Name;type=$_.Current.ControlType.ProgrammaticName;enabled=$_.Current.IsEnabled;id=$_.Current.AutomationId}
            })
            $visible | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $EvidenceRoot 'ribbon-visible.json') -Encoding UTF8
            throw ('Visible Ribbon item missing: '+$name)
        }
        $homeTab=Find-VisibleNamed ([string][char]0xD648)
        $homeTab.GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
        Start-Sleep -Milliseconds 300
        $all=Find-VisibleNamed (-join @([char]0xC804,[char]0xCCB4,[char]0xAE30,[char]0xB2A5))
        if(-not $all.Current.IsEnabled){throw 'Home all-functions disabled'}
        try{$all.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Expand()}
        catch{$all.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()}
        Start-Sleep -Milliseconds 300
        $save=Find-VisibleNamed (-join @([char]0xC800,[char]0xC7A5)) -Expandable
        try{$save.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Expand()}
        catch{$save.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()}
        Start-Sleep -Milliseconds 300
        $manner=Find-VisibleNamed (-join @([char]0xB9E4,[char]0xB108,' ',[char]0xC800,[char]0xC7A5))
        if(-not $manner.Current.IsEnabled){
            $result.disabled_menu_availability=[string]$app.Run($macro+'NxRouteAvailabilitySummary','feature:NX-FILE-MANNER-SAVE')
            throw 'Home dynamic feature disabled after normal startup'
        }
        $result.home_dynamic_menu='PASS'
        try{$save.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Collapse()}catch{}
        try{$all.GetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern).Collapse()}catch{}
    }
    $connected=[bool]$app.Run($macro+'NxHostBridgeCompatible')
    if(-not $connected) { throw 'Installed VBA/DLL handshake failed' }
    $result.bridge='PASS'
    $addins=$app.COMAddIns;$addin=$addins.Item('LH.NxHost.Connect')
    if(-not [bool]$addin.Connect) { throw 'Excel COM add-in is not connected' }
    $result.com_addin_connected=$true
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addin)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addins)
    $addin=$null;$addins=$null
    if([int]($ExpectedVersion -replace '^v0\.01_r','') -ge 64) { foreach($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
        $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
        $key=$null
        try {
            $key=$base.OpenSubKey('Software\Classes\LH.NxHost.HwpxExportService\CLSID')
            if($null -eq $key -or [string]$key.GetValue('') -ne '{1434C649-18AA-4435-B0D2-2BD81B548A01}') { throw 'Installed HWPX service registration missing' }
        } finally { if($null -ne $key){$key.Dispose()};$base.Dispose() }
    }
    $result.hwpx_service_registered='PASS'
    } else { $result.hwpx_service_registered='NOT_APPLICABLE_TO_OLD_RELEASE' }
    if([int]($ExpectedVersion -replace '^v0\.01_r','') -ge 65) {
        $services=@{
            'LH.NxHost.PicturePreviewService'='{C812F023-E912-49AA-98B7-75DC86501902}'
            'LH.NxHost.WorkbookCompareService'='{C812F023-E912-49AA-98B7-75DC86501904}'
        }
        if([int]($ExpectedVersion -replace '^v0\.01_r','') -ge 67) {
            $services['LH.NxHost.WorkbookCompareResultsService']='{C812F023-E912-49AA-98B7-75DC86501906}'
            $resultsService=$app.Run($macro+'NxHostCreateWorkbookCompareResults')
            if($null -eq $resultsService){throw 'Installed compare results handshake failed'}
            try { $resultsService.Close();$result.compare_results_bridge='PASS' }
            finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($resultsService);$resultsService=$null }
        }
        foreach($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
            $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
            try { foreach($entry in $services.GetEnumerator()) {
                $key=$base.OpenSubKey('Software\Classes\'+$entry.Key+'\CLSID')
                try { if($null -eq $key -or [string]$key.GetValue('') -ne $entry.Value){throw ('Installed service registration missing: '+$entry.Key)} }
                finally { if($null -ne $key){$key.Dispose()} }
            } } finally { $base.Dispose() }
        }
        $result.r101_service_registration='PASS'
    } else { $result.r101_service_registration='NOT_APPLICABLE_TO_OLD_RELEASE' }
    if([int]($ExpectedVersion -replace '^v0\.01_r','') -ge 75) {
        . (Join-Path $PSScriptRoot 'Test-R75InstalledIntegrity.ps1')
        $result.runtime_integrity=Test-R75InstalledIntegrity $app $macro $EvidenceRoot
    }
    if($SeedPreferences -or $VerifyPreferences) {
        if([string]::IsNullOrWhiteSpace($env:LHEXCEL_PROFILE_ROOT)){throw 'Preference test requires an isolated profile'}
        $route='feature:NX-DATA-NORMALIZE'
        if($SeedPreferences) {
            [void]$app.Run($macro+'NxFavoritesAdd',$route)
            [void]$app.Run($macro+'NxShortcutsAssign',$route,'Ctrl+Alt+Shift','Q')
            [void]$app.Run($macro+'NxFocusResetSettings')
        }
        if(-not [bool]$app.Run($macro+'NxFavoritesContains',$route)){throw 'Favorite was not preserved'}
        if([string]$app.Run($macro+'NxShortcutsBindingDisplayForRoute',$route) -ne 'Ctrl+Alt+Shift+Q'){throw 'Shortcut was not preserved'}
        $result.preferences='PASS'
    }
    $startup=Join-Path $env:APPDATA 'Microsoft/Excel/XLSTART'
    $activeXlam=@(Get-ChildItem -LiteralPath $startup -Filter '*.xlam' -ErrorAction SilentlyContinue | Where-Object {$_.Name -like '*v0.01_r*.xlam'})
    if($activeXlam.Count -ne 1 -or $activeXlam[0].Name -ne $name){throw 'Duplicate or unexpected automatic XLAM'}
    $result.single_autoload='PASS'
    $result.status='PASS'
} catch {
    $result.status='FAIL';$result.failure=$_.Exception.Message
} finally {
    foreach($value in @($addin,$addins)) { if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)} }
    $addin=$null;$addins=$null
    if($null -ne $app -and $null -ne $owner -and $owner.owned) {
        # Only the newly started empty test workbook may be closed without saving.
        $books=$app.Workbooks
        $safe=$true
        for($i=1;$i -le $books.Count;$i++) {
            $book=$books.Item($i)
            if(-not $book.IsAddin -and ([string]$book.Path -ne '' -or -not [bool]$book.Saved)) { $safe=$false }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($books)
        if($safe) { $app.Quit() } else { $result.failure='Unexpected workbook preserved';$result.status='HOLD' }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app);$app=$null
        [GC]::Collect();[GC]::WaitForPendingFinalizers()
        if($safe) { $exit=Stop-ExactProcessAfterGrace -Process $owner.process -Label 'r101 installed Excel' -GraceMs 30000 -Detailed; $result.cleanup=$exit.exit_mode }
    }
    $exitDeadline=[DateTime]::UtcNow.AddSeconds(10)
    while(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0 -and [DateTime]::UtcNow -lt $exitDeadline) { Start-Sleep -Milliseconds 200 }
    if($installed -and -not $KeepInstalled -and (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0)) {
        try { Invoke-Setup '--uninstall';$result.uninstall='PASS' }
        catch { $result.uninstall='FAIL';$result.status='FAIL';$result.failure=$_.Exception.Message }
    }
    if($KeepInstalled -and $installed -and $result.cleanup -eq 'NATURAL'){$result.uninstall='RETAINED_FOR_UPGRADE_FIXTURE'}
    if($result.uninstall -ne 'PASS' -and $result.uninstall -ne 'RETAINED_FOR_UPGRADE_FIXTURE') { $result.status='HOLD' }
    if($result.cleanup -ne 'NATURAL'){$result.status='FAIL'}
    $env:LHEXCEL_PROFILE_ROOT=$previousProfile
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'installed-acceptance.json'),($result|ConvertTo-Json -Depth 8),(New-Object Text.UTF8Encoding($false)))
}
$result|ConvertTo-Json -Depth 8
if($result.status -ne 'PASS') { exit 1 }
