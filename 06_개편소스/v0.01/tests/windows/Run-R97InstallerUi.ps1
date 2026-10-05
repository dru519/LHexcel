param([Parameter(Mandatory=$true)][string]$SetupExe,[Parameter(Mandatory=$true)][string]$EvidenceRoot,
    [ValidatePattern('^v0\.01_r[0-9]+$')][string]$ExpectedVersion='v0.01_r97')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot '../../build/Excel-ProcessLifecycle.ps1')
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
Add-Type -AssemblyName System.Drawing
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Fresh evidence required'}
if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Existing Excel preserved'}
[void](New-Item -ItemType Directory -Path $EvidenceRoot)
$result=[ordered]@{status='RUNNING';update_ui='NOT_RUN';launch='NOT_RUN';control_panel_remove='NOT_RUN';exe_remove='NOT_RUN';post_remove_excel='NOT_RUN'}
$setup=$null;$excel=$null;$binding=$null
$previousProfile=$env:LHEXCEL_PROFILE_ROOT
$env:LHEXCEL_PROFILE_ROOT=Join-Path $env:TEMP ('nx-r97-ui-'+[Guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $env:LHEXCEL_PROFILE_ROOT)
function Find-Button([int]$ProcessId,[string]$Name,[string]$Id='', [int]$Seconds=30,[string]$Title=''){
    $condition=[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId)
    $deadline=[DateTime]::UtcNow.AddSeconds($Seconds)
    do {
        $roots=[Windows.Automation.AutomationElement]::RootElement.FindAll([Windows.Automation.TreeScope]::Children,$condition)
        foreach($root in $roots){
            if($Title -ne '' -and $root.Current.Name -ne $Title){
                $dialog=$root.FindFirst([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.AndCondition]::new(
                    [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,$Title),
                    [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::Window)))
                if($null -eq $dialog){continue}
                $root=$dialog
            }
            $buttons=$root.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::Button))
            foreach($button in $buttons){if(-not $button.Current.IsOffscreen -and $button.Current.IsEnabled -and (($Id -ne '' -and $button.Current.AutomationId -eq $Id) -or ($Id -eq '' -and $button.Current.Name -eq $Name))){return $button}}
        }
        Start-Sleep -Milliseconds 100
    }while([DateTime]::UtcNow -lt $deadline)
    throw ('Button missing: '+$Name+'/'+$Id)
}
function Click([object]$Button){$Button.GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern).Invoke()}
function Capture([int]$ProcessId,[string]$Name){
    $root=[Windows.Automation.AutomationElement]::RootElement.FindFirst([Windows.Automation.TreeScope]::Children,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId))
    $rect=$root.Current.BoundingRectangle
    $bitmap=New-Object Drawing.Bitmap([int]$rect.Width,[int]$rect.Height)
    $graphics=[Drawing.Graphics]::FromImage($bitmap)
    try{$graphics.CopyFromScreen([int]$rect.X,[int]$rect.Y,0,0,$bitmap.Size);$bitmap.Save((Join-Path $EvidenceRoot ($Name+'.png')))}finally{$graphics.Dispose();$bitmap.Dispose()}
}
function Next([int]$ProcessId){Click (Find-Button $ProcessId ((-join @([char]0xB2E4,[char]0xC74C))+' >'))}
function Choose([int]$ProcessId,[string]$Action){
    $root=[Windows.Automation.AutomationElement]::RootElement.FindFirst([Windows.Automation.TreeScope]::Children,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,$ProcessId))
    $choices=$root.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::RadioButton))
    $selected=@($choices | Where-Object {$_.Current.Name.StartsWith($Action) -and $_.Current.IsEnabled -and -not $_.Current.IsOffscreen})
    if($selected.Count -ne 1){throw ('Choice missing or ambiguous: '+$Action)}
    $selected[0].GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
}
function Confirm-Removal([int]$ProcessId){
    $brand=-join @([char]0xB0B4,[char]0xC5D1,[char]0xC140)
    $install=-join @([char]0xC124,[char]0xCE58)
    $remove=-join @([char]0xC81C,[char]0xAC70)
    $ok=-join @([char]0xD655,[char]0xC778)
    Click (Find-Button $ProcessId $remove)
    Click (Find-Button $ProcessId (-join @([char]0xB9C8,[char]0xCE68)) -Seconds 60)
}
function Close-OwnedExcel {
    if($null -eq $script:excel){return}
    $books=$script:excel.Workbooks
    try {
        for($i=1;$i -le $books.Count;$i++){
            $book=$books.Item($i)
            try{if(-not $book.IsAddin -and ($book.Path -ne '' -or -not $book.Saved)){throw 'Unexpected workbook preserved'}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}
        }
    }finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($books)}
    $script:excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:excel);$script:excel=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    $closed=Stop-ExactProcessAfterGrace -Process $script:binding.process -Label 'r97 installed UI Excel' -GraceMs 15000 -Detailed
    if($closed.exit_mode -ne 'NATURAL'){throw 'Excel did not close naturally'}
    $script:binding=$null
}
function Assert-Removed {
    $name=(-join @([char]0xB0B4,[char]0xC5D1,[char]0xC140))+' '+$ExpectedVersion+'.xlam'
    if(Test-Path -LiteralPath (Join-Path $env:APPDATA ('Microsoft/Excel/XLSTART/'+$name))){throw 'Startup XLAM remains'}
    foreach($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)){
        $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,$view)
        try{foreach($path in @('Software\Microsoft\Windows\CurrentVersion\Uninstall\LHExcel','Software\Microsoft\Office\Excel\Addins\LH.NxHost.Connect','Software\Classes\LH.NxHost.Connect')){
            $key=$base.OpenSubKey($path);if($null -ne $key){$key.Dispose();throw ('Registration remains: '+$path)}
        }}finally{$base.Dispose()}
    }
}
try{
    $baseline=Get-ExcelProcessBaseline
    $setup=Start-Process -FilePath $SetupExe -PassThru -WindowStyle Normal
    $historyTitle=-join @([char]0xC774,[char]0xC804,' ',[char]0xC124,[char]0xCE58,' ',[char]0xAE30,[char]0xB85D,' ',[char]0xC815,[char]0xB9AC)
    Click (Find-Button $setup.Id $historyTitle)
    Click (Find-Button $setup.Id (-join @([char]0xC804,[char]0xCCB4,' ',[char]0xC120,[char]0xD0DD)))
    Click (Find-Button $setup.Id (-join @([char]0xC804,[char]0xCCB4,' ',[char]0xD574,[char]0xC81C)))
    $closeLabel=-join @([char]0xB2EB,[char]0xAE30)
    Click (Find-Button $setup.Id $closeLabel -Title $historyTitle)
    $result.history_dialog='PASS'
    Capture $setup.Id 'welcome'
    Next $setup.Id
    Choose $setup.Id (-join @([char]0xC5C5,[char]0xB370,[char]0xC774,[char]0xD2B8))
    Next $setup.Id
    # Review precedes mutation; Finish honors the visible Excel launch checkbox.
    Capture $setup.Id 'update-review'
    Click (Find-Button $setup.Id (-join @([char]0xC5C5,[char]0xB370,[char]0xC774,[char]0xD2B8)))
    Click (Find-Button $setup.Id (-join @([char]0xB9C8,[char]0xCE68)) -Seconds 60)
    if(-not $setup.WaitForExit(10000)){throw 'Setup did not finish after launch'}
    $result.update_ui='PASS'
    $deadline=[DateTime]::UtcNow.AddSeconds(30)
    do{
        try{$excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');$binding=Get-ExactExcelProcessOwnership $excel $baseline 'r97 launch';if($binding.owned){break}}catch{}
        if($null -ne $excel){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel);$excel=$null}
        Start-Sleep -Milliseconds 100
    }while([DateTime]::UtcNow -lt $deadline)
    if($null -eq $excel -or -not $binding.owned){throw 'Prompt did not start owned Excel'}
    $name=(-join @([char]0xB0B4,[char]0xC5D1,[char]0xC140))+' '+$ExpectedVersion+'.xlam'
    $version=''
    $deadline=[DateTime]::UtcNow.AddSeconds(30)
    do {
        try{if($excel.Ready){$version=[string]$excel.Run(("'"+$name+"'!NxProductVersionText"));break}}catch{}
        Start-Sleep -Milliseconds 100
    }while([DateTime]::UtcNow -lt $deadline)
    if(-not $version.EndsWith($ExpectedVersion)){throw 'Launched wrong add-in'}
    if(-not [bool]$excel.Run(("'"+$name+"'!NxHostBridgeCompatible"))){throw 'Installed DLL handshake failed'}
    $result.bridge='PASS'
    $result.launch='PASS';Close-OwnedExcel
    # A status check returns to maintenance instead of terminating the installer.
    $setup=Start-Process -FilePath $SetupExe -PassThru -WindowStyle Normal
    Next $setup.Id
    Choose $setup.Id (-join @([char]0xC124,[char]0xCE58,' ',[char]0xC0C1,[char]0xD0DC,' ',[char]0xD655,[char]0xC778))
    Next $setup.Id
    Click (Find-Button $setup.Id (-join @([char]0xD655,[char]0xC778,' ',[char]0xC2DC,[char]0xC791)))
    Click (Find-Button $setup.Id (-join @([char]0xC124,[char]0xCE58,' ',[char]0xAD00,[char]0xB9AC,[char]0xB85C,' ',[char]0xB3CC,[char]0xC544,[char]0xAC00,[char]0xAE30)) -Seconds 60)
    $result.check_returns_to_management='PASS'
    Choose $setup.Id (-join @([char]0xBCF5,[char]0xAD6C))
    Next $setup.Id
    Click (Find-Button $setup.Id (-join @([char]0xBCF5,[char]0xAD6C)))
    Click (Find-Button $setup.Id (-join @([char]0xB9C8,[char]0xCE68)) -Seconds 60)
    if(-not $setup.WaitForExit(15000)){throw 'Repair did not finish'}
    $result.repair_ui='PASS'
    $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::Registry64)
    $key=$base.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Uninstall\LHExcel')
    try{$command=[string]$key.GetValue('UninstallString')}finally{$key.Dispose();$base.Dispose()}
    if($command -notmatch '^"([^"]+)" --uninstall$'){throw 'Unexpected Control Panel command'}
    $setup=Start-Process -FilePath $Matches[1] -ArgumentList '--uninstall' -PassThru -WindowStyle Normal
    Confirm-Removal $setup.Id
    if(-not $setup.WaitForExit(15000)){throw 'Control Panel removal did not finish'}
    Assert-Removed;$result.control_panel_remove='PASS'
    $session=Start-ExactInteractiveExcel (Get-ExcelProcessBaseline) 'r97 after removal'
    $excel=$session.excel;$binding=$session.binding
    $books=$excel.Workbooks
    try{for($i=1;$i -le $books.Count;$i++){$book=$books.Item($i);try{if($book.Name -like ('*'+$ExpectedVersion+'*')){throw 'Removed XLAM loaded'}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($books)}
    $addins=$excel.COMAddIns
    try{for($i=1;$i -le $addins.Count;$i++){$addin=$addins.Item($i);try{if($addin.ProgId -eq 'LH.NxHost.Connect'){throw 'Removed COM add-in loaded'}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addin)}}}finally{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addins)}
    Close-OwnedExcel;$result.post_remove_excel='PASS'
    # First-install back/cancel must not create an installed entry point.
    $setup=Start-Process -FilePath $SetupExe -PassThru -WindowStyle Normal
    Next $setup.Id
    Click (Find-Button $setup.Id ('< '+(-join @([char]0xB4A4,[char]0xB85C))))
    Next $setup.Id
    Click (Find-Button $setup.Id (-join @([char]0xCDE8,[char]0xC18C)))
    if(-not $setup.WaitForExit(15000)){throw 'Cancelled install did not exit'}
    Assert-Removed;$result.fresh_back_cancel='PASS'
    $setup=Start-Process -FilePath $SetupExe -ArgumentList '--uninstall' -PassThru -WindowStyle Normal
    Click (Find-Button $setup.Id (-join @([char]0xB9C8,[char]0xCE68)))
    if(-not $setup.WaitForExit(15000)){throw 'Repeated removal did not exit'}
    Assert-Removed;$result.repeated_removal='PASS'
    $setup=Start-Process -FilePath $SetupExe -PassThru -WindowStyle Normal
    Next $setup.Id
    Capture $setup.Id 'fresh-install-review'
    Click (Find-Button $setup.Id (-join @([char]0xC124,[char]0xCE58)))
    $finish=Find-Button $setup.Id (-join @([char]0xB9C8,[char]0xCE68)) -Seconds 60
    $root=[Windows.Automation.AutomationElement]::RootElement.FindFirst([Windows.Automation.TreeScope]::Children,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,$setup.Id))
    $checkbox=$root.FindFirst([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ControlTypeProperty,[Windows.Automation.ControlType]::CheckBox))
    $toggle=$checkbox.GetCurrentPattern([Windows.Automation.TogglePattern]::Pattern)
    if($toggle.Current.ToggleState -ne [Windows.Automation.ToggleState]::Off){$toggle.Toggle()}
    Click $finish
    if(-not $setup.WaitForExit(15000)){throw 'Fresh install did not finish'}
    if(@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count){throw 'Excel launched despite opt out'}
    $result.fresh_install_opt_out='PASS'
    $setup=Start-Process -FilePath $SetupExe -PassThru -WindowStyle Normal
    Next $setup.Id
    Choose $setup.Id (-join @([char]0xC81C,[char]0xAC70))
    Next $setup.Id
    Confirm-Removal $setup.Id
    if(-not $setup.WaitForExit(15000)){throw 'EXE removal did not finish'}
    Assert-Removed;$result.exe_remove='PASS';$result.status='PASS'
}catch{$result.status='FAIL';$result.failure=$_.Exception.Message}
finally{
    if($null -ne $excel -and $null -ne $binding -and $binding.owned){try{Close-OwnedExcel}catch{$result.cleanup_failure=$_.Exception.Message}}
    $env:LHEXCEL_PROFILE_ROOT=$previousProfile
    [IO.File]::WriteAllText((Join-Path $EvidenceRoot 'installer-ui.json'),($result|ConvertTo-Json -Depth 5),(New-Object Text.UTF8Encoding($false)))
}
$result|ConvertTo-Json -Depth 5
if($result.status -ne 'PASS'){exit 1}
