function Invoke-R72HangulPicker([object]$Excel,[int]$ExcelPid,[string]$Macro,[string]$Action,[string]$Address) {
    $watcher=Start-Job -ArgumentList $ExcelPid,$Action,$Address -ScriptBlock {
        param($ownedPid,$action,$address)
        Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes,System.Windows.Forms
        Add-Type 'using System; using System.Runtime.InteropServices; public class NxPickerWindow { [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd,uint msg,IntPtr wp,IntPtr lp); }'
        $deadline=[DateTime]::UtcNow.AddSeconds(30)
        $dialog=$null
        do {
            $windows=[Windows.Automation.AutomationElement]::RootElement.FindAll(
                [Windows.Automation.TreeScope]::Children,
                (New-Object Windows.Automation.PropertyCondition([Windows.Automation.AutomationElement]::ProcessIdProperty,[int]$ownedPid)))
            foreach($window in $windows){
                if($window.Current.ClassName -eq 'bosa_sdm_XL9'){$dialog=$window;break}
            }
            if($null -eq $dialog){Start-Sleep -Milliseconds 100}
        }while($null -eq $dialog -and [DateTime]::UtcNow -lt $deadline)
        if($null -eq $dialog){throw 'Owned range picker was not observed'}
        $receipt=[ordered]@{title=$dialog.Current.Name;pid=$ownedPid;action=$action;address=$address;controls=@();status='FAIL'}
        try{
            $controls=$dialog.FindAll([Windows.Automation.TreeScope]::Descendants,[Windows.Automation.Condition]::TrueCondition)
            $edits=@()
            foreach($control in $controls){
                $receipt.controls+=@{name=$control.Current.Name;type=$control.Current.ControlType.ProgrammaticName;id=$control.Current.AutomationId}
                if($control.Current.ClassName -eq 'EDTBX' -and $control.Current.IsEnabled){$edits+=,$control}
            }
            if($action -eq 'accept'){
                if($edits.Count -ne 1){throw ('Expected one range edit, observed '+$edits.Count)}
                if(-not $edits[0].Current.Name.EndsWith($address)){throw 'Picker default does not match the selected range'}
                [void][NxPickerWindow]::PostMessage([IntPtr]$dialog.Current.NativeWindowHandle,0x100,[IntPtr]13,[IntPtr]1)
                [void][NxPickerWindow]::PostMessage([IntPtr]$dialog.Current.NativeWindowHandle,0x102,[IntPtr]13,[IntPtr]1)
                [void][NxPickerWindow]::PostMessage([IntPtr]$dialog.Current.NativeWindowHandle,0x101,[IntPtr]13,[IntPtr]1)
            }else{
                [void][NxPickerWindow]::PostMessage([IntPtr]$dialog.Current.NativeWindowHandle,0x100,[IntPtr]27,[IntPtr]1)
            }
            $receipt.status='PASS'
        }catch{
            $receipt.error=$_.Exception.Message
            [void][NxPickerWindow]::PostMessage([IntPtr]$dialog.Current.NativeWindowHandle,0x100,[IntPtr]27,[IntPtr]1)
        }
        [pscustomobject]$receipt
    }
    try{
        [void]$Excel.Run($Macro)
        if($null -eq (Wait-Job $watcher -Timeout 40)){throw 'Picker watcher did not finish'}
        $result=@(Receive-Job $watcher -ErrorAction Stop)
        if($result.Count -ne 1 -or $result[0].status -ne 'PASS'){throw ($result | ConvertTo-Json -Depth 5 -Compress)}
        return $result[0]
    }finally{
        if($watcher.State -eq 'Running'){Stop-Job $watcher}
        Remove-Job $watcher -Force
    }
}
