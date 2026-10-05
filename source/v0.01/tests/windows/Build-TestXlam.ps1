param(
    [Parameter(Mandatory=$true)][ValidateSet('RedProbe','Green')][string]$Mode,
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [string]$Suite = 'Core',
    [string]$ExpectedCompileFailure = '',
    [string]$ExtensionManifest = '',
    [string]$SuiteOwner = '',
    [string]$RunId = '',
    [string]$SourceDigest = '',
    [string]$SnapshotDigest = '',
    [string]$EvidenceRoot = ''
)

function Release-ComObject([object]$Value) { if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } }
function Initialize-NativeWindowApi {
    if (-not ('NxBuild.NativeWindowApi' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
namespace NxBuild { public static class NativeWindowApi {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
}}
'@
    }
}
function Get-ExcelPid([object]$Excel) {
    Initialize-NativeWindowApi
    [uint32]$excelProcessId=0
    [void][NxBuild.NativeWindowApi]::GetWindowThreadProcessId([IntPtr]$Excel.Hwnd,[ref]$excelProcessId)
    return $excelProcessId
}
. (Join-Path $PSScriptRoot '../../build/Excel-ProcessLifecycle.ps1')
function Test-ForbiddenSource([string]$FullPath) {
    $separator = '[\\/]'
    $forbidden = @(
        ('src' + $separator + 'vba' + $separator + 'ui'),
        ('src' + $separator + 'vba' + $separator + 'document'),
        ('src' + $separator + 'ribbon'),
        ('custom' + 'UI'),
        ('\.xlam$')
    )
    foreach ($pattern in $forbidden) {
        if ($FullPath -match $pattern) { throw 'Forbidden product boundary source rejected' }
    }
}
function Assert-ExactJsonProperties([object]$Object, [string[]]$Expected, [string]$Context) {
    if ($null -eq $Object) { throw ($Context + ' is missing') }
    $actual = @($Object.PSObject.Properties.Name | Sort-Object)
    $wanted = @($Expected | Sort-Object)
    if (($actual -join '|') -ne ($wanted -join '|')) { throw ($Context + ' property set rejected') }
}
function Assert-JsonUserFormLayout([object]$Layout, [string]$LayoutPath) {
    $rootKeys = @('schema_version','name','caption','width','height','client_width','client_height','start_up_position','code_path','default_control','cancel_control','controls')
    $rootOptionalKeys = @('dialog_id','variant','dynamic_regions')
    $commonKeys = @('type','name','left','top','width','height','tab_index','tab_stop','enabled')
    $sharedOptionalKeys = @('tag','control_tip_text','font_name','font_size','font_bold','fore_color','back_color')
    $typeOptionalKeys = @{
        TextBox = @('locked','multi_line','text_align','selectable')
        ListBox = @('items','selected_index','multi_select','column_count','column_widths')
    }
    $extraKeys = @{
        Label = @('caption')
        TextBox = @('value')
        CommandButton = @('caption')
        ComboBox = @('items','selected_index','style','match_required')
        ListBox = @()
        CheckBox = @('caption','value')
    }
    $rootPropertyNames = @($Layout.PSObject.Properties.Name)
    $presentRootOptionalKeys = @($rootOptionalKeys | Where-Object { $rootPropertyNames -contains $_ })
    Assert-ExactJsonProperties $Layout @($rootKeys + $presentRootOptionalKeys) 'UserForm layout'
    if ($Layout.schema_version -is [bool] -or -not ($Layout.schema_version -is [int]) -or $Layout.schema_version -ne 2) { throw 'UserForm layout schema rejected' }
    if (-not ($Layout.name -is [string]) -or [string]$Layout.name -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { throw 'UserForm name rejected' }
    if (-not ($Layout.caption -is [string]) -or [string]::IsNullOrWhiteSpace([string]$Layout.caption)) { throw 'UserForm caption rejected' }
    foreach ($name in @('width','height','client_width','client_height','start_up_position')) { if (-not ($Layout.$name -is [int])) { throw ('UserForm integer property rejected: ' + $name) } }
    if (($rootPropertyNames -contains 'dialog_id') -xor ($rootPropertyNames -contains 'variant')) { throw 'UserForm dialog metadata pair rejected' }
    if ($rootPropertyNames -contains 'dialog_id' -and ([string]$Layout.dialog_id -notmatch '^NX-DLG-[A-Z0-9-]+$' -or [string]$Layout.variant -notmatch '^[a-z][a-z0-9-]+$')) { throw 'UserForm dialog metadata rejected' }
    $isControlFreeOverlay = ($rootPropertyNames -contains 'dialog_id' -and [string]$Layout.dialog_id -eq 'NX-DLG-FOCUS-OVERLAY' -and [string]$Layout.variant -eq 'owned-modeless-overlay')
    if ($Layout.width -le 0 -or $Layout.height -le 0 -or $Layout.client_width -le 0 -or $Layout.client_height -le 0) { throw 'UserForm geometry rejected' }
    if ($isControlFreeOverlay) {
        if ($Layout.start_up_position -ne 0 -or $null -eq $Layout.controls -or -not ($Layout.controls -is [array]) -or @($Layout.controls).Count -ne 0 -or $null -ne $Layout.default_control -or $null -ne $Layout.cancel_control) { throw 'Control-free overlay contract rejected' }
    } else {
        if ($Layout.start_up_position -ne 1) { throw 'UserForm geometry rejected' }
        if ($null -eq $Layout.controls -or -not ($Layout.controls -is [array]) -or @($Layout.controls).Count -eq 0) { throw 'UserForm controls missing' }
        if (-not ($Layout.default_control -is [string]) -or [string]$Layout.default_control -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { throw 'UserForm default control rejected' }
    }
    if (-not ($Layout.code_path -is [string]) -or [string]$Layout.code_path -notmatch '^src/vba/(?:frame/ui|features/[a-z][a-z0-9_]*/ui)/[A-Za-z][A-Za-z0-9_]*\.vba$') { throw 'UserForm code path rejected' }
    if ($null -ne $Layout.cancel_control -and (-not ($Layout.cancel_control -is [string]) -or [string]$Layout.cancel_control -notmatch '^[A-Za-z][A-Za-z0-9_]*$')) { throw 'UserForm cancel control rejected' }
    if ($null -ne $Layout.cancel_control -and [string]$Layout.default_control -eq [string]$Layout.cancel_control) { throw 'UserForm default and cancel controls must differ' }
    $seen = @{}
    $interactiveTabs = @{}
    $controlTypes = @{}
    foreach ($item in @($Layout.controls)) {
        $type = [string]$item.type
        if (-not $extraKeys.ContainsKey($type)) { throw 'UserForm control type rejected' }
        $propertyNames = @($item.PSObject.Properties.Name)
        $allowedOptionalKeys = @($sharedOptionalKeys)
        if ($typeOptionalKeys.ContainsKey($type)) { $allowedOptionalKeys += @($typeOptionalKeys[$type]) }
        $presentOptionalKeys = @($allowedOptionalKeys | Where-Object { $propertyNames -contains $_ })
        Assert-ExactJsonProperties $item @($commonKeys + $extraKeys[$type] + $presentOptionalKeys) ('UserForm control ' + $type)
        if (-not ($item.name -is [string]) -or [string]$item.name -notmatch '^[A-Za-z][A-Za-z0-9_]*$' -or $seen.ContainsKey([string]$item.name)) { throw 'UserForm control name rejected' }
        $seen[[string]$item.name] = $true
        $controlTypes[[string]$item.name] = $type
        foreach ($name in @('left','top','width','height','tab_index')) { if (-not ($item.$name -is [int])) { throw ('UserForm control integer property rejected: ' + $name) } }
        foreach ($name in @('tab_stop','enabled')) { if (-not ($item.$name -is [bool])) { throw ('UserForm control boolean property rejected: ' + $name) } }
        foreach ($name in @('tag','control_tip_text','font_name')) {
            if ($propertyNames -contains $name -and (-not ($item.$name -is [string]) -or [string]::IsNullOrWhiteSpace([string]$item.$name))) { throw ('UserForm control string property rejected: ' + $name) }
        }
        foreach ($name in @('fore_color','back_color')) {
            if ($propertyNames -contains $name -and ($item.$name -is [bool] -or -not ($item.$name -is [int]))) { throw ('UserForm control color property rejected: ' + $name) }
        }
        if ($propertyNames -contains 'font_size' -and ($item.font_size -is [bool] -or -not ($item.font_size -is [ValueType]) -or [double]$item.font_size -le 0)) { throw 'UserForm control font size rejected' }
        if ($propertyNames -contains 'font_bold' -and -not ($item.font_bold -is [bool])) { throw 'UserForm control font bold rejected' }
        if ($item.left -lt 0 -or $item.top -lt 0 -or $item.width -le 0 -or $item.height -le 0 -or $item.left + $item.width -gt $Layout.client_width -or $item.top + $item.height -gt $Layout.client_height) { throw 'UserForm control bounds rejected' }
        if ($item.tab_index -lt 0 -and ($type -ne 'Label' -or $item.tab_index -ne -1 -or $item.tab_stop)) { throw 'UserForm control tab index rejected' }
        if ($item.tab_stop) {
            $tabKey = [string]$item.tab_index
            if ($interactiveTabs.ContainsKey($tabKey)) { throw 'UserForm interactive tab index rejected' }
            $interactiveTabs[$tabKey] = $true
        }
        switch ($type) {
            'Label' {
                if (-not ($item.caption -is [string]) -or $item.tab_stop -or -not $item.enabled) { throw 'UserForm Label contract rejected' }
            }
            'TextBox' {
                if (-not ($item.value -is [string])) { throw 'UserForm TextBox value rejected' }
                foreach ($name in @('locked','multi_line')) { if ($propertyNames -contains $name -and -not ($item.$name -is [bool])) { throw ('UserForm TextBox boolean property rejected: ' + $name) } }
                $isLocked = if ($propertyNames -contains 'locked') { [bool]$item.locked } else { $false }
                if ($propertyNames -contains 'text_align' -and [string]$item.text_align -notin @('left','center','right')) { throw 'UserForm TextBox alignment rejected' }
                if ($propertyNames -contains 'selectable' -and -not ($item.selectable -is [bool])) { throw 'UserForm TextBox selectable contract rejected' }
                if ($isLocked -and $item.tab_stop -and (-not ($propertyNames -contains 'selectable') -or -not [bool]$item.selectable)) { throw 'UserForm display TextBox contract rejected' }
                if ($propertyNames -contains 'selectable' -and [bool]$item.selectable -and (-not $isLocked -or -not $item.tab_stop)) { throw 'UserForm selectable TextBox must be locked and tabbable' }
            }
            'CommandButton' {
                $isAccessibleColorSwatch = ($item.caption -is [string] -and $item.caption.Length -eq 0 -and $propertyNames -contains 'back_color' -and $propertyNames -contains 'control_tip_text')
                if (-not ($item.caption -is [string]) -or ([string]::IsNullOrWhiteSpace([string]$item.caption) -and -not $isAccessibleColorSwatch)) { throw 'UserForm CommandButton caption rejected' }
            }
            'CheckBox' {
                if (-not ($item.caption -is [string]) -or [string]::IsNullOrWhiteSpace([string]$item.caption) -or -not ($item.value -is [bool])) { throw 'UserForm CheckBox contract rejected' }
            }
            'ComboBox' {
                if ($null -eq $item.items -or -not ($item.items -is [array]) -or @($item.items).Count -lt 1 -or @($item.items).Count -gt 20) { throw 'UserForm ComboBox items rejected' }
                $choices = @{}
                foreach ($choice in @($item.items)) {
                    if (-not ($choice -is [string]) -or $choice.Length -ne $choice.Trim().Length -or $choice.Trim().Length -lt 1 -or $choice.Trim().Length -gt 80 -or $choice.IndexOf([char]0) -ge 0 -or $choice -match '[\r\n]' -or $choices.ContainsKey($choice)) { throw 'UserForm ComboBox item rejected' }
                    $choices[$choice] = $true
                }
                if ($item.selected_index -is [bool] -or -not ($item.selected_index -is [int]) -or $item.selected_index -lt 0 -or $item.selected_index -ge @($item.items).Count) { throw 'UserForm ComboBox selected index rejected' }
                if ($item.style -ne 'drop_down_list' -or -not ($item.match_required -is [bool]) -or -not $item.match_required -or -not $item.tab_stop -or -not $item.enabled) { throw 'UserForm ComboBox behavior rejected' }
            }
            'ListBox' {
                $hasItems = $propertyNames -contains 'items'
                $hasSelectedIndex = $propertyNames -contains 'selected_index'
                if ($hasItems -xor $hasSelectedIndex) { throw 'UserForm ListBox static metadata pair rejected' }
                if ($hasItems) {
                    if ($null -eq $item.items -or -not ($item.items -is [array]) -or @($item.items).Count -lt 1 -or @($item.items).Count -gt 100) { throw 'UserForm ListBox items rejected' }
                    $choices = @{}
                    foreach ($choice in @($item.items)) {
                        if (-not ($choice -is [string]) -or $choice.Length -ne $choice.Trim().Length -or $choice.Trim().Length -lt 1 -or $choice.Trim().Length -gt 160 -or $choice.IndexOf([char]0) -ge 0 -or $choice -match '[\r\n]' -or $choices.ContainsKey($choice)) { throw 'UserForm ListBox item rejected' }
                        $choices[$choice] = $true
                    }
                    if ($item.selected_index -is [bool] -or -not ($item.selected_index -is [int]) -or $item.selected_index -lt 0 -or $item.selected_index -ge @($item.items).Count) { throw 'UserForm ListBox selected index rejected' }
                }
                if ($propertyNames -contains 'multi_select' -and -not ($item.multi_select -is [bool])) { throw 'UserForm ListBox multi-select rejected' }
                if (($propertyNames -contains 'column_count') -xor ($propertyNames -contains 'column_widths')) { throw 'UserForm ListBox column metadata pair rejected' }
                if ($propertyNames -contains 'column_count' -and ($item.column_count -is [bool] -or -not ($item.column_count -is [int]) -or $item.column_count -lt 1 -or $item.column_count -gt 10 -or -not ($item.column_widths -is [string]) -or [string]::IsNullOrWhiteSpace([string]$item.column_widths))) { throw 'UserForm ListBox column metadata rejected' }
                if (-not $item.tab_stop -or -not $item.enabled) { throw 'UserForm ListBox behavior rejected' }
            }
        }
    }
    if ($null -ne $Layout.default_control -and (-not $controlTypes.ContainsKey([string]$Layout.default_control) -or $controlTypes[[string]$Layout.default_control] -ne 'CommandButton')) { throw 'UserForm default control reference rejected' }
    if ($null -ne $Layout.cancel_control -and (-not $controlTypes.ContainsKey([string]$Layout.cancel_control) -or $controlTypes[[string]$Layout.cancel_control] -ne 'CommandButton')) { throw 'UserForm cancel control reference rejected' }
}
function Set-VbComponentProperty([object]$Component, [string]$Name, [object]$Value) {
    $properties = $null; $property = $null
    try {
        $properties = $Component.Properties
        $property = $properties.Item($Name)
        [void]$property.GetType().InvokeMember(
            'Value',
            [Reflection.BindingFlags]::SetProperty,
            $null,
            $property,
            @($Value)
        )
    } catch {
        $valueType = if ($null -eq $Value) { 'null' } else { $Value.GetType().FullName }
        throw ('UserForm component property failed: name=' + $Name + '; value_type=' + $valueType + '; cause=' + $_.Exception.Message)
    } finally {
        Release-ComObject $property
        Release-ComObject $properties
        $property = $null; $properties = $null
    }
}
function Get-VbComponentPropertyDouble([object]$Component, [string]$Name) {
    $properties = $null; $property = $null
    try {
        $properties = $Component.Properties
        $property = $properties.Item($Name)
        return [Convert]::ToDouble($property.Value, [Globalization.CultureInfo]::InvariantCulture)
    } catch {
        throw ('UserForm component property read failed: name=' + $Name + '; cause=' + $_.Exception.Message)
    } finally {
        Release-ComObject $property
        Release-ComObject $properties
        $property = $null; $properties = $null
    }
}
function Set-VbUserFormDimensions([object]$Component, [object]$Definition) {
    $outerWidth = [double]$Definition.width
    $outerHeight = [double]$Definition.height
    $clientWidth = [double]$Definition.client_width
    $clientHeight = [double]$Definition.client_height
    Set-VbComponentProperty $Component 'Width' $outerWidth
    Set-VbComponentProperty $Component 'Height' $outerHeight
    $insideWidth = Get-VbComponentPropertyDouble $Component 'InsideWidth'
    $insideHeight = Get-VbComponentPropertyDouble $Component 'InsideHeight'
    Set-VbComponentProperty $Component 'Width' ([Math]::Round([Math]::Max($outerWidth, $outerWidth + $clientWidth - $insideWidth), 2))
    Set-VbComponentProperty $Component 'Height' ([Math]::Round([Math]::Max($outerHeight, $outerHeight + $clientHeight - $insideHeight), 2))
    if ((Get-VbComponentPropertyDouble $Component 'InsideWidth') + 0.5 -lt $clientWidth -or (Get-VbComponentPropertyDouble $Component 'InsideHeight') + 0.5 -lt $clientHeight) {
        throw 'UserForm client dimensions could not be satisfied'
    }
}
function Set-JsonFormProperties([object]$Object, [object]$Definition, [bool]$IsForm) {
    if ($IsForm) {
        Set-VbComponentProperty $Object 'Caption' ([string]$Definition.caption)
        Set-VbUserFormDimensions $Object $Definition
        Set-VbComponentProperty $Object 'StartUpPosition' ([int]$Definition.start_up_position)
        return
    }
    $propertyStage = 'Left'
    try {
    $Object.Left = [double]$Definition.left
    $propertyStage = 'Top'
    $Object.Top = [double]$Definition.top
    $propertyStage = 'Width'
    $Object.Width = [double]$Definition.width
    $propertyStage = 'Height'
    $Object.Height = [double]$Definition.height
    $propertyStage = 'TabStop'
    $Object.TabStop = [bool]$Definition.tab_stop
    if ([int]$Definition.tab_index -ge 0) { $propertyStage = 'TabIndex'; $Object.TabIndex = [int]$Definition.tab_index }
    $propertyStage = 'Enabled'
    $Object.Enabled = [bool]$Definition.enabled
    $propertyNames = @($Definition.PSObject.Properties.Name)
    if ($propertyNames -contains 'tag') { $propertyStage = 'Tag'; $Object.Tag = [string]$Definition.tag }
    if ($propertyNames -contains 'control_tip_text') { $propertyStage = 'ControlTipText'; $Object.ControlTipText = [string]$Definition.control_tip_text }
    if ($propertyNames -contains 'fore_color') { $propertyStage = 'ForeColor'; $Object.ForeColor = [int]$Definition.fore_color }
    if ($propertyNames -contains 'back_color') { $propertyStage = 'BackColor'; $Object.BackColor = [int]$Definition.back_color }
    if ($propertyNames -contains 'font_name') { $propertyStage = 'FontName'; $Object.FontName = [string]$Definition.font_name }
    if ($propertyNames -contains 'font_size') { $propertyStage = 'FontSize'; $Object.FontSize = [double]$Definition.font_size }
    if ($propertyNames -contains 'font_bold') { $propertyStage = 'FontBold'; $Object.FontBold = [bool]$Definition.font_bold }
    switch ([string]$Definition.type) {
        'Label' { $propertyStage = 'Label.Caption'; $Object.Caption = [string]$Definition.caption }
        'TextBox' {
            $propertyStage = 'TextBox.Value'
            $Object.Value = [string]$Definition.value
            if ($propertyNames -contains 'locked') { $propertyStage = 'TextBox.Locked'; $Object.Locked = [bool]$Definition.locked }
            if ($propertyNames -contains 'multi_line') { $propertyStage = 'TextBox.MultiLine'; $Object.MultiLine = [bool]$Definition.multi_line }
            if ($propertyNames -contains 'text_align') {
                $propertyStage = 'TextBox.TextAlign'
                $Object.TextAlign = switch ([string]$Definition.text_align) { 'left' { 1 }; 'center' { 2 }; 'right' { 3 } }
            }
        }
        'CommandButton' { $propertyStage = 'CommandButton.Caption'; $Object.Caption = [string]$Definition.caption }
        'CheckBox' { $propertyStage = 'CheckBox.Caption'; $Object.Caption = [string]$Definition.caption; $propertyStage = 'CheckBox.Value'; $Object.Value = [bool]$Definition.value }
        'ComboBox' {
            $propertyStage = 'ComboBox.Style'
            $Object.Style = 2
            $propertyStage = 'ComboBox.MatchRequired'
            $Object.MatchRequired = $true
            $propertyStage = 'ComboBox.Items'
            foreach ($choice in @($Definition.items)) { [void]$Object.AddItem([string]$choice) }
            $propertyStage = 'ComboBox.ListIndex'
            $Object.ListIndex = [int]$Definition.selected_index
        }
        'ListBox' {
            if ($propertyNames -contains 'multi_select') { $propertyStage = 'ListBox.MultiSelect'; $Object.MultiSelect = if ([bool]$Definition.multi_select) { 1 } else { 0 } }
            if ($propertyNames -contains 'column_count') { $propertyStage = 'ListBox.ColumnCount'; $Object.ColumnCount = [int]$Definition.column_count; $propertyStage = 'ListBox.ColumnWidths'; $Object.ColumnWidths = [string]$Definition.column_widths }
            if ($propertyNames -contains 'items') {
                $propertyStage = 'ListBox.Items'
                foreach ($choice in @($Definition.items)) { [void]$Object.AddItem([string]$choice) }
                $propertyStage = 'ListBox.ListIndex'
                $Object.ListIndex = [int]$Definition.selected_index
            }
        }
        default { throw 'Unsupported UserForm control type' }
    }
    } catch {
        throw ('UserForm JSON property failed: property=' + $propertyStage + '; cause=' + $_.Exception.Message)
    }
}
function Import-JsonUserForm([object]$Components, [string]$DataRoot, [string]$LayoutPath, [hashtable]$AllowedPaths, [hashtable]$ClaimedCodePaths, [Collections.Generic.List[object]]$EvidenceForms) {
    $layoutText = [IO.File]::ReadAllText($LayoutPath, [Text.Encoding]::UTF8)
    $layout = $layoutText | ConvertFrom-Json
    Assert-JsonUserFormLayout $layout $LayoutPath
    $codePath = [IO.Path]::GetFullPath((Join-Path $DataRoot ([string]$layout.code_path)))
    $layoutFull = [IO.Path]::GetFullPath($LayoutPath)
    $layoutDirectory = [IO.Path]::GetDirectoryName($layoutFull)
    $layoutBaseName = [IO.Path]::GetFileName($layoutFull) -replace '\.form\.json$', ''
    if (-not [StringComparer]::OrdinalIgnoreCase.Equals([IO.Path]::GetDirectoryName($codePath), $layoutDirectory)) { throw 'UserForm code path must be an owner sibling' }
    if (-not [StringComparer]::Ordinal.Equals([IO.Path]::GetFileNameWithoutExtension($codePath), [string]$layout.name)) { throw 'UserForm code basename rejected' }
    if (-not [StringComparer]::Ordinal.Equals($layoutBaseName, [string]$layout.name)) { throw 'UserForm layout basename rejected' }
    $codeKey = $codePath.ToLowerInvariant()
    if (-not $AllowedPaths.ContainsKey($codeKey) -or [IO.Path]::GetExtension($codePath).ToLowerInvariant() -ne '.vba') { throw 'UserForm code path is not manifest-bound' }
    if ($ClaimedCodePaths.ContainsKey($codeKey)) { throw 'UserForm code path is claimed more than once' }
    $ClaimedCodePaths[$codeKey] = $true
    $code = [IO.File]::ReadAllText($codePath, [Text.Encoding]::UTF8)
    if ($code -notmatch '^Option Explicit(?:\r\n|\r|\n)' -or $code -match '(?m)^(VERSION|Begin |Attribute )|OleObjectBlob') { throw 'UserForm code module boundary rejected' }
    $component = $null; $designer = $null; $control = $null; $codeModule = $null; $keyboardControl = $null
    $comboEvidence = @()
    $listEvidence = @()
    $stage = 'component-add'; $currentType = ''; $currentName = ''
    try {
        $component = $Components.Add(3)
        $stage = 'component-name'
        $component.Name = [string]$layout.name
        $stage = 'form-properties'
        Set-JsonFormProperties -Object $component -Definition $layout -IsForm $true
        $stage = 'designer-open'
        $designer = $component.Designer
        if ($null -eq $designer) { throw 'JSON UserForm designer is unavailable' }
        foreach ($item in @($layout.controls)) {
            $currentType = [string]$item.type
            $currentName = [string]$item.name
            $stage = 'control-add:' + [string]$item.type + ':' + [string]$item.name
            $progId = switch ([string]$item.type) {
                'Label' { 'Forms.Label.1' }
                'TextBox' { 'Forms.TextBox.1' }
                'CommandButton' { 'Forms.CommandButton.1' }
                'ComboBox' { 'Forms.ComboBox.1' }
                'ListBox' { 'Forms.ListBox.1' }
                'CheckBox' { 'Forms.CheckBox.1' }
                default { throw 'Unsupported JSON UserForm control type' }
            }
            $control = $designer.Controls.Add($progId, [string]$item.name, $true)
            $stage = 'control-properties:' + [string]$item.type + ':' + [string]$item.name
            Set-JsonFormProperties -Object $control -Definition $item -IsForm $false
            if ([string]$item.type -eq 'ComboBox') {
                $comboEvidence += [pscustomobject][ordered]@{
                    name=[string]$item.name
                    prog_id='Forms.ComboBox.1'
                    items=@($item.items | ForEach-Object { [string]$_ })
                    selected_index=[int]$item.selected_index
                    style_value=2
                    match_required=$true
                    set_sequence=@('style','match_required','items','selected_index')
                }
            }
            if ([string]$item.type -eq 'ListBox') {
                $listPropertyNames=@($item.PSObject.Properties.Name)
                $listItems=if($listPropertyNames -contains 'items'){@($item.items|ForEach-Object {[string]$_})}else{@()}
                $listIndex=if($listPropertyNames -contains 'selected_index'){[int]$item.selected_index}else{-1}
                $listMulti=if($listPropertyNames -contains 'multi_select'){[bool]$item.multi_select}else{$false}
                $listSetSequence=@()
                foreach($propertyName in @('multi_select','column_count','column_widths','items','selected_index')){if($listPropertyNames -contains $propertyName){$listSetSequence+=$propertyName}}
                $listEvidence += [pscustomobject][ordered]@{
                    name=[string]$item.name
                    prog_id='Forms.ListBox.1'
                    items=@($listItems)
                    selected_index=$listIndex
                    multi_select=$listMulti
                    multi_select_value=if($listMulti){1}else{0}
                    set_sequence=@($listSetSequence)
                }
            }
            Release-ComObject $control
            $control = $null
        }
        if ($null -ne $layout.default_control) {
            $currentType = 'CommandButton'
            $currentName = [string]$layout.default_control
            $stage = 'default-control:' + $currentName
            $keyboardControl = $designer.Controls.Item($currentName)
            $keyboardControl.Default = $true
            Release-ComObject $keyboardControl
            $keyboardControl = $null
        }
        if ($null -ne $layout.cancel_control) {
            $currentName = [string]$layout.cancel_control
            $stage = 'cancel-control:' + $currentName
            $keyboardControl = $designer.Controls.Item($currentName)
            $keyboardControl.Cancel = $true
            Release-ComObject $keyboardControl
            $keyboardControl = $null
        }
        $currentType = ''; $currentName = ''
        $stage = 'code-module-open'
        $codeModule = $component.CodeModule
        $stage = 'code-module-add'
        [void]$codeModule.AddFromString(($code -replace '\r\n|\r|\n', "`r`n"))
        $formEvidence=[pscustomobject][ordered]@{
            name=[string]$layout.name
            layout_path=Get-EvidenceRelativePath $DataRoot $layoutFull
            layout_sha256=Get-Sha256 $layoutFull
            code_path=Get-EvidenceRelativePath $DataRoot $codePath
            code_sha256=Get-Sha256 $codePath
            controls_created=@($layout.controls).Count
            default_control=if($null -eq $layout.default_control){$null}else{[string]$layout.default_control}
            cancel_control=if($null -eq $layout.cancel_control){$null}else{[string]$layout.cancel_control}
            comboboxes=@($comboEvidence)
        }
        $formEvidence | Add-Member -NotePropertyName listboxes -NotePropertyValue @($listEvidence)
        [void]$EvidenceForms.Add($formEvidence)
    } catch {
        $failure = $_.Exception.Message
        if ($null -ne $component) { try { $Components.Remove($component) } catch {} }
        throw ('JSON UserForm build failed: form=' + [string]$layout.name + '; stage=' + $stage + '; control_type=' + $currentType + '; control_name=' + $currentName + '; cause=' + $failure)
    } finally {
        Release-ComObject $control
        Release-ComObject $keyboardControl
        Release-ComObject $codeModule
        Release-ComObject $designer
        Release-ComObject $component
        $control = $null; $keyboardControl = $null; $codeModule = $null; $designer = $null; $component = $null
    }
}
function Import-ExactVbaFiles([object]$Workbook, [string[]]$Files, [string]$Root, [Collections.Generic.List[object]]$EvidenceForms) {
    if ($null -eq $Workbook) { throw 'Workbook is required' }
    if ($null -eq $Files -or $Files.Count -eq 0) { throw 'Empty file allowlist rejected' }
    $seen = @{}; $allowedPaths = @{}; $claimedCodePaths = @{}; $vbaCount = 0
    foreach ($file in $Files) {
        if ([string]::IsNullOrWhiteSpace($file) -or $file -match '[*?]') { throw 'Wildcard source path rejected' }
        if (-not (Test-Path -PathType Leaf -LiteralPath $file)) { throw 'Non-file source path rejected' }
        $full = [IO.Path]::GetFullPath($file)
        if ($seen.ContainsKey($full)) { throw 'Duplicate source path rejected' }
        $extension = [IO.Path]::GetExtension($full).ToLowerInvariant()
        $isFormJson = $full.ToLowerInvariant().EndsWith('.form.json')
        if ($extension -notin @('.bas', '.cls', '.vba') -and -not $isFormJson) { throw 'Unsupported VBA build input rejected' }
        Test-ForbiddenSource $full
        $seen[$full] = $true
        $allowedPaths[$full.ToLowerInvariant()] = $true
        if ($extension -eq '.vba') { $vbaCount += 1 }
    }
    ${project} = $null
    $components = $null
    $importRoot = Join-Path ([IO.Path]::GetTempPath()) ('LHexcelVbaImport-' + [Guid]::NewGuid().ToString('N'))
    try {
        [void](New-Item -ItemType Directory -Force -Path $importRoot)
        ${project} = $Workbook.VBProject
        $components = ${project}.VBComponents
        foreach ($file in $Files) {
            $component = $null
            $full = [IO.Path]::GetFullPath($file)
            $importPath = Join-Path $importRoot ([IO.Path]::GetFileName($full))
            if (Test-Path -LiteralPath $importPath) { throw 'Duplicate VBA import filename rejected' }
            $extension = [IO.Path]::GetExtension($full).ToLowerInvariant()
            if ($full.ToLowerInvariant().EndsWith('.form.json')) {
                Import-JsonUserForm -Components $components -DataRoot $Root -LayoutPath $full -AllowedPaths $allowedPaths -ClaimedCodePaths $claimedCodePaths -EvidenceForms $EvidenceForms
            } elseif ($extension -eq '.vba') {
                continue
            } else {
                $sourceText = [IO.File]::ReadAllText($full,[Text.Encoding]::UTF8)
                $normalized = $sourceText -replace '\r\n|\r|\n',"`r`n"
                [IO.File]::WriteAllText($importPath,$normalized,[Text.Encoding]::Default)
                try { $component = $components.Import($importPath) }
                finally { Release-ComObject $component; $component = $null }
            }
        }
        if ($claimedCodePaths.Count -ne $vbaCount) { throw 'Every manifest-bound UserForm code path must be claimed exactly once' }
    } finally {
        Release-ComObject $components
        Release-ComObject ${project}
        $components = $null
        ${project} = $null
        Remove-Item -LiteralPath $importRoot -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $importRoot) { throw 'VBA import temp residue' }
    }
}
function Get-Sha256([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant() }
function Get-TextSha256([string]$Text) { $sha = [Security.Cryptography.SHA256]::Create(); try { ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').ToLowerInvariant() } finally { $sha.Dispose() } }
function Write-AtomicUserFormBuildEvidence([string]$Path, [object]$Evidence) {
    $directory = Split-Path -Parent $Path
    [void](New-Item -ItemType Directory -Force -Path $directory)
    $temporary = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
    $writer = New-Object IO.StreamWriter($temporary, $false, (New-Object Text.UTF8Encoding($false)))
    try {
        $json = $Evidence | ConvertTo-Json -Depth 12 -Compress
        $writer.Write($json + "`n")
        $writer.Flush()
        $writer.BaseStream.Flush($true)
    } finally {
        $writer.Close()
    }
    if(Test-Path -LiteralPath $Path){[IO.File]::Replace($temporary,$Path,$null)}else{[IO.File]::Move($temporary,$Path)}
}
function Get-RootRelativePath([string]$Root, [string]$Path) {
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    $pathFull = [IO.Path]::GetFullPath($Path)
    $prefix = $rootFull + [IO.Path]::DirectorySeparatorChar
    if (-not $pathFull.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Source path escapes repository root' }
    $relative = $pathFull.Substring($prefix.Length).Replace('\', '/')
    if ([string]::IsNullOrWhiteSpace($relative) -or $relative -match '(^|/)\.\.(/|$)') { throw 'Source relative path rejected' }
    return $relative
}
function Get-EvidenceRelativePath([string]$Root, [string]$Path) {
    $rootFull=[IO.Path]::GetFullPath($Root).TrimEnd([IO.Path]::DirectorySeparatorChar,[IO.Path]::AltDirectorySeparatorChar)
    $pathFull=[IO.Path]::GetFullPath($Path)
    $prefix=$rootFull+[IO.Path]::DirectorySeparatorChar
    if(-not $pathFull.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)){throw 'Evidence source path escapes repository root'}
    $relative=$pathFull.Substring($prefix.Length).Replace('\','/')
    if([string]::IsNullOrWhiteSpace($relative) -or $relative -match '(^|/)\.\.(/|$)'){throw 'Evidence source relative path rejected'}
    return $relative
}
function Get-SourceRecords([string]$Root, [string[]]$Paths) {
    $records = @()
    foreach ($path in $Paths) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'Missing source path rejected' }
        $records += [pscustomobject]@{ relative_path = Get-RootRelativePath $Root $path; sha256 = Get-Sha256 $path }
    }
    return @($records)
}
function Get-AllowlistDigest([object[]]$Records) {
    $canonical = (($Records | ForEach-Object { $_.relative_path + "`n" + $_.sha256 + "`n" }) -join '')
    return Get-TextSha256 $canonical
}
function Assert-ExactSources([string[]]$Actual, [string[]]$Expected) {
    if ($Actual.Count -ne $Expected.Count) { throw 'Source allowlist count mismatch' }
    $seen = @{}
    for ($index = 0; $index -lt $Expected.Count; $index++) {
        $full = [IO.Path]::GetFullPath($Actual[$index])
        if ($full -ne [IO.Path]::GetFullPath($Expected[$index])) { throw 'Source allowlist order mismatch' }
        if ($seen.ContainsKey($full)) { throw 'Duplicate source path rejected' }
        $forbiddenBoundary = '[*?]|src[\\/]vba[\\/](ui|document)|ribbon|' + ('custom' + 'UI') + '|\.xlam$'
        if ($full -match $forbiddenBoundary) { throw 'Forbidden product boundary source rejected' }
        $seen[$full] = $true
    }
}
function Read-ExtensionManifestGraph([string]$ManifestPath, [string]$Root, [string]$Suite, [string]$SuiteOwner, [string[]]$Base, [int]$Depth, [hashtable]$State) {
    if ($Depth -gt 8) { throw 'Extension manifest dependency depth exceeded' }
    if ([string]::IsNullOrWhiteSpace($ManifestPath) -or -not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) { throw 'Feature suite requires an extension manifest file' }
    $manifestFull = [IO.Path]::GetFullPath($ManifestPath)
    $manifestRelative = Get-RootRelativePath $Root $manifestFull
    if ($manifestRelative -notmatch '^tests/manifests/[A-Za-z][A-Za-z0-9_-]*\.json$') { throw 'Extension dependency manifest path rejected' }
    $manifestKey = $manifestFull.ToLowerInvariant()
    if ($State.Visiting.ContainsKey($manifestKey)) { throw 'Extension manifest dependency cycle rejected' }
    if ($State.Loaded.ContainsKey($manifestKey)) { throw 'Extension manifest duplicate dependency rejected' }
    $State.Visiting[$manifestKey] = $true
    try {
        $manifest = Get-Content -LiteralPath $manifestFull -Raw | ConvertFrom-Json
        Assert-ExactJsonProperties $manifest @('schema_version','suite','owner_phase','base_allowlist_digest','dependencies','modules') 'Extension manifest'
        if ($manifest.schema_version -is [bool] -or -not ($manifest.schema_version -is [int]) -or $manifest.schema_version -ne 2 -or $manifest.suite -ne $Suite) { throw 'Extension manifest schema or requested suite rejected' }
        if ($manifest.owner_phase -ne $SuiteOwner) { throw 'Extension manifest owner must exactly equal suite owner' }
        $baseRecords = Get-SourceRecords $Root $Base
        if ($manifest.base_allowlist_digest -ne (Get-AllowlistDigest $baseRecords)) { throw 'Extension manifest base allowlist digest rejected' }
        $result = New-Object System.Collections.Generic.List[string]
        foreach ($dependency in @($manifest.dependencies)) {
            Assert-ExactJsonProperties $dependency @('path','sha256','suite','owner_phase') 'Extension dependency'
            if (-not ($dependency.path -is [string]) -or [string]$dependency.path -notmatch '^tests/manifests/[A-Za-z][A-Za-z0-9_-]*\.json$') { throw 'Extension dependency path rejected' }
            if (-not ($dependency.sha256 -is [string]) -or [string]$dependency.sha256 -notmatch '^[0-9a-fA-F]{64}$') { throw 'Extension dependency SHA-256 rejected' }
            if (-not ($dependency.suite -is [string]) -or [string]$dependency.suite -notmatch '^[A-Za-z][A-Za-z0-9_-]*$' -or -not ($dependency.owner_phase -is [string]) -or [string]$dependency.owner_phase -notmatch '^[a-z][a-z0-9_-]*$') { throw 'Extension dependency suite or owner rejected' }
            $State.DependencyCount += 1
            if ($State.DependencyCount -gt 8) { throw 'Extension manifest dependency count exceeded' }
            $dependencyFull = [IO.Path]::GetFullPath((Join-Path $Root ([string]$dependency.path)))
            if (-not (Test-Path -LiteralPath $dependencyFull -PathType Leaf) -or (Get-Sha256 $dependencyFull) -ne ([string]$dependency.sha256).ToLowerInvariant()) { throw 'Extension dependency current SHA-256 rejected' }
            $dependencyModules = @(Read-ExtensionManifestGraph $dependencyFull $Root ([string]$dependency.suite) ([string]$dependency.owner_phase) $Base ($Depth + 1) $State)
            foreach ($dependencyModule in $dependencyModules) { [void]$result.Add($dependencyModule) }
        }
        foreach ($entry in @($manifest.modules)) {
            Assert-ExactJsonProperties $entry @('path','sha256','owner_phase') 'Extension module'
            if ($null -eq $entry -or -not ($entry.path -is [string]) -or [string]::IsNullOrWhiteSpace($entry.path) -or -not ($entry.sha256 -is [string]) -or [string]$entry.sha256 -notmatch '^[0-9a-fA-F]{64}$' -or $entry.owner_phase -ne $manifest.owner_phase) { throw 'Extension manifest path, SHA-256, or phase rejected' }
            if ([IO.Path]::IsPathRooted($entry.path) -or $entry.path -match '(^|[\\/])\.\.([\\/]|$)') { throw 'Extension manifest repository-relative path rejected' }
            $full = [IO.Path]::GetFullPath((Join-Path $Root $entry.path))
            $relative = Get-RootRelativePath $Root $full
            $relativeKey = $relative.ToLowerInvariant()
            $extension = [IO.Path]::GetExtension($relative).ToLowerInvariant()
            $isFormJson = $relativeKey.EndsWith('.form.json')
            if ($extension -notin @('.bas', '.cls', '.vba') -and -not $isFormJson) { throw 'Extension manifest build input type rejected' }
            if ($extension -eq '.vba' -or $isFormJson) {
                $expectedUi = if ($manifest.owner_phase -eq 'frame') { 'src/vba/frame/ui' } else { 'src/vba/features/' + [string]$manifest.owner_phase + '/ui' }
                $parent = ([IO.Path]::GetDirectoryName($relative)).Replace('\', '/')
                if (-not [StringComparer]::OrdinalIgnoreCase.Equals($parent, $expectedUi)) { throw 'UserForm input crosses the exact owner UI path' }
            }
            if ($State.Modules.ContainsKey($relativeKey)) { throw 'Extension manifest duplicate path or base shadowing rejected' }
            if (-not (Test-Path -LiteralPath $full -PathType Leaf) -or (Get-Sha256 $full) -ne ([string]$entry.sha256).ToLowerInvariant()) { throw 'Extension manifest current SHA-256 rejected' }
            if ($relative -notmatch ('(^|/)' + [regex]::Escape([string]$manifest.owner_phase) + '(/|$)')) { throw 'Extension manifest cross-phase exact owner-phase boundary rejected' }
            $State.Modules[$relativeKey] = $true
            [void]$result.Add($full)
        }
        if (@($manifest.modules).Count -eq 0) { throw 'Extension manifest requires modules' }
        $State.Loaded[$manifestKey] = $true
        return @($result)
    } finally {
        [void]$State.Visiting.Remove($manifestKey)
    }
}

$dataRoot = [IO.Path]::GetFullPath($DataRoot)
if (-not (Test-Path -LiteralPath $dataRoot -PathType Container)) { throw 'DataRoot directory missing' }
$buildEvidencePath = $null
# Active UserForm suites remain covered by exact owner manifests.
if ($Mode -eq 'Green' -and $Suite -in @('AI','Template','Data','Draw','File','Calculator','Symbols')) {
    if ($RunId -notmatch '^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$' -or $SourceDigest -notmatch '^[0-9a-f]{64}$' -or $SnapshotDigest -notmatch '^[0-9a-f]{64}$' -or [string]::IsNullOrWhiteSpace($EvidenceRoot)) { throw 'UserForm build evidence binding inputs rejected' }
    $buildEvidencePath = Join-Path (Join-Path ([IO.Path]::GetFullPath($EvidenceRoot)) 'build') ($Suite + '.UserFormBuild.json')
    Remove-Item -LiteralPath $buildEvidencePath -Force -ErrorAction SilentlyContinue
    if(Test-Path -LiteralPath $buildEvidencePath){throw 'Stale UserForm build evidence could not be removed'}
}
$controlRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$newBlank = Join-Path $controlRoot 'build/New-BlankXlam.ps1'
$redSources = @(
    Join-Path -Path $dataRoot -ChildPath 'tests/vba/T_Core.bas'
    Join-Path -Path $dataRoot -ChildPath 'tests/vba/NxTestHarness.bas'
)
$greenSources = @(
    (Join-Path $dataRoot 'src/vba/core/NxTypes.bas'), (Join-Path $dataRoot 'src/vba/core/NxConstants.bas'), (Join-Path $dataRoot 'src/vba/core/NxFormulaInterop.bas'), (Join-Path $dataRoot 'src/vba/core/NxUserErrors.bas'), (Join-Path $dataRoot 'src/vba/core/INxFeatureCommand.cls'),
    (Join-Path $dataRoot 'src/vba/core/CNxExecutionContext.cls'), (Join-Path $dataRoot 'src/vba/core/CNxPlan.cls'), (Join-Path $dataRoot 'src/vba/core/CNxResult.cls'),
    (Join-Path $dataRoot 'src/vba/core/CNxStateGuard.cls'), (Join-Path $dataRoot 'src/vba/core/CNxSideEffect.cls'), (Join-Path $dataRoot 'src/vba/core/CNxCellRollbackJournal.cls'),
    (Join-Path $dataRoot 'src/vba/core/NxContextFactory.bas'), (Join-Path $dataRoot 'src/vba/core/NxCommandRunner.bas'), (Join-Path $dataRoot 'src/vba/core/CNxApproval.cls'),
    (Join-Path $dataRoot 'src/vba/core/CNxSideEffectDispatcher.cls'), (Join-Path $dataRoot 'src/vba/core/CNxLimitsPolicy.cls'),
    (Join-Path $dataRoot 'src/vba/core/CNxRangeSelectionSession.cls'), (Join-Path $dataRoot 'src/vba/core/NxRangePicker.bas'),
    (Join-Path $dataRoot 'tests/vba/NxTestHarness.bas'), (Join-Path $dataRoot 'tests/vba/T_Core.bas'), (Join-Path $dataRoot 'tests/vba/CFakeFeatureCommand.cls')
)

if ($Mode -eq 'RedProbe') {
    if ($Suite -ne 'Core' -or $ExpectedCompileFailure -ne 'CNxStateGuard' -or -not [string]::IsNullOrWhiteSpace($ExtensionManifest)) { throw 'RedProbe is Core-only with the exact expected failure' }
    $allowlist = $redSources
} elseif ($Suite -eq 'Core') {
    if (-not [string]::IsNullOrWhiteSpace($ExtensionManifest) -or -not [string]::IsNullOrWhiteSpace($ExpectedCompileFailure)) { throw 'Core rejects any extension or expected compile failure' }
    $allowlist = $greenSources
} else {
    if (-not [string]::IsNullOrWhiteSpace($ExpectedCompileFailure)) { throw 'Feature green does not accept a compile failure symbol' }
    $state = @{ Visiting=@{}; Loaded=@{}; Modules=@{}; DependencyCount=0 }
    foreach ($basePath in $greenSources) { $state.Modules[(Get-RootRelativePath $dataRoot $basePath).ToLowerInvariant()] = $true }
    $allowlist = @($greenSources + (Read-ExtensionManifestGraph $ExtensionManifest $dataRoot $Suite $SuiteOwner $greenSources 0 $state))
}
Assert-ExactSources $allowlist $allowlist

$artifact = Join-Path ([IO.Path]::GetTempPath()) ('NxTask3-' + [Guid]::NewGuid().ToString('N') + '-' + $Mode + '-' + $Suite + '-TestHost.xlam')
$created = & $newBlank -Path $artifact
if ($created -ne [IO.Path]::GetFullPath($artifact)) { throw 'Blank XLAM creation returned an unexpected artifact' }
$excel = $null; $workbooks = $null; $book = $null; $excelProcess=$null; $excelPid=0; $cleanupFailures = @()
$formEvidence = New-Object 'System.Collections.Generic.List[object]'
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excelPid=Get-ExcelPid $excel
    $excelProcess=Get-Process -Id $excelPid -ErrorAction Stop
    [void]$excelProcess.Handle
    [uint32]$excelHwndPid=0
    [void][NxBuild.NativeWindowApi]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelHwndPid)
    if($excelHwndPid -ne $excelPid){throw 'build host Excel instance is not exact-bound'}
    $workbooks = $excel.Workbooks
    $book = $workbooks.Open($artifact)
    Import-ExactVbaFiles -Workbook $book -Files $allowlist -Root $dataRoot -EvidenceForms $formEvidence
    $book.Save()
} finally {
    try { if ($book) { $book.Close($false) } } catch { $cleanupFailures += $_.Exception.Message }
    try { if ($excel) { $excel.Quit() } } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $book } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $workbooks } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $excel } catch { $cleanupFailures += $_.Exception.Message }
    $book = $null; $workbooks = $null; $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if($null -ne $excelProcess){
        $excelCleanup=Stop-ExactProcessAfterGrace $excelProcess 'build host Excel'
        if(-not [string]::IsNullOrWhiteSpace($excelCleanup)){$cleanupFailures += $excelCleanup}
        try{$excelProcess.Dispose()}catch{$cleanupFailures += $_.Exception.Message}
        $excelProcess=$null
    }
}
if ($cleanupFailures.Count -ne 0) { throw ('Build host cleanup failed: ' + ($cleanupFailures -join ' | ')) }
$records = Get-SourceRecords $dataRoot $allowlist
$artifactSha256 = Get-Sha256 $artifact
$buildEvidenceSha256 = $null
if($null -ne $buildEvidencePath){
    $manifestFull=[IO.Path]::GetFullPath($ExtensionManifest)
    $UserFormBuildEvidence=[ordered]@{
        schema_version=1
        run_id=$RunId
        suite=$Suite
        owner=$SuiteOwner
        mode=$Mode
        source_digest=$SourceDigest
        snapshot_digest=$SnapshotDigest
        extension_manifest_path=Get-EvidenceRelativePath $dataRoot $manifestFull
        extension_manifest_sha256=Get-Sha256 $manifestFull
        artifact_path=[IO.Path]::GetFullPath($artifact)
        artifact_sha256=$artifactSha256
        forms=$formEvidence.ToArray()
        created_utc=[DateTime]::UtcNow.ToString('o')
    }
    Write-AtomicUserFormBuildEvidence $buildEvidencePath $UserFormBuildEvidence
    $buildEvidenceSha256=Get-Sha256 $buildEvidencePath
}
[pscustomobject]@{ ArtifactPath = [IO.Path]::GetFullPath($artifact); SourceAllowlist = $records; SourceAllowlistDigest = Get-AllowlistDigest $records; HostSha256 = $artifactSha256; HostFileFormat = 55; Mode = $Mode; Suite = $Suite; BuildEvidencePath=$buildEvidencePath; BuildEvidenceSha256=$buildEvidenceSha256 }
