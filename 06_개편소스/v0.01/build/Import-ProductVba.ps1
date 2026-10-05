param(
    [Parameter(Mandatory=$true)][string]$DataRoot,
    [Parameter(Mandatory=$true)][string]$XlamPath,
    [string]$ManifestPath = (Join-Path $PSScriptRoot 'manifests/Product.json')
)

Set-StrictMode -Version Latest
$root = [IO.Path]::GetFullPath($DataRoot)
$rootPrefix = $root.TrimEnd([char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)) + [IO.Path]::DirectorySeparatorChar
$xlam = [IO.Path]::GetFullPath($XlamPath)
$manifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($manifest.schema_version -ne 3 -or $manifest.suite -ne 'Product' -or $manifest.owner_phase -ne 'product') { throw 'Product manifest identity rejected' }
$records = @($manifest.modules)
$bindings = @($manifest.form_bindings)
if ($records.Count -eq 0 -or $bindings.Count -eq 0) { throw 'Product manifest component inventory rejected' }
$standardCount = @($records | Where-Object { $_.kind -eq 'standard_module' }).Count
$classCount = @($records | Where-Object { $_.kind -eq 'class_module' }).Count
$layoutCount = @($records | Where-Object { $_.kind -eq 'form_layout' }).Count
$formCodeCount = @($records | Where-Object { $_.kind -eq 'form_code' }).Count
if ($standardCount -eq 0 -or $classCount -eq 0 -or $layoutCount -eq 0 -or $formCodeCount -eq 0 -or ($standardCount + $classCount + $layoutCount + $formCodeCount) -ne $records.Count) { throw 'Product manifest kind inventory rejected' }
if ($layoutCount -ne $bindings.Count -or $formCodeCount -ne $bindings.Count) { throw 'Product form component inventory rejected' }
if ($records.path -match '^(?:tests/|tests\\)') { throw 'Product manifest contains forbidden test paths' }
if ($records.path -match '^src[\\/]vba[\\/]features[\\/]hwpx[\\/]' -or $records.path -match '(?:^|[\\/])(?:CNxRoster|NxRoster|FNxRoster)') { throw 'Product manifest contains removed HWPX or roster paths' }
$seen = @{}
$expectedNameSeen = @{}
$expectedProductComponents = @()
foreach ($entry in $records) {
    $rel = ([string]$entry.path).Replace('\','/')
    if ($seen.ContainsKey($rel)) { throw ('Duplicate product module: ' + $rel) }
    $seen[$rel] = $true
    if ($rel -notmatch '^src/vba/[A-Za-z0-9_./-]+(?:\.form\.json|\.(?:bas|cls|vba))$') { throw ('Product module path rejected: ' + $rel) }
    $full = [IO.Path]::GetFullPath((Join-Path $root $rel))
    if (-not $full.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Product module escaped root' }
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw ('Product module missing: ' + $rel) }
    $actual = (Get-FileHash -LiteralPath $full -Algorithm SHA256).Hash.ToLowerInvariant()
    if ($actual -ne ([string]$entry.sha256).ToLowerInvariant()) { throw ('Product module hash mismatch: ' + $rel) }
    if ([string]$entry.kind -eq 'form_code') { continue }
    $expectedName = $null
    $expectedType = 0
    switch ([string]$entry.kind) {
        'standard_module' {
            $nameMatch = [regex]::Match([IO.File]::ReadAllText($full, [Text.Encoding]::UTF8), '(?m)^Attribute VB_Name = "([A-Za-z][A-Za-z0-9_]*)"\r?$')
            if (-not $nameMatch.Success) { throw ('Product standard module name missing: ' + $rel) }
            $expectedName = $nameMatch.Groups[1].Value; $expectedType = 1
        }
        'class_module' {
            $nameMatch = [regex]::Match([IO.File]::ReadAllText($full, [Text.Encoding]::UTF8), '(?m)^Attribute VB_Name = "([A-Za-z][A-Za-z0-9_]*)"\r?$')
            if (-not $nameMatch.Success) { throw ('Product class module name missing: ' + $rel) }
            $expectedName = $nameMatch.Groups[1].Value; $expectedType = 2
        }
        'form_layout' {
            $layoutIdentity = Get-Content -LiteralPath $full -Raw -Encoding UTF8 | ConvertFrom-Json
            $expectedName = [string]$layoutIdentity.name; $expectedType = 3
            if ($expectedName -notmatch '^[A-Za-z][A-Za-z0-9_]*$') { throw ('Product UserForm name rejected: ' + $rel) }
        }
        default { throw ('Product manifest component kind rejected: ' + [string]$entry.kind) }
    }
    if ($expectedNameSeen.ContainsKey($expectedName)) { throw ('Duplicate Product VBA component name: ' + $expectedName) }
    $expectedNameSeen[$expectedName] = $true
    $expectedProductComponents += [pscustomobject]@{name=$expectedName;type=$expectedType;path=$rel}
}
if ($expectedProductComponents.Count -ne ($records.Count - $formCodeCount)) { throw 'Product manifest component projection rejected' }
$bindingSeen = @{}
foreach ($binding in $bindings) {
    $bindingKey = ([string]$binding.layout) + '|' + ([string]$binding.code)
    if ($bindingSeen.ContainsKey($bindingKey)) { throw 'Duplicate Product form binding rejected' }
    $bindingSeen[$bindingKey] = $true
    $layout = @($records | Where-Object { $_.path -eq [string]$binding.layout -and $_.kind -eq 'form_layout' })
    $code = @($records | Where-Object { $_.path -eq [string]$binding.code -and $_.kind -eq 'form_code' })
    if ($layout.Count -ne 1 -or $code.Count -ne 1 -or [string]$layout[0].code_path -ne [string]$binding.code) { throw 'Product form binding rejected' }
}
foreach ($layout in @($records | Where-Object { $_.kind -eq 'form_layout' })) {
    if (@($bindings | Where-Object { $_.layout -eq [string]$layout.path -and $_.code -eq [string]$layout.code_path }).Count -ne 1) { throw 'Product form layout binding bijection rejected' }
}
foreach ($code in @($records | Where-Object { $_.kind -eq 'form_code' })) {
    if (@($bindings | Where-Object { $_.code -eq [string]$code.path }).Count -ne 1) { throw 'Product form code binding bijection rejected' }
}

function Release-ComObject([object]$Value) { if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } }
function Set-ComProperty([object]$Component, [string]$Name, [object]$Value) {
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
        throw ('Product UserForm component property failed: name=' + $Name + '; value_type=' + $valueType + '; cause=' + $_.Exception.Message)
    } finally {
        Release-ComObject $property
        Release-ComObject $properties
    }
}
function Get-ComPropertyDouble([object]$Component, [string]$Name) {
    $properties = $null; $property = $null
    try {
        $properties = $Component.Properties
        $property = $properties.Item($Name)
        return [Convert]::ToDouble($property.Value, [Globalization.CultureInfo]::InvariantCulture)
    } catch {
        throw ('Product UserForm component property read failed: name=' + $Name + '; cause=' + $_.Exception.Message)
    } finally {
        Release-ComObject $property
        Release-ComObject $properties
    }
}
function Set-UserFormDimensions([object]$Component, [object]$Layout) {
    $outerWidth = [double]$Layout.width
    $outerHeight = [double]$Layout.height
    $clientWidth = [double]$Layout.client_width
    $clientHeight = [double]$Layout.client_height
    Set-ComProperty $Component 'Width' $outerWidth
    Set-ComProperty $Component 'Height' $outerHeight
    $insideWidth = Get-ComPropertyDouble $Component 'InsideWidth'
    $insideHeight = Get-ComPropertyDouble $Component 'InsideHeight'
    $adjustedWidth = [Math]::Max($outerWidth, $outerWidth + $clientWidth - $insideWidth)
    $adjustedHeight = [Math]::Max($outerHeight, $outerHeight + $clientHeight - $insideHeight)
    Set-ComProperty $Component 'Width' ([Math]::Round($adjustedWidth, 2))
    Set-ComProperty $Component 'Height' ([Math]::Round($adjustedHeight, 2))
    if ((Get-ComPropertyDouble $Component 'InsideWidth') + 0.5 -lt $clientWidth -or (Get-ComPropertyDouble $Component 'InsideHeight') + 0.5 -lt $clientHeight) {
        throw 'Product UserForm client dimensions could not be satisfied'
    }
}
function Get-ProductVbaCodeBody([string]$Path, [string]$Kind, [string]$ExpectedName) {
    $raw = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    $normalized = $raw.Replace("`r`n", "`n").Replace("`r", "`n")
    $lines = @($normalized -split "`n")
    $bodyStart = 0
    $nameAttribute = 'Attribute VB_Name = "' + $ExpectedName + '"'
    if ($Kind -eq 'standard_module') {
        if ($lines.Count -lt 2 -or $lines[0] -cne $nameAttribute) { throw ('Product standard module envelope rejected: ' + $Path) }
        $bodyStart = 1
    } elseif ($Kind -eq 'class_module') {
        if ($lines.Count -lt 6 -or
            $lines[0] -cne 'VERSION 1.0 CLASS' -or
            $lines[1] -cne 'BEGIN' -or
            $lines[2] -cne "  MultiUse = -1  'True" -or
            $lines[3] -cne 'END' -or
            $lines[4] -cne $nameAttribute) { throw ('Product class module envelope rejected: ' + $Path) }
        $bodyStart = 5
    } else { throw ('Product VBA code-body kind rejected: ' + $Kind) }
    $bodyLines = New-Object System.Collections.Generic.List[string]
    for ($lineIndex = $bodyStart; $lineIndex -lt $lines.Count; $lineIndex++) {
        if ($lines[$lineIndex] -match '^Attribute\s') { throw ('Product VBA code body contains forbidden Attribute: ' + $Path) }
        [void]$bodyLines.Add([string]$lines[$lineIndex])
    }
    $body = [string]::Join("`r`n", $bodyLines)
    if ([string]::IsNullOrWhiteSpace($body)) { throw ('Product VBA code body empty: ' + $Path) }
    return $body
}

. (Join-Path $PSScriptRoot 'Excel-ProcessLifecycle.ps1')
$excel = $null; $books = $null; $book = $null; $project = $null; $components = $null; $references = $null; $vbe = $null; $bars = $null; $compileControl = $null; $compileTarget = $null; $compileWatcher = $null; $excelProcess = $null; $excelBinding = $null; $excelOwned = $false
$cleanupFailures = @()
try {
    $excelBaseline = Get-ExcelProcessBaseline
    $excel = New-Object -ComObject Excel.Application
    $excelBinding = Get-ExactExcelProcessOwnership $excel $excelBaseline 'Product VBA import Excel'
    $excelOwned = [bool]$excelBinding.owned
    if (-not $excelOwned) { throw ('Product VBA import Excel ownership rejected: ' + [string]$excelBinding.failure) }
    $excel.Visible = $false
    $excelProcess = $excelBinding.process
    $books = $excel.Workbooks
    $book = $books.Open($xlam, $false, $false)
    if ($null -eq $book -or -not [bool]$book.IsAddin -or [bool]$book.ReadOnly -or [IO.Path]::GetFullPath([string]$book.FullName) -ne $xlam) { throw 'Product VBA import workbook identity rejected' }
    try { $project = $book.VBProject } catch { throw ('Product VBOM access blocked by Trust Center or AccessVBOM policy: ' + $_.Exception.Message) }
    $components = $project.VBComponents
    # Rename only document components of this verified build-owned add-in.
    # ThisWorkbook is intrinsic VBA syntax; it is independent of CodeName.
    $documentComponent = $components.Item([string]$book.CodeName)
    try { $documentComponent.Name = 'NxProductWorkbook' }
    finally { Release-ComObject $documentComponent }
    $sheetOrdinal = 0
    foreach ($productSheet in $book.Worksheets) {
        $sheetOrdinal++
        $documentComponent = $components.Item([string]$productSheet.CodeName)
        try { $documentComponent.Name = ('NxProductSheet{0:D3}' -f $sheetOrdinal) }
        finally { Release-ComObject $documentComponent; Release-ComObject $productSheet }
    }
    foreach ($entry in $records) {
        $full = [IO.Path]::GetFullPath((Join-Path $root ([string]$entry.path)))
        if ([string]$entry.kind -eq 'form_code') { continue }
        $expectedImportedComponents = @($expectedProductComponents | Where-Object { $_.path -eq [string]$entry.path })
        if ($expectedImportedComponents.Count -ne 1) { throw ('Product VBA import identity missing: ' + [string]$entry.path) }
        $expectedImportedComponent = $expectedImportedComponents[0]
        $component = $null; $designer = $null; $controls = $null; $codeModule = $null
        try {
            if ([string]$entry.kind -eq 'form_layout') {
                $layout = Get-Content -LiteralPath $full -Raw -Encoding UTF8 | ConvertFrom-Json
                $codeEntry = @($records | Where-Object { $_.kind -eq 'form_code' -and $_.path -eq [string]$entry.code_path })
                if ($codeEntry.Count -ne 1) { throw ('Form layout/code binding rejected: ' + [string]$entry.path) }
                $codeFull = [IO.Path]::GetFullPath((Join-Path $root ([string]$codeEntry[0].path)))
                $component = $components.Add(3)
                $component.Name = [string]$layout.name
                Set-ComProperty $component 'Caption' ([string]$layout.caption)
                Set-UserFormDimensions $component $layout
                Set-ComProperty $component 'StartUpPosition' ([int]$layout.start_up_position)
                $designer = $component.Designer
                $controls = $designer.Controls
                foreach ($controlDef in @($layout.controls)) {
                    $progId = switch ([string]$controlDef.type) {
                        'Label' { 'Forms.Label.1' }; 'TextBox' { 'Forms.TextBox.1' }; 'CommandButton' { 'Forms.CommandButton.1' }; 'ComboBox' { 'Forms.ComboBox.1' }; 'ListBox' { 'Forms.ListBox.1' }; 'CheckBox' { 'Forms.CheckBox.1' }; default { throw ('Unsupported product UserForm control: ' + $controlDef.type) }
                    }
                    $control = $null
                    try {
                        $control = $controls.Add($progId, [string]$controlDef.name, $true)
                        if ($controlDef.PSObject.Properties.Name -contains 'integral_height') {
                            if ([string]$controlDef.type -ne 'ListBox' -or $controlDef.integral_height -isnot [bool]) { throw ('Product ListBox integral height contract rejected: ' + [string]$controlDef.name) }
                            # Disable row snapping before size/font assignments can alter the persisted height.
                            $control.IntegralHeight = [bool]$controlDef.integral_height
                        }
                        $control.Left = [double]$controlDef.left; $control.Top = [double]$controlDef.top; $control.Width = [double]$controlDef.width; $control.Height = [double]$controlDef.height
                        $tabIndex = [int]$controlDef.tab_index
                        $tabStop = [bool]$controlDef.tab_stop
                        if ($tabStop -and $tabIndex -lt 0) { throw ('Product UserForm tab index rejected: ' + [string]$controlDef.name) }
                        $control.TabStop = $tabStop; $control.Enabled = [bool]$controlDef.enabled
                        if ($controlDef.PSObject.Properties.Name -contains 'caption') { $control.Caption = [string]$controlDef.caption }
                        if ($controlDef.PSObject.Properties.Name -contains 'tag') {
                            if ($controlDef.tag -isnot [string] -or [string]$controlDef.type -in @('ComboBox','ListBox')) { throw ('Product UserForm tag contract rejected: ' + [string]$controlDef.name) }
                            $control.Tag = [string]$controlDef.tag
                        }
                        if ($controlDef.PSObject.Properties.Name -contains 'control_tip_text') {
                            if ($controlDef.control_tip_text -isnot [string]) { throw ('Product UserForm control tip contract rejected: ' + [string]$controlDef.name) }
                            $control.ControlTipText = [string]$controlDef.control_tip_text
                        }
                        if ($controlDef.PSObject.Properties.Name -contains 'fore_color') {
                            if ($controlDef.fore_color -isnot [ValueType]) { throw ('Product UserForm foreground color contract rejected: ' + [string]$controlDef.name) }
                            $control.ForeColor = [int]$controlDef.fore_color
                        }
                        if ($controlDef.PSObject.Properties.Name -contains 'back_color') {
                            if ($controlDef.back_color -isnot [ValueType]) { throw ('Product UserForm background color contract rejected: ' + [string]$controlDef.name) }
                            $control.BackColor = [int]$controlDef.back_color
                        }
                        if ($controlDef.PSObject.Properties.Name -contains 'font_name') {
                            if ($controlDef.font_name -isnot [string] -or [string]::IsNullOrWhiteSpace([string]$controlDef.font_name)) { throw ('Product UserForm font name contract rejected: ' + [string]$controlDef.name) }
                            $control.FontName = [string]$controlDef.font_name
                        }
                        if ($controlDef.PSObject.Properties.Name -contains 'font_size') {
                            if ($controlDef.font_size -isnot [ValueType] -or [double]$controlDef.font_size -le 0) { throw ('Product UserForm font size contract rejected: ' + [string]$controlDef.name) }
                            $control.FontSize = [double]$controlDef.font_size
                        }
                        if ($controlDef.PSObject.Properties.Name -contains 'font_bold') {
                            if ($controlDef.font_bold -isnot [bool]) { throw ('Product UserForm font bold contract rejected: ' + [string]$controlDef.name) }
                            $control.FontBold = [bool]$controlDef.font_bold
                        }
                        switch ([string]$controlDef.type) {
                            'TextBox' {
                                if ($controlDef.PSObject.Properties.Name -contains 'value') { $control.Value = [string]$controlDef.value }
                                elseif ($controlDef.PSObject.Properties.Name -contains 'text') { $control.Value = [string]$controlDef.text }
                                if ($controlDef.PSObject.Properties.Name -contains 'locked') { $control.Locked = [bool]$controlDef.locked }
                                if ($controlDef.PSObject.Properties.Name -contains 'multi_line') { $control.MultiLine = [bool]$controlDef.multi_line }
                                if ($controlDef.PSObject.Properties.Name -contains 'text_align') {
                                    $control.TextAlign = switch ([string]$controlDef.text_align) {
                                        'left' { 1 }; 'center' { 2 }; 'right' { 3 }; default { throw ('Product TextBox text alignment rejected: ' + [string]$controlDef.name) }
                                    }
                                }
                            }
                            'CheckBox' { $control.Value = [bool]$controlDef.value }
                            'ComboBox' {
                                if ([string]$controlDef.style -ne 'drop_down_list' -or -not [bool]$controlDef.match_required) { throw ('Product ComboBox contract rejected: ' + [string]$controlDef.name) }
                                $control.Style = 2
                                $control.MatchRequired = $true
                                $comboItems = @($controlDef.items | ForEach-Object { [string]$_ })
                                if (@($comboItems | Where-Object { $_.Contains('|') -or $_.Contains([string][char]31) }).Count -ne 0) { throw ('Product ComboBox metadata separator rejected: ' + [string]$controlDef.name) }
                                $control.Tag = 'nxcombo1|' + [string]([int]$controlDef.selected_index) + '|' + ($comboItems -join [char]31)
                            }
                            'ListBox' {
                                if ($controlDef.PSObject.Properties.Name -contains 'column_count') { $control.ColumnCount = [int]$controlDef.column_count }
                                if ($controlDef.PSObject.Properties.Name -contains 'column_widths') { $control.ColumnWidths = [string]$controlDef.column_widths }
                                if ($controlDef.PSObject.Properties.Name -contains 'multi_select') { $control.MultiSelect = if ([bool]$controlDef.multi_select) { 1 } else { 0 } }
                                if ($controlDef.PSObject.Properties.Name -contains 'items') {
                                    $listItems = @($controlDef.items | ForEach-Object { [string]$_ })
                                    if (@($listItems | Where-Object { $_.Contains('|') -or $_.Contains([string][char]31) }).Count -ne 0) { throw ('Product ListBox metadata separator rejected: ' + [string]$controlDef.name) }
                                    $control.Tag = 'nxlist1|' + [string]([int]$controlDef.selected_index) + '|' + ($listItems -join [char]31)
                                }
                            }
                        }
                        if ([string]$controlDef.type -ne 'TextBox' -and $controlDef.PSObject.Properties.Name -contains 'text_align') { throw ('Product TextBox text alignment rejected: ' + [string]$controlDef.name) }
                    } finally { Release-ComObject $control }
                }
                # Assign in descending order only after every control exists. MSForms moves
                # neighboring controls whenever TabIndex changes, so inline assignment is not stable.
                foreach ($tabDef in @($layout.controls | Where-Object { [bool]$_.tab_stop } | Sort-Object { [int]$_.tab_index } -Descending)) {
                    $tabControl = $null
                    try {
                        $tabControl = $controls.Item([string]$tabDef.name)
                        $tabControl.TabIndex = [int]$tabDef.tab_index
                    } finally { Release-ComObject $tabControl }
                }
                if ($null -ne $layout.default_control) {
                    $defaultControl = $null
                    try { $defaultControl = $controls.Item([string]$layout.default_control); $defaultControl.Default = $true } finally { Release-ComObject $defaultControl }
                }
                if ($null -ne $layout.cancel_control) {
                    $cancelControl = $null
                    try { $cancelControl = $controls.Item([string]$layout.cancel_control); $cancelControl.Cancel = $true } finally { Release-ComObject $cancelControl }
                }
                $codeModule = $component.CodeModule
                $codeModule.AddFromString([IO.File]::ReadAllText($codeFull, [Text.Encoding]::UTF8))
            } else {
                $moduleCode = Get-ProductVbaCodeBody $full ([string]$entry.kind) ([string]($expectedImportedComponent.name))
                $component = $components.Add([int]($expectedImportedComponent.type))
                $component.Name = [string]($expectedImportedComponent.name)
                $codeModule = $component.CodeModule
                $codeModule.AddFromString($moduleCode)
            }
            if ($null -eq $component) { throw ('VBA import returned no component: ' + $entry.path) }
            # Read back the manifest-bound component before inventory admission.
            $importedName = [string]($component.Name)
            $importedType = [int]($component.Type)
            if ($importedName -ine [string]($expectedImportedComponent.name) -or $importedType -ne [int]($expectedImportedComponent.type)) {
                throw ('Product VBA imported component identity mismatch: path=' + [string]$entry.path + '; expected=' + [string]($expectedImportedComponent.name) + ':' + [string]($expectedImportedComponent.type) + '; actual=' + $importedName + ':' + [string]$importedType)
            }
        } finally {
            Release-ComObject $codeModule
            Release-ComObject $controls
            Release-ComObject $designer
            Release-ComObject $component
        }
    }

    $actualProductComponents = @()
    for ($componentIndex = 1; $componentIndex -le [int]$components.Count; $componentIndex++) {
        $inventoryComponent = $null
        try {
            $inventoryComponent = $components.Item($componentIndex)
            $componentType = [int]$inventoryComponent.Type
            if ($componentType -in @(1,2,3)) {
                $actualProductComponents += [pscustomobject]@{name=[string]($inventoryComponent.Name);type=$componentType}
            }
        } finally { Release-ComObject $inventoryComponent }
    }
    if ($actualProductComponents.Count -ne $expectedProductComponents.Count) {
        throw ('Product-owned VBA component inventory mismatch: actual=' + [string]$actualProductComponents.Count + '; expected=' + [string]$expectedProductComponents.Count)
    }
    foreach ($expectedComponent in $expectedProductComponents) {
        # VBA identifiers are case-insensitive and VBE may normalize their casing
        # while importing an otherwise identical component.  The manifest's
        # case-insensitive uniqueness check keeps this comparison bijective.
        $matches = @($actualProductComponents | Where-Object { $_.name -ieq [string]$expectedComponent.name -and [int]$_.type -eq [int]$expectedComponent.type })
        if ($matches.Count -ne 1) {
            $prefixLength = [Math]::Min(6, ([string]$expectedComponent.name).Length)
            $prefix = ([string]$expectedComponent.name).Substring(0, $prefixLength)
            $candidates = @($actualProductComponents | Where-Object { [int]$_.type -eq [int]$expectedComponent.type -and $_.name -like ($prefix + '*') } | ForEach-Object { ([string]$_.name) + ':' + ([string]$_.type) })
            throw ('Product-owned VBA component inventory mismatch: expected=' + [string]$expectedComponent.name + ':' + [string]$expectedComponent.type + '; actual candidates=' + ($candidates -join ','))
        }
    }
    foreach ($actualComponent in $actualProductComponents) {
        $matches = @($expectedProductComponents | Where-Object { $_.name -ieq [string]$actualComponent.name -and [int]$_.type -eq [int]$actualComponent.type })
        if ($matches.Count -ne 1) { throw ('Unexpected importable Product VBA component: ' + [string]$actualComponent.name) }
    }

    $references = $project.References
    $referenceCount = [int]$references.Count
    $brokenReferenceCount = 0
    for ($referenceIndex = 1; $referenceIndex -le $referenceCount; $referenceIndex++) {
        $reference = $null
        try {
            $reference = $references.Item($referenceIndex)
            if ([bool]$reference.IsBroken) { $brokenReferenceCount++ }
        } finally { Release-ComObject $reference }
    }
    if ($brokenReferenceCount -ne 0) { throw ('Product VBA broken_reference_count=' + $brokenReferenceCount) }

    $vbe = $excel.VBE
    $bars = $vbe.CommandBars
    $compileControl = Find-VbeCompileControl $bars
    if ($null -eq $compileControl -or [int]$compileControl.Id -ne 578) { throw 'Product VBA compile control 578 unavailable' }
    $compileTarget = $components.Item('NxGeneratedProductRegistry')
    $compileTarget.Activate()
    $activeProjectPath = [IO.Path]::GetFullPath([string]$vbe.ActiveVBProject.FileName)
    if ($activeProjectPath -ne $xlam) { throw 'Product VBA compile active project identity mismatch' }
    if (-not [bool]$compileControl.Enabled) { throw 'Product VBA compile control was disabled before required build compile' }
    $compileWatcher = Start-ExactOwnedCompileDialogWatcher -ExcelPid ([int]$excelBinding.pid) -ExcelHwnd ([Int64]$excelBinding.hwnd) -VbeHwnd ([Int64]$vbe.MainWindow.HWnd)
    $compileExecuteError = $null
    try { $compileControl.Execute() } catch { $compileExecuteError = $_.Exception.Message }
    $watchCompleted = Wait-Job -Job $compileWatcher -Timeout 30
    if ($null -eq $watchCompleted) { Stop-Job -Job $compileWatcher -ErrorAction SilentlyContinue; throw 'Product VBA compile dialog watcher timed out' }
    $watchResult = @(Receive-Job -Job $compileWatcher -Wait -AutoRemoveJob -ErrorAction Stop | Select-Object -First 1)
    $compileWatcher = $null
    if (-not [string]::IsNullOrWhiteSpace([string]$compileExecuteError)) { throw ('Product VBA compile 578 execution failed: ' + $compileExecuteError) }
    if ($watchResult.Count -ne 1 -or [string]$watchResult[0].status -ne 'NO_DIALOG') {
        $dialogText = if ($watchResult.Count -eq 1) { [string]$watchResult[0].dialog_text } else { 'watcher returned no result' }
        $compilePane = $null; $compileModule = $null
        $compileModuleName = 'UNAVAILABLE'; [int]$compileLine = 0; [int]$compileColumn = 0; [int]$compileEndLine = 0; [int]$compileEndColumn = 0
        $compileSource = 'UNAVAILABLE'; $compilePaneError = $null
        try {
            $compilePane = $vbe.ActiveCodePane
            if ($null -eq $compilePane) { throw 'ActiveCodePane unavailable after compile dialog' }
            $compileModule = $compilePane.CodeModule
            $compileModuleName = [string]($compileModule.Name)
            $compilePane.GetSelection([ref]$compileLine, [ref]$compileColumn, [ref]$compileEndLine, [ref]$compileEndColumn)
            if ($compileLine -gt 0) { $compileSource = ([string]($compileModule.Lines($compileLine, 1))).Replace("`r", ' ').Replace("`n", ' ') }
        } catch { $compilePaneError = $_.Exception.Message }
        finally { Release-ComObject $compileModule; Release-ComObject $compilePane }
        $compileLocation = 'compile_module=' + $compileModuleName + '; compile_line=' + [string]$compileLine + '; compile_column=' + [string]$compileColumn + '; compile_source=' + $compileSource
        if (-not [string]::IsNullOrWhiteSpace([string]$compilePaneError)) { $compileLocation += '; compile_pane_error=' + $compilePaneError }
        throw ('Product VBA compile dialog detected: ' + $dialogText + "`n" + $compileLocation)
    }
    if ([bool]$compileControl.Enabled) { throw 'Product VBA compile did not reach saved disabled state' }
    $book.Save()
} finally {
    if ($null -ne $compileWatcher) { try { Stop-Job -Job $compileWatcher -ErrorAction SilentlyContinue; Remove-Job -Job $compileWatcher -Force -ErrorAction SilentlyContinue } catch { $cleanupFailures += $_.Exception.Message } }
    try { Release-ComObject $compileTarget } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $compileControl } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $bars } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $vbe } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $references } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $components } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $project } catch { $cleanupFailures += $_.Exception.Message }
    try { if ($book) { $book.Close($false) } } catch { $cleanupFailures += $_.Exception.Message }
    if ($excelOwned -and $null -ne $excel) {
        try { [void](Assert-ExactExcelProcessOwnership $excelBinding) } catch { $cleanupFailures += $_.Exception.Message; $excelOwned = $false }
    }
    try { if ($excelOwned -and $null -ne $excel) { $excel.Quit() } } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $book } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $books } catch { $cleanupFailures += $_.Exception.Message }
    try { Release-ComObject $excel } catch { $cleanupFailures += $_.Exception.Message }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($excelOwned -and $null -ne $excelProcess) {
        $cleanup = Stop-ExactProcessAfterGrace -Process $excelProcess -Label 'Product VBA import Excel' -GraceMs 10000 -Detailed
        if ($cleanup.failure) { $cleanupFailures += [string]$cleanup.failure }
    }
    if ($null -ne $excelProcess) { try { $excelProcess.Dispose() } catch { $cleanupFailures += $_.Exception.Message } }
    if ($cleanupFailures.Count -gt 0) { throw ('Product VBA import cleanup failed: ' + ($cleanupFailures -join ' | ')) }
}
Wait-ExactArtifactReadable $xlam 10000
