<#
.SYNOPSIS
Builds a HWPX table from JSON using only Windows PowerShell 5.1 and .NET built-ins.

.DESCRIPTION
CLI contract: -InputPath, -OutputPath, and -TemplatePath are required for generation.
Use -VerifyPath to verify an existing HWPX package. Exit codes: 0 success, 2 argument
error, 3 input JSON error, 4 template/package error, 6 generated output verification
error, 1 unexpected error.
#>
param(
    [string]$InputPath = "",
    [string]$OutputPath = "",
    [string]$TemplatePath = "",
    [string]$VerifyPath = ""
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"
$script:PictureLayout = $null

$script:HP = "http://www.hancom.co.kr/hwpml/2011/paragraph"
$script:HH = "http://www.hancom.co.kr/hwpml/2011/head"
$script:HC = "http://www.hancom.co.kr/hwpml/2011/core"
$script:TABLE_WIDTH = 47341
$script:DEFAULT_ROW_HEIGHT = 1865
$script:HEADER_ROW_HEIGHT = 2200
$script:MIN_BODY_ROW_HEIGHT = 1500
$script:CELL_MARGIN_HORIZONTAL = 340
$script:CELL_MARGIN_VERTICAL = 120
$script:MIN_COLUMN_SHARE = 0.08
$script:MAX_COLUMN_SHARE = 0.42
$script:INNER_BORDER_COLOR = "#808080"
$script:OUTER_BORDER_COLOR = "#000000"
$script:MAX_CELLS = 16384
$script:TABLE_FONT_SIZE_PT = 13
$script:TABLE_TITLE_ROW = $true
$script:TABLE_CHAR_PR_IDS = @("0", "2", "3", "10")
$script:WON_TEXT = ([char]0xC6D0).ToString()
$script:JUNGGOTHIC_TEXT = ([char]0xC911).ToString() + ([char]0xACE0).ToString() + ([char]0xB515).ToString()
$script:HANYANG_JUNGGOTHIC_TEXT = ([char]0xD55C).ToString() + ([char]0xC591).ToString() + $script:JUNGGOTHIC_TEXT
$script:HAMSOROM_DODUM_TEXT = ([char]0xD568).ToString() + ([char]0xCD08).ToString() + ([char]0xB86C).ToString() + ([char]0xB3CB).ToString() + ([char]0xC6C0).ToString()
$script:MALGUN_GOTHIC_TEXT = ([char]0xB9D1).ToString() + ([char]0xC740).ToString() + " " + ([char]0xACE0).ToString() + ([char]0xB515).ToString()
$script:GULIM_TEXT = ([char]0xAD74).ToString() + ([char]0xB9BC).ToString()
$script:FORMAL_STYLE_TEXT = ([char]0xACF5).ToString() + ([char]0xBB38).ToString() + ([char]0xD615).ToString()
$script:SIMPLE_STYLE_TEXT = ([char]0xAC04).ToString() + ([char]0xACB0).ToString() + ([char]0xD615).ToString()
$script:TABLE_FONT_NAME = $script:JUNGGOTHIC_TEXT
$script:TABLE_STYLE = $script:FORMAL_STYLE_TEXT
$script:ALLOWED_TABLE_FONTS = @($script:JUNGGOTHIC_TEXT, $script:HAMSOROM_DODUM_TEXT, $script:MALGUN_GOTHIC_TEXT, $script:GULIM_TEXT)
$script:JUNGGOTHIC_FONT_IDS = @{
    HANGUL = "3"; LATIN = "2"; HANJA = "3"; JAPANESE = "3"; OTHER = "2"; SYMBOL = "3"; USER = "2"
}
$script:JUNGGOTHIC_TYPE_INFO = @{
    familyType = "FCAT_GOTHIC"; weight = "0"; proportion = "0"; contrast = "0"; strokeVariation = "0";
    armStyle = "0"; letterform = "0"; midline = "0"; xHeight = "0"
}

function New-HwpxError([string]$Code, [string]$Message, [int]$ExitCode) {
    $exception = [System.Exception]::new($Message)
    $exception.Data["Code"] = $Code
    $exception.Data["ExitCode"] = $ExitCode
    return $exception
}

function Resolve-RequiredPath([string]$PathValue, [string]$Name) {
    if ([string]::IsNullOrWhiteSpace($PathValue)) {
        throw (New-HwpxError "ARGS_MISSING" "$Name is required." 2)
    }
    return [IO.Path]::GetFullPath($PathValue)
}

function New-XmlDocument {
    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    return $doc
}

function Load-XmlBytes([byte[]]$Bytes) {
    $doc = New-XmlDocument
    $text = [Text.Encoding]::UTF8.GetString($Bytes)
    $doc.LoadXml($text)
    return $doc
}

function New-NamespaceManager([System.Xml.XmlDocument]$Doc) {
    $ns = [System.Xml.XmlNamespaceManager]::new($Doc.NameTable)
    $ns.AddNamespace("hp", $script:HP)
    $ns.AddNamespace("hh", $script:HH)
    $ns.AddNamespace("hc", $script:HC)
    return ,$ns
}

function New-HpElement([System.Xml.XmlDocument]$Doc, [string]$LocalName, [hashtable]$Attributes) {
    $node = $Doc.CreateElement("hp", $LocalName, $script:HP)
    if ($null -ne $Attributes) {
        foreach ($key in $Attributes.Keys) {
            $node.SetAttribute([string]$key, [string]$Attributes[$key])
        }
    }
    return $node
}

function New-HhElement([System.Xml.XmlDocument]$Doc, [string]$LocalName, [hashtable]$Attributes) {
    $node = $Doc.CreateElement("hh", $LocalName, $script:HH)
    if ($null -ne $Attributes) {
        foreach ($key in $Attributes.Keys) {
            $node.SetAttribute([string]$key, [string]$Attributes[$key])
        }
    }
    return $node
}

function New-HcElement([System.Xml.XmlDocument]$Doc, [string]$LocalName, [hashtable]$Attributes) {
    $node = $Doc.CreateElement("hc", $LocalName, $script:HC)
    if ($null -ne $Attributes) {
        foreach ($key in $Attributes.Keys) {
            $node.SetAttribute([string]$key, [string]$Attributes[$key])
        }
    }
    return $node
}

function ConvertTo-XmlBytes([System.Xml.XmlDocument]$Doc) {
    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = [System.Text.UTF8Encoding]::new($false)
    $settings.OmitXmlDeclaration = $false
    $settings.Indent = $false
    $stream = New-Object System.IO.MemoryStream
    $writer = [System.Xml.XmlWriter]::Create($stream, $settings)
    try {
        $Doc.Save($writer)
        $writer.Flush()
        return $stream.ToArray()
    } finally {
        $writer.Close()
        $stream.Dispose()
    }
}

function Read-Payload([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw (New-HwpxError "INPUT_NOT_FOUND" "Input JSON was not found: $Path" 3)
    }
    try {
        $json = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
        if ($json.Length -gt 0 -and [int][char]$json[0] -eq 0xFEFF) {
            $json = $json.Substring(1)
        }
        $payload = $json | ConvertFrom-Json
    } catch {
        throw (New-HwpxError "INPUT_JSON_INVALID" "Input JSON could not be parsed: $($_.Exception.Message)" 3)
    }
    if ($payload.PSObject.Properties.Name -contains 'picture_layout') {
        $payload = Initialize-PictureLayout $payload
    }
    if (-not ($payload.PSObject.Properties.Name -contains "schema_version") -or [int]$payload.schema_version -ne 1) {
        throw (New-HwpxError "INPUT_SCHEMA_INVALID" "schema_version must be 1." 3)
    }
    foreach ($forbiddenName in @("source_path", "workbook_path", "python_path")) {
        if ($payload.PSObject.Properties.Name -contains $forbiddenName) {
            throw (New-HwpxError "INPUT_METADATA_REJECTED" "Source metadata is not accepted: $forbiddenName" 3)
        }
    }
    $rows = 0
    $cols = 0
    if ($payload.PSObject.Properties.Name -contains "rows") { $rows = [int]$payload.rows }
    elseif ($payload.PSObject.Properties.Name -contains "rowCount") { $rows = [int]$payload.rowCount }
    if ($payload.PSObject.Properties.Name -contains "columns" -and $payload.columns -isnot [System.Array]) { $cols = [int]$payload.columns }
    elseif ($payload.PSObject.Properties.Name -contains "cols") { $cols = [int]$payload.cols }
    elseif ($payload.PSObject.Properties.Name -contains "columnCount") { $cols = [int]$payload.columnCount }
    if ($rows -lt 1 -or $cols -lt 1) {
        throw (New-HwpxError "INPUT_DIMENSIONS_INVALID" "Rows and columns must both be at least 1." 3)
    }
    if (($rows * $cols) -gt $script:MAX_CELLS) {
        throw (New-HwpxError "INPUT_TOO_LARGE" "Table cell count must not exceed $($script:MAX_CELLS)." 3)
    }
    if ($payload.PSObject.Properties.Name -contains "settings" -and $null -ne $payload.settings) {
        $settings = $payload.settings
        if ($settings.PSObject.Properties.Name -contains "font_name") {
            $fontName = [string]$settings.font_name
            if ($script:ALLOWED_TABLE_FONTS -notcontains $fontName) {
                throw (New-HwpxError "SETTINGS_FONT_INVALID" "The selected Korean font is not approved." 3)
            }
            $script:TABLE_FONT_NAME = $fontName
        }
        if ($settings.PSObject.Properties.Name -contains "font_type" -and [string]$settings.font_type -ne "HFT") {
            throw (New-HwpxError "SETTINGS_FONT_TYPE_INVALID" "Only HFT font type is accepted." 3)
        }
        if ($settings.PSObject.Properties.Name -contains "font_size_pt") {
            $fontSize = [int]$settings.font_size_pt
            if ($fontSize -lt 9 -or $fontSize -gt 24) { throw (New-HwpxError "SETTINGS_FONT_SIZE_INVALID" "Font size must be between 9 and 24pt." 3) }
            $script:TABLE_FONT_SIZE_PT = $fontSize
        }
        if ($settings.PSObject.Properties.Name -contains "title_row") { $script:TABLE_TITLE_ROW = [bool]$settings.title_row }
        if ($settings.PSObject.Properties.Name -contains "table_style") {
            $tableStyle = [string]$settings.table_style
            if ($tableStyle -ne $script:FORMAL_STYLE_TEXT -and $tableStyle -ne $script:SIMPLE_STYLE_TEXT) {
                throw (New-HwpxError "SETTINGS_TABLE_STYLE_INVALID" "The selected table style is not approved." 3)
            }
            $script:TABLE_STYLE = $tableStyle
            if ($tableStyle -eq $script:SIMPLE_STYLE_TEXT) {
                $script:INNER_BORDER_COLOR = "#C0C0C0"
                $script:OUTER_BORDER_COLOR = "#808080"
                $script:HEADER_ROW_HEIGHT = 2000
            }
        }
    }
    return [pscustomobject]@{ Raw = $payload; Rows = $rows; Cols = $cols }
}

function Get-JsonArray($Object, [string[]]$Names) {
    foreach ($name in $Names) {
        if ($Object.PSObject.Properties.Name -contains $name) {
            $value = $Object.$name
            if ($null -eq $value) { return @() }
            return @($value)
        }
    }
    return @()
}

function Normalize-Widths($PayloadInfo) {
    $cols = [int]$PayloadInfo.Cols
    $rawWidths = New-Object System.Collections.Generic.List[double]
    $items = Get-JsonArray $PayloadInfo.Raw @("column_widths", "columns")
    foreach ($item in $items) {
        $width = 1.0
        try {
            if ($item -is [ValueType]) {
                $width = [Math]::Max([double]$item, 0.1)
            } elseif ($null -ne $item -and $item.PSObject.Properties.Name -contains "width") {
                $width = [Math]::Max([double]$item.width, 0.1)
            }
        } catch { $width = 1.0 }
        $rawWidths.Add($width)
    }
    while ($rawWidths.Count -lt $cols) { $rawWidths.Add(1.0) }
    while ($rawWidths.Count -gt $cols) { $rawWidths.RemoveAt($rawWidths.Count - 1) }
    $total = 0.0
    foreach ($width in $rawWidths) { $total += $width }
    if ($total -le 0) { $total = [double]$cols }
    $minShare = [Math]::Min($script:MIN_COLUMN_SHARE, 0.7 / [double]$cols)
    $maxShare = [Math]::Max(0.18, [Math]::Min($script:MAX_COLUMN_SHARE, 1.0 - (($cols - 1) * $minShare)))
    $shares = New-Object System.Collections.Generic.List[double]
    foreach ($width in $rawWidths) {
        $shares.Add([Math]::Min([Math]::Max(($width / $total), $minShare), $maxShare))
    }
    $adjustedTotal = 0.0
    foreach ($share in $shares) { $adjustedTotal += $share }
    if ($adjustedTotal -le 0) { $adjustedTotal = 1.0 }
    $widths = New-Object System.Collections.Generic.List[int]
    $sum = 0
    foreach ($share in $shares) {
        $value = [Math]::Max(1, [int][Math]::Round($script:TABLE_WIDTH * $share / $adjustedTotal))
        $widths.Add($value)
        $sum += $value
    }
    $widths[$widths.Count - 1] = $widths[$widths.Count - 1] + ($script:TABLE_WIDTH - $sum)
    return @($widths)
}

function Normalize-Heights($PayloadInfo) {
    $rows = [int]$PayloadInfo.Rows
    $items = @(Get-JsonArray $PayloadInfo.Raw @("row_heights", "rowHeights", "rowsMeta"))
    $heights = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $items.Count; $i++) {
        $height = if ($i -eq 0 -and $script:TABLE_TITLE_ROW) { $script:HEADER_ROW_HEIGHT } else { $script:DEFAULT_ROW_HEIGHT }
        try {
            if ($items[$i] -is [ValueType]) {
                $pointHeight = [Math]::Max([double]$items[$i], 8.0)
                $minimum = if ($i -eq 0 -and $script:TABLE_TITLE_ROW) { $script:HEADER_ROW_HEIGHT } else { $script:MIN_BODY_ROW_HEIGHT }
                $height = [Math]::Max($minimum, [int][Math]::Round($pointHeight * 100.0))
            } elseif ($null -ne $items[$i] -and $items[$i].PSObject.Properties.Name -contains "height") {
                $pointHeight = [Math]::Max([double]$items[$i].height, 8.0)
                $minimum = if ($i -eq 0) { $script:HEADER_ROW_HEIGHT } else { $script:MIN_BODY_ROW_HEIGHT }
                $height = [Math]::Max($minimum, [int][Math]::Round($pointHeight * 100.0))
            }
        } catch {}
        $heights.Add($height)
    }
    while ($heights.Count -lt $rows) {
        $heights.Add($(if ($heights.Count -eq 0 -and $script:TABLE_TITLE_ROW) { $script:HEADER_ROW_HEIGHT } else { $script:DEFAULT_ROW_HEIGHT }))
    }
    while ($heights.Count -gt $rows) { $heights.RemoveAt($heights.Count - 1) }
    return @($heights)
}

function Get-CellValues($PayloadInfo) {
    $values = @{}
    foreach ($cell in (Get-JsonArray $PayloadInfo.Raw @("cells"))) {
        try {
            $row = if ($cell.PSObject.Properties.Name -contains "row") { [int]$cell.row - 1 } else { [int]$cell.r - 1 }
            $col = if ($cell.PSObject.Properties.Name -contains "column") { [int]$cell.column - 1 } else { [int]$cell.c - 1 }
            if ($row -lt 0 -or $col -lt 0 -or $row -ge $PayloadInfo.Rows -or $col -ge $PayloadInfo.Cols) {
                throw (New-HwpxError "INPUT_CELL_RANGE" "Cell coordinate is outside the declared table." 3)
            }
            if ($values.ContainsKey("$row,$col")) {
                throw (New-HwpxError "INPUT_CELL_DUPLICATE" "Duplicate cell coordinate: $($row + 1),$($col + 1)" 3)
            }
            $text = ""
            if ($cell.PSObject.Properties.Name -contains "text") { $text = $cell.text }
            elseif ($cell.PSObject.Properties.Name -contains "v") { $text = $cell.v }
            if ($null -eq $text) { $text = "" }
            # Strip display padding only; preserve internal spacing and line breaks.
            $values["$row,$col"] = ([string]$text).Trim([char[]]@(32, 9, 160, 12288))
        } catch {
            if ($_.Exception.Data.Contains("ExitCode")) { throw }
            throw (New-HwpxError "INPUT_CELL_INVALID" "Cell payload is invalid: $($_.Exception.Message)" 3)
        }
    }
    return $values
}

function Get-MergeMaps($PayloadInfo) {
    $starts = @{}
    $covered = @{}
    $occupied = @{}
    foreach ($merge in (Get-JsonArray $PayloadInfo.Raw @("merges"))) {
        try {
            $row = [int]$merge.row - 1
            $col = [int]$merge.column - 1
            $rowSpan = [int]$merge.row_span
            $colSpan = [int]$merge.column_span
            if ($row -lt 0 -or $col -lt 0 -or $rowSpan -lt 1 -or $colSpan -lt 1) {
                throw (New-HwpxError "INPUT_MERGE_INVALID" "Merge coordinates and spans must be positive." 3)
            }
            if (($row + $rowSpan) -gt $PayloadInfo.Rows -or ($col + $colSpan) -gt $PayloadInfo.Cols) {
                throw (New-HwpxError "INPUT_MERGE_RANGE" "Merge is outside the declared table." 3)
            }
            for ($mergeRow = $row; $mergeRow -lt ($row + $rowSpan); $mergeRow++) {
                for ($mergeCol = $col; $mergeCol -lt ($col + $colSpan); $mergeCol++) {
                    $key = "$mergeRow,$mergeCol"
                    if ($occupied.ContainsKey($key)) {
                        throw (New-HwpxError "INPUT_MERGE_OVERLAP" "Overlapping merges are not accepted." 3)
                    }
                    $occupied[$key] = $true
                    if ($mergeRow -ne $row -or $mergeCol -ne $col) { $covered[$key] = $true }
                }
            }
            $starts["$row,$col"] = [pscustomobject]@{ RowSpan = $rowSpan; ColSpan = $colSpan }
        } catch {
            if ($_.Exception.Data.Contains("ExitCode")) { throw }
            throw (New-HwpxError "INPUT_MERGE_INVALID" "Merge payload is invalid: $($_.Exception.Message)" 3)
        }
    }
    return [pscustomobject]@{ Starts = $starts; Covered = $covered }
}

function Test-NumericText([string]$Value) {
    $text = $Value.Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    $text = $text.Replace(",", "").Replace(" ", "")
    if ($text.StartsWith("(") -and $text.EndsWith(")")) {
        $text = "-" + $text.Substring(1, $text.Length - 2)
    }
    foreach ($prefix in @([string][char]0x20A9, "\")) {
        if ($text.StartsWith($prefix)) { $text = $text.Substring($prefix.Length) }
    }
    foreach ($suffix in @($script:WON_TEXT, "%")) {
        if ($text.EndsWith($suffix)) { $text = $text.Substring(0, $text.Length - $suffix.Length) }
    }
    $number = 0.0
    return [double]::TryParse($text, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)
}

function Get-ParagraphStyle([int]$Row, [string]$Value) {
    if ($script:TABLE_TITLE_ROW -and $Row -eq 0) { return [pscustomobject]@{ Para = "11"; Char = "10" } }
    if (Test-NumericText $Value) { return [pscustomobject]@{ Para = "15"; Char = "2" } }
    return [pscustomobject]@{ Para = "11"; Char = "2" }
}

function Get-BorderFillId([int]$Row, [int]$Col, [int]$Rows, [int]$Cols) {
    if ($Cols -eq 1) {
        if ($script:TABLE_TITLE_ROW -and $Row -eq 0) { return 17 }
        if ($Row -eq ($Rows - 1)) { return 19 }
        return 18
    }
    $firstCol = $Col -eq 0
    $lastCol = $Col -eq ($Cols - 1)
    if ($script:TABLE_TITLE_ROW -and $Row -eq 0) {
        if ($firstCol) { return 7 }
        if ($lastCol) { return 9 }
        if ($Col -eq 1) { return 8 }
        return 10
    }
    if ($Row -eq ($Rows - 1)) {
        if ($firstCol) { return 5 }
        if ($lastCol) { return 15 }
        if ($Col -eq 1) { return 14 }
        return 16
    }
    if ($firstCol) { return 6 }
    if ($lastCol) { return 12 }
    if ($Col -eq 1) { return 11 }
    return 13
}

function Set-TextNode([System.Xml.XmlNode]$Node, [string]$Value) {
    while ($Node.HasChildNodes) { [void]$Node.RemoveChild($Node.FirstChild) }
    [void]$Node.AppendChild($Node.OwnerDocument.CreateTextNode($Value))
}

function New-TemplateClone([System.Xml.XmlDocument]$Doc, [System.Xml.XmlNode]$TemplateNode, [string]$FallbackLocalName) {
    if ($null -ne $TemplateNode) {
        return $Doc.ImportNode($TemplateNode, $true)
    }
    return New-HpElement $Doc $FallbackLocalName @{}
}

function New-CellFromTemplate([System.Xml.XmlDocument]$Doc, [System.Xml.XmlNode]$TemplateCell, [int]$Row, [int]$Col, [int]$Rows, [int]$Cols, [int]$Width, [int]$Height, [int]$RowSpan, [int]$ColSpan, [string]$Value, [int]$ParagraphId) {
    $ns = New-NamespaceManager $Doc
    $tc = New-TemplateClone $Doc $TemplateCell "tc"
    foreach ($pair in @{
        name = ""; header = "0"; hasMargin = "0"; protect = "0"; editable = "0"; dirty = "0";
        borderFillIDRef = [string](Get-BorderFillId $Row $Col $Rows $Cols)
    }.GetEnumerator()) {
        $tc.SetAttribute([string]$pair.Key, [string]$pair.Value)
    }

    $subList = $tc.SelectSingleNode("hp:subList", $ns)
    if ($null -eq $subList) {
        $subList = New-HpElement $Doc "subList" @{
            id = ""; textDirection = "HORIZONTAL"; lineWrap = "BREAK"; vertAlign = "CENTER";
            linkListIDRef = "0"; linkListNextIDRef = "0"; textWidth = "0"; textHeight = "0";
            hasTextRef = "0"; hasNumRef = "0"
        }
        [void]$tc.PrependChild($subList)
    }
    foreach ($pair in @{
        id = ""; textDirection = "HORIZONTAL"; lineWrap = "BREAK"; vertAlign = "CENTER";
        linkListIDRef = "0"; linkListNextIDRef = "0"; textWidth = "0"; textHeight = "0";
        hasTextRef = "0"; hasNumRef = "0"
    }.GetEnumerator()) {
        $subList.SetAttribute([string]$pair.Key, [string]$pair.Value)
    }

    foreach ($oldPara in @($subList.SelectNodes("hp:p", $ns))) {
        [void]$subList.RemoveChild($oldPara)
    }
    $style = Get-ParagraphStyle $Row $Value
    $p = New-HpElement $Doc "p" @{
        id = [string]$ParagraphId; paraPrIDRef = $style.Para; styleIDRef = "0";
        pageBreak = "0"; columnBreak = "0"; merged = "0"
    }
    $run = New-HpElement $Doc "run" @{ charPrIDRef = $style.Char }
    $t = New-HpElement $Doc "t" @{}
    Set-TextNode $t $Value
    [void]$run.AppendChild($t)
    [void]$p.AppendChild($run)
    [void]$subList.AppendChild($p)

    $cellAddr = $tc.SelectSingleNode("hp:cellAddr", $ns)
    if ($null -eq $cellAddr) { $cellAddr = $tc.AppendChild((New-HpElement $Doc "cellAddr" @{})) }
    $cellAddr.SetAttribute("colAddr", [string]$Col)
    $cellAddr.SetAttribute("rowAddr", [string]$Row)

    $cellSpan = $tc.SelectSingleNode("hp:cellSpan", $ns)
    if ($null -eq $cellSpan) { $cellSpan = $tc.AppendChild((New-HpElement $Doc "cellSpan" @{})) }
    $cellSpan.SetAttribute("colSpan", [string]$ColSpan)
    $cellSpan.SetAttribute("rowSpan", [string]$RowSpan)

    $cellSz = $tc.SelectSingleNode("hp:cellSz", $ns)
    if ($null -eq $cellSz) { $cellSz = $tc.AppendChild((New-HpElement $Doc "cellSz" @{})) }
    $cellSz.SetAttribute("width", [string]$Width)
    $cellSz.SetAttribute("height", [string]$Height)

    $cellMargin = $tc.SelectSingleNode("hp:cellMargin", $ns)
    if ($null -eq $cellMargin) { $cellMargin = $tc.AppendChild((New-HpElement $Doc "cellMargin" @{})) }
    $cellMargin.SetAttribute("left", [string]$script:CELL_MARGIN_HORIZONTAL)
    $cellMargin.SetAttribute("right", [string]$script:CELL_MARGIN_HORIZONTAL)
    $cellMargin.SetAttribute("top", [string]$script:CELL_MARGIN_VERTICAL)
    $cellMargin.SetAttribute("bottom", [string]$script:CELL_MARGIN_VERTICAL)
    return $tc
}

function New-EmptyParagraph([System.Xml.XmlDocument]$Doc, [int]$ParagraphId) {
    $p = New-HpElement $Doc "p" @{
        id = [string]$ParagraphId; paraPrIDRef = "14"; styleIDRef = "0";
        pageBreak = "0"; columnBreak = "0"; merged = "0"
    }
    $run = New-HpElement $Doc "run" @{ charPrIDRef = "0" }
    [void]$run.AppendChild((New-HpElement $Doc "t" @{}))
    [void]$p.AppendChild($run)
    return $p
}

function Build-SectionXml($PayloadInfo, [byte[]]$TemplateSectionBytes) {
    $oldDoc = Load-XmlBytes $TemplateSectionBytes
    $oldNs = New-NamespaceManager $oldDoc
    $oldNs.AddNamespace("hs", "http://www.hancom.co.kr/hwpml/2011/section")
    $doc = Load-XmlBytes $TemplateSectionBytes
    $ns = New-NamespaceManager $doc
    $section = $doc.DocumentElement
    while ($section.HasChildNodes) { [void]$section.RemoveChild($section.FirstChild) }

    $secPara = $oldDoc.SelectSingleNode("/hs:sec/hp:p[.//hp:secPr]", $oldNs)
    if ($null -eq $secPara) {
        throw (New-HwpxError "TEMPLATE_SECTION_INVALID" "Template Contents/section0.xml does not contain a section properties paragraph." 4)
    }
    $importedSecPara = $doc.ImportNode($secPara, $true)
    foreach ($textNode in @($importedSecPara.SelectNodes(".//hp:t", $ns))) {
        if ($null -ne $textNode.InnerText -and $textNode.InnerText.Length -gt 0) {
            Set-TextNode $textNode ($textNode.InnerText.Replace("12pt", "13pt"))
        }
    }
    [void]$section.AppendChild($importedSecPara)

    $templateTbl = $oldDoc.SelectSingleNode("//hp:tbl", $oldNs)
    if ($null -eq $templateTbl) {
        throw (New-HwpxError "TEMPLATE_TABLE_MISSING" "Template Contents/section0.xml does not contain a table to clone." 4)
    }
    $templateRun = $templateTbl.ParentNode
    $templateParagraph = $templateRun.ParentNode
    $templateRow = $templateTbl.SelectSingleNode("hp:tr", $oldNs)
    $templateCell = $templateTbl.SelectSingleNode("hp:tr/hp:tc", $oldNs)
    if ($null -eq $templateCell) {
        throw (New-HwpxError "TEMPLATE_CELL_MISSING" "Template table does not contain a cell to clone." 4)
    }

    $rows = [int]$PayloadInfo.Rows
    $cols = [int]$PayloadInfo.Cols
    $widths = Normalize-Widths $PayloadInfo
    $heights = Normalize-Heights $PayloadInfo
    $values = Get-CellValues $PayloadInfo
    $mergeMaps = Get-MergeMaps $PayloadInfo
    $tableHeight = 0
    foreach ($height in $heights) { $tableHeight += [int]$height }

    $p = New-TemplateClone $doc $templateParagraph "p"
    $p.SetAttribute("id", "1000000002")
    $p.SetAttribute("paraPrIDRef", "0")
    $p.SetAttribute("styleIDRef", "0")
    $p.SetAttribute("pageBreak", "0")
    $p.SetAttribute("columnBreak", "0")
    $p.SetAttribute("merged", "0")
    while ($p.HasChildNodes) { [void]$p.RemoveChild($p.FirstChild) }
    $run = New-TemplateClone $doc $templateRun "run"
    $run.SetAttribute("charPrIDRef", "0")
    while ($run.HasChildNodes) { [void]$run.RemoveChild($run.FirstChild) }
    $tbl = $doc.ImportNode($templateTbl, $true)
    foreach ($oldRow in @($tbl.SelectNodes("hp:tr", $ns))) {
        [void]$tbl.RemoveChild($oldRow)
    }
    foreach ($pair in @{
        id = "1000000999"; zOrder = "6"; numberingType = "TABLE"; textWrap = "TOP_AND_BOTTOM";
        textFlow = "BOTH_SIDES"; lock = "0"; dropcapstyle = "None"; pageBreak = "CELL";
        repeatHeader = "1"; rowCnt = [string]$rows; colCnt = [string]$cols; cellSpacing = "0";
        borderFillIDRef = "3"; noAdjust = "0"
    }.GetEnumerator()) {
        $tbl.SetAttribute([string]$pair.Key, [string]$pair.Value)
    }

    $sz = $tbl.SelectSingleNode("hp:sz", $ns)
    if ($null -eq $sz) { $sz = $tbl.PrependChild((New-HpElement $doc "sz" @{})) }
    $sz.SetAttribute("width", [string]$script:TABLE_WIDTH)
    $sz.SetAttribute("widthRelTo", "ABSOLUTE")
    $sz.SetAttribute("height", [string]$tableHeight)
    $sz.SetAttribute("heightRelTo", "ABSOLUTE")
    $sz.SetAttribute("protect", "0")

    $pos = $tbl.SelectSingleNode("hp:pos", $ns)
    if ($null -eq $pos) { [void]$tbl.InsertAfter((New-HpElement $doc "pos" @{}), $sz); $pos = $tbl.SelectSingleNode("hp:pos", $ns) }
    foreach ($pair in @{
        treatAsChar = "1"; affectLSpacing = "0"; flowWithText = "1"; allowOverlap = "0"; holdAnchorAndSO = "0";
        vertRelTo = "PARA"; horzRelTo = "COLUMN"; vertAlign = "TOP"; horzAlign = "LEFT"; vertOffset = "0"; horzOffset = "0"
    }.GetEnumerator()) { $pos.SetAttribute([string]$pair.Key, [string]$pair.Value) }

    $outMargin = $tbl.SelectSingleNode("hp:outMargin", $ns)
    if ($null -eq $outMargin) { [void]$tbl.InsertAfter((New-HpElement $doc "outMargin" @{}), $pos); $outMargin = $tbl.SelectSingleNode("hp:outMargin", $ns) }
    foreach ($name in @("left", "right", "top", "bottom")) { $outMargin.SetAttribute($name, "0") }
    $inMargin = $tbl.SelectSingleNode("hp:inMargin", $ns)
    if ($null -eq $inMargin) { [void]$tbl.InsertAfter((New-HpElement $doc "inMargin" @{}), $outMargin); $inMargin = $tbl.SelectSingleNode("hp:inMargin", $ns) }
    $inMargin.SetAttribute("left", [string]$script:CELL_MARGIN_HORIZONTAL)
    $inMargin.SetAttribute("right", [string]$script:CELL_MARGIN_HORIZONTAL)
    $inMargin.SetAttribute("top", [string]$script:CELL_MARGIN_VERTICAL)
    $inMargin.SetAttribute("bottom", [string]$script:CELL_MARGIN_VERTICAL)

    $paragraphId = 1000000100
    for ($row = 0; $row -lt $rows; $row++) {
        $tr = New-TemplateClone $doc $templateRow "tr"
        while ($tr.HasChildNodes) { [void]$tr.RemoveChild($tr.FirstChild) }
        for ($col = 0; $col -lt $cols; $col++) {
            $key = "$row,$col"
            if ($mergeMaps.Covered.ContainsKey($key)) { continue }
            $rowSpan = 1
            $colSpan = 1
            if ($mergeMaps.Starts.ContainsKey($key)) {
                $rowSpan = [int]$mergeMaps.Starts[$key].RowSpan
                $colSpan = [int]$mergeMaps.Starts[$key].ColSpan
            }
            $cellWidth = 0
            for ($widthIndex = $col; $widthIndex -lt ($col + $colSpan); $widthIndex++) { $cellWidth += [int]$widths[$widthIndex] }
            $cellHeight = 0
            for ($heightIndex = $row; $heightIndex -lt ($row + $rowSpan); $heightIndex++) { $cellHeight += [int]$heights[$heightIndex] }
            $value = if ($values.ContainsKey($key)) { [string]$values[$key] } else { "" }
            [void]$tr.AppendChild((New-CellFromTemplate $doc $templateCell $row $col $rows $cols $cellWidth $cellHeight $rowSpan $colSpan $value $paragraphId))
            $paragraphId++
        }
        [void]$tbl.AppendChild($tr)
    }
    [void]$run.AppendChild($tbl)
    [void]$p.AppendChild($run)
    [void]$section.AppendChild($p)
    [void]$section.AppendChild((New-EmptyParagraph $doc 1000000003))
    return ConvertTo-XmlBytes $doc
}

function Ensure-BorderFill([System.Xml.XmlDocument]$Doc, [System.Xml.XmlNode]$BorderFills, [string]$Id, [string]$Fill) {
    $ns = New-NamespaceManager $Doc
    $existing = $BorderFills.SelectSingleNode("hh:borderFill[@id='$Id']", $ns)
    if ($null -ne $existing) { return $existing }
    $borderFill = New-HhElement $Doc "borderFill" @{ id = $Id; threeD = "0"; shadow = "0"; centerLine = "NONE"; breakCellSeparateLine = "0" }
    [void]$borderFill.AppendChild((New-HhElement $Doc "slash" @{ type = "NONE"; Crooked = "0"; isCounter = "0" }))
    [void]$borderFill.AppendChild((New-HhElement $Doc "backSlash" @{ type = "NONE"; Crooked = "0"; isCounter = "0" }))
    foreach ($name in @("leftBorder", "rightBorder", "topBorder", "bottomBorder", "diagonal")) {
        [void]$borderFill.AppendChild((New-HhElement $Doc $name @{ type = "NONE"; width = "0.1 mm"; color = $script:OUTER_BORDER_COLOR }))
    }
    if (-not [string]::IsNullOrWhiteSpace($Fill)) {
        $fillBrush = New-HcElement $Doc "fillBrush" @{}
        [void]$fillBrush.AppendChild((New-HcElement $Doc "winBrush" @{ faceColor = $Fill; hatchColor = "#000000"; alpha = "0" }))
        [void]$borderFill.AppendChild($fillBrush)
    }
    [void]$BorderFills.AppendChild($borderFill)
    return $borderFill
}

function Set-BorderSide([System.Xml.XmlDocument]$Doc, [System.Xml.XmlNode]$BorderFill, [string]$SideName, [string]$LineType, [string]$Width, [string]$Color) {
    $ns = New-NamespaceManager $Doc
    $side = $BorderFill.SelectSingleNode("hh:$SideName", $ns)
    if ($null -eq $side) {
        $side = New-HhElement $Doc $SideName @{}
        [void]$BorderFill.AppendChild($side)
    }
    $side.SetAttribute("type", $LineType)
    $side.SetAttribute("width", $Width)
    $side.SetAttribute("color", $Color)
}

function Patch-HeaderXml([byte[]]$Bytes) {
    $doc = Load-XmlBytes $Bytes
    $ns = New-NamespaceManager $doc
    foreach ($fontface in @($doc.SelectNodes("//hh:fontface", $ns))) {
        $lang = $fontface.GetAttribute("lang")
        if (-not $script:JUNGGOTHIC_FONT_IDS.ContainsKey($lang)) { continue }
        $targetId = $script:JUNGGOTHIC_FONT_IDS[$lang]
        $font = $fontface.SelectSingleNode("hh:font[@id='$targetId']", $ns)
        if ($null -eq $font) {
            $font = New-HhElement $doc "font" @{ id = $targetId; isEmbedded = "0" }
            [void]$fontface.AppendChild($font)
        }
        $font.SetAttribute("face", $script:TABLE_FONT_NAME)
        $font.SetAttribute("type", "HFT")
        $font.SetAttribute("isEmbedded", "0")
        $typeInfo = $font.SelectSingleNode("hh:typeInfo", $ns)
        if ($null -eq $typeInfo) {
            $typeInfo = New-HhElement $doc "typeInfo" @{}
            [void]$font.AppendChild($typeInfo)
        }
        foreach ($key in $script:JUNGGOTHIC_TYPE_INFO.Keys) {
            $typeInfo.SetAttribute([string]$key, [string]$script:JUNGGOTHIC_TYPE_INFO[$key])
        }
    }
    foreach ($font in @($doc.SelectNodes("//hh:font", $ns))) {
        if ($font.GetAttribute("face") -eq $script:HANYANG_JUNGGOTHIC_TEXT) {
            $font.SetAttribute("face", $script:TABLE_FONT_NAME)
            $font.SetAttribute("type", "HFT")
        }
    }
    foreach ($charPr in @($doc.SelectNodes("//hh:charPr", $ns))) {
        if ($script:TABLE_CHAR_PR_IDS -notcontains $charPr.GetAttribute("id")) { continue }
        $charPr.SetAttribute("height", [string]($script:TABLE_FONT_SIZE_PT * 100))
        $fontRef = $charPr.SelectSingleNode("hh:fontRef", $ns)
        if ($null -ne $fontRef) {
            foreach ($pair in @{
                hangul = $script:JUNGGOTHIC_FONT_IDS["HANGUL"]; latin = $script:JUNGGOTHIC_FONT_IDS["LATIN"];
                hanja = $script:JUNGGOTHIC_FONT_IDS["HANJA"]; japanese = $script:JUNGGOTHIC_FONT_IDS["JAPANESE"];
                other = $script:JUNGGOTHIC_FONT_IDS["OTHER"]; symbol = $script:JUNGGOTHIC_FONT_IDS["SYMBOL"];
                user = $script:JUNGGOTHIC_FONT_IDS["USER"]
            }.GetEnumerator()) {
                $fontRef.SetAttribute([string]$pair.Key, [string]$pair.Value)
            }
        }
    }
    foreach ($paraId in @("11", "15", "17")) {
        $paraPr = $doc.SelectSingleNode("//hh:paraPr[@id='$paraId']", $ns)
        if ($null -eq $paraPr) { continue }
        $margin = $paraPr.SelectSingleNode("hh:margin", $ns)
        if ($null -eq $margin) {
            $margin = New-HhElement $doc "margin" @{}
            [void]$paraPr.AppendChild($margin)
        }
        foreach ($name in @("intent", "left", "right", "prev", "next")) {
            $target = $margin.SelectSingleNode("hc:$name", $ns)
            if ($null -eq $target) {
                $target = New-HcElement $doc $name @{}
                [void]$margin.AppendChild($target)
            }
            $target.SetAttribute("value", "0")
            $target.SetAttribute("unit", "HWPUNIT")
        }
    }
    $borderFills = $doc.SelectSingleNode("//hh:borderFills", $ns)
    if ($null -ne $borderFills) {
        $headerFill = if ($script:TABLE_STYLE -eq $script:FORMAL_STYLE_TEXT) { "#EEF3F8" } else { "" }
        $headerBorderFill = Ensure-BorderFill $doc $borderFills "17" $headerFill
        if ($script:TABLE_STYLE -eq $script:SIMPLE_STYLE_TEXT) {
            $existingFill = $headerBorderFill.SelectSingleNode("hc:fillBrush", $ns)
            if ($null -ne $existingFill) { [void]$headerBorderFill.RemoveChild($existingFill) }
        }
        [void](Ensure-BorderFill $doc $borderFills "18" "")
        [void](Ensure-BorderFill $doc $borderFills "19" "")
        $borderFills.SetAttribute("itemCnt", [string]$borderFills.ChildNodes.Count)
        $styles = @{
            "7" = @("NONE", "SOLID", "SOLID", "DOUBLE_SLIM", "0.1 mm", "0.3 mm", "0.5 mm")
            "8" = @("SOLID", "SOLID", "SOLID", "DOUBLE_SLIM", "0.1 mm", "0.3 mm", "0.5 mm")
            "9" = @("SOLID", "NONE", "SOLID", "DOUBLE_SLIM", "0.1 mm", "0.3 mm", "0.5 mm")
            "10" = @("SOLID", "SOLID", "SOLID", "DOUBLE_SLIM", "0.1 mm", "0.3 mm", "0.5 mm")
            "6" = @("NONE", "SOLID", "SOLID", "SOLID", "0.1 mm", "0.1 mm", "0.1 mm")
            "11" = @("SOLID", "SOLID", "SOLID", "SOLID", "0.1 mm", "0.1 mm", "0.1 mm")
            "12" = @("SOLID", "NONE", "SOLID", "SOLID", "0.1 mm", "0.1 mm", "0.1 mm")
            "13" = @("SOLID", "SOLID", "SOLID", "SOLID", "0.1 mm", "0.1 mm", "0.1 mm")
            "5" = @("NONE", "SOLID", "SOLID", "SOLID", "0.1 mm", "0.1 mm", "0.3 mm")
            "14" = @("SOLID", "SOLID", "SOLID", "SOLID", "0.1 mm", "0.1 mm", "0.3 mm")
            "15" = @("SOLID", "NONE", "SOLID", "SOLID", "0.1 mm", "0.1 mm", "0.3 mm")
            "16" = @("SOLID", "SOLID", "SOLID", "SOLID", "0.1 mm", "0.1 mm", "0.3 mm")
            "17" = @("NONE", "NONE", "SOLID", "DOUBLE_SLIM", "0.1 mm", "0.3 mm", "0.5 mm")
            "18" = @("NONE", "NONE", "SOLID", "SOLID", "0.1 mm", "0.1 mm", "0.1 mm")
            "19" = @("NONE", "NONE", "SOLID", "SOLID", "0.1 mm", "0.1 mm", "0.3 mm")
        }
        if ($script:TABLE_STYLE -eq $script:SIMPLE_STYLE_TEXT) {
            foreach ($id in @($styles.Keys)) {
                $style = $styles[$id]
                $style[5] = "0.1 mm"
                $style[6] = if ($style[3] -eq "DOUBLE_SLIM") { "0.5 mm" } else { "0.1 mm" }
            }
        }
        foreach ($id in $styles.Keys) {
            $borderFill = $borderFills.SelectSingleNode("hh:borderFill[@id='$id']", $ns)
            if ($null -eq $borderFill) { continue }
            $style = $styles[$id]
            $topColor = if (@("7", "8", "9", "10", "17") -contains $id) { $script:OUTER_BORDER_COLOR } else { $script:INNER_BORDER_COLOR }
            $bottomColor = if (@("5", "14", "15", "16", "19") -contains $id) { $script:OUTER_BORDER_COLOR } else { $script:INNER_BORDER_COLOR }
            $topWidth = if (@("7", "8", "9", "10", "17") -contains $id) { $style[5] } else { $style[4] }
            Set-BorderSide $doc $borderFill "leftBorder" $style[0] $style[4] $script:INNER_BORDER_COLOR
            Set-BorderSide $doc $borderFill "rightBorder" $style[1] $style[4] $script:INNER_BORDER_COLOR
            Set-BorderSide $doc $borderFill "topBorder" $style[2] $topWidth $topColor
            Set-BorderSide $doc $borderFill "bottomBorder" $style[3] $style[6] $bottomColor
            Set-BorderSide $doc $borderFill "diagonal" "NONE" $style[4] $script:INNER_BORDER_COLOR
        }
    }
    return ConvertTo-XmlBytes $doc
}

function Get-PreviewTextBytes($PayloadInfo) {
    $values = Get-CellValues $PayloadInfo
    $lines = New-Object System.Collections.Generic.List[string]
    for ($row = 0; $row -lt $PayloadInfo.Rows; $row++) {
        $items = New-Object System.Collections.Generic.List[string]
        for ($col = 0; $col -lt $PayloadInfo.Cols; $col++) {
            $key = "$row,$col"
            $items.Add($(if ($values.ContainsKey($key)) { [string]$values[$key] } else { "" }))
        }
        $lines.Add([string]::Join("`t", $items.ToArray()))
    }
    return [Text.Encoding]::UTF8.GetBytes(([string]::Join("`n", $lines.ToArray()) + "`n"))
}

function Copy-Stream([IO.Stream]$InputStream, [IO.Stream]$OutputStream) {
    $buffer = New-Object byte[] 81920
    while ($true) {
        $read = $InputStream.Read($buffer, 0, $buffer.Length)
        if ($read -le 0) { break }
        $OutputStream.Write($buffer, 0, $read)
    }
}

function Read-ZipEntryBytes([System.IO.Compression.ZipArchiveEntry]$Entry) {
    $stream = $Entry.Open()
    $memory = New-Object System.IO.MemoryStream
    try {
        Copy-Stream $stream $memory
        return $memory.ToArray()
    } finally {
        $memory.Dispose()
        $stream.Dispose()
    }
}

function Initialize-Crc32Type {
    if ($null -ne ("LHExcel.Crc32" -as [type])) { return }
    $source = @'
using System;

namespace LHExcel {
    public static class Crc32 {
        public static uint Compute(byte[] data) {
            uint crc = 0xffffffffu;
            for (int index = 0; index < data.Length; index++) {
                crc ^= data[index];
                for (int bit = 0; bit < 8; bit++) {
                    crc = (crc & 1u) != 0 ? (crc >> 1) ^ 0xedb88320u : crc >> 1;
                }
            }
            return crc ^ 0xffffffffu;
        }
    }
}
'@
    [void](Add-Type -TypeDefinition $source -Language CSharp)
}

function Write-StoredZip($Entries, [string]$Output) {
    Initialize-Crc32Type
    $now = [DateTime]::Now
    $year = [Math]::Max(1980, [Math]::Min(2107, $now.Year))
    $dosTime = [uint16](($now.Hour -shl 11) -bor ($now.Minute -shl 5) -bor [Math]::Floor($now.Second / 2))
    $dosDate = [uint16]((($year - 1980) -shl 9) -bor ($now.Month -shl 5) -bor $now.Day)
    $utf8Flag = [uint16]0x0800
    $records = New-Object System.Collections.Generic.List[object]
    $stream = [IO.File]::Open($Output, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    $writer = [IO.BinaryWriter]::new($stream, [Text.Encoding]::UTF8, $true)
    try {
        foreach ($item in $Entries) {
            [byte[]]$nameBytes = [Text.Encoding]::UTF8.GetBytes([string]$item.Name)
            [byte[]]$data = $item.Data
            [uint32]$crc = [LHExcel.Crc32]::Compute($data)
            [uint32]$size = $data.Length
            [uint32]$offset = $stream.Position
            $writer.Write([uint32]0x04034b50)
            $writer.Write([uint16]20)
            $writer.Write($utf8Flag)
            $writer.Write([uint16]0)
            $writer.Write($dosTime)
            $writer.Write($dosDate)
            $writer.Write($crc)
            $writer.Write($size)
            $writer.Write($size)
            $writer.Write([uint16]$nameBytes.Length)
            $writer.Write([uint16]0)
            $writer.Write($nameBytes)
            $writer.Write($data)
            $records.Add([pscustomobject]@{
                NameBytes = $nameBytes; Crc = $crc; Size = $size; Offset = $offset
            })
        }

        [uint32]$centralOffset = $stream.Position
        foreach ($record in $records) {
            $writer.Write([uint32]0x02014b50)
            $writer.Write([uint16]20)
            $writer.Write([uint16]20)
            $writer.Write($utf8Flag)
            $writer.Write([uint16]0)
            $writer.Write($dosTime)
            $writer.Write($dosDate)
            $writer.Write([uint32]$record.Crc)
            $writer.Write([uint32]$record.Size)
            $writer.Write([uint32]$record.Size)
            $writer.Write([uint16]$record.NameBytes.Length)
            $writer.Write([uint16]0)
            $writer.Write([uint16]0)
            $writer.Write([uint16]0)
            $writer.Write([uint16]0)
            $writer.Write([uint32]0)
            $writer.Write([uint32]$record.Offset)
            $writer.Write([byte[]]$record.NameBytes)
        }
        [uint32]$centralSize = $stream.Position - $centralOffset
        [uint16]$entryCount = $records.Count
        $writer.Write([uint32]0x06054b50)
        $writer.Write([uint16]0)
        $writer.Write([uint16]0)
        $writer.Write($entryCount)
        $writer.Write($entryCount)
        $writer.Write($centralSize)
        $writer.Write($centralOffset)
        $writer.Write([uint16]0)
        $writer.Flush()
    } finally {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function Initialize-PictureLayout($payload) {
    $layout = $payload.picture_layout
    foreach ($field in @('width_cm','height_cm','columns','rows','titles','sort','files')) {
        if ($layout.PSObject.Properties.Name -notcontains $field) { throw "PICTURE_FIELD: $field" }
    }
    $width = [double]$layout.width_cm; $height = [double]$layout.height_cm
    $columns = [double]$layout.columns; $rows = [double]$layout.rows
    if (-not ($width -ge 0.5 -and $width -le 16.7 -and $height -ge 0.5 -and $height -le 24)) { throw 'PICTURE_SIZE' }
    if ($columns -ne [Math]::Truncate($columns) -or $rows -ne [Math]::Truncate($rows) -or $columns -lt 1 -or $columns -gt 8 -or $rows -lt 1 -or $rows -gt 12) { throw 'PICTURE_LAYOUT' }
    if ($layout.titles -isnot [bool] -or @('title','taken') -notcontains $layout.sort) { throw 'PICTURE_OPTIONS' }
    $captionHeight = if ($layout.titles) { 0.7 } else { 0 }
    $titleFileNames = $true
    if ($layout.PSObject.Properties.Name -contains 'title_filenames') {
        if ($layout.title_filenames -isnot [bool]) { throw 'PICTURE_OPTIONS' }
        $titleFileNames = $layout.title_filenames
    }
    if ($width * $columns -gt 16.7 -or ($height + $captionHeight) * $rows -gt 24) { throw 'PICTURE_PAGE_SIZE: 한 쪽의 그림 배치가 가로 16.7cm, 세로 24cm를 넘습니다.' }
    $paths = @($layout.files)
    if ($paths.Count -lt 1 -or $paths.Count -gt 200) { throw 'PICTURE_COUNT: 1~200개를 선택하세요.' }
    Add-Type -AssemblyName System.Drawing
    $images = New-Object System.Collections.Generic.List[object]
    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $totalBytes = 0L
    foreach ($path in $paths) {
        if ($path -isnot [string] -or $path.Length -gt 1024 -or -not [IO.Path]::IsPathRooted($path) -or $path.StartsWith('\\?\') -or $path.StartsWith('\\.\') -or $path.IndexOf(':',2) -ge 0) { throw 'PICTURE_PATH' }
        $full = [IO.Path]::GetFullPath($path)
        if (-not $seen.Add($full)) { continue }
        if (@('.png','.jpg','.jpeg','.bmp') -notcontains [IO.Path]::GetExtension($full).ToLowerInvariant()) { throw 'PICTURE_TYPE' }
        $stream = [IO.File]::Open($full,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        try {
            if ($stream.Length -lt 4 -or $stream.Length -gt 32MB) { throw 'PICTURE_FILE_SIZE' }
            $image = [Drawing.Image]::FromStream($stream,$false,$true)
            try {
                if ([long]$image.Width * $image.Height -gt 16000000) { throw 'PICTURE_PIXEL_LIMIT' }
                $taken = [DateTime]::MaxValue
                if ($image.PropertyIdList -contains 36867) {
                    $value = [Text.Encoding]::ASCII.GetString($image.GetPropertyItem(36867).Value).Trim([char]0)
                    $parsed = [DateTime]::MinValue
                    if ([DateTime]::TryParseExact($value,'yyyy:MM:dd HH:mm:ss',[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$parsed)) { $taken = $parsed }
                }
                if ($image.PropertyIdList -contains 274) {
                    $orientation = [BitConverter]::ToUInt16($image.GetPropertyItem(274).Value,0)
                    $rotations = @(0,0,4,2,6,5,1,7,3)
                    if ($orientation -ge 2 -and $orientation -le 8) { $image.RotateFlip([Drawing.RotateFlipType]$rotations[$orientation]) }
                }
                $buffer = New-Object IO.MemoryStream
                $normalized = New-Object Drawing.Bitmap $image.Width,$image.Height,([Drawing.Imaging.PixelFormat]::Format32bppArgb)
                try {
                    $graphics = [Drawing.Graphics]::FromImage($normalized)
                    try {
                        $graphics.CompositingMode = [Drawing.Drawing2D.CompositingMode]::SourceCopy
                        $graphics.DrawImageUnscaled($image,0,0)
                    } finally { $graphics.Dispose() }
                    $normalized.Save($buffer,[Drawing.Imaging.ImageFormat]::Png)
                    $bytes = $buffer.ToArray(); $totalBytes += $bytes.Length
                    if ($totalBytes -gt 128MB) { throw 'PICTURE_TOTAL_SIZE' }
                    $images.Add([pscustomobject]@{Title=[IO.Path]::GetFileNameWithoutExtension($full); Path=$full; Taken=$taken; Bytes=$bytes})
                } finally { $normalized.Dispose(); $buffer.Dispose() }
            } finally { $image.Dispose() }
        } finally { $stream.Dispose() }
    }
    # Stable invariant title order, then stable capture-time order. Missing EXIF sorts last.
    $ordered = $images.ToArray()
    for ($i=1; $i -lt $ordered.Count; $i++) {
        $item=$ordered[$i]; $j=$i-1
        while ($j -ge 0) {
            $comparison=0
            if ($layout.sort -eq 'taken') { $comparison=[DateTime]::Compare($ordered[$j].Taken,$item.Taken) }
            if ($comparison -eq 0) { $comparison=[StringComparer]::OrdinalIgnoreCase.Compare($ordered[$j].Title,$item.Title) }
            if ($comparison -eq 0) { $comparison=[StringComparer]::Ordinal.Compare($ordered[$j].Path,$item.Path) }
            if ($comparison -le 0) { break }
            $ordered[$j+1]=$ordered[$j]; $j--
        }
        $ordered[$j+1]=$item
    }
    $factor=if ($layout.titles) {2} else {1}
    $pages=[int][Math]::Ceiling($ordered.Count / ($columns*$rows))
    $script:PictureLayout = [pscustomobject]@{Images=$ordered; Width=[int][Math]::Round($width*7200/2.54); Height=[int][Math]::Round($height*7200/2.54); Columns=[int]$columns; Rows=[int]$rows; Factor=$factor; Pages=$pages; TitleFileNames=$titleFileNames}
    return [pscustomobject]@{schema_version=1; rows=([int]$rows*$factor*$pages); columns=[int]$columns; row_heights=@(for($i=0;$i -lt $rows*$factor*$pages;$i++){20}); cells=@(); merges=@(); settings=[pscustomobject]@{font_name='중고딕';font_type='HFT';font_size_pt=9;title_row=$false;table_style='간결형'}}
}

function Apply-PictureLayout([byte[]]$SectionBytes, [byte[]]$HeaderBytes, [byte[]]$ManifestBytes) {
    $layout=$script:PictureLayout
    $section=Load-XmlBytes $SectionBytes; $header=Load-XmlBytes $HeaderBytes; $manifest=Load-XmlBytes $ManifestBytes
    $sns=New-NamespaceManager $section; $hns=New-NamespaceManager $header
    $sns.AddNamespace('hs','http://www.hancom.co.kr/hwpml/2011/section')
    foreach ($textNode in @($section.SelectNodes('/hs:sec/hp:p[.//hp:secPr]//hp:t',$sns))) {
        Set-TextNode $textNode ''
    }
    $fills=$header.SelectSingleNode('//hh:borderFills',$hns)
    $nextFill=1
    foreach ($fill in $fills.ChildNodes) { if ($fill -is [Xml.XmlElement]) { $nextFill=[Math]::Max($nextFill,1+[int]$fill.GetAttribute('id')) } }
    $emptyFill=$nextFill
    $manifestNode=$manifest.SelectSingleNode('//*[local-name()="manifest"]')
    if ($null -eq $manifestNode) { throw 'PICTURE_MANIFEST' }
    $binaries=New-Object System.Collections.Generic.List[object]
    for ($index=-1; $index -lt $layout.Images.Count; $index++) {
        $fill=New-HhElement $header 'borderFill' @{id=($emptyFill+$index+1);threeD='0';shadow='0';centerLine='NONE';breakCellSeparateLine='0'}
        foreach ($slash in @('slash','backSlash')) { [void]$fill.AppendChild((New-HhElement $header $slash @{type='NONE';Crooked='0';isCounter='0'})) }
        foreach ($side in @('leftBorder','rightBorder','topBorder','bottomBorder','diagonal')) { [void]$fill.AppendChild((New-HhElement $header $side @{type=$(if($side -eq 'diagonal'){'NONE'}else{'SOLID'});width='0.1 mm';color='#808080'})) }
        if ($index -ge 0) {
            $id='nxphoto'+($index+1); $entry='BinData/'+$id+'.png'
            $brush=New-HcElement $header 'fillBrush' @{}
            $imageBrush=New-HcElement $header 'imgBrush' @{mode='TOTAL'}
            [void]$imageBrush.AppendChild((New-HcElement $header 'img' @{binaryItemIDRef=$id;bright='0';contrast='0';effect='REAL_PIC';alpha='0'}))
            [void]$brush.AppendChild($imageBrush); [void]$fill.AppendChild($brush)
            $item=$manifest.CreateElement('opf','item',$manifestNode.NamespaceURI)
            $item.SetAttribute('id',$id); $item.SetAttribute('href',$entry); $item.SetAttribute('media-type','image/png'); $item.SetAttribute('isEmbeded','1')
            [void]$manifestNode.AppendChild($item)
            $binaries.Add([pscustomobject]@{Name=$entry;Data=$layout.Images[$index].Bytes})
        }
        [void]$fills.AppendChild($fill)
    }
    $fills.SetAttribute('itemCnt',[string]$fills.ChildNodes.Count)
    $table=$section.SelectSingleNode('//hp:tbl',$sns)
    $paragraph=$table.ParentNode.ParentNode
    $allRows=@($table.SelectNodes('hp:tr',$sns))
    foreach ($row in $allRows) { [void]$table.RemoveChild($row) }
    $groupRows=$layout.Rows*$layout.Factor
    $groupHeight=$layout.Rows*($layout.Height+$(if($layout.Factor -eq 2){1984}else{0}))
    for ($page=0; $page -lt $layout.Pages; $page++) {
        $p=$paragraph.CloneNode($true); $p.SetAttribute('id',[string](1100000000+$page)); $p.SetAttribute('pageBreak',$(if($page -eq 0){'0'}else{'1'}))
        $t=$p.SelectSingleNode('.//hp:tbl',$sns)
        $t.SetAttribute('id',[string](1200000000+$page)); $t.SetAttribute('rowCnt',[string]$groupRows); $t.SetAttribute('repeatHeader','0')
        $size=$t.SelectSingleNode('hp:sz',$sns); $size.SetAttribute('width',[string]($layout.Width*$layout.Columns)); $size.SetAttribute('height',[string]$groupHeight)
        for ($r=0; $r -lt $groupRows; $r++) {
            $row=$allRows[$page*$groupRows+$r]
            $isCaption=$layout.Factor -eq 2 -and $r%2 -eq 1
            foreach ($cell in $row.SelectNodes('hp:tc',$sns)) {
                $addr=$cell.SelectSingleNode('hp:cellAddr',$sns); $col=[int]$addr.GetAttribute('colAddr'); $addr.SetAttribute('rowAddr',[string]$r)
                $index=$page*$layout.Rows*$layout.Columns+[int][Math]::Floor($r/$layout.Factor)*$layout.Columns+$col
                $cell.SetAttribute('borderFillIDRef',[string]$(if(-not $isCaption -and $index -lt $layout.Images.Count){$emptyFill+1+$index}else{$emptyFill}))
                $size=$cell.SelectSingleNode('hp:cellSz',$sns); $size.SetAttribute('width',[string]$layout.Width); $size.SetAttribute('height',[string]$(if($isCaption){1984}else{$layout.Height}))
                $cell.SetAttribute('hasMargin','1')
                $margin=$cell.SelectSingleNode('hp:cellMargin',$sns)
                foreach($side in @('left','right','top','bottom')) { $margin.SetAttribute($side,'0') }
                $text=$cell.SelectSingleNode('.//hp:t',$sns)
                $text.InnerText=$(if($isCaption -and $layout.TitleFileNames -and $index -lt $layout.Images.Count){$layout.Images[$index].Title}else{''})
            }
            [void]$t.AppendChild($row)
        }
        [void]$paragraph.ParentNode.InsertBefore($p,$paragraph)
    }
    [void]$paragraph.ParentNode.RemoveChild($paragraph)
    return @{Section=(ConvertTo-XmlBytes $section); Header=(ConvertTo-XmlBytes $header); Manifest=(ConvertTo-XmlBytes $manifest); Binaries=$binaries}
}

function Write-Hwpx($PayloadInfo, [string]$Template, [string]$Output) {
    if (-not (Test-Path -LiteralPath $Template -PathType Leaf)) {
        throw (New-HwpxError "TEMPLATE_NOT_FOUND" "Template HWPX was not found: $Template" 4)
    }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $outDir = [IO.Path]::GetDirectoryName($Output)
    if (-not [string]::IsNullOrWhiteSpace($outDir) -and -not (Test-Path -LiteralPath $outDir)) {
        [void][IO.Directory]::CreateDirectory($outDir)
    }
    $sourceStream = [IO.File]::Open($Template, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $source = [System.IO.Compression.ZipArchive]::new($sourceStream, [IO.Compression.ZipArchiveMode]::Read)
        try {
            $sectionEntry = $source.GetEntry("Contents/section0.xml")
            $headerEntry = $source.GetEntry("Contents/header.xml")
            $mimetypeEntry = $source.GetEntry("mimetype")
            if ($null -eq $sectionEntry -or $null -eq $headerEntry -or $null -eq $mimetypeEntry) {
                throw (New-HwpxError "TEMPLATE_PACKAGE_INVALID" "Template must contain mimetype, Contents/header.xml, and Contents/section0.xml." 4)
            }
            $sectionXml = Build-SectionXml $PayloadInfo (Read-ZipEntryBytes $sectionEntry)
            $headerXml = Patch-HeaderXml (Read-ZipEntryBytes $headerEntry)
            $previewText = Get-PreviewTextBytes $PayloadInfo
            $picturePackage=$null
            if ($null -ne $script:PictureLayout) {
                $picturePackage=Apply-PictureLayout $sectionXml $headerXml (Read-ZipEntryBytes $source.GetEntry('Contents/content.hpf'))
                $sectionXml=$picturePackage.Section; $headerXml=$picturePackage.Header
                $previewText=[Text.Encoding]::UTF8.GetBytes(($script:PictureLayout.Images.Title -join "`n"))
            }
            $outputEntries = New-Object System.Collections.Generic.List[object]
            $outputEntries.Add([pscustomobject]@{ Name = "mimetype"; Data = (Read-ZipEntryBytes $mimetypeEntry) })
            foreach ($entry in $source.Entries) {
                if ($entry.FullName -eq "mimetype") { continue }
                $data = if ($entry.FullName -eq "Contents/section0.xml") {
                    $sectionXml
                } elseif ($entry.FullName -eq "Contents/header.xml") {
                    $headerXml
                } elseif ($entry.FullName -eq "Preview/PrvText.txt") {
                    $previewText
                } elseif ($null -ne $picturePackage -and $entry.FullName -eq 'Contents/content.hpf') {
                    $picturePackage.Manifest
                } else {
                    Read-ZipEntryBytes $entry
                }
                $outputEntries.Add([pscustomobject]@{ Name = $entry.FullName; Data = $data })
            }
            if ($null -ne $picturePackage) {
                foreach ($binary in $picturePackage.Binaries) { $outputEntries.Add($binary) }
                if (Test-Path -LiteralPath $Output) { throw 'PICTURE_OUTPUT_EXISTS' }
            } elseif (Test-Path -LiteralPath $Output) { Remove-Item -LiteralPath $Output -Force }
            Write-StoredZip $outputEntries $Output
        } finally {
            $source.Dispose()
        }
    } finally {
        $sourceStream.Dispose()
    }
}

function Verify-Hwpx([string]$Path) {
    $result = [ordered]@{
        path = $Path; ok = $false; mimetype_first = $false; mimetype_stored = $false;
        has_header_xml = $false; has_section_xml = $false; table_count = 0; cell_count = 0; errors = @()
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        $result.errors = @("HWPX file does not exist")
        return $result
    }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $archive = [System.IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Read)
        try {
            if ($archive.Entries.Count -lt 1) {
                $result.errors = @("HWPX package is empty")
                return $result
            }
            $first = $archive.Entries[0]
            $result.mimetype_first = $first.FullName -eq "mimetype"
            $result.mimetype_stored = $result.mimetype_first -and $first.CompressedLength -eq $first.Length
            $headerEntry = $archive.GetEntry("Contents/header.xml")
            $sectionEntry = $archive.GetEntry("Contents/section0.xml")
            $result.has_header_xml = $null -ne $headerEntry
            $result.has_section_xml = $null -ne $sectionEntry
            $errors = New-Object System.Collections.Generic.List[string]
            if ($null -eq $headerEntry) { $errors.Add("Contents/header.xml is missing") }
            if ($null -eq $sectionEntry) { $errors.Add("Contents/section0.xml is missing") }
            if ($errors.Count -eq 0) {
                $section = Load-XmlBytes (Read-ZipEntryBytes $sectionEntry)
                $ns = New-NamespaceManager $section
                $result.table_count = $section.SelectNodes("//hp:tbl", $ns).Count
                $result.cell_count = $section.SelectNodes("//hp:tc", $ns).Count
                if (-not $result.mimetype_first) { $errors.Add("mimetype is not the first ZIP entry") }
                if (-not $result.mimetype_stored) { $errors.Add("mimetype is not stored without compression") }
                if ($result.table_count -lt 1) { $errors.Add("table element is missing") }
                if ($result.cell_count -lt 1) { $errors.Add("table cells are missing") }
            }
            $result.errors = @($errors)
            $result.ok = $errors.Count -eq 0
            return $result
        } finally {
            $archive.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

function Main {
    if (-not [string]::IsNullOrWhiteSpace($VerifyPath)) {
        $verifyFullPath = Resolve-RequiredPath $VerifyPath "VerifyPath"
        $verification = Verify-Hwpx $verifyFullPath
        [Console]::Out.WriteLine(($verification | ConvertTo-Json -Depth 8 -Compress))
        if (-not $verification.ok) { return 6 }
        return 0
    }
    $inputFullPath = Resolve-RequiredPath $InputPath "InputPath"
    $outputFullPath = Resolve-RequiredPath $OutputPath "OutputPath"
    $templateFullPath = Resolve-RequiredPath $TemplatePath "TemplatePath"
    $payload = Read-Payload $inputFullPath
    Write-Hwpx $payload $templateFullPath $outputFullPath
    $verification = Verify-Hwpx $outputFullPath
    if (-not $verification.ok) {
        throw (New-HwpxError "OUTPUT_VERIFY_FAILED" ("Generated HWPX verification failed: " + ([string]::Join("; ", @($verification.errors)))) 6)
    }
    [Console]::Out.WriteLine(([ordered]@{ status = "PASS"; output = $outputFullPath; verification = $verification } | ConvertTo-Json -Depth 8 -Compress))
    return 0
}

try {
    exit (Main)
} catch {
    $code = 1
    if ($_.Exception.Data.Contains("ExitCode")) { $code = [int]$_.Exception.Data["ExitCode"] }
    $errorCode = if ($_.Exception.Data.Contains("Code")) { [string]$_.Exception.Data["Code"] } else { "UNHANDLED_ERROR" }
    [Console]::Error.WriteLine(("ERROR[{0}] line {1}: {2}" -f $errorCode, $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message))
    [Console]::Error.WriteLine($_.ScriptStackTrace)
    exit $code
}
