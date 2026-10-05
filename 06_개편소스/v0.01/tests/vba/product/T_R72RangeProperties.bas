Attribute VB_Name = "T_R72RangeProperties"
Option Explicit

Public Function Names() As String
    Names = "font_name|font_size|font_bold|font_italic|font_underline|font_strike|font_rgb|font_theme|font_tint|fill_rgb|fill_pattern|fill_tint|pattern_rgb|align|valign|orientation|indent|addindent|shrink|reading|wrap|locked|hidden|numberformat|border_line|border_weight|border_rgb|border_tint|diagonal|theme_same_rgb|uniform_90000"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, changed As Range
    Dim before As String, after As String, detail As String, started As Double
    Dim handle As Integer, elapsed As Double
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("B2:K11")
    If name = "uniform_90000" Then Set target = book.Worksheets(1).Range("A1:AX1800")
    target.Value2 = "test": target.Font.Name = "맑은 고딕": target.Font.Size = 9
    If name = "theme_same_rgb" Then target.Font.ThemeColor = xlThemeColorAccent1
    Set changed = target.Cells(target.Rows.Count, target.Columns.Count)
    before = Aggregate(target)
    Select Case name
        Case "font_name": changed.Font.Name = "Arial"
        Case "font_size": changed.Font.Size = 15
        Case "font_bold": changed.Font.Bold = True
        Case "font_italic": changed.Font.Italic = True
        Case "font_underline": changed.Font.Underline = xlUnderlineStyleDouble
        Case "font_strike": changed.Font.Strikethrough = True
        Case "font_rgb": changed.Font.Color = RGB(123, 87, 213)
        Case "font_theme": changed.Font.ThemeColor = xlThemeColorAccent2
        Case "font_tint": changed.Font.TintAndShade = 0.3
        Case "fill_rgb": changed.Interior.Color = RGB(255, 200, 150)
        Case "fill_pattern": changed.Interior.Pattern = xlPatternGray50
        Case "fill_tint": changed.Interior.Pattern = xlSolid: changed.Interior.TintAndShade = 0.2
        Case "pattern_rgb": changed.Interior.PatternColor = vbRed
        Case "align": changed.HorizontalAlignment = xlRight
        Case "valign": changed.VerticalAlignment = xlTop
        Case "orientation": changed.Orientation = 30
        Case "indent": changed.IndentLevel = 2
        Case "addindent": changed.HorizontalAlignment = xlDistributed: changed.AddIndent = True
        Case "shrink": changed.ShrinkToFit = True
        Case "reading": changed.ReadingOrder = xlRTL
        Case "wrap": changed.WrapText = True
        Case "locked": changed.Locked = False
        Case "hidden": changed.FormulaHidden = True
        Case "numberformat": changed.NumberFormat = "0.000"
        Case "border_line": changed.Borders(xlEdgeLeft).LineStyle = xlContinuous
        Case "border_weight": changed.Borders(xlEdgeLeft).Weight = xlThick
        Case "border_rgb": changed.Borders(xlEdgeLeft).Color = vbBlue
        Case "border_tint": changed.Borders(xlEdgeLeft).TintAndShade = 0.5
        Case "diagonal": changed.Borders(xlDiagonalUp).LineStyle = xlContinuous
        Case "theme_same_rgb": changed.Font.Color = changed.Font.Color
        Case "uniform_90000"
            started = Timer
            after = Aggregate(target)
            elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400#
            handle = FreeFile
            Open Environ$("LHEXCEL_PROFILE_ROOT") & "\range-properties.tsv" For Append As #handle
            Print #handle, "uniform_90000" & vbTab & Format$(elapsed * 1000#, "0.000")
            Close #handle
            If before <> after Then Err.Raise 5, , "Uniform aggregate drift"
            GoTo Passed
    End Select
    after = Aggregate(target)
    If before = after Then Err.Raise 5, , "Mixed property invisible to aggregate"
Passed:
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
End Function

Private Function Aggregate(ByVal target As Range) As String
    Dim item As Variant, edge As Variant, result As String
    For Each item In Array("NumberFormat", "HorizontalAlignment", "VerticalAlignment", "Orientation", "AddIndent", "IndentLevel", "ShrinkToFit", "ReadingOrder", "WrapText", "Locked", "FormulaHidden")
        result = result & ReadProperty(target, CStr(item))
    Next item
    For Each item In Array("Name", "Size", "Bold", "Italic", "Underline", "Strikethrough", "Color", "ColorIndex", "ThemeColor", "TintAndShade")
        result = result & ReadProperty(target.Font, CStr(item))
    Next item
    For Each item In Array("Pattern", "PatternColor", "PatternColorIndex", "Color", "ColorIndex", "ThemeColor", "TintAndShade")
        result = result & ReadProperty(target.Interior, CStr(item))
    Next item
    For Each edge In Array(xlDiagonalDown, xlDiagonalUp, xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        For Each item In Array("Weight", "Color", "ColorIndex", "ThemeColor", "TintAndShade", "LineStyle")
            result = result & ReadProperty(target.Borders(edge), CStr(item))
        Next item
    Next edge
    Aggregate = result
End Function

Private Function ReadProperty(ByVal item As Object, ByVal name As String) As String
    Dim value As Variant
    On Error GoTo Unavailable
    value = CallByName(item, name, VbGet)
    If IsNull(value) Then
        ReadProperty = "|" & name & "=NULL"
    Else
        ReadProperty = "|" & name & "=" & VarType(value) & ":" & CStr(value)
    End If
    Exit Function
Unavailable:
    ReadProperty = "|" & name & "=ERR" & Err.Number
    Err.Clear
End Function
