Attribute VB_Name = "T_R71Serialization"
Option Explicit
Public Function Names() As String
    Names = "value|formula|numberformat|font_name|font_size|bold|italic|underline|strike|font_color|font_tint|pattern|pattern_color|fill_color|fill_tint|align|valign|orientation|indent|shrink|reading|wrap|locked|hiddenformula|border_line|border_weight|border_color|border_tint|diagonal|merge|benchmark_90000"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, cell As Range, before As String, after As String
    Dim context As CNxExecutionContext, oldChanged As Boolean, xmlChanged As Boolean
    Dim started As Double, elapsed As Double, detail As String, saved As Boolean
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("B2:D5")
    target.Value2 = "test"
    target.Cells(1, 1).Formula = "=1+2"
    Set cell = target.Cells(4, 3)
    target.Select
    If name = "benchmark_90000" Then
        Set target = book.Worksheets(1).Range("A1:AX1800")
        target.Value2 = "benchmark": target.Font.Name = "맑은 고딕": target.Font.Size = 9
        target.Cells(2, 2).Formula = "=1+2": target.Cells(1800, 50).Value2 = "last"
        saved = book.Saved
        started = Timer
        before = CStr(target.Value(xlRangeValueXMLSpreadsheet))
        elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400#
        after = CStr(target.Value(xlRangeValueXMLSpreadsheet))
        If before <> after Then Err.Raise 5, , "XML not deterministic"
        If saved <> book.Saved Then Err.Raise 5, , "XML changed Saved state"
        RunCase = "PASS|" & name & "|xml_ms=" & Format$(elapsed * 1000#, "0.000") & "|chars=" & Len(before)
        GoTo CleanUp
    End If
    before = CStr(target.Value(xlRangeValueXMLSpreadsheet))
    If before <> CStr(target.Value(xlRangeValueXMLSpreadsheet)) Then Err.Raise 5, , "XML not deterministic"
    Set context = NxContextFactory.CaptureCurrent()
    Select Case name
        Case "value": cell.Value2 = "changed"
        Case "formula": cell.Formula = "=7+9"
        Case "numberformat": cell.NumberFormat = "0.00"
        Case "font_name": cell.Font.Name = "Arial"
        Case "font_size": cell.Font.Size = 15
        Case "bold": cell.Font.Bold = True
        Case "italic": cell.Font.Italic = True
        Case "underline": cell.Font.Underline = xlUnderlineStyleDouble
        Case "strike": cell.Font.Strikethrough = True
        Case "font_color": cell.Font.Color = vbRed
        Case "font_tint": cell.Font.TintAndShade = 0.4
        Case "pattern": cell.Interior.Pattern = xlPatternGray50
        Case "pattern_color": cell.Interior.PatternColor = vbRed
        Case "fill_color": cell.Interior.Color = vbYellow
        Case "fill_tint": cell.Interior.TintAndShade = 0.3
        Case "align": cell.HorizontalAlignment = xlRight
        Case "valign": cell.VerticalAlignment = xlTop
        Case "orientation": cell.Orientation = 30
        Case "indent": cell.IndentLevel = 2
        Case "shrink": cell.ShrinkToFit = True
        Case "reading": cell.ReadingOrder = xlRTL
        Case "wrap": cell.WrapText = True
        Case "locked": cell.Locked = False
        Case "hiddenformula": cell.FormulaHidden = True
        Case "border_line": cell.Borders(xlEdgeLeft).LineStyle = xlContinuous
        Case "border_weight": cell.Borders(xlEdgeLeft).Weight = xlThick
        Case "border_color": cell.Borders(xlEdgeLeft).Color = vbBlue
        Case "border_tint": cell.Borders(xlEdgeLeft).TintAndShade = 0.5
        Case "diagonal": cell.Borders(xlDiagonalUp).LineStyle = xlContinuous
        Case "merge": target.Worksheet.Range("C5:D5").ClearContents: target.Worksheet.Range("C5:D5").Merge
    End Select
    after = CStr(target.Value(xlRangeValueXMLSpreadsheet))
    oldChanged = Not context.SourceStateMatchesCurrent()
    xmlChanged = (before <> after)
    If oldChanged And Not xmlChanged Then
        RunCase = "FAIL|" & name & "|CANDIDATE_MISS"
    Else
        RunCase = "PASS|" & name & "|legacy_changed=" & CStr(oldChanged) & "|xml_changed=" & CStr(xmlChanged)
    End If
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
End Function
