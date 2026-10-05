Attribute VB_Name = "T_R71FastSnapshot"
Option Explicit
Public Function Names() As String
    Names = "volatile|multi_area|partial_merge|saved_state|error_values|spill|style_roundtrip|theme_tint|pattern_tint|automatic_color|source_array|add_indent|border_colorindex|fill_colorindex|outside_style"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, context As CNxExecutionContext, changed As Range
    Dim oldSaved As Boolean, detail As String
    Dim values As Variant
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("B2:D5")
    target.Value2 = "test": Set changed = target.Cells(4, 3)
    If name = "source_array" Then
        Set target = book.Worksheets(1).Range("A1:AX1800")
        target.Value2 = "benchmark": target.Cells(1, 1).Value = CVErr(xlErrNA)
        values = target.Value2
        If Not NxDataSpecialSourceMatches(target, values, 1800, 50) Then Err.Raise 5, , "array match"
        target.Cells(1800, 50).Value2 = "changed"
        If NxDataSpecialSourceMatches(target, values, 1800, 50) Then Err.Raise 5, , "last array change missed"
        If values(1800, 50) <> "benchmark" Then Err.Raise 5, , "input array changed"
        RunCase = "PASS|" & name
        GoTo CleanUp
    End If
    Select Case name
        Case "volatile": target.Cells(1, 1).Formula = "=RAND()"
        Case "multi_area": Set target = Union(target, book.Worksheets(1).Range("G9:H10")): Set changed = book.Worksheets(1).Range("H10")
        Case "partial_merge": book.Worksheets(1).Range("A1:C3").ClearContents: book.Worksheets(1).Range("A1:C3").Merge
        Case "error_values": target.Cells(1, 1).Value = CVErr(xlErrNA)
        Case "spill": target.ClearContents: target.Cells(1, 1).Formula2 = "=SEQUENCE(2,2)"
        Case "theme_tint": changed.Font.ThemeColor = xlThemeColorAccent1: changed.Font.TintAndShade = 0.2
        Case "pattern_tint": changed.Interior.Pattern = xlPatternGray50: changed.Interior.PatternColor = vbRed
        Case "automatic_color": changed.Font.ColorIndex = xlColorIndexAutomatic
    End Select
    target.Select
    If name = "saved_state" Then
        book.SaveAs Environ$("LHEXCEL_PROFILE_ROOT") & "\snapshot-clean.xlsx", xlOpenXMLWorkbook
        If Not book.Saved Then Err.Raise 5, , "clean workbook fixture"
    End If
    oldSaved = book.Saved
    Set context = NxContextFactory.CaptureCurrent()
    If name = "style_roundtrip" Then SaveXml target, "before"
    If oldSaved <> book.Saved Then Err.Raise 5, , "capture altered Saved"
    If Not context.SourceStateMatchesCurrent Then Err.Raise 5, , "unchanged source rejected"
    Select Case name
        Case "volatile"
            book.Worksheets(1).Calculate
            If Not context.SourceStateMatchesCurrent Then Err.Raise 5, , "formula cache drift"
            target.Cells(1, 1).Formula = "=RAND()+1"
        Case "style_roundtrip"
            changed.Font.Bold = True: changed.Font.Bold = False
            SaveXml target, "after"
            If Not context.SourceStateMatchesCurrent Then Err.Raise 5, , "style pool drift"
            changed.Value2 = "changed"
        Case "theme_tint": changed.Font.TintAndShade = 0.5
        Case "pattern_tint": changed.Interior.PatternTintAndShade = 0.4
        Case "automatic_color": changed.Font.Color = vbBlack
        Case "spill": target.Cells(1, 1).Formula2 = "=SEQUENCE(2,2,2)"
        Case "add_indent": changed.HorizontalAlignment = xlDistributed: changed.AddIndent = True
        Case "border_colorindex": changed.Borders(xlEdgeLeft).ColorIndex = 3
        Case "fill_colorindex": changed.Interior.ColorIndex = 6
        Case "outside_style"
            book.Worksheets(1).Range("Z100").Font.Bold = True
            If Not context.SourceStateMatchesCurrent Then Err.Raise 5, , "unrelated style invalidated source"
            changed.Value2 = "changed"
        Case Else: changed.Value2 = "changed"
    End Select
    If context.SourceStateMatchesCurrent Then Err.Raise 5, , "source drift missed"
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
End Function

Private Sub SaveXml(ByVal target As Range, ByVal label As String)
    Dim document As Object
    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    document.LoadXML CStr(target.Value(xlRangeValueXMLSpreadsheet))
    document.Save Environ$("LHEXCEL_PROFILE_ROOT") & "\style-" & label & ".xml"
End Sub
