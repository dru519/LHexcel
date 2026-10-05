Attribute VB_Name = "T_R68QuickFormat"
Option Explicit
Private mUndoFixture As Workbook
Private mUndoStage As String

Private Function QuickKeys() As Variant
    QuickKeys = Split("FONT-SIZE-9|FONT-SIZE-10|FONT-SIZE-12|FONT-SIZE-15|FONT-COLOR-BLACK|FONT-COLOR-RED|FONT-COLOR-BLUE|FONT-COLOR-GREEN|FILL-GRAY|FILL-LIGHT-RED|FILL-LIGHT-YELLOW|FILL-LIGHT-GREEN", "|")
End Function

Public Function Names() As String
    Dim key As Variant, mode As Variant, result As String
    For Each mode In Array("single", "multi", "protected")
        For Each key In QuickKeys()
            If Len(result) > 0 Then result = result & "|"
            result = result & CStr(mode) & "_" & CStr(key)
        Next key
    Next mode
    Names = result & "|shape_selection|undo_details|whole_row_fill|whole_column_fill|privacy_report"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim fixture As Workbook, sheet As Worksheet, selected As Range, cell As Range
    Dim shape As Shape, parts As Variant, key As String, mode As String, keys As Variant
    Dim index As Long, found As Boolean, errorNumber As Long, detail As String
    Dim previousEvents As Boolean, previousScreen As Boolean, previousCalculation As XlCalculation
    On Error GoTo Failed
    previousEvents = Application.EnableEvents
    previousScreen = Application.ScreenUpdating
    previousCalculation = Application.Calculation
    Set fixture = Workbooks.Add(xlWBATWorksheet)
    Set sheet = fixture.Worksheets(1)
    Initialize sheet.Range("A1:E5")
    sheet.Range("B2").Formula = "=12+30"
    sheet.Range("D4").Formula = "=8*7"
    sheet.Range("B2").AddComment "keep note"
    NxResetCommandRegistry
    If name = "privacy_report" Then
        CheckPrivacyReport fixture
    ElseIf name = "whole_row_fill" Or name = "whole_column_fill" Then
        If name = "whole_row_fill" Then
            Set selected = sheet.Rows(2)
        Else
            Set selected = sheet.Columns(2)
        End If
        selected.Select
        NxRouteCommand "NX-CMD-STYLE-FILL-LIGHT-GREEN"
        Require selected.Interior.Color = RGB(201, 240, 208), "whole axis fill incomplete"
        Require sheet.Range("B2").Formula = "=12+30", "whole axis changed formula"
        NxQuickFormatUndoLast
        Require sheet.Range("B2").Interior.Color = RGB(231, 235, 240), "whole axis undo lost original fill"
        Require selected.Cells(selected.Rows.Count, selected.Columns.Count).Interior.ColorIndex = xlColorIndexNone, "whole axis undo lost empty tail"
    ElseIf name = "undo_details" Then
        CheckUndoDetails sheet
    ElseIf name = "shape_selection" Then
        Set shape = sheet.Shapes.AddShape(msoShapeRectangle, 8, 8, 30, 30)
        shape.Select
        On Error Resume Next
        NxRouteCommand "NX-CMD-STYLE-FONT-SIZE-9"
        errorNumber = Err.Number
        Err.Clear
        On Error GoTo Failed
        Require errorNumber <> 0, "shape selection was not rejected"
        CheckCell sheet.Range("B2"), -1
    Else
        parts = Split(name, "_", 2)
        Require UBound(parts) = 1, "unknown case"
        mode = CStr(parts(0)): key = CStr(parts(1))
        Require mode = "single" Or mode = "multi" Or mode = "protected", "unknown mode"
        keys = QuickKeys()
        For index = LBound(keys) To UBound(keys)
            If CStr(keys(index)) = key Then found = True: Exit For
        Next index
        Require found, "unknown format"
        Set selected = sheet.Range("B2")
        If mode = "multi" Then Set selected = Application.Union(sheet.Range("B2:B3"), sheet.Range("D4:D5"))
        selected.Select
        If mode = "protected" Then sheet.Protect "r68-test", AllowFormattingCells:=True
        On Error Resume Next
        NxRouteCommand "NX-CMD-STYLE-" & key
        errorNumber = Err.Number
        detail = Err.Description
        Err.Clear
        On Error GoTo Failed
        If mode = "protected" Then
            Require errorNumber <> 0, "protected sheet accepted formatting"
            sheet.Unprotect "r68-test"
            ' The handler must also reject a direct call, even if Excel allows formatting.
            On Error Resume Next
            sheet.Protect "r68-test", AllowFormattingCells:=True
            NxCmdStyle "NX_STYLE_" & Replace(key, "-", "_")
            errorNumber = Err.Number
            Err.Clear
            On Error GoTo Failed
            Require errorNumber <> 0, "direct protected call accepted formatting"
            sheet.Unprotect "r68-test"
            For Each cell In selected.Cells
                CheckCell cell, -1
            Next cell
        Else
            Require errorNumber = 0, "route failed: " & detail
            For Each cell In selected.Cells
                CheckCell cell, index
            Next cell
            NxQuickFormatUndoLast
            For Each cell In selected.Cells
                CheckCell cell, -1
            Next cell
        End If
        Require Selection.Address = selected.Address, "selection changed"
    End If
    CheckCell sheet.Range("C3"), -1
    Require sheet.Range("B2").Formula = "=12+30", "formula changed"
    Require sheet.Range("D4").Formula = "=8*7", "second-area formula changed"
    Require sheet.Range("B2").Comment.Text = "keep note", "comment changed"
    Require Application.EnableEvents = previousEvents, "events not restored"
    Require Application.ScreenUpdating = previousScreen, "screen updating not restored"
    Require Application.Calculation = previousCalculation, "calculation not restored"
    fixture.Close False
    RunCase = "PASS|" & name
    Exit Function
Failed:
    detail = Err.Description & " / " & mUndoStage
    On Error Resume Next
    If Not fixture Is Nothing Then fixture.Close False
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousScreen
    Application.Calculation = previousCalculation
    RunCase = "FAIL|" & name & "|" & detail
End Function

Private Sub CheckUndoDetails(ByVal sheet As Worksheet)
    Dim r As Range, theme As Long, tint As Double, started As Double
    mUndoStage = "mixed sizes"
    Set r = sheet.Range("G1:H3")
    r.Value2 = 12
    r.Font.Size = 11
    r.Cells(1, 1).Font.Size = 18
    r.Select
    NxRouteCommand "NX-CMD-STYLE-FONT-SIZE-9"
    NxQuickFormatUndoLast
    Require r.Cells(1, 1).Font.Size = 18 And r.Cells(2, 1).Font.Size = 11, "mixed sizes"
    mUndoStage = "mixed colors"
    r.Font.Color = vbBlue
    r.Cells(1, 1).Font.Color = vbGreen
    NxRouteCommand "NX-CMD-STYLE-FONT-COLOR-RED"
    NxQuickFormatUndoLast
    Require r.Cells(1, 1).Font.Color = vbGreen And r.Cells(2, 1).Font.Color = vbBlue, "mixed font colors"
    r.Interior.Color = vbYellow
    r.Cells(1, 1).Interior.Pattern = xlNone
    NxRouteCommand "NX-CMD-STYLE-FILL-GRAY"
    NxQuickFormatUndoLast
    Require r.Cells(1, 1).Interior.Pattern = xlNone And r.Cells(2, 1).Interior.Color = vbYellow, "mixed fill"
    mUndoStage = "font theme"
    r.Font.ThemeColor = xlThemeColorAccent2
    r.Font.TintAndShade = 0.4
    NxRouteCommand "NX-CMD-STYLE-FONT-COLOR-BLUE"
    NxQuickFormatUndoLast
    Require r.Font.ThemeColor = xlThemeColorAccent2 And Abs(r.Font.TintAndShade - 0.4) < 0.001, "font theme"
    mUndoStage = "no fill"
    r.Interior.Pattern = xlNone
    NxRouteCommand "NX-CMD-STYLE-FILL-GRAY"
    NxQuickFormatUndoLast
    Require r.Interior.Pattern = xlNone, "no fill"
    mUndoStage = "pattern fill"
    r.Interior.ThemeColor = xlThemeColorAccent3
    r.Interior.TintAndShade = 0.2
    r.Interior.Pattern = xlGray16
    r.Interior.PatternColor = vbRed
    NxRouteCommand "NX-CMD-STYLE-FILL-GRAY"
    NxQuickFormatUndoLast
    Require r.Interior.ThemeColor = xlThemeColorAccent3 And r.Interior.Pattern = xlGray16, "fill theme/pattern"
    Require r.Interior.PatternColor = vbRed, "pattern color"
    mUndoStage = "rich text"
    Set r = sheet.Range("G5")
    r.Value2 = "Abcd"
    r.Font.Size = 11
    r.Characters(1, 1).Font.Size = 18
    r.Select
    NxRouteCommand "NX-CMD-STYLE-FONT-SIZE-9"
    NxQuickFormatUndoLast
    Require r.Characters(1, 1).Font.Size = 18 And r.Characters(2, 1).Font.Size = 11, "rich text sizes"
    mUndoStage = "large range"
    Set r = sheet.Range("J1:BG1800")
    r.Font.Size = 11
    r.Select
    started = Timer
    NxRouteCommand "NX-CMD-STYLE-FONT-SIZE-9"
    NxQuickFormatUndoLast
    Require r.Font.Size = 11, "large range undo"
    Require Timer - started < 10, "uniform 90000-cell undo too slow"
End Sub

Public Sub PrepareNativeUndo()
    Set mUndoFixture = Workbooks.Add(xlWBATWorksheet)
    With mUndoFixture.Worksheets(1).Range("A1:B2")
        .Value2 = 12
        .Font.Size = 11
        .Select
    End With
    NxRouteCommand "NX-CMD-STYLE-FONT-SIZE-15"
End Sub

Public Function VerifyNativeUndo() As String
    VerifyNativeUndo = "FAIL|native_undo|font size"
    If mUndoFixture.Worksheets(1).Range("A1:B2").Font.Size = 11 Then VerifyNativeUndo = "PASS|native_undo"
    mUndoFixture.Close False
    Set mUndoFixture = Nothing
End Function

Private Sub CheckPrivacyReport(ByVal sourceBook As Workbook)
    Dim findings As New Collection, incomplete As New Collection, report As Worksheet
    Dim source As Worksheet, hidden As Worksheet, failure As Long, detail As String
    On Error GoTo Failed
    Set source = sourceBook.Worksheets(1)
    source.Name = "개인'정보"
    Set hidden = sourceBook.Worksheets.Add
    hidden.Name = "숨김자료": hidden.Visible = xlSheetHidden
    Require NxPrivacyAuditKind("보고서", "") = "", "unlabelled Korean word flagged"
    Require NxPrivacyAuditKind("sample@example.invalid", "") = "", "email must be excluded"
    Require NxPrivacyAuditKind("2024-06-17", "") = "", "business date flagged"
    Require NxPrivacyAuditKind("홍길동", "이름") = "이름", "labelled name missed"
    Require NxPrivacyAuditKind("1990-06-17", "생년월일") = "생년월일", "birth date missed"
    Require NxPrivacyAuditKind("1990-02-31", "생년월일") = "", "invalid birth date flagged"
    Require NxPrivacyAuditKind("900617-1234567", "") = "주민등록번호", "resident pattern missed"
    Require NxPrivacyAuditKind("991332-1234567", "") = "", "invalid resident date flagged"
    Require NxPrivacyAuditKind("010-1234-5678", "") = "연락처", "phone missed"
    Require NxPrivacyAuditKind("099-1234-5678", "") = "", "invalid phone prefix flagged"
    Require NxPrivacyAuditKind("금액01012345678원", "") = "", "embedded amount flagged"
    source.Range("G1").Value2 = "성명": source.Range("G2").Value2 = "홍길동"
    Dim audited As Collection, misses As Collection, checked As Double
    Set audited = NxPrivacyScanWorkbook(sourceBook, checked, misses)
    Require audited.Count = 1, "header-based workbook classification"
    source.Range("G1:G2").Clear
    findings.Add Array("이름", source.Name, 2, "B2", "표시", "홍길동")
    findings.Add Array("전화번호", hidden.Name, 1, "A1", "숨김", "010-0000-0000")
    Set report = NxPrivacyCreateWorkbookReport(sourceBook, findings, incomplete, 2)
    Require report.Range("B7").Value2 = "홍길동", "privacy value column"
    Require report.Range("A1").Value2 = "개인정보", "privacy title"
    Require report.Range("D6").Value2 = "", "privacy extra columns"
    Require report.Range("C7").Hyperlinks.Count = 1, "privacy source link missing"
    Require report.Range("C7").Hyperlinks(1).SubAddress = source.Range("B2").Address(True, True, xlA1, True), "privacy unsaved source target"
    Require report.Range("C8").Hyperlinks.Count = 0, "privacy hidden link must not mislead"
    Require hidden.Visible = xlSheetHidden, "privacy changed hidden source"
    Require InStr(CStr(report.Range("C8").Value2), "[숨김") > 0, "privacy hidden explanation missing"
    Require report.Range("C7").Font.Underline = xlUnderlineStyleSingle, "privacy link appearance missing"
    report.Range("C7").Hyperlinks(1).Follow
    Require ActiveWorkbook Is sourceBook, "privacy unsaved link opened wrong workbook"
    Require ActiveSheet Is source, "privacy unsaved link opened wrong sheet"
    Require ActiveCell.Address = "$B$2", "privacy unsaved link opened wrong cell"
    report.Parent.Close False
    sourceBook.SaveAs Environ$("LHEXCEL_PROFILE_ROOT") & "\privacy-source.xlsx", xlOpenXMLWorkbook
    Set report = NxPrivacyCreateWorkbookReport(sourceBook, findings, incomplete, 2)
    Require Len(report.Range("C7").Hyperlinks(1).Address) > 0, "privacy saved file path missing"
    report.Range("C7").Hyperlinks(1).Follow
    Require ActiveWorkbook Is sourceBook, "privacy saved link opened wrong workbook"
    Require ActiveSheet Is source, "privacy saved link opened wrong sheet"
    Require ActiveCell.Address = "$B$2", "privacy saved link opened wrong cell"
    report.Parent.Close False
    source.Activate
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not report Is Nothing Then report.Parent.Close False
    On Error GoTo 0
    Err.Raise failure, "CheckPrivacyReport", detail
End Sub

Private Sub Initialize(ByVal cells As Range)
    With cells
        .Value2 = "keep value"
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Italic = True
        .Font.Underline = xlUnderlineStyleSingle
        .Font.Strikethrough = True
        .Font.Color = RGB(83, 62, 121)
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(231, 235, 240)
        .NumberFormat = "0.000"
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Locked = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(25, 45, 65)
        .RowHeight = 23
        .ColumnWidth = 14
    End With
End Sub

Private Sub CheckCell(ByVal cell As Range, ByVal formatIndex As Long)
    Dim expectedSize As Long, expectedFont As Long, expectedFill As Long, edge As Variant
    expectedSize = 11
    expectedFont = RGB(83, 62, 121)
    expectedFill = RGB(231, 235, 240)
    Select Case formatIndex
        Case 0: expectedSize = 9
        Case 1: expectedSize = 10
        Case 2: expectedSize = 12
        Case 3: expectedSize = 15
        Case 4: expectedFont = RGB(0, 0, 0)
        Case 5: expectedFont = RGB(255, 0, 0)
        Case 6: expectedFont = RGB(0, 0, 255)
        Case 7: expectedFont = RGB(0, 128, 0)
        Case 8: expectedFill = RGB(219, 219, 219)
        Case 9: expectedFill = RGB(255, 202, 208)
        Case 10: expectedFill = RGB(255, 236, 161)
        Case 11: expectedFill = RGB(201, 240, 208)
    End Select
    Require cell.Font.Size = expectedSize, "font size"
    Require cell.Font.Color = expectedFont, "font color"
    Require cell.Interior.Color = expectedFill, "fill color"
    Require cell.Font.Name = "Arial", "font name"
    Require cell.Font.Bold And cell.Font.Italic And cell.Font.Strikethrough, "font flags"
    Require cell.Font.Underline = xlUnderlineStyleSingle, "underline"
    Require cell.Interior.Pattern = xlSolid, "fill pattern"
    Require cell.NumberFormat = "0.000", "number format"
    Require cell.HorizontalAlignment = xlRight, "horizontal alignment"
    Require cell.VerticalAlignment = xlCenter, "vertical alignment"
    Require cell.WrapText And cell.Locked, "wrap/locked"
    Require cell.RowHeight = 23 And cell.ColumnWidth = 14, "dimensions"
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        Require cell.Borders(edge).LineStyle = xlContinuous, "border style"
        Require cell.Borders(edge).Weight = xlThin, "border weight"
        Require cell.Borders(edge).Color = RGB(25, 45, 65), "border color"
    Next edge
    If Not cell.HasFormula Then Require cell.Value2 = "keep value", "value changed"
End Sub

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 2268, "T_R68QuickFormat", detail
End Sub
