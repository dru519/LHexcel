Attribute VB_Name = "T_R94FormulaTools"
Option Explicit
Private mPanel As FNxFormulaTools, mPanelBook As Workbook

Public Function Names() As String
    Names = "notes_keep|notes_replace|notes_undo|report_sheet|report_book|no_formula|large_reference|partial_reference|protected|popup_notes|popup_report|navigator_icon|thousand_formulas|notes_richtext|reference_status|ten_thousand_formulas|sparse_selection|business_formula"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, sheet As Worksheet, report As Worksheet, other As Worksheet, candidate As Workbook
    Dim panel As FNxFormulaTools, image As Object
    Dim failure As Long, detail As String, count As Long, stage As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet): Set sheet = book.Worksheets(1)
    sheet.Name = "원본 자료"
    sheet.Range("A1:A3").Value2 = 7
    sheet.Range("B1").Formula = "=SUM(A1:A3)"
    sheet.Range("B2").Formula = "=A1*2"
    If name = "notes_keep" Then
        sheet.Range("B1").AddComment "기존 설명"
        NxFormulaNotesApply sheet.Range("B1:B3"), False
        Require sheet.Range("B1").Comment.Text = "기존 설명" & vbCrLf & "=SUM(A1:A3)", "keep text"
        Require Not sheet.Range("B1").Comment.Visible And sheet.Range("B3").Comment Is Nothing, "hidden/formula only"
        NxFormulaNotesApply sheet.Range("B1"), False
        Require sheet.Range("B1").Comment.Text = "기존 설명" & vbCrLf & "=SUM(A1:A3)", "no duplicate"
    ElseIf name = "notes_replace" Then
        sheet.Range("B1").AddComment "기존 설명"
        NxFormulaNotesApply sheet.Range("B1"), True, True
        Require sheet.Range("B1").Comment.Text = "=SUM(A1:A3)" And sheet.Range("B1").Comment.Visible, "replace/show"
    ElseIf name = "notes_undo" Then
        sheet.Range("B1").AddComment "기존 설명"
        sheet.Range("B1").Comment.Visible = True
        NxFormulaNotesApply sheet.Range("B1:B2"), False
        NxFormulaNotesUndoLast
        Require sheet.Range("B1").Comment.Text = "기존 설명" And sheet.Range("B1").Comment.Visible, "restore note"
        Require sheet.Range("B2").Comment Is Nothing, "remove created"
    ElseIf name = "notes_richtext" Then
        stage = "fixture"
        sheet.Range("B1").AddComment "기존 설명과 검토 의견"
        With sheet.Range("B1").Comment.Shape.TextFrame
            .Characters(1, 2).Font.Bold = True
            .Characters(1, 2).Font.Color = RGB(255, 0, 0)
            .Characters(4, 2).Font.Italic = True
            .Characters(4, 2).Font.Size = 14
            .Characters(8, 2).Font.Underline = xlUnderlineStyleSingle
        End With
        stage = "append"
        NxFormulaNotesApply sheet.Range("B1"), False
        For Each candidate In Application.Workbooks
            If Not candidate.IsAddin Then Require candidate.Windows(1).Visible, "hidden note backup exposed as document"
        Next candidate
        stage = "check append"
        RequireNoteFonts sheet.Range("B1")
        stage = "undo append"
        NxFormulaNotesUndoLast
        Require sheet.Range("B1").Comment.Text = "기존 설명과 검토 의견", "rich text undo text"
        RequireNoteFonts sheet.Range("B1")
        stage = "replace"
        NxFormulaNotesApply sheet.Range("B1"), True, True
        stage = "undo replace"
        NxFormulaNotesUndoLast
        RequireNoteFonts sheet.Range("B1")
    ElseIf name = "report_sheet" Or name = "report_book" Then
        Set report = NxFormulaReferenceReport(sheet.Range("B1:B2"), name = "report_book")
        Require report.Range("A1").Value2 = "수식 셀" And report.Range("E1").Value2 = "확인 상태", "korean headers"
        Require report.Range("B2").Value2 = "=SUM(A1:A3)" And Not report.Range("B2").HasFormula, "formula text"
        Require InStr(report.Range("C2").Value2, "$A$1:$A$3") > 0, "range group"
        Require report.Hyperlinks.Count = 4, "links"
        Require InStr(report.Cells(2, 1).Hyperlinks(1).ScreenTip, "원본 파일") > 0, "link recovery guidance"
        report.Cells(2, 1).Hyperlinks(1).Follow NewWindow:=False
        Require ActiveSheet Is sheet, "source link target"
        Require Selection.Address = "$B$1", "source link cell"
        report.Activate
        report.Cells(2, 3).Hyperlinks(1).Follow NewWindow:=False
        Require ActiveSheet Is sheet And Selection.Address = "$A$1:$A$3", "precedent link range"
        report.Activate
        Require report.Range("A1").Font.Size = 9 And report.Range("A1:A3").Borders(xlEdgeLeft).LineStyle = xlNone, "open table"
        If name = "report_book" Then
            Require Not report.Parent Is book, "new book"
            Require report.Parent.Path = "" And book.Worksheets.Count = 1, "unsaved/unchanged source"
        Else
            Require report.Parent Is book, "new sheet"
        End If
    ElseIf name = "no_formula" Then
        count = Workbooks.Count
        On Error Resume Next
        Set report = NxFormulaReferenceReport(sheet.Range("D:D"), True)
        failure = Err.Number: Err.Clear
        On Error GoTo Failed
        Require failure <> 0 And Workbooks.Count = count, "no empty output"
    ElseIf name = "large_reference" Then
        sheet.Range("B1").Formula = "=SUM(A:A)"
        Set report = NxFormulaReferenceReport(sheet.Range("B1"), True)
        Require report.Cells(report.Rows.Count, 1).End(xlUp).Row = 2 And report.Range("D2").Value2 = "1,048,576개 셀", "whole column collapsed"
        Require IsEmpty(report.Range("A3").Value2), "no extra data row"
    ElseIf name = "partial_reference" Then
        Set other = book.Worksheets.Add
        other.Name = "다른 자료"
        sheet.Range("B1").Formula = "=A1+'다른 자료'!A1"
        Set report = NxFormulaReferenceReport(sheet.Range("B1"), True)
        Require InStr(report.Range("E2").Value2, "별도 확인") > 0 Or InStr(report.Range("E2").Value2, "확인 불가") > 0, "partial not false complete"
    ElseIf name = "protected" Then
        sheet.Protect "test"
        On Error Resume Next
        NxFormulaNotesApply sheet.Range("B1"), False
        failure = Err.Number: Err.Clear
        On Error GoTo Failed
        Require failure <> 0 And sheet.Range("B1").Comment Is Nothing, "protected notes"
        sheet.Unprotect "test"
        book.Protect "test", True
        On Error Resume Next
        Set report = NxFormulaReferenceReport(sheet.Range("B1"), False)
        failure = Err.Number: Err.Clear
        On Error GoTo Failed
        Require failure <> 0 And book.Worksheets.Count = 1, "protected structure"
        book.Unprotect "test"
    ElseIf name = "popup_notes" Or name = "popup_report" Then
        Set panel = New FNxFormulaTools
        panel.BindTarget sheet.Range("B1"), name = "popup_report"
        Require panel.Controls("cmdExecute").Enabled And panel.Controls("cboOption").ListCount = 2, "direct options"
        Require panel.Controls("cmdCancel").Cancel, "escape"
        panel.NxProbeExecute
        Set panel = Nothing
        If name = "popup_notes" Then
            Require Not sheet.Range("B1").Comment Is Nothing, "popup apply"
        Else
            Set report = ActiveSheet
            Require Not report.Parent Is book, "popup default new workbook"
        End If
    ElseIf name = "thousand_formulas" Then
        sheet.Range("B1:B1000").FormulaR1C1 = "=RC[-1]*2"
        Set report = NxFormulaReferenceReport(sheet.Range("B1:B1000"), True)
        Require report.Cells(report.Rows.Count, 1).End(xlUp).Row = 1001, "complete report"
        Require Not report.Range("B2:B1001").HasFormula And report.Hyperlinks.Count = 2000, "bulk text and links"
    ElseIf name = "ten_thousand_formulas" Then
        sheet.Range("B1:B10000").FormulaR1C1 = "=IFERROR(RC[-1]*1.1,0)"
        Set report = NxFormulaReferenceReport(sheet.Range("B1:B10000"), True)
        Require report.Cells(report.Rows.Count, 1).End(xlUp).Row = 10001, "large output complete"
        Require report.Hyperlinks.Count = 20000 And Not report.Range("B10001").HasFormula, "large output links/text"
    ElseIf name = "sparse_selection" Then
        sheet.Range("B2").ClearContents
        sheet.Range("AX1800").Formula = "=SUM(A1:A3)"
        Set report = NxFormulaReferenceReport(sheet.Range("A1:AX1800"), True)
        Require report.Cells(report.Rows.Count, 1).End(xlUp).Row = 3, "90000 cells only two formulas"
    ElseIf name = "business_formula" Then
        sheet.Range("B1").Formula = "=IFERROR(SUMIF(A1:A3,"">0"",C1:C3),0)"
        Set report = NxFormulaReferenceReport(sheet.Range("B1"), True)
        Require report.Range("B2").Value2 = sheet.Range("B1").Formula And Not report.Range("B2").HasFormula, "business formula unchanged"
        Require report.Hyperlinks.Count >= 2, "business references"
    ElseIf name = "reference_status" Then
        Require InStr(NxFormulaReferenceStatus("=A1+'다른 자료'!A1", True), "다른 시트 참조") > 0, "cross sheet status"
        Require InStr(NxFormulaReferenceStatus("='C:\[원본.xlsx]자료'!A1", False), "외부 파일 참조") > 0, "external status"
        Require InStr(NxFormulaReferenceStatus("=INDIRECT(""A1"")", False), "동적 참조") > 0, "indirect status"
        Require InStr(NxFormulaReferenceStatus("=OFFSET(A1,1,0)", True), "동적 참조") > 0, "offset status"
        Require InStr(NxFormulaReferenceStatus("=SUM(매출[금액])", True), "표 구조 참조") > 0, "table status"
        Require NxFormulaReferenceStatus("=IF(A1>0,""!"",""["")", True) = "같은 시트의 직접 참조 확인", "ignore quoted punctuation"
    ElseIf name = "navigator_icon" Then
        Set image = Application.CommandBars.GetImageMso("FindDialog", 16, 16)
        Require Not image Is Nothing, "icon available"
    Else
        Err.Raise 5, , "unknown case"
    End If
    If Not report Is Nothing Then
        If Not report.Parent Is book Then report.Parent.Close False
    End If
    book.Close False
    RunCase = "PASS|" & name
    Exit Function
Failed:
    detail = stage & ": " & Err.Source & ": " & Err.Description
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    If Not report Is Nothing Then
        If Not report.Parent Is book Then report.Parent.Close False
    End If
    If Not book Is Nothing Then book.Close False
    RunCase = "FAIL|" & name & "|" & detail
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise 5, "T_R94FormulaTools", detail
End Sub

Private Sub RequireNoteFonts(ByVal cell As Range)
    With cell.Comment.Shape.TextFrame
        Require .Characters(1, 2).Font.Bold And .Characters(1, 2).Font.Color = RGB(255, 0, 0), "bold/red restored"
        Require .Characters(4, 2).Font.Italic And .Characters(4, 2).Font.Size = 14, "italic/size restored"
        Require CBool(.Characters(8, 2).Font.Underline), "underline restored"
    End With
End Sub

Public Sub OpenPanel(ByVal references As Boolean)
    Set mPanelBook = Workbooks.Add(xlWBATWorksheet)
    mPanelBook.Worksheets(1).Range("A1").Formula = "=SUM(B1:B3)"
    Set mPanel = New FNxFormulaTools
    mPanel.BindTarget mPanelBook.Worksheets(1).Range("A1"), references
    mPanel.Show vbModeless
End Sub
Public Function Caption() As String
    Caption = mPanel.Caption
End Function
Public Sub ClosePanel()
    Unload mPanel
    Set mPanel = Nothing
    mPanelBook.Close False
    Set mPanelBook = Nothing
End Sub

Public Sub PrepareNativeUndo()
    Set mPanelBook = Workbooks.Add(xlWBATWorksheet)
    mPanelBook.Worksheets(1).Range("A1").Formula = "=1+2"
    NxFormulaNotesApply mPanelBook.Worksheets(1).Range("A1"), False
End Sub
Public Function VerifyNativeUndo() As String
    Require mPanelBook.Worksheets(1).Range("A1").Comment Is Nothing, "native undo note"
    mPanelBook.Close False
    Set mPanelBook = Nothing
    VerifyNativeUndo = "PASS|native_notes_undo"
End Function
