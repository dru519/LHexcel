Attribute VB_Name = "T_R94Functions"
Option Explicit
Private mPanel As FNxFunctionWrap
Private mPanelBook As Workbook
Public FailAfterFirst As Boolean

Public Function Names() As String
    Names = "formula_preservation|numeric_modes|errors|reapply|selection|undo|rollback|popup_round|popup_iferror|popup_precision"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, sheet As Worksheet, r As Range, mode As Variant, digits As Variant
    Dim expected As Double, count As Long, failure As Long, detail As String, original As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet): Set sheet = book.Worksheets(1)
    Set r = sheet.Range("B2")
    If name = "formula_preservation" Then
        original = "=SUM($A$1,A2,IF(A3=""x,y"",2,3))"
        Require NxWrappedFormula(original, "ROUND", "2") = "=ROUND(SUM($A$1,A2,IF(A3=""x,y"",2,3)),2)", "references/string"
        Require NxWrappedFormula("=ROUND(IF(A1=""x,y"",SUM(A2,A3),0),2)", "ROUND", "0") = "=ROUND(IF(A1=""x,y"",SUM(A2,A3),0),0)", "nested reapply"
        Require NxWrappedFormula("=ROUND(A1,2)+SUM(A2,A3)", "ROUND", "0") = "=ROUND(ROUND(A1,2)+SUM(A2,A3),0)", "outer expression"
        Require NxWrappedFormula("=ROUND(SUM({1,2;3,4}),2)", "ROUND", "0") = "=ROUND(SUM({1,2;3,4}),0)", "array constant"
        Require NxFunctionArgument(True, 2, "a""b") = """a""""b""", "quote escaping"
    ElseIf name = "numeric_modes" Then
        For Each mode In Array("ROUND", "ROUNDUP", "ROUNDDOWN")
            For Each digits In Array("4", "3", "2", "1", "0", "-1", "-2", "-3", "-4", "-5")
                r.Value2 = -1110.456
                NxFunctionApply r, CStr(mode), CStr(digits), False
                expected = sheet.Evaluate("=" & CStr(mode) & "(-1110.456," & CStr(digits) & ")")
                Require r.HasFormula And Abs(r.Value2 - expected) < 0.00001, "numeric " & mode & digits
            Next digits
        Next mode
        On Error Resume Next
        original = NxFunctionArgument(False, 0, "1.5")
        failure = Err.Number: Err.Clear
        On Error GoTo Failed
        Require failure <> 0, "invalid digits"
    ElseIf name = "errors" Then
        r.Formula = "=1/0"
        NxFunctionApply r, "ROUND", "2", False
        Require IsError(r.Value2) And InStr(r.Formula, "IFERROR") = 0, "round must preserve error"
        NxFunctionApply r, "IFERROR", NxFunctionArgument(True, 0, ""), False
        Require r.Value2 = "", "blank error result"
        NxFunctionApply r, "IFERROR", "0", False
        Require r.Value2 = 0, "numeric error result"
        NxFunctionApply r, "IFERROR", NxFunctionArgument(True, 2, "확인"), False
        Require r.Value2 = "확인", "text error result"
    ElseIf name = "reapply" Then
        r.Formula = "=ROUND(12.345,2)"
        NxFunctionApply r, "ROUND", "0", False
        Require r.Formula = "=ROUND(12.345,0)", "same outer replacement"
        NxFunctionApply r, "ROUNDUP", "1", False
        Require r.Formula = "=ROUNDUP(ROUND(12.345,0),1)", "different outer wrapper"
        r.Formula = "=IFERROR(1/0,0)"
        NxFunctionApply r, "ROUND", "2", False
        Require r.Formula = "=ROUND(IFERROR(1/0,0),2)", "existing iferror preserved"
    ElseIf name = "selection" Then
        sheet.Range("A1:A4").Value2 = 12.34
        sheet.Range("A5").Value2 = "text"
        sheet.Rows(2).Hidden = True
        sheet.Range("A4").FormulaArray = "=SUM(1,2)"
        NxFunctionApply sheet.Range("A1:A6"), "ROUND", "0", False
        Require sheet.Range("A1").HasFormula And Not sheet.Range("A2").HasFormula, "hidden exclusion"
        Require sheet.Range("A4").HasArray And sheet.Range("A5").Value2 = "text" And IsEmpty(sheet.Range("A6")), "excluded values"
        NxFunctionApply Application.Union(sheet.Range("A2"), sheet.Range("A3")), "ROUND", "1", True
        Require sheet.Range("A2").HasFormula, "hidden inclusion/multiarea"
        sheet.Protect "test"
        On Error Resume Next
        NxFunctionApply sheet.Range("A1"), "ROUND", "0", False
        failure = Err.Number: Err.Clear
        On Error GoTo Failed
        Require failure <> 0, "protected write"
        sheet.Unprotect "test"
    ElseIf name = "undo" Then
        sheet.Range("A1").Value2 = 12.345
        sheet.Range("A2").Formula = "=$A$1*2"
        sheet.Range("A1:A2").NumberFormat = "0.000%"
        NxFunctionApply sheet.Range("A1:A2"), "ROUND", "0", False
        NxFunctionUndoLast
        Require sheet.Range("A1").Value2 = 12.345 And Not sheet.Range("A1").HasFormula, "constant undo"
        Require sheet.Range("A2").Formula = "=$A$1*2", "formula undo"
        Require sheet.Range("A1:A2").NumberFormat = "0.000%", "format preserved"
    ElseIf name = "rollback" Then
        sheet.Range("A1").Value2 = 12.345
        sheet.Range("A2").Formula = "=$A$1*2"
        FailAfterFirst = True
        On Error Resume Next
        NxFunctionApply sheet.Range("A1:A2"), "ROUND", "0", False
        failure = Err.Number: Err.Clear
        On Error GoTo Failed
        Require failure <> 0 And Not FailAfterFirst, "fault not exercised"
        Require sheet.Range("A1").Value2 = 12.345 And Not sheet.Range("A1").HasFormula, "partial constant rollback"
        Require sheet.Range("A2").Formula = "=$A$1*2", "partial formula rollback"
    ElseIf name = "popup_precision" Then
        Set mPanel = New FNxFunctionWrap
        r.Value2 = 123456.789
        mPanel.BindTarget r, False
        Require mPanel.Controls("cboPrecision").ListCount = 10, "precision presets"
        Require mPanel.Controls("lblSample").BackColor = vbWhite, "white preview"
        For count = 0 To 8
            mPanel.Controls("cboPrecision").ListIndex = count
            Require mPanel.Controls("txtArgument").Value = CStr(4 - count), "preset mapping"
            Require Not mPanel.Controls("txtArgument").Enabled, "preset input state"
        Next count
        mPanel.Controls("cboPrecision").ListIndex = 9
        Require mPanel.Controls("txtArgument").Enabled, "custom input state"
        mPanel.Controls("txtArgument").Value = "1.5"
        Require Not mPanel.Controls("cmdExecute").Enabled, "fractional precision rejection"
        mPanel.Controls("txtArgument").Value = "-5"
        Require mPanel.Controls("cmdExecute").Enabled, "custom precision accepted"
        mPanel.NxProbeExecute
        Set mPanel = Nothing
        Require r.Formula = "=ROUND(123456.789,-5)", "custom precision execution"
    ElseIf name = "popup_round" Or name = "popup_iferror" Then
        Set mPanel = New FNxFunctionWrap
        r.Formula = "=1/0"
        mPanel.BindTarget r, name = "popup_iferror"
        Require mPanel.Controls("cmdExecute").Enabled, "execute requires preview"
        Require InStr(mPanel.Controls("lblSample").Caption, "적용 1개") > 0, "automatic preview"
        mPanel.NxProbeExecute
        Set mPanel = Nothing
        If name = "popup_round" Then
            Require r.Formula = "=ROUND(1/0,2)" And IsError(r.Value2), "round popup"
        Else
            Require r.Formula = "=IFERROR(1/0,"""")" And r.Value2 = "", "iferror popup"
        End If
    Else
        Err.Raise 5, , "unknown case"
    End If
    book.Close False
    RunCase = "PASS|" & name
    Exit Function
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    RunCase = "FAIL|" & name & "|" & detail
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise 5, "T_R94Functions", detail
End Sub

Public Sub OpenPanel(ByVal errorMode As Boolean)
    Set mPanelBook = Workbooks.Add(xlWBATWorksheet)
    mPanelBook.Worksheets(1).Range("A1").Formula = "=1234.567/2"
    Set mPanel = New FNxFunctionWrap
    mPanel.BindTarget mPanelBook.Worksheets(1).Range("A1"), errorMode
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
