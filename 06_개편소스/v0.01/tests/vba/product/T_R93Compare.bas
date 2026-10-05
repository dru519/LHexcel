Attribute VB_Name = "T_R93Compare"
Option Explicit
Public Function Names() As String
    Names = "sparse_90000|dense_5000|semantics|scalar|rollover|cancel"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim a As Workbook, b As Workbook, report As Workbook, detail As Worksheet
    Dim source As Range, result As CNxResult, item As Worksheet, expected As Long, actual As Long
    Dim root As String, pa As String, pb As String, rows As Long, cols As Long, r As Long
    Dim values As Variant, started As Double, elapsed As Double, h As Integer, oldScreen As Boolean
    Dim path As String, sheetName As String, address As String, issue As String
    Dim sample As Long, sampleCount As Long
    Dim progress As C_R93CompareProgress, priorStatus As Variant, priorCancel As XlEnableCancelKey
    Dim bookCount As Long, cancelled As Long
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating: Application.ScreenUpdating = False
    root = Environ$("LHEXCEL_PROFILE_ROOT")
    pa = root & "\" & name & "-a.xlsx": pb = root & "\" & name & "-b.xlsx"
    Set a = Workbooks.Add(xlWBATWorksheet): Set b = Workbooks.Add(xlWBATWorksheet)
    a.Worksheets(1).Name = "자료'표": b.Worksheets(1).Name = "자료'표"
    rows = 100: cols = 50
    If name = "sparse_90000" Then rows = 1800
    If name = "rollover" Then rows = 660
    If name = "semantics" Then rows = 4: cols = 2
    If name = "scalar" Then rows = 1: cols = 1
    a.Worksheets(1).Cells(1, 1).Resize(rows, cols).Value2 = 1
    b.Worksheets(1).Cells(1, 1).Resize(rows, cols).Value2 = 2
    expected = rows * cols
    If name = "sparse_90000" Then
        values = a.Worksheets(1).Cells(1, 1).Resize(rows, cols).Value2
        For r = 1 To rows Step 2: values(r, 1) = 2: Next r
        b.Worksheets(1).Cells(1, 1).Resize(rows, cols).Value2 = values
        expected = 900
    End If
    If name = "semantics" Then
        a.Worksheets(1).Range("A1").Formula = "=1+2"
        b.Worksheets(1).Range("A1").Value2 = "'=1+2"
        a.Worksheets(1).Range("A2").Value = CVErr(xlErrNA)
        b.Worksheets(1).Range("A2").Value = CVErr(xlErrDiv0)
        a.Worksheets(1).Range("A3").NumberFormat = "0.00"
    End If
    a.SaveAs pa, xlOpenXMLWorkbook: b.SaveAs pb, xlOpenXMLWorkbook
    If name = "cancel" Then
        Set progress = New C_R93CompareProgress
        progress.StopStage = "원본 이동 링크 작성"
        If InStr(NxProductVersionText(), "r92") > 0 Then progress.StopStage = "셀 비교"
        priorStatus = Application.StatusBar: priorCancel = Application.EnableCancelKey
        bookCount = Workbooks.Count
        NxWorkbookCompareBindProgress progress
        On Error Resume Next
        Set result = NxWorkbookCompareReport(pa, pb, vbNullString, False)
        cancelled = Err.Number: Err.Clear
        On Error GoTo Failed
        If cancelled <> 18 Or Not progress.Fired Then Err.Raise 5, , "cancel event not honored"
        If Workbooks.Count <> bookCount Or Not a.Saved Or Not b.Saved Then Err.Raise 5, , "cancel left output or changed source"
        If Application.StatusBar <> priorStatus Or Application.EnableCancelKey <> priorCancel Then Err.Raise 5, , "cancel state not restored"
        RunCase = "PASS|" & name
        GoTo Clean
    End If
    sampleCount = 1
    If name = "sparse_90000" Or name = "dense_5000" Then sampleCount = 3
    For sample = 1 To sampleCount
    actual = 0
    started = Timer
    Set result = NxWorkbookCompareReport(pa, pb, vbNullString, False)
    NxWorkbookCompareFinishReport True
    elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400#
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
    Set report = ActiveWorkbook
    If report Is a Or report Is b Then Err.Raise 5, , "missing output"
    If report.Worksheets("비교요약").Range("B4").Value2 <> expected Then Err.Raise 5, , "difference count"
    If InStr(NxProductVersionText(), "r93") > 0 Then CheckOpenTable report.Worksheets("비교요약").Range("A5:F6")
    For Each item In report.Worksheets
        If Left$(item.Name, 3) = "셀차이" Then
            r = item.Cells(item.Rows.Count, 1).End(xlUp).Row - 1
            actual = actual + r
            If InStr(NxProductVersionText(), "r93") > 0 Then CheckOpenTable item.Range("A1:O" & CStr(r + 1))
            If item.Hyperlinks.Count <> r * 2 + 1 Then Err.Raise 5, , "link count"
            If Not NxWorkbookCompareNavigateSnapshot(item, 2, 0) Then Err.Raise 5, , "snapshot link resolve"
            If ActiveCell.Address(False, False) <> CStr(item.Range("B2").Value2) Then Err.Raise 5, , "link target"
            If name = "semantics" Then
                If item.Range("F2").HasFormula Then Err.Raise 5, , "report formula injection"
                If InStr(CStr(item.Range("F2").Value2), "=1+2") = 0 Then Err.Raise 5, , "formula text lost"
            End If
        End If
    Next item
    If actual <> expected Then Err.Raise 5, , "buffer tail lost"
    If Not a.Saved Or Not b.Saved Then Err.Raise 5, , "source saved changed"
    h = FreeFile
    Open root & "\compare-r93.tsv" For Append As #h
    Print #h, name & vbTab & Format$(elapsed, "0.000") & vbTab & expected
    Close #h: h = 0
    report.Close False: Set report = Nothing
    Next sample
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & "|" & Err.Description
Clean:
    On Error Resume Next
    If h <> 0 Then Close #h
    NxWorkbookCompareFinishReport False
    If Not report Is Nothing Then report.Close False
    If Not a Is Nothing Then a.Close False
    If Not b Is Nothing Then b.Close False
    Application.ScreenUpdating = oldScreen
End Function

Private Sub CheckOpenTable(ByVal target As Range)
    If target.Borders(xlEdgeLeft).LineStyle <> xlNone Or target.Borders(xlEdgeRight).LineStyle <> xlNone Then Err.Raise 5, , "report sides not open"
    If target.Borders(xlEdgeBottom).Weight <> xlThick Or target.Rows(1).Borders(xlEdgeBottom).Weight <> xlThick Then Err.Raise 5, , "report rules not thick"
    If target.Font.Size <> 9 Or target.Rows(1).Interior.Color <> NxDrawTableHeaderFill("MONO") Then Err.Raise 5, , "report style mismatch"
End Sub
