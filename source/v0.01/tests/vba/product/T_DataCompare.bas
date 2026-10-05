Attribute VB_Name = "T_DataCompare"
Option Explicit
Public Function NxTestFileCompareRoute() As String
    Dim a As Workbook, b As Workbook, source As Workbook, report As Workbook
    Dim options As Object, result As CNxResult, leftPath As String, rightPath As String
    Dim before As Long, sheet As Worksheet, links As Long, form As FNxWorkbookCompare
    On Error GoTo Failed
    Set source = ActiveWorkbook
    leftPath = Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator & "r88-left.xlsx"
    rightPath = Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator & "r88-right.xlsx"
    Set a = Workbooks.Add(xlWBATWorksheet): a.Worksheets(1).Name = "공통"
    a.Worksheets(1).Range("A1:C3").Value2 = 1
    Set sheet = a.Worksheets.Add: sheet.Name = "기준에만": sheet.Range("A1").Value2 = 7
    a.SaveAs leftPath, xlOpenXMLWorkbook: a.Close False: Set a = Nothing
    Set b = Workbooks.Add(xlWBATWorksheet): b.Worksheets(1).Name = "공통"
    b.Worksheets(1).Range("A1:C3").Value2 = 1: b.Worksheets(1).Range("B2").Value2 = 9
    b.SaveAs rightPath, xlOpenXMLWorkbook: b.Close False: Set b = Nothing
    source.Activate: before = Workbooks.Count
    Set form = New FNxWorkbookCompare: form.BindFeature
    If form.Caption <> "내엑셀 - 파일 비교" Or Not form.cmdExecute.Enabled Then Err.Raise 5, , "file popup/direct execution unavailable"
    Unload form: Set form = Nothing
    Set options = CreateObject("Scripting.Dictionary")
    options.Add "base_path", leftPath: options.Add "compare_path", rightPath
    options.Add "compare_formats", False
    Set result = NxFileRun("NX-FILE-FILE-COMPARE", "", workflowOptions:=options)
    If result Is Nothing Then Err.Raise 5, , "file route missing result"
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
    If Workbooks.Count <> before + 1 Then Err.Raise 5, , "file route input leak or missing report"
    Set report = ActiveWorkbook
    If report Is source Or Len(report.Path) <> 0 Then Err.Raise 5, , "report not unsaved"
    For Each sheet In report.Worksheets: links = links + sheet.Hyperlinks.Count: Next sheet
    If links = 0 Then Err.Raise 5, , "report navigation links missing"
    report.Close False: Set report = Nothing
    NxTestFileCompareRoute = "PASS|file-route|unsaved|links=" & links
    Exit Function
Failed:
    NxTestFileCompareRoute = "FAIL|file-route|" & Err.Number & "|" & Err.Description
    On Error Resume Next
    If Not form Is Nothing Then Unload form
    If Not report Is Nothing Then report.Close False
    If Not a Is Nothing Then a.Close False
    If Not b Is Nothing Then b.Close False
    NxWorkbookCompareFinishReport False
End Function
Public Sub NxTestDataCompareUiData()
    Dim book As Workbook, sheet As Worksheet
    Set book = ActiveWorkbook
    Set sheet = book.Worksheets(1)
    sheet.Range("A1:C3").Value2 = 10
    sheet.Range("E1:G3").Value2 = 10
    sheet.Range("F2").Value2 = 20
    sheet.Range("A1:C3").Select
    book.Worksheets.Add(After:=sheet).Range("A1:C3").Value2 = 20
    sheet.Activate
    book.Saved = True
End Sub
Public Function NxTestDataCompareCancel() As String
    Dim source As Workbook, sheet As Worksheet, probe As CNxDataCompareCancelProbe, result As CNxResult
    Dim options As Object, bookCount As Long, started As Double, elapsed As Double
    On Error GoTo Failed
    Set source = Workbooks.Add(xlWBATWorksheet): Set sheet = source.Worksheets(1)
    sheet.Range("A1:AX20000").Value2 = 1
    source.Saved = True: sheet.Range("A1").Select: bookCount = Workbooks.Count
    Set probe = New CNxDataCompareCancelProbe
    NxDataCompareBind probe
    Set options = CreateObject("Scripting.Dictionary")
    options.Add "data_compare", True: options.Add "left_range", sheet.Range("A1:AX20000")
    options.Add "right_range", sheet.Range("A1:AX20000"): options.Add "compare_formulas", False
    started = Timer
    Set result = NxFileRun("NX-FILE-WORKBOOK-COMPARE", "", workflowOptions:=options)
    elapsed = Timer - started
    If result.Outcome = NxSuccess Then Err.Raise 5, , "cancel ignored"
    If Workbooks.Count <> bookCount Or Not source.Saved Then Err.Raise 5, , "cancel residue/source"
    source.Close False
    NxTestDataCompareCancel = "PASS|cancel|" & elapsed & "|" & probe.Calls
    Exit Function
Failed:
    NxTestDataCompareCancel = "FAIL|cancel|" & Err.Number & "|" & Err.Description
    On Error Resume Next
    NxDataCompareFinish False
    If Not source Is Nothing Then source.Close False
End Function

Public Function NxTestDataCompareEdges() As String
    Dim source As Workbook, output As Workbook, ls As Worksheet, rs As Worksheet, l As Range, r As Range
    Dim options As Object, result As CNxResult, form As FNxDataCompare, stage As String, registry As CNxFeatureRegistry
    On Error GoTo Failed
    Set registry = NxCreateProductRegistry()
    If registry.FeatureCount <> NxProductSurfaceFeatureCount() Then Err.Raise 5, , "feature registry"
    Set source = Workbooks.Add(xlWBATWorksheet): Set ls = source.Worksheets(1)
    Set rs = source.Worksheets.Add(After:=ls)
    ls.Range("B2").Value2 = 1: rs.Range("D4").Value2 = 1
    ls.Range("C2").NumberFormat = "@": ls.Range("C2").Value2 = "=1+1"
    rs.Range("E4").NumberFormat = "@": rs.Range("E4").Value2 = "=1+1"
    ls.Range("B3").Value2 = CVErr(xlErrNA): rs.Range("D5").Value2 = CVErr(xlErrNA)
    ls.Range("C3").Value2 = "001": rs.Range("E5").Value2 = 1
    ls.Range("B4").Value2 = 5
    Set l = ls.Range("B2:C4"): Set r = rs.Range("D4:F5")
    ls.Rows(3).Hidden = True: rs.Columns(4).Hidden = True
    ls.Protect: rs.Protect
    source.Saved = True: ls.Activate: ls.Range("A1").Select
    stage = "range"
    Set options = CreateObject("Scripting.Dictionary")
    options.Add "data_compare", True: options.Add "left_range", l: options.Add "right_range", r: options.Add "compare_formulas", False
    Set result = NxFileRun("NX-FILE-WORKBOOK-COMPARE", "", workflowOptions:=options)
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
    Set output = ActiveWorkbook
    If Not source.Saved Then Err.Raise 5, , "protected source changed"
    If output.Worksheets(1).Range("B5").HasFormula Then Err.Raise 5, , "formula injection"
    If output.Worksheets(1).Range("B5").Value2 <> "=1+1" Then Err.Raise 5, , "literal changed"
    If output.Worksheets(1).Range("A6").Interior.Color = RGB(255, 199, 206) Then Err.Raise 5, , "equal error marked"
    If output.Worksheets(1).Range("C7").Interior.Color = RGB(255, 199, 206) Then Err.Raise 5, , "both absent marked"
    If output.Worksheets(2).Range("A7").Interior.Color <> RGB(255, 199, 206) Then Err.Raise 5, , "missing target unmarked"
    output.Close False: Set output = Nothing
    stage = "sheet"
    ls.Unprotect: rs.Unprotect: ls.Cells.Clear: rs.Cells.Clear
    ls.Range("A1").Formula = "=1+1": rs.Range("A1").Formula = "=2"
    Set l = NxDataCompareUsedRange(ls): Set r = NxDataCompareUsedRange(rs)
    source.Saved = True: ls.Activate: ls.Range("A1").Select
    Set options.Item("left_range") = l: Set options.Item("right_range") = r
    options.Item("compare_formulas") = True
    Set result = NxFileRun("NX-FILE-SHEET-COMPARE", "", workflowOptions:=options)
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
    Set output = ActiveWorkbook
    If output.Worksheets(1).Range("A5").Interior.Color <> RGB(255, 199, 206) Then Err.Raise 5, , "formula difference"
    output.Close False: Set output = Nothing
    stage = "forms"
    Set form = New FNxDataCompare: form.BindFeature "NX-FILE-WORKBOOK-COMPARE"
    If Not form.Controls("txtLeft").Visible Or form.Controls("cboLeft").Visible Or Not form.Controls("cmdLeft").Visible Then Err.Raise 5, , "range form"
    If Not form.Controls("txtRight").Visible Or form.Controls("cboRight").Visible Or Not form.Controls("cmdRight").Visible Then Err.Raise 5, , "right range form"
    Unload form
    Set form = New FNxDataCompare: form.BindFeature "NX-FILE-SHEET-COMPARE"
    If Not form.Controls("cboLeft").Visible Or form.Controls("txtLeft").Visible Or form.Controls("cmdLeft").Visible Then Err.Raise 5, , "sheet form"
    If Not form.Controls("cboRight").Visible Or form.Controls("txtRight").Visible Or form.Controls("cmdRight").Visible Then Err.Raise 5, , "right sheet form"
    Unload form
    source.Close False
    NxTestDataCompareEdges = "PASS|edges": Exit Function
Failed:
    NxTestDataCompareEdges = "FAIL|edges|" & stage & "|" & Err.Number & "|" & Err.Description
    On Error Resume Next
    NxDataCompareFinish False
    If Not output Is Nothing Then output.Close False
    If Not source Is Nothing Then source.Close False
End Function

Public Function NxTestDataCompare(ByVal cells As Long, ByVal pattern As String) As String
    Dim source As Workbook, output As Workbook, leftSheet As Worksheet, rightSheet As Worksheet
    Dim leftRange As Range, rightRange As Range, options As Object, result As CNxResult
    Dim values() As Variant, expected() As Variant, actual As Variant, row As Long, col As Long, rows As Long
    Dim differences As Long, started As Double, elapsed As Double, metrics As String
    Dim priorScreen As Boolean, priorEvents As Boolean, priorCalculation As XlCalculation
    Dim errText As String, stage As String
    On Error GoTo Failed
    priorScreen = Application.ScreenUpdating: priorEvents = Application.EnableEvents: priorCalculation = Application.Calculation
    rows = cells \ 50
    Set source = Workbooks.Add(xlWBATWorksheet)
    Set leftSheet = source.Worksheets(1)
    Set rightSheet = source.Worksheets.Add(After:=leftSheet)
    ReDim values(1 To rows, 1 To 50): ReDim expected(1 To rows, 1 To 50)
    For row = 1 To rows
        For col = 1 To 50
            values(row, col) = row * 50 + col
            expected(row, col) = values(row, col)
            If pattern = "all" Or (pattern = "sparse" And row Mod 100 = 0 And col = 25) Or (pattern = "checker" And (row + col) Mod 2 = 0) Then
                expected(row, col) = -values(row, col): differences = differences + 1
            End If
        Next col
    Next row
    Set leftRange = leftSheet.Range("A1").Resize(rows, 50): Set rightRange = rightSheet.Range("A1").Resize(rows, 50)
    leftRange.Value2 = values: rightRange.Value2 = expected
    source.Saved = True: leftSheet.Activate: leftSheet.Range("A1").Select
    If cells = 1000000 And pattern = "same" Then leftRange.Select
    Set options = CreateObject("Scripting.Dictionary")
    options.Add "data_compare", True
    options.Add "left_range", leftRange: options.Add "right_range", rightRange: options.Add "compare_formulas", False
    stage = "run": started = Timer
    Set result = NxFileRun("NX-FILE-WORKBOOK-COMPARE", vbNullString, workflowOptions:=options)
    elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400#
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
    metrics = NxDataCompareMetrics()
    stage = "check"
    Set output = ActiveWorkbook
    If output Is source Then Err.Raise 5, , "result not activated"
    If output.Worksheets.Count <> 2 Or Len(output.Path) <> 0 Then Err.Raise 5, , "result structure"
    If CLng(Split(metrics, "|")(4)) <> differences Then Err.Raise 5, , "difference count"
    actual = output.Worksheets(1).Range("A5").Resize(rows, 50).Value2
    For row = 1 To rows
        For col = 1 To 50
            If CStr(actual(row, col)) <> CStr(values(row, col)) Then Err.Raise 5, , "left value"
        Next col
    Next row
    actual = output.Worksheets(2).Range("A5").Resize(rows, 50).Value2
    For row = 1 To rows
        For col = 1 To 50
            If CStr(actual(row, col)) <> CStr(expected(row, col)) Then Err.Raise 5, , "right value"
        Next col
    Next row
    If pattern = "all" Then
        If output.Worksheets(1).Range("A5").Resize(rows, 50).Interior.Color <> RGB(255, 199, 206) Then Err.Raise 5, , "all paint"
    ElseIf pattern = "sparse" Then
        If output.Worksheets(1).Cells(104, 25).Interior.Color <> RGB(255, 199, 206) Then Err.Raise 5, , "sparse paint"
        If output.Worksheets(2).Cells(104, 25).Interior.Color <> RGB(255, 199, 206) Then Err.Raise 5, , "right paint"
        If output.Worksheets(1).Range("A5").Interior.Color = RGB(255, 199, 206) Then Err.Raise 5, , "false paint"
    ElseIf pattern = "same" Then
        If output.Worksheets(1).Range("A5").Resize(rows, 50).Interior.Color = RGB(255, 199, 206) Then Err.Raise 5, , "same paint"
    End If
    If Not source.Saved Then Err.Raise 5, , "source dirty"
    actual = leftRange.Value2
    For row = 1 To rows
        For col = 1 To 50
            If actual(row, col) <> values(row, col) Then Err.Raise 5, , "source mutation"
        Next col
    Next row
    If Application.ScreenUpdating <> priorScreen Or Application.EnableEvents <> priorEvents Or Application.Calculation <> priorCalculation Then Err.Raise 5, , "application state"
    output.Close False: source.Close False
    NxTestDataCompare = "PASS|" & cells & "|" & pattern & "|" & elapsed & "|" & metrics
    Exit Function
Failed:
    errText = Err.Number & "|" & Err.Description
    On Error Resume Next
    NxDataCompareFinish False
    If Not output Is Nothing Then output.Close False
    If Not source Is Nothing Then source.Close False
    NxTestDataCompare = "FAIL|" & stage & "|" & errText
End Function
