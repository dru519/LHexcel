Attribute VB_Name = "T_R93Outputs"
Option Explicit

Public Function Names() As String
    Names = "style_contract|range_compare_90000|date_book|age_book|analysis_book|privacy_book|consolidate_rows"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim a As Workbook, b As Workbook, outputBook As Workbook, sheet As Worksheet
    Dim source As Range, output As Range, reference As Range, result As CNxResult
    Dim request As CNxDataRequest, options As CNxDataSpecialOptions
    Dim findings As New Collection, incomplete As New Collection, values As Variant
    Dim root As String, pa As String, pb As String, pc As String, edge As Variant
    Dim oldEvents As Boolean, oldAlerts As Boolean, oldScreen As Boolean, rejected As Long
    On Error GoTo Failed
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts: oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False: Application.DisplayAlerts = False: Application.ScreenUpdating = False
    root = Environ$("LHEXCEL_PROFILE_ROOT")
    Set a = Workbooks.Add(xlWBATWorksheet)
    Set source = a.Worksheets(1).Range("B2:C4")
    source.Value2 = 12.5: source.NumberFormat = "0.00"
    Select Case name
        Case "style_contract"
            Set output = a.Worksheets(1).Range("B2:C4")
            Set reference = a.Worksheets(1).Range("E2:F4")
            reference.Value2 = 12.5: reference.NumberFormat = "0.00"
            output.Cells(2, 1).Formula = "=1+2"
            a.Worksheets(1).Hyperlinks.Add output.Cells(3, 1), "", "'" & a.Worksheets(1).Name & "'!A1"
            NxDrawApplyBusinessTable reference, False, True, True, False
            NxDrawStyleGeneratedTable a, output
            CheckOpen output, True
            For Each edge In Array(xlEdgeLeft, xlEdgeRight, xlEdgeTop, xlEdgeBottom, xlInsideHorizontal, xlInsideVertical)
                If output.Borders(edge).LineStyle <> reference.Borders(edge).LineStyle Then Err.Raise 5, , "border parity"
            Next edge
            If output.Cells(1, 1).Value2 <> 12.5 Or output.NumberFormat <> "0.00" Then Err.Raise 5, , "values/formats changed"
            If output.Cells(2, 1).Formula <> "=1+2" Or output.Cells(3, 1).Hyperlinks.Count <> 1 Then Err.Raise 5, , "formula/link lost"
            pa = root & "\style-owner.xlsx": a.SaveAs pa, xlOpenXMLWorkbook
            On Error Resume Next
            NxDrawStyleGeneratedTable a, output
            rejected = Err.Number: Err.Clear
            On Error GoTo Failed
            If rejected = 0 Or Not a.Saved Then Err.Raise 5, , "saved owner not rejected before mutation"
        Case "range_compare_90000"
            Set b = Workbooks.Add(xlWBATWorksheet)
            Set source = a.Worksheets(1).Range("B2").Resize(1800, 50)
            source.Value2 = 1: b.Worksheets(1).Range("B2").Resize(1800, 50).Value2 = 1
            b.Worksheets(1).Range("C3").Value2 = 2
            a.Saved = True: b.Saved = True
            Set result = NxDataCompareCreate("NX-FILE-WORKBOOK-COMPARE", source, b.Worksheets(1).Range("B2").Resize(1800, 50))
            RequireSuccess result
            NxDataCompareFinish True
            Set outputBook = ActiveWorkbook
            For Each sheet In outputBook.Worksheets
                Set output = sheet.Range("A5").Resize(1800, 50)
                CheckOpen output, False
                If output.Cells(2, 2).Interior.Color <> RGB(255, 199, 206) Then Err.Raise 5, , "difference color lost"
                If output.Cells(1, 1).Interior.Pattern <> xlNone Or output.Cells(1800, 50).Value2 <> 1 Then Err.Raise 5, , "equal cells changed"
            Next sheet
            If Not a.Saved Or Not b.Saved Or source.NumberFormat = "@" Then Err.Raise 5, , "source changed"
        Case "date_book", "age_book"
            source.NumberFormat = "@": source.Value2 = "240617"
            If name = "age_book" Then source.Value2 = "20000617"
            a.Activate: source.Select: a.Saved = True
            If name = "date_book" Then
                Set result = NxDataRunNormalizationForTest(source, "날짜 정리", "yyyy-mm-dd", "새 통합문서")
            Else
                Set options = New CNxDataSpecialOptions
                options.Configure "새 통합문서", False, "2026-09-15", "숫자", False, False, False, "빈칸", "자동", "*"
                Set result = NxDataSpecialRunOptions(NX_FEATURE_DATA_AGE, source, options, True)
            End If
            RequireSuccess result
            Set outputBook = ActiveWorkbook
            If outputBook Is a Then Err.Raise 5, , "missing output book"
            Set output = outputBook.Worksheets(1).Range("A1:B3")
            CheckOpen output, False
            If name = "date_book" Then
                If Format$(output.Cells(3, 2).Value, "yyyy-mm-dd") <> "2024-06-17" Or output.Cells(3, 2).NumberFormat <> "yyyy-mm-dd" Then Err.Raise 5, , "date format lost"
            Else
                If CStr(output.Cells(3, 2).Value2) <> "26" Then Err.Raise 5, , "age lost"
            End If
            If Not a.Saved Then Err.Raise 5, , "source dirtied"
        Case "analysis_book"
            source.Rows(1).Value2 = "제목": a.Activate: source.Select: a.Saved = True
            Set request = New CNxDataRequest
            request.ConfigureAnalysis "NX-DATA-UNIQUE-COUNT", source, True, False, True, True, True, True, "원본 순서", 1, 2, True, False, True, "새창으로 보기"
            values = NxDataAnalysisValues(request)
            Set result = NxDataCreateAnalysisResult(request, values, sheet)
            RequireSuccess result
            Set outputBook = sheet.Parent
            Set output = sheet.Range("A1").Resize(UBound(values, 1), UBound(values, 2))
            CheckOpen output, True
            If output.Columns(output.Columns.Count).NumberFormat <> "0.0%" Or Not a.Saved Then Err.Raise 5, , "ratio/source preservation"
        Case "privacy_book"
            findings.Add Array("연락처", "합성자료", 2, "A2", "보임", "=literal")
            a.Saved = True
            Set sheet = NxPrivacyCreateWorkbookReport(a, findings, incomplete, 6)
            Set outputBook = sheet.Parent
            CheckOpen sheet.Range("A6:F7"), True
            If sheet.Range("F7").Value2 <> "=literal" Or sheet.Range("F7").HasFormula Or Not a.Saved Then Err.Raise 5, , "privacy literal/source changed"
        Case "consolidate_rows"
            Set b = Workbooks.Add(xlWBATWorksheet)
            a.Worksheets(1).Cells.Clear
            a.Worksheets(1).Range("A1:B1").Value2 = "제목": a.Worksheets(1).Range("A2:B3").Value2 = 1
            b.Worksheets(1).Range("A1:B1").Value2 = "제목": b.Worksheets(1).Range("A2:B3").Value2 = 2
            pa = root & "\consolidate-a.xlsx": pb = root & "\consolidate-b.xlsx": pc = root & "\consolidate-output.xlsx"
            a.SaveAs pa, xlOpenXMLWorkbook: b.SaveAs pb, xlOpenXMLWorkbook
            a.Close False: Set a = Nothing: b.Close False: Set b = Nothing
            Set result = NxRunConsolidation(Array(pa, pb), pc, False, False, "ROWS", True)
            RequireSuccess result
            Set outputBook = Workbooks.Open(pc)
            CheckOpen outputBook.Worksheets(1).Range("A1:B5"), True
            If outputBook.Worksheets(1).Range("B5").Value2 <> 2 Then Err.Raise 5, , "consolidation tail lost"
        Case Else: Err.Raise 5, , "Unknown case"
    End Select
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & "|" & Err.Description
Clean:
    On Error Resume Next
    NxDataCompareFinish False
    If Not outputBook Is Nothing Then outputBook.Close False
    If Not a Is Nothing Then a.Close False
    If Not b Is Nothing Then b.Close False
    Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts: Application.ScreenUpdating = oldScreen
End Function

Private Sub CheckOpen(ByVal target As Range, ByVal hasHeader As Boolean)
    If target.Borders(xlEdgeLeft).LineStyle <> xlNone Or target.Borders(xlEdgeRight).LineStyle <> xlNone Then Err.Raise 5, , "sides not open"
    If target.Borders(xlEdgeBottom).Weight <> xlThick Or target.Font.Size <> 9 Then Err.Raise 5, , "body style mismatch"
    If hasHeader Then
        If target.Rows(1).Borders(xlEdgeBottom).Weight <> xlThick Or target.Rows(1).Interior.Color <> RGB(230, 230, 230) Then Err.Raise 5, , "header style mismatch"
    End If
End Sub

Private Sub RequireSuccess(ByVal result As CNxResult)
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
End Sub
