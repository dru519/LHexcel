Attribute VB_Name = "T_R62DataAnalysis"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "R62UniqueRecordCountsAndRatios"
    names.Add "R62UniqueColumnsAndBlankSentinel"
    names.Add "R62CompareCrossMembershipAndDense"
    names.Add "R62CompareRowsAndHidden"
    names.Add "R62AnalysisRejectsStaleValuesAndFormulas"
    names.Add "R62AnalysisLiteralNewWorkbookAndSourcePreservation"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "R62UniqueRecordCountsAndRatios": UniqueRecordCountsAndRatios
        Case "R62UniqueColumnsAndBlankSentinel": UniqueColumnsAndBlankSentinel
        Case "R62CompareCrossMembershipAndDense": CompareCrossMembershipAndDense
        Case "R62CompareRowsAndHidden": CompareRowsAndHidden
        Case "R62AnalysisRejectsStaleValuesAndFormulas": RejectsStaleValuesAndFormulas
        Case "R62AnalysisLiteralNewWorkbookAndSourcePreservation": LiteralNewWorkbookAndSourcePreservation
        Case Else: NxRaiseContractError "Unknown R62 analysis test"
    End Select
End Sub

Private Sub UniqueRecordCountsAndRatios()
    Dim values(1 To 5, 1 To 2) As Variant, result As Variant
    values(1, 1) = "이름": values(1, 2) = "부서"
    values(2, 1) = " 홍  길동 ": values(2, 2) = "A"
    values(3, 1) = "홍 길동": values(3, 2) = "A"
    values(4, 1) = "홍 길동": values(4, 2) = "B"
    result = NxDataAnalysisUniqueMatrix(values, False, True, True, True, True, True, "원본 순서")
    NxTestHarness.AssertTrue UBound(result, 1) = 3 And UBound(result, 2) = 4, "unique table shape"
    NxTestHarness.AssertTrue result(1, 3) = "건수" And result(1, 4) = "비율", "unique table headers"
    NxTestHarness.AssertTrue result(2, 1) = "홍 길동" And result(2, 3) = 2, "whole-record grouping"
    NxTestHarness.AssertTrue Abs(CDbl(result(2, 4)) - 2# / 3#) < 0.0000001, "ratio denominator excludes blank records"
    result = NxDataAnalysisUniqueMatrix(values, False, False, True, True, True, False, "값 내림차순")
    NxTestHarness.AssertTrue UBound(result, 1) = 2 And result(2, 2) = 3, "first-column grouping"
End Sub

Private Sub UniqueColumnsAndBlankSentinel()
    Dim values(1 To 2, 1 To 4) As Variant, blanks(1 To 2, 1 To 1) As Variant, result As Variant
    values(1, 1) = "이름": values(2, 1) = "부서"
    values(1, 2) = "가": values(2, 2) = "A"
    values(1, 3) = "가": values(2, 3) = "A"
    values(1, 4) = "나": values(2, 4) = "B"
    result = NxDataAnalysisUniqueMatrix(values, True, True, True, True, True, False, "값 내림차순")
    NxTestHarness.AssertTrue result(2, 1) = "나" And result(3, 3) = 2, "column records and descending order"
    blanks(2, 1) = "(빈값)"
    result = NxDataAnalysisUniqueMatrix(blanks, False, True, False, False, False, False, "원본 순서")
    NxTestHarness.AssertTrue UBound(result, 1) = 3, "empty cell must not collide with literal sentinel"
    NxTestHarness.AssertTrue result(2, 2) = 1 And result(3, 2) = 1, "separate blank and literal counts"
End Sub

Private Sub CompareCrossMembershipAndDense()
    Dim values(1 To 5, 1 To 2) As Variant, result As Variant
    values(1, 1) = "기준": values(1, 2) = "대조"
    values(2, 1) = "A": values(2, 2) = "b"
    values(3, 1) = "B": values(3, 2) = "C"
    values(4, 1) = " X-1 ": values(4, 2) = "x1"
    values(5, 2) = "a"
    result = NxDataAnalysisCompareMatrix(values, False, True, 1, 2, True, False)
    NxTestHarness.AssertTrue UBound(result, 2) = 4, "column comparison includes original two columns"
    NxTestHarness.AssertTrue result(2, 3) = "중복" And result(2, 4) = "중복", "membership is cross-list, not same-row equality"
    NxTestHarness.AssertTrue result(3, 4) = "고유" And result(5, 3) = "", "unique and empty states"
    NxTestHarness.AssertTrue result(4, 3) = "고유", "hyphen preserved without dense option"
    result = NxDataAnalysisCompareMatrix(values, False, True, 1, 2, True, True)
    NxTestHarness.AssertTrue result(4, 3) = "중복" And result(4, 4) = "중복", "dense comparison option"
    values(3, 2) = CVErr(xlErrNA)
    result = NxDataAnalysisCompareMatrix(values, False, True, 1, 2, True, True)
    NxTestHarness.AssertTrue result(3, 4) = "error", "error value must not appear as a blank category"
End Sub

Private Sub CompareRowsAndHidden()
    Dim values(1 To 2, 1 To 4) As Variant, result As Variant
    values(1, 1) = "기준": values(2, 1) = "대조"
    values(1, 2) = "A": values(2, 2) = "B"
    values(1, 3) = "B": values(2, 3) = "C"
    values(2, 4) = "A"
    result = NxDataAnalysisCompareMatrix(values, True, True, 1, 2, True, False, "0001", False)
    NxTestHarness.AssertTrue UBound(result, 1) = 2 And UBound(result, 2) = 4, "row comparison produces two result rows"
    NxTestHarness.AssertTrue result(1, 2) = "고유" And result(2, 2) = "중복", "hidden item omitted from membership"
    NxTestHarness.AssertTrue result(2, 4) = "", "hidden output category remains blank"
End Sub

Private Sub RejectsStaleValuesAndFormulas()
    Dim book As Workbook, source As Range, request As CNxDataRequest
    Dim rejected As Boolean, failureNumber As Long, detail As String
    On Error GoTo Failed
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set source = book.Worksheets(1).Range("A1:A2")
    source.Value2 = "same": source.Select
    Set request = AnalysisRequest(source)
    source.Cells(2, 1).Value2 = "changed"
    On Error Resume Next
    request.RequireFreshAnalysis
    rejected = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    NxTestHarness.AssertTrue rejected, "stale value was not rejected"
    source.Value2 = "same"
    source.Cells(1, 1).Formula = "=1+1"
    Set request = AnalysisRequest(source)
    source.Cells(1, 1).Formula = "=2"
    On Error Resume Next
    request.RequireFreshAnalysis
    rejected = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    NxTestHarness.AssertTrue rejected, "same-value formula replacement was not rejected"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: detail = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_R62DataAnalysis", detail
End Sub

Private Sub LiteralNewWorkbookAndSourcePreservation()
    Dim book As Workbook, resultBook As Workbook, source As Range, request As CNxDataRequest, values As Variant
    Dim result As CNxResult, formulas As Range, before As Variant, failureNumber As Long, detail As String
    Dim createdOutput As Worksheet
    On Error GoTo Failed
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set source = book.Worksheets(1).Range("A1:A2")
    source.NumberFormat = "@": source.Value2 = "=1+1": source.Select
    before = NxDataAnalysisRangeMatrix(source, True)
    Set request = AnalysisRequest(source)
    values = NxDataAnalysisValues(request)
    Set result = NxDataCreateAnalysisResult(request, values, createdOutput)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess And Not result.SourceChanged, "new workbook analysis result failed"
    Set resultBook = Application.ActiveWorkbook
    NxTestHarness.AssertTrue Not resultBook Is book, "result overwrote source workbook"
    NxTestHarness.AssertTrue NxDataAnalysisMatrixMatches(before, NxDataAnalysisRangeMatrix(source, True)), "source formulas changed"
    On Error Resume Next
    Set formulas = resultBook.Worksheets(1).UsedRange.SpecialCells(xlCellTypeFormulas)
    Err.Clear
    On Error GoTo Failed
    NxTestHarness.AssertTrue formulas Is Nothing, "literal output became executable formula"
    NxTestHarness.AssertTrue resultBook.Worksheets(1).Range("A2").Value2 = "=1+1", "literal text was not preserved"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: detail = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    If Not resultBook Is Nothing Then
        If Not resultBook Is book Then resultBook.Close SaveChanges:=False
    End If
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_R62DataAnalysis", detail
End Sub

Private Function AnalysisRequest(ByVal source As Range) As CNxDataRequest
    Dim request As New CNxDataRequest
    request.ConfigureAnalysis "NX-DATA-UNIQUE-COUNT", source, False, False, True, True, True, False, _
        "원본 순서", 1, 1, False, False, True, "새창으로 보기"
    Set AnalysisRequest = request
End Function
