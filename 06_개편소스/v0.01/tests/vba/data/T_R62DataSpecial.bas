Attribute VB_Name = "T_R62DataSpecial"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestR62AgeInputAndReference"
    names.Add "TestR62MoneyOptionsAndPrecision"
    names.Add "TestR62MaskOptionsAndAuditIsolation"
    names.Add "TestR62SpecialOutputModes"
    names.Add "TestR62SpecialPreviewAndCollision"
    names.Add "TestR62NormalizeOutputVisibility"
    Set TestNames = names
End Function

Public Sub RunAll()
    Dim name As Variant
    For Each name In TestNames: RunCase CStr(name): Next name
End Sub

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestR62AgeInputAndReference": TestAge
        Case "TestR62MoneyOptionsAndPrecision": TestMoney
        Case "TestR62MaskOptionsAndAuditIsolation": TestMask
        Case "TestR62SpecialOutputModes": TestOutputs
        Case "TestR62SpecialPreviewAndCollision": TestPreview
        Case "TestR62NormalizeOutputVisibility": TestNormalize
        Case Else: NxRaiseContractError "Unknown r62 special case: " & name
    End Select
End Sub

Private Sub TestAge()
    Dim parsed As Date, options As New CNxDataSpecialOptions
    Assert NxAgeValue("1990-05-20", DateSerial(2026, 5, 19)) = 35, "Before birthday"
    Assert NxAgeValue(900520, DateSerial(2026, 5, 20)) = 36, "Numeric YYMMDD misread as serial"
    Assert NxAgeValue("90년 5월 20일", DateSerial(2026, 5, 20)) = 36, "Korean birth date"
    Assert NxAgeValue("900520-*******", DateSerial(2026, 5, 20)) = 36, "Masked resident number"
    Assert NxAgeValue("000101-3123456", DateSerial(2026, 5, 20)) = 26, "Resident-number century"
    Assert Not NxAgeTryParseDate("2026-02-30", Date, parsed), "Invalid date rolled over"
    Assert IsError(NxAgeValue("2027-01-01", DateSerial(2026, 9, 6))), "Future birthday accepted"
    options.Configure "새 시트 만들기", False, "2026-05-20", "만 00세", True, True, True, "빈칸", "자동", "*"
    Assert options.TransformValue(NX_FEATURE_DATA_AGE, "900520") = "만 36세", "Korean age option"
End Sub

Private Sub TestMoney()
    Assert NxKoreanMoneyWithOptions("1,234원", True, True, True, "빈칸") = "금1,234원정(금일천이백삼십사원정)", "Money donor options"
    Assert NxKoreanMoneyWithOptions("0", False, False, False, "빈칸") = "", "Zero blank"
    Assert NxKoreanMoneyWithOptions("0", False, False, False, "영원") = "영원", "Zero Korean"
    Assert NxKoreanMoneyWithOptions("0", True, True, True, "error") = "error", "Zero error"
    Assert NxKoreanMoneyWithOptions("-1", True, True, True, "영원") = "error", "Negative amount"
    Assert NxKoreanMoneyWithOptions("1.5", True, True, True, "영원") = "error", "Fractional amount"
    Assert InStr(NxKoreanMoneyWithOptions("9007199254740993", False, False, True, "영원"), "9,007,199,254,740,993원(") = 1, "Decimal text precision lost"
End Sub

Private Sub TestMask()
    Assert NxPrivacyMaskWithOptions("900520-1123456", "주민등록번호", "*") = "900520-*******", "Resident mask retained gender"
    Assert NxPrivacyMaskWithOptions("홍길동", "이름", "O") = "홍O동", "Name endpoints"
    Assert NxPrivacyMaskWithOptions("010-1234-5678", "휴대폰번호", "O") = "010-OOOO-5678", "Phone character"
    Assert NxPrivacyMaskWithOptions("h@example.invalid", "이메일", "*") = "h***@example.invalid", "Email minimum mask"
    Assert NxPrivacyMaskWithOptions("not personal", "자동", "*") = "error", "Unmatched input"
    Assert NxPrivacyMaskedText("900520-1123456") = "900520-1******", "Audit API behavior changed"
End Sub

Private Sub TestOutputs()
    Dim book As Workbook, resultBook As Workbook, source As Range, options As CNxDataSpecialOptions
    Dim result As CNxResult, mode As Variant, detail As String
    On Error GoTo Failed
    For Each mode In Array("빈 결과 열/행", "원본 변경", "새 시트 만들기", "새 통합문서")
        Set book = Workbooks.Add(xlWBATWorksheet)
        Set source = book.Worksheets(1).Range("A1:A2")
        source.NumberFormat = "@": source.Cells(1, 1).Value2 = "900520": source.Cells(2, 1).Value2 = "000101-3123456"
        source.Select
        Set options = New CNxDataSpecialOptions
        options.Configure CStr(mode), False, "2026-05-20", "숫자", True, True, True, "빈칸", "자동", "*"
        Set result = NxDataSpecialRunOptions(NX_FEATURE_DATA_AGE, source, options, True)
        Assert result.Outcome = NxSuccess, "Output mode failed: " & CStr(mode) & " | " & result.Recovery
        Select Case CStr(mode)
            Case "빈 결과 열/행"
                Assert source.Cells(1, 1).Value2 = "900520" And source.Offset(0, 1).Cells(1, 1).Value2 = 36, "Adjacent output"
            Case "원본 변경"
                Assert source.Cells(1, 1).Value2 = 36 And source.Cells(1, 1).NumberFormat = "General", "Original age format"
            Case "새 시트 만들기"
                Assert book.Worksheets.Count = 2 And ActiveSheet.Range("A1").Value2 = 36, "New sheet output"
                Assert source.Cells(1, 1).Value2 = "900520", "New sheet changed source"
            Case "새 통합문서"
                Set resultBook = ActiveWorkbook
                Assert Not resultBook Is book, "New workbook reused source"
                Assert resultBook.Worksheets(1).Range("A1").Value2 = 36, "New workbook output"
                Assert source.Cells(1, 1).Value2 = "900520", "New workbook changed source"
                resultBook.Close SaveChanges:=False: Set resultBook = Nothing
        End Select
        book.Close SaveChanges:=False: Set book = Nothing
    Next mode
CleanUp:
    On Error Resume Next
    If Not resultBook Is Nothing Then resultBook.Close SaveChanges:=False
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description: Err.Clear
    Resume CleanUp
End Sub

Private Sub TestPreview()
    Dim book As Workbook, source As Range, options As New CNxDataSpecialOptions
    Dim result As CNxResult, text As String, detail As String, collision As Boolean
    Dim captured As CNxExecutionContext
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet): Set source = book.Worksheets(1).Range("A1:C1")
    source.NumberFormat = "@": source.Cells(1, 1).Value2 = "=title": source.Cells(1, 2).Value2 = "홍길동": source.Cells(1, 3).Value2 = "김철수"
    options.Configure "빈 결과 열/행", True, "", "만 00세", True, True, True, "빈칸", "이름", "*"
    book.Saved = True: source.Select
    text = NxDataSpecialPreview(NX_FEATURE_DATA_PRIVACY_MASK, source, options)
    Assert book.Saved And InStr(text, "홍*동") > 0, "Preview changed source or missed transformation"
    Set result = NxDataSpecialRunOptions(NX_FEATURE_DATA_PRIVACY_MASK, source, options, True)
    Assert result.Outcome = NxSuccess, "Horizontal output failed: " & result.Recovery
    Assert source.Offset(1, 0).Cells(1, 1).Value2 = "=title" And Not source.Offset(1, 0).Cells(1, 1).HasFormula, "Header formula injection"
    Assert source.Offset(1, 0).Cells(1, 2).Value2 = "홍*동", "Horizontal result placement"
    On Error Resume Next
    Set result = NxDataSpecialRunOptions(NX_FEATURE_DATA_PRIVACY_MASK, source, options, True)
    collision = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    Assert collision And source.Cells(1, 2).Value2 = "홍길동", "Occupied output was overwritten"
    source.Select: Set captured = NxContextFactory.CaptureCurrent()
    source.Cells(1, 2).Value2 = "변경됨"
    Assert Not captured.SourceStateMatchesCurrent, "Stale source state accepted"
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description: Err.Clear
    Resume CleanUp
End Sub

Private Sub Assert(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then NxRaiseContractError message
End Sub

Private Sub TestNormalize()
    Dim book As Workbook, source As Range, result As CNxResult, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet): Set source = book.Worksheets(1).Range("A1:A2")
    source.NumberFormat = "@": source.Cells(1, 1).Value2 = "26년 9월 6일": source.Cells(2, 1).Value2 = "20260907"
    source.Select
    Set result = NxDataRunNormalizationForTest(source, "날짜 정리", "yyyymmdd", "새 시트 만들기")
    Assert result.Outcome = NxSuccess, "Normalize output: " & result.Recovery
    Assert Not ActiveSheet Is source.Worksheet, "Guard hid normalization output"
    Assert ActiveSheet.Range("A1").Text = "20260906", "Korean normalize date format | " & NormalizeCellEvidence(ActiveSheet.Range("A1"))
    Assert VarType(ActiveSheet.Range("A1").Value2) = vbDouble, "Normalized date must be an Excel serial"
    Assert ActiveSheet.Range("A2").Text = "20260907", "Eight-digit date did not preserve its calendar day"
    Assert source.Cells(1, 1).Value2 = "26년 9월 6일", "Normalization changed source"
    source.Cells(1, 1).Value2 = "=1+1": source.Cells(2, 1).Value2 = "+SUM(1,2)"
    source.Worksheet.Activate
    source.Select
    Set result = NxDataRunNormalizationForTest(source, "기본 정리", "yyyymmdd", "새 시트 만들기")
    Assert result.Outcome = NxSuccess, "Literal normalization failed: " & result.Recovery
    Assert Not ActiveSheet.Range("A1").HasFormula And Not ActiveSheet.Range("A2").HasFormula, "Normalized text executed as a formula"
    Assert ActiveSheet.Range("A1").Value2 = "=1+1", "Formula-like literal was changed"
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description: Err.Clear
    Resume CleanUp
End Sub

Private Function NormalizeCellEvidence(ByVal cell As Range) As String
    NormalizeCellEvidence = cell.Address(External:=True) & " | Value2=" & CStr(cell.Value2) & _
        " | TypeName=" & TypeName(cell.Value2) & " | NumberFormat=" & CStr(cell.NumberFormat) & _
        " | Text=" & CStr(cell.Text)
End Function
