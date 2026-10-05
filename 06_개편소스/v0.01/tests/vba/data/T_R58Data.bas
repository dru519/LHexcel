Attribute VB_Name = "T_R58Data"
Option Explicit

' Direct native entrypoint used only by the instrumented Product.xlam copy.
Public Sub RunAll()
    Dim book As Workbook, source As Range, preview As String, groupPreview As String, beforeSaved As Boolean
    Dim result As CNxResult, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set source = book.Worksheets(1).Range("A1:A3")
    source.NumberFormat = "@"
    source.Cells(1, 1).Value2 = " Acme "
    source.Cells(2, 1).Value2 = "acme"
    source.Cells(3, 1).Value2 = "=1+1"
    book.Saved = True
    beforeSaved = book.Saved
    book.Activate
    source.Worksheet.Activate
    source.Select

    groupPreview = NxDataPreviewText("NX-DATA-UNIQUE-COUNT", source.Cells(1, 1).Resize(2, 1), False, 1, True, False, True)
    T_R58DataAssert InStr(1, groupPreview, "고유 그룹 1개", vbBinaryCompare) > 0, "Trim/case grouping did not retain the original values"
    T_R58DataAssert book.Saved = beforeSaved And CStr(source.Cells(1, 1).Value2) = " Acme ", "Group preview had a source side effect"

    ' Group preview used only A1:A2; normalization is a separate A1:A3 request.
    source.Select
    T_R58DataAssert Application.Selection.Address = source.Address, "Normalize fixture selection is not exact"
    preview = NxDataNormalizationPreview(source, "기본 정리", "yyyy-mm-dd", "원본 변경")
    T_R58DataAssert InStr(1, preview, "값 변경 1개", vbBinaryCompare) > 0, "Normalize preview changed-count is not exact"
    T_R58DataAssert InStr(1, preview, "앞 3개 셀", vbBinaryCompare) > 0, "Normalize preview sample count is not bounded/exact"
    T_R58DataAssert InStr(1, preview, """ Acme "" → ""Acme""", vbBinaryCompare) > 0, "Normalize preview did not retain before/after text"
    T_R58DataAssert book.Saved = beforeSaved And CStr(source.Cells(1, 1).Value2) = " Acme ", "Normalize preview had a source side effect"

    Set result = NxDataRunNormalizationForTest(source, "기본 정리", "yyyy-mm-dd", "원본 변경")
    T_R58DataAssert Not result Is Nothing, "Normalize execution returned no result"
    T_R58DataAssert result.Outcome = NxSuccess, "Normalize execution did not complete"
    T_R58DataAssert CStr(source.Cells(1, 1).Value2) = "Acme" And CStr(source.Cells(2, 1).Value2) = "acme", "Normalize output changed grouped original values unexpectedly"
    T_R58DataAssert Not source.Cells(3, 1).HasFormula And CStr(source.Cells(3, 1).Value2) = "=1+1", "Formula-like text was not preserved literally"
    T_R58DataAssert Not book.Saved, "Source-mode normalize did not mark the owned temporary book changed"
    TestPrivacyWorkbookFindingsAndIncomplete
    TestPrivacyRunnerSuccessPreservesSourceTarget
    TestPrivacyRunnerIncompletePreservesSourceTarget
    TestPrivacyOwnedReportDiscardPreservesOtherWorkbooks
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Private Sub TestPrivacyWorkbookFindingsAndIncomplete()
    Dim book As Workbook, sheet As Worksheet, hidden As Worksheet, veryHidden As Worksheet, report As Worksheet
    Dim findings As Collection, incomplete As Collection, checked As Double, detail As String, scanDetail As String, issue As Variant, savedAfterScan As Boolean
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1): sheet.Name = "표시"
    Set hidden = book.Worksheets.Add(After:=sheet): hidden.Name = "숨김"
    Set veryHidden = book.Worksheets.Add(After:=hidden): veryHidden.Name = "매우숨김"
    sheet.Range("A1:B2").NumberFormat = "@"
    sheet.Range("A1").Value2 = "010-1234-5678"
    sheet.Range("B2").Value2 = "=literal@example.invalid"
    hidden.Range("A1").Value2 = "010-2222-3333"
    veryHidden.Range("B2").Value2 = "성명"
    veryHidden.Range("B3").Value2 = "홍길동"
    sheet.Activate
    hidden.Visible = xlSheetHidden: veryHidden.Visible = xlSheetVeryHidden
    hidden.Protect
    book.Saved = True
    Set findings = NxPrivacyScanWorkbook(book, checked, incomplete)
    savedAfterScan = book.Saved
    scanDetail = "checked=" & CStr(checked) & "; findings=" & CStr(findings.Count) & "; incomplete=" & CStr(incomplete.Count)
    For Each issue In incomplete
        scanDetail = scanDetail & "; " & CStr(issue(0)) & ": " & CStr(issue(1))
    Next issue
    If incomplete.Count > 0 Then scanDetail = scanDetail & PrivacySpecialCellsDetail(hidden)
    T_R58DataAssert checked = 5 And findings.Count = 3 And incomplete.Count = 0, "Hidden/VeryHidden or labelled-name finding was omitted: " & scanDetail
    Set report = NxPrivacyCreateWorkbookReport(book, findings, incomplete, checked)
    T_R58DataAssert Not report.Parent Is book, "Privacy report mutated the source workbook"
    T_R58DataAssert book.Saved And book.Worksheets.Count = 3, "Privacy scan changed the original Saved state or sheet inventory: afterScan=" & CStr(savedAfterScan) & "; afterReport=" & CStr(book.Saved) & "; sheets=" & CStr(book.Worksheets.Count)
    T_R58DataAssert hidden.Visible = xlSheetHidden And veryHidden.Visible = xlSheetVeryHidden, "Privacy scan unhid source sheets"
    T_R58DataAssert CStr(report.Cells(8, 2).Value2) = "010-2222-3333", "Privacy report omitted the original suspect value"
    T_R58DataAssert Not report.Cells(8, 2).HasFormula, "Suspect value became an executable formula"
    report.Parent.Close SaveChanges:=False: Set report = Nothing
    hidden.Unprotect
    hidden.Range("C3").Value2 = "010-3333-4444"
    hidden.Protect
    book.Saved = True
    Set findings = NxPrivacyScanWorkbook(book, checked, incomplete)
    T_R58DataAssert checked = 6 And findings.Count = 4 And incomplete.Count = 0, "Sparse protected worksheet scan is incomplete"
    T_R58DataAssert hidden.Visible = xlSheetHidden And hidden.ProtectContents And book.Saved, "Protected scan changed source state"
    sheet.Range("XFD1048576").NumberFormat = "0.00"
    sheet.Range("C4").Formula = "="""""
    book.Saved = True
    Set findings = NxPrivacyScanWorkbook(book, checked, incomplete)
    T_R58DataAssert checked = 7 And findings.Count = 4 And incomplete.Count = 0, "Sparse full-sheet extent or empty formula was not counted exactly"
    T_R58DataAssert book.Saved And hidden.ProtectContents, "Sparse scan changed source state"
    hidden.Unprotect
    sheet.Cells.Clear: hidden.Cells.Clear: veryHidden.Cells.Clear
    Set findings = NxPrivacyScanWorkbook(book, checked, incomplete)
    Set report = NxPrivacyCreateWorkbookReport(book, findings, incomplete, checked)
    T_R58DataAssert checked = 0 And incomplete.Count = 0, "Empty workbook scan is not complete"
    T_R58DataAssert InStr(CStr(report.Range("B3").Value2), "발견되지 않았습니다") > 0, "Zero-findings message is missing"
    report.Parent.Close SaveChanges:=False: Set report = Nothing
    sheet.Range("A1:A200001").Value2 = "ordinary-value"
    Set findings = NxPrivacyScanWorkbook(book, checked, incomplete)
    T_R58DataAssert incomplete.Count = 1 And findings.Count = 0, "Oversized sheet was silently accepted"
    Set report = NxPrivacyCreateWorkbookReport(book, findings, incomplete, checked)
    T_R58DataAssert InStr(CStr(report.Range("B3").Value2), "검사 미완료") > 0, "Incomplete scan was reported as no privacy findings"
CleanUp:
    On Error Resume Next
    If Not report Is Nothing Then report.Parent.Close SaveChanges:=False
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Private Function PrivacySpecialCellsDetail(ByVal sheet As Worksheet) As String
    Dim used As Range, values As Range, constantsDetail As String, formulasDetail As String
    Set used = sheet.UsedRange
    On Error Resume Next
    Set values = used.SpecialCells(xlCellTypeConstants)
    constantsDetail = "constantsError=" & CStr(Err.Number)
    If Not values Is Nothing Then constantsDetail = constantsDetail & ",count=" & CStr(values.CountLarge) & ",address=" & values.Address
    Set values = Nothing
    Err.Clear
    Set values = used.SpecialCells(xlCellTypeFormulas)
    formulasDetail = "formulasError=" & CStr(Err.Number)
    If Not values Is Nothing Then formulasDetail = formulasDetail & ",count=" & CStr(values.CountLarge) & ",address=" & values.Address
    On Error GoTo 0
    PrivacySpecialCellsDetail = "; used=" & used.Address & "; countA=" & CStr(Application.WorksheetFunction.CountA(used)) & "; " & constantsDetail & "; " & formulasDetail
End Function

Public Sub TestPrivacyRunnerSuccessPreservesSourceTarget()
    RunPrivacyRunnerCase False
End Sub

Public Sub TestPrivacyRunnerIncompletePreservesSourceTarget()
    RunPrivacyRunnerCase True
End Sub

Private Sub RunPrivacyRunnerCase(ByVal incompleteCase As Boolean)
    Dim book As Workbook, sheet As Worksheet, hidden As Worksheet, veryHidden As Worksheet, report As Worksheet
    Dim commandObject As New CNxDataSpecialCommand, context As CNxExecutionContext, result As CNxResult
    Dim beforeBooks As Long, originalAddress As String, detail As String, originalValue As Variant
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1)
    Set hidden = book.Worksheets.Add(After:=sheet)
    Set veryHidden = book.Worksheets.Add(After:=hidden)
    If incompleteCase Then
        sheet.Range("A1:A200001").Value2 = "ordinary-value"
    Else
        sheet.Range("A1").Value2 = "010-1234-5678"
        hidden.Range("A1").Value2 = "privacy@example.invalid"
        veryHidden.Range("B3").Value2 = "홍길동"
    End If
    sheet.Activate
    hidden.Visible = xlSheetHidden: veryHidden.Visible = xlSheetVeryHidden
    sheet.Range("C3").Select
    originalAddress = Application.Selection.Address
    originalValue = sheet.Range("A1").Value2
    book.Saved = True
    beforeBooks = Application.Workbooks.Count
    Set context = NxContextFactory.CaptureCurrent()
    commandObject.ConfigureWorkbookScan book
    Set result = RunPrivacyApprovalFixture(commandObject)
    T_R58DataAssert Not result Is Nothing, "Privacy approved runner returned no result"
    T_R58DataAssert result.Target = context.WorkbookIdentity, "Privacy result did not retain the approved source target"
    T_R58DataAssert Not result.SourceChanged, "Privacy runner reported a source mutation"
    If incompleteCase Then
        T_R58DataAssert result.Outcome = NxPartialFailure, "Incomplete privacy scan failed result validation"
        T_R58DataAssert result.Stage = "scan_incomplete", "Incomplete privacy scan lost its diagnostic stage"
    Else
        T_R58DataAssert result.Outcome = NxSuccess, "Privacy scan failed result validation"
    End If
    Set report = commandObject.ScanReport
    T_R58DataAssert Not report Is Nothing, "Privacy runner did not retain its owned report"
    T_R58DataAssert Not report.Parent Is book, "Privacy runner created the report in the source"
    T_R58DataAssert Application.Workbooks.Count = beforeBooks + 1, "Privacy runner created an unexpected workbook count"
    T_R58DataAssert book.Saved And book.Worksheets.Count = 3, "Privacy runner changed source Saved state or sheets"
    T_R58DataAssert sheet.Range("A1").Value2 = originalValue, "Privacy runner changed source data"
    T_R58DataAssert hidden.Visible = xlSheetHidden And veryHidden.Visible = xlSheetVeryHidden, "Privacy runner changed hidden states"
    T_R58DataAssert Application.ActiveWorkbook Is book, "Privacy runner did not restore the source workbook"
    T_R58DataAssert Application.Selection.Address = originalAddress, "Privacy runner did not restore the source selection"
    If incompleteCase Then
        T_R58DataAssert InStr(CStr(report.Range("B3").Value2), "검사 미완료") > 0, "Incomplete report was lost during runner reconciliation"
    End If
CleanUp:
    On Error Resume Next
    commandObject.DiscardScanReport
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Public Sub TestPrivacyOwnedReportDiscardPreservesOtherWorkbooks()
    Dim book As Workbook, otherBook As Workbook, report As Worksheet
    Dim commandObject As New CNxDataSpecialCommand, context As CNxExecutionContext, result As CNxResult
    Dim beforeBooks As Long, rejected As Boolean, originalEvents As Boolean, detail As String
    On Error GoTo Failed
    Set otherBook = Workbooks.Add(xlWBATWorksheet)
    otherBook.Worksheets(1).Range("A1").Value2 = "unrelated-fixture"
    otherBook.Saved = True
    Set book = Workbooks.Add(xlWBATWorksheet)
    book.Worksheets(1).Range("A1").Value2 = "privacy@example.invalid"
    book.Worksheets(1).Range("C3").Select
    book.Saved = True
    beforeBooks = Application.Workbooks.Count
    originalEvents = Application.EnableEvents
    Set context = NxContextFactory.CaptureCurrent()
    commandObject.ConfigureWorkbookScan book
    Set result = RunPrivacyApprovalFixture(commandObject)
    T_R58DataAssert result.Outcome = NxSuccess, "Privacy cleanup fixture failed before report validation"
    Set report = commandObject.ScanReport
    T_R58DataAssert Not report Is Nothing, "Privacy cleanup fixture did not create a report"
    On Error Resume Next
    result.ValidateFor NX_FEATURE_DATA_PRIVACY_SCAN, context.WorkbookIdentity & "|injected-invalid-target"
    rejected = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    T_R58DataAssert rejected, "Injected post-report result validation failure was not rejected"
    ' The controller uses this same command-owned discard after a failed runner result.
    commandObject.DiscardScanReport
    commandObject.DiscardScanReport
    Set report = commandObject.ScanReport
    T_R58DataAssert report Is Nothing, "Discard retained the owned report reference"
    T_R58DataAssert Application.Workbooks.Count = beforeBooks, "Discard did not remove exactly the created report"
    T_R58DataAssert book.Saved And CStr(book.Worksheets(1).Range("A1").Value2) = "privacy@example.invalid", "Discard changed the source"
    T_R58DataAssert otherBook.Saved And CStr(otherBook.Worksheets(1).Range("A1").Value2) = "unrelated-fixture", "Discard changed or closed an unrelated workbook"
    T_R58DataAssert Application.ActiveWorkbook Is book, "Discard changed the restored active source"
    T_R58DataAssert Application.Selection.Address = "$C$3", "Discard changed the restored source selection"
    T_R58DataAssert Application.EnableEvents = originalEvents, "Discard failed to restore events"
CleanUp:
    On Error Resume Next
    commandObject.DiscardScanReport
    If Not book Is Nothing Then book.Close SaveChanges:=False
    If Not otherBook Is Nothing Then otherBook.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Private Function RunPrivacyApprovalFixture(ByVal commandObject As CNxDataSpecialCommand) As CNxResult
    Dim command As INxFeatureCommand, router As New CNxExecutionRouter, ticket As CNxExecutionTicket
    ' Approval fixture only: use the production dispatcher/result validation, not UI dialog automation.
    Set command = commandObject
    Set ticket = router.Prepare(NxDataFeatureDefinition(NX_FEATURE_DATA_PRIVACY_SCAN), command)
    Select Case ticket.Decision.ResolvedGrade
        Case NxExecutionFast: Set RunPrivacyApprovalFixture = router.ExecuteFast(ticket)
        Case NxExecutionGuarded: Set RunPrivacyApprovalFixture = router.ExecuteGuarded(ticket, True)
        Case NxExecutionPlanned: Set RunPrivacyApprovalFixture = NxDataRunPlannedFrameForTest(command, ticket, router, "privacy approval fixture")
        Case Else: NxRaiseContractError "Privacy approval fixture received an invalid execution grade"
    End Select
End Function

Private Sub T_R58DataAssert(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then NxRaiseContractError message
End Sub
