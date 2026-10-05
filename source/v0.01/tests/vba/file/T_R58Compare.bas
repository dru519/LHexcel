Attribute VB_Name = "T_R58Compare"
Option Explicit

Private mPreparedReportPath As String
Private mPreparedBasePath As String
Private mPreparedComparePath As String

Public Function PrepareConstantsReport() As String
    Dim baseBook As Workbook, compareBook As Workbook
    Dim basePath As String, comparePath As String, outputPath As String, detail As String, failed As Boolean
    On Error GoTo Failed
    mPreparedReportPath = vbNullString
    T_R58CompareTraceStage "constants:start"
    basePath = T_R58ComparePath("base"): comparePath = T_R58ComparePath("compare"): outputPath = T_R58ComparePath("report")
    T_R58CompareTraceStage "constants:paths_allocated"
    Set baseBook = Workbooks.Add(xlWBATWorksheet): Set compareBook = Workbooks.Add(xlWBATWorksheet)
    T_R58CompareTraceStage "constants:books_created"
    baseBook.Worksheets(1).Name = "검사'표"
    compareBook.Worksheets(1).Name = "검사'표"
    With baseBook.Worksheets(1)
        .Range("A1").Value2 = "기준 상수"
        .Range("A2").Formula = "=1+1"
        .Range("A3").Value = CVErr(xlErrNA)
        .Range("A4").NumberFormat = "0.00": .Range("A4").Interior.Color = RGB(255, 255, 0)
        .Range("A6:B6").Merge: .Range("A6").Value2 = "병합"
        .Range("A7").Value2 = 1: .Range("A8").Formula = "=1/A7"
        .Calculate
    End With
    T_R58CompareTraceStage "constants:base_populated"
    With compareBook.Worksheets(1)
        .Range("A1").Value2 = "비교 상수"
        .Range("A2").Value2 = 2
        .Range("A3").Value = CVErr(xlErrValue)
        .Range("A4").NumberFormat = "0": .Range("A4").Interior.Color = RGB(0, 255, 0)
        .Range("A6").Value2 = "병합"
        .Range("A7").Value2 = 0: .Range("A8").Formula = "=1/A7"
        .Calculate
    End With
    T_R58CompareTraceStage "constants:compare_populated"
    baseBook.SaveAs basePath, xlOpenXMLWorkbook: compareBook.SaveAs comparePath, xlOpenXMLWorkbook
    T_R58CompareTraceStage "constants:fixtures_saved"
    baseBook.Close SaveChanges:=False: compareBook.Close SaveChanges:=False
    Set baseBook = Nothing: Set compareBook = Nothing
    T_R58CompareTraceStage "constants:fixtures_closed"
    T_R58CompareTraceStage "constants:compare_report:before"
    Call NxWorkbookCompareReport(basePath, comparePath, outputPath, True)
    T_R58CompareTraceStage "constants:compare_report:after"
    T_R58CompareAssert Len(Dir$(outputPath)) > 0, "Comparison report was not materialized"
    mPreparedReportPath = outputPath
    mPreparedBasePath = basePath: mPreparedComparePath = comparePath
    PrepareConstantsReport = outputPath
    T_R58CompareTraceStage "constants:report_published"
CleanUp:
    On Error Resume Next
    If Not baseBook Is Nothing Then baseBook.Close SaveChanges:=False
    If Not compareBook Is Nothing Then compareBook.Close SaveChanges:=False
    If failed Then
        T_R58CompareDelete basePath: T_R58CompareDelete comparePath: T_R58CompareDelete outputPath
        mPreparedReportPath = vbNullString: mPreparedBasePath = vbNullString: mPreparedComparePath = vbNullString
    End If
    On Error GoTo 0
    If failed Then NxRaiseContractError detail
    Exit Function
Failed:
    detail = Err.Description
    T_R58CompareTraceStage "constants:failed"
    failed = True
    Resume CleanUp
End Function

Public Sub VerifyOpenedConstantsReport(ByVal reportPath As String)
    Dim reportBook As Workbook, rowIndex As Long, sameFormulaResultReported As Boolean, detail As String
    Dim sourcePath As String, sheetName As String, cellAddress As String, issue As String
    On Error GoTo Failed
    T_R58CompareTraceStage "constants:verify_opened:start"
    Set reportBook = T_R58CompareFindOpenBook(reportPath)
    T_R58CompareAssert Not reportBook Is Nothing, "Comparison report was not open outside producer macro"
    detail = Join(Application.Transpose(reportBook.Worksheets("셀차이").Range("C2:C20").Value2), "|")
    T_R58CompareAssert InStr(1, detail, "값", vbTextCompare) > 0, "Constant difference missing"
    T_R58CompareAssert InStr(1, detail, "수식/상수", vbTextCompare) > 0, "Formula/constant difference missing"
    T_R58CompareAssert InStr(1, detail, "병합", vbTextCompare) > 0, "Merge difference missing"
    T_R58CompareAssert InStr(1, detail, "서식", vbTextCompare) > 0, "Formatted blank difference missing"
    With reportBook.Worksheets("셀차이")
        T_R58CompareAssert .Range("H2").Hyperlinks.Count = 1, "Base original-cell link missing"
        T_R58CompareAssert .Range("I2").Hyperlinks.Count = 1, "Compared original-cell link missing"
        T_R58CompareAssert .Range("H2").Hyperlinks(1).SubAddress = "'셀차이'!H2", "Base link does not remain on its report anchor"
        T_R58CompareAssert .Range("I2").Hyperlinks(1).SubAddress = "'셀차이'!I2", "Compared link does not remain on its report anchor"
        T_R58CompareAssert Len(.Range("H2").Hyperlinks(1).Address) = 0 And Len(.Range("I2").Hyperlinks(1).Address) = 0, "Report contains an automatic external-file hyperlink"
        T_R58CompareAssert NxWorkbookCompareResolveLink(reportBook.Worksheets("셀차이"), .Range("H2").Hyperlinks(1), sourcePath, sheetName, cellAddress, issue), "Base original mapping was rejected: " & issue
        T_R58CompareAssert StrComp(sourcePath, mPreparedBasePath, vbTextCompare) = 0, "Base mapping points to a snapshot instead of the original"
        T_R58CompareAssert sheetName = "검사'표" And cellAddress = "A1", "Base mapping lost the exact original sheet or cell"
        T_R58CompareAssert NxWorkbookCompareResolveLink(reportBook.Worksheets("셀차이"), .Range("I2").Hyperlinks(1), sourcePath, sheetName, cellAddress, issue), "Compared original mapping was rejected: " & issue
        T_R58CompareAssert StrComp(sourcePath, mPreparedComparePath, vbTextCompare) = 0, "Compared mapping points to a snapshot instead of the original"
        T_R58CompareAssert sheetName = "검사'표" And cellAddress = "A1", "Compared mapping lost the exact original sheet or cell"
        T_R58CompareAssert Not .Range("H2").HasFormula And Not .Range("I2").HasFormula, "Navigation was emitted as an executable formula"
        For rowIndex = 2 To .UsedRange.Rows.Count
            If CStr(.Cells(rowIndex, 2).Value2) = "A8" Then _
                sameFormulaResultReported = (InStr(1, CStr(.Cells(rowIndex, 3).Value2), "값", vbBinaryCompare) > 0)
        Next rowIndex
    End With
    T_R58CompareAssert sameFormulaResultReported, "Same formula with a different evaluated result/error was omitted"
    T_R58CompareTraceStage "constants:verify_opened:complete"
    Exit Sub
Failed:
    detail = Err.Description
    T_R58CompareTraceStage "constants:verify_opened:failed"
    On Error GoTo 0
    NxRaiseContractError detail
End Sub

Public Sub CleanupPreparedConstantsReport(ByVal reportPath As String)
    T_R58CompareAssert Len(mPreparedReportPath) > 0 And StrComp(reportPath, mPreparedReportPath, vbTextCompare) = 0 And T_R58CompareIsOwnedReportPath(reportPath), "Unowned comparison report cleanup rejected"
    T_R58CompareDelete reportPath
    T_R58CompareDelete mPreparedBasePath: T_R58CompareDelete mPreparedComparePath
    T_R58CompareAssert Len(Dir$(reportPath)) = 0, "Owned comparison report cleanup failed"
    mPreparedReportPath = vbNullString
    mPreparedBasePath = vbNullString: mPreparedComparePath = vbNullString
    T_R58CompareTraceStage "constants:report_cleanup_complete"
End Sub

Public Sub RunRemaining()
    T_R58CompareTraceStage "run_remaining:start"
    TestBoundedFormattedBlankLimit
    T_R58CompareTraceStage "run_all:formatted_blank_limit:complete"
    TestUserOwnedBookRemainsOpenAndUnsaved
    T_R58CompareTraceStage "run_all:user_owned:complete"
    TestFailedOpenLeavesNoReport
    T_R58CompareTraceStage "run_all:failed_open:complete"
End Sub

Private Sub TestBoundedFormattedBlankLimit()
    Dim baseBook As Workbook, compareBook As Workbook, basePath As String, comparePath As String, outputPath As String
    Dim failed As Boolean, detail As String
    On Error GoTo Failed
    T_R58CompareTraceStage "formatted_blank_limit:start"
    basePath = T_R58ComparePath("limit-base"): comparePath = T_R58ComparePath("limit-compare"): outputPath = T_R58ComparePath("limit-report")
    Set baseBook = Workbooks.Add(xlWBATWorksheet): Set compareBook = Workbooks.Add(xlWBATWorksheet)
    T_R58CompareTraceStage "formatted_blank_limit:books_created"
    baseBook.Worksheets(1).Range("A200001").NumberFormat = "0.00"
    baseBook.SaveAs basePath, xlOpenXMLWorkbook: compareBook.SaveAs comparePath, xlOpenXMLWorkbook
    baseBook.Close SaveChanges:=False: compareBook.Close SaveChanges:=False
    Set baseBook = Nothing: Set compareBook = Nothing
    T_R58CompareTraceStage "formatted_blank_limit:fixtures_closed"
    On Error Resume Next
    Call NxWorkbookCompareReport(basePath, comparePath, outputPath, True)
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    T_R58CompareAssert failed, "Formatted blank UsedRange above 200,000 cells was accepted"
    T_R58CompareAssert Len(Dir$(outputPath)) = 0, "Preflight limit created a report"
    T_R58CompareTraceStage "formatted_blank_limit:assertions_complete"
CleanUp:
    On Error Resume Next
    If Not baseBook Is Nothing Then baseBook.Close SaveChanges:=False
    If Not compareBook Is Nothing Then compareBook.Close SaveChanges:=False
    T_R58CompareDelete basePath: T_R58CompareDelete comparePath: T_R58CompareDelete outputPath
    On Error GoTo 0
    T_R58CompareTraceStage "formatted_blank_limit:cleanup_complete"
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    T_R58CompareTraceStage "formatted_blank_limit:failed"
    Resume CleanUp
End Sub

Private Sub TestUserOwnedBookRemainsOpenAndUnsaved()
    Dim baseBook As Workbook, compareBook As Workbook, retainedBook As Workbook
    Dim basePath As String, comparePath As String, outputPath As String, detail As String
    On Error GoTo Failed
    T_R58CompareTraceStage "user_owned:start"
    basePath = T_R58ComparePath("owned-base"): comparePath = T_R58ComparePath("owned-compare"): outputPath = T_R58ComparePath("owned-report")
    Set baseBook = Workbooks.Add(xlWBATWorksheet): baseBook.Worksheets(1).Range("A1").Value2 = "disk"
    T_R58CompareTraceStage "user_owned:base_created"
    baseBook.SaveAs basePath, xlOpenXMLWorkbook
    baseBook.Worksheets(1).Range("A1").Value2 = "memory-unsaved"
    Set retainedBook = baseBook
    Set compareBook = Workbooks.Add(xlWBATWorksheet): compareBook.Worksheets(1).Range("A1").Value2 = "other"
    compareBook.SaveAs comparePath, xlOpenXMLWorkbook: compareBook.Close SaveChanges:=False: Set compareBook = Nothing
    T_R58CompareTraceStage "user_owned:fixtures_closed"
    Call NxWorkbookCompareReport(basePath, comparePath, outputPath, False)
    T_R58CompareAssert retainedBook Is baseBook, "User-owned workbook identity changed"
    T_R58CompareAssert Not baseBook.Saved, "User-owned workbook Saved state changed"
    T_R58CompareAssert baseBook.Worksheets(1).Range("A1").Value2 = "memory-unsaved", "User-owned in-memory value changed"
    T_R58CompareTraceStage "user_owned:assertions_complete"
CleanUp:
    On Error Resume Next
    If Not baseBook Is Nothing Then baseBook.Close SaveChanges:=False
    If Not compareBook Is Nothing Then compareBook.Close SaveChanges:=False
    T_R58CompareDelete basePath: T_R58CompareDelete comparePath: T_R58CompareDelete outputPath
    On Error GoTo 0
    T_R58CompareTraceStage "user_owned:cleanup_complete"
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    T_R58CompareTraceStage "user_owned:failed"
    Resume CleanUp
End Sub

Private Sub TestFailedOpenLeavesNoReport()
    Dim compareBook As Workbook, missingPath As String, comparePath As String, outputPath As String, failed As Boolean, detail As String
    On Error GoTo Failed
    T_R58CompareTraceStage "failed_open:start"
    missingPath = T_R58ComparePath("missing")
    comparePath = T_R58ComparePath("failed-open-compare"): outputPath = T_R58ComparePath("failed-open-report")
    Set compareBook = Workbooks.Add(xlWBATWorksheet): compareBook.SaveAs comparePath, xlOpenXMLWorkbook: compareBook.Close SaveChanges:=False: Set compareBook = Nothing
    T_R58CompareTraceStage "failed_open:fixture_closed"
    On Error Resume Next
    Call NxWorkbookCompareReport(missingPath, comparePath, outputPath, False)
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo Failed
    T_R58CompareAssert failed, "Missing source did not fail"
    T_R58CompareAssert Len(Dir$(outputPath)) = 0, "Failed open created a report"
    T_R58CompareTraceStage "failed_open:assertions_complete"
CleanUp:
    On Error Resume Next
    If Not compareBook Is Nothing Then compareBook.Close SaveChanges:=False
    T_R58CompareDelete comparePath: T_R58CompareDelete outputPath
    On Error GoTo 0
    T_R58CompareTraceStage "failed_open:cleanup_complete"
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    T_R58CompareTraceStage "failed_open:failed"
    Resume CleanUp
End Sub

Private Function T_R58CompareFindOpenBook(ByVal reportPath As String) As Workbook
    Dim candidate As Workbook
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, reportPath, vbTextCompare) = 0 Then
            Set T_R58CompareFindOpenBook = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Function T_R58CompareIsOwnedReportPath(ByVal reportPath As String) As Boolean
    Dim expectedPrefix As String
    expectedPrefix = Environ$("TEMP") & "\LHexcel-r58-compare-report-"
    T_R58CompareIsOwnedReportPath = (LCase$(Left$(reportPath, Len(expectedPrefix))) = LCase$(expectedPrefix))
End Function

Private Function T_R58ComparePath(ByVal label As String) As String
    T_R58ComparePath = Environ$("TEMP") & "\LHexcel-r58-compare-" & label & "-" & Replace$(NxCreateRunUuid(), "-", vbNullString) & ".xlsx"
End Function

Private Sub T_R58CompareAssert(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then NxRaiseContractError message
End Sub

Private Sub T_R58CompareDelete(ByVal filePath As String)
    If Len(filePath) = 0 Then Exit Sub
    On Error Resume Next
    If Len(Dir$(filePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then Kill filePath
    On Error GoTo 0
End Sub

Private Sub T_R58CompareTraceStage(ByVal stageName As String)
    Dim tracePath As String, handle As Integer
    tracePath = Environ$("LHEXCEL_R58_COMPARE_TRACE")
    If Len(tracePath) = 0 Then Exit Sub
    On Error Resume Next
    handle = FreeFile
    Open tracePath For Append Access Write As #handle
    Print #handle, "{""local_time"":""" & Format$(Now, "yyyy-mm-dd\Thh:nn:ss") & """,""stage"":""" & stageName & """}"
    Close #handle
    On Error GoTo 0
End Sub
