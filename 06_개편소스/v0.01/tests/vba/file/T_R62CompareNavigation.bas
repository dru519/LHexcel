Attribute VB_Name = "T_R62CompareNavigation"
Option Explicit

Private mBasePath As String
Private mComparePath As String
Private mReportPath As String

' Native runner opens the published report between PrepareReport and VerifyOpenedReport.
' This also proves that links survive save/close/reopen and need no report VBA project.
Public Function PrepareReport() As String
    Dim baseBook As Workbook, compareBook As Workbook, sheet As Worksheet, issue As String
    On Error GoTo Failed
    If Len(mReportPath) > 0 Then NxRaiseContractError "이전 r62 비교 검증 자료를 먼저 정리하세요."
    mBasePath = FixturePath("base"): mComparePath = FixturePath("compare"): mReportPath = FixturePath("report")
    Set baseBook = Workbooks.Add(xlWBATWorksheet)
    Set compareBook = Workbooks.Add(xlWBATWorksheet)
    PopulateFixture baseBook, "기준"
    PopulateFixture compareBook, "비교"
    Set sheet = baseBook.Worksheets.Add(After:=baseBook.Worksheets(baseBook.Worksheets.Count))
    sheet.Name = "기준에만": sheet.Range("A1").Value2 = "기준 전용"
    baseBook.SaveAs Filename:=mBasePath, FileFormat:=xlOpenXMLWorkbook
    compareBook.SaveAs Filename:=mComparePath, FileFormat:=xlOpenXMLWorkbook
    baseBook.Close SaveChanges:=False: Set baseBook = Nothing
    compareBook.Close SaveChanges:=False: Set compareBook = Nothing
    VerifyLockedSourceInspection mBasePath, True
    VerifyLockedMalformedSource
    Call NxWorkbookCompareReport(mBasePath, mComparePath, mReportPath, False)
    Require Len(Dir$(mReportPath)) > 0, "보고서가 생성되지 않았습니다."
    PrepareReport = mReportPath
    Exit Function
Failed:
    issue = Err.Description
    On Error Resume Next
    If Not baseBook Is Nothing Then baseBook.Close SaveChanges:=False
    If Not compareBook Is Nothing Then compareBook.Close SaveChanges:=False
    On Error GoTo 0
    NxRaiseContractError issue
End Function

Private Sub VerifyLockedSourceInspection(ByVal filePath As String, ByVal shouldAccept As Boolean)
    Dim sourceHandle As Integer, beforeSha As String, issue As String, failure As String
    Dim priorSecurity As MsoAutomationSecurity, priorEvents As Boolean, beforeCount As Long
    priorSecurity = Application.AutomationSecurity: priorEvents = Application.EnableEvents
    beforeCount = Application.Workbooks.Count
    On Error GoTo Failed
    beforeSha = NxTemplateFileSha256(filePath)
    sourceHandle = FreeFile
    Open filePath For Binary Access Read Lock Write As #sourceHandle
    RequireLockedWriterDenied filePath
    issue = NxSafeWorkbookInputIssue(filePath)
    Require CBool((Len(issue) = 0) = shouldAccept), "Locked source inspection result rejected: " & issue
    RequireLockedWriterDenied filePath
    Require Application.Workbooks.Count = beforeCount, "Inspection opened a workbook"
    Require Application.AutomationSecurity = priorSecurity And Application.EnableEvents = priorEvents, "Inspection changed macro security or events"
    Close #sourceHandle: sourceHandle = 0
    Require NxTemplateFileSha256(filePath) = beforeSha, "Inspection changed locked original bytes"
    Exit Sub
Failed:
    failure = Err.Description
    On Error Resume Next
    If sourceHandle <> 0 Then Close #sourceHandle
    Application.AutomationSecurity = priorSecurity: Application.EnableEvents = priorEvents
    On Error GoTo 0
    NxRaiseContractError failure
End Sub

Private Sub RequireLockedWriterDenied(ByVal filePath As String)
    Dim writerHandle As Integer, blockedNumber As Long
    writerHandle = FreeFile
    On Error Resume Next
    Err.Clear
    ' Opening for binary write does not modify bytes; no Put/write is performed.
    Open filePath For Binary Access Write Shared As #writerHandle
    blockedNumber = Err.Number
    If blockedNumber = 0 Then Close #writerHandle
    Err.Clear
    On Error GoTo 0
    Require CBool(blockedNumber = 55 Or blockedNumber = 70), "Original write-denial lock was not retained"
End Sub

Private Sub VerifyLockedMalformedSource()
    Dim invalidPath As String, handle As Integer, signature As String, failure As String
    On Error GoTo Failed
    invalidPath = FixturePath("locked-invalid")
    Require Not CreateObject("Scripting.FileSystemObject").FileExists(invalidPath), "Malformed fixture destination already exists"
    signature = "not-an-ooxml-package"
    handle = FreeFile
    Open invalidPath For Binary Access Write As #handle
    Put #handle, 1, signature
    Close #handle: handle = 0
    VerifyLockedSourceInspection invalidPath, False
    Kill invalidPath
    Exit Sub
Failed:
    failure = Err.Description
    On Error Resume Next
    If handle <> 0 Then Close #handle
    On Error GoTo 0
    ' Retain the uniquely named fixture on failure for the runner's evidence inventory.
    NxRaiseContractError failure
End Sub

Private Sub PopulateFixture(ByVal book As Workbook, ByVal text As String)
    Dim sheet As Worksheet
    Set sheet = book.Worksheets(1): sheet.Name = "검사'표"
    sheet.Range("A1").Value2 = text
    sheet.Range("D7").Value2 = text & " 정확한 셀"
    sheet.Range("C9").Value2 = text & " 숨김 행"
    sheet.Rows(9).Hidden = True
    sheet.Range("E4").Value2 = text & " 숨김 열"
    sheet.Columns(5).Hidden = True
    Set sheet = book.Worksheets.Add(After:=sheet): sheet.Name = "숨김원본"
    sheet.Range("A1").Value2 = text & " 숨김 시트": sheet.Visible = xlSheetVeryHidden
End Sub

Public Sub VerifyOpenedReport(ByVal reportPath As String)
    Dim reportBook As Workbook, sourceBook As Workbook, detailSheet As Worksheet, summarySheet As Worksheet
    Dim detailRow As Long, summaryRow As Long, issue As String, resolvedPath As String, resolvedSheet As String, resolvedCell As String
    Dim priorSecurity As MsoAutomationSecurity, priorEvents As Boolean, reportSaved As Boolean
    On Error GoTo Failed
    priorSecurity = Application.AutomationSecurity: priorEvents = Application.EnableEvents
    Set reportBook = RequirePreparedReport(reportPath)
    Set detailSheet = reportBook.Worksheets("셀차이"): Set summarySheet = reportBook.Worksheets("비교요약")
    detailRow = FindDetailRow(detailSheet, "검사'표", "D7")
    reportSaved = reportBook.Saved
    Require summarySheet.Range("B2").Hyperlinks.Count = 0, "경로 표시가 외부 링크입니다."
    Require summarySheet.Range("B3").Hyperlinks.Count = 0, "비교 경로 표시가 외부 링크입니다."
    Require NxWorkbookCompareResolveLink(detailSheet, detailSheet.Cells(detailRow, 8).Hyperlinks(1), resolvedPath, resolvedSheet, resolvedCell, issue), issue
    Require resolvedPath = mBasePath And resolvedSheet = "검사'표" And resolvedCell = "D7", "행과 기준 원본 매핑이 일치하지 않습니다."
    Require NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(detailRow, 8).Hyperlinks(1), issue), issue
    Set sourceBook = FindOpenBook(mBasePath)
    Require Not sourceBook Is Nothing, "실제 기준 파일이 열리지 않았습니다."
    Require sourceBook.ReadOnly, "닫힌 원본을 읽기 전용으로 열지 않았습니다."
    Require Application.ActiveWorkbook Is sourceBook, "기준 원본이 활성화되지 않았습니다."
    Require Application.ActiveSheet.Name = "검사'표", "원본 시트명이 달라졌습니다."
    Require Application.ActiveCell.Address(False, False) = "D7", "정확한 원본 셀로 이동하지 않았습니다."
    Require sourceBook.Worksheets("검사'표").Range("D7").Value2 = "기준 정확한 셀", "기준 값이 바뀌었습니다."
    Require Application.AutomationSecurity = priorSecurity And Application.EnableEvents = priorEvents, "원본 이동 후 Excel 상태가 달라졌습니다."
    sourceBook.Close SaveChanges:=False: Set sourceBook = Nothing
    Require NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(detailRow, 9).Hyperlinks(1), issue), issue
    Set sourceBook = FindOpenBook(mComparePath)
    Require Application.ActiveWorkbook Is sourceBook, "비교 원본이 활성화되지 않았습니다."
    Require Application.ActiveCell.Address(False, False) = "D7", "비교 원본의 다른 셀로 이동했습니다."
    sourceBook.Close SaveChanges:=False: Set sourceBook = Nothing
    summaryRow = FindSummaryRow(summarySheet, "기준에만")
    Require summarySheet.Cells(summaryRow, 6).Hyperlinks.Count = 0, "없는 비교 시트에 대체 링크가 생성되었습니다."
    Require NxWorkbookCompareNavigate(summarySheet, summarySheet.Cells(summaryRow, 5).Hyperlinks(1), issue), issue
    Require Application.ActiveSheet.Name = "기준에만" And Application.ActiveCell.Address(False, False) = "A1", "요약 링크가 해당 시트 A1이 아닙니다."
    Set sourceBook = FindOpenBook(mBasePath): sourceBook.Close SaveChanges:=False: Set sourceBook = Nothing

    ' Exercise the actual Application SheetFollowHyperlink event, not just its resolver.
    NxRouteAvailabilityStartEvents
    Application.EnableEvents = True
    reportBook.Activate: detailSheet.Activate
    detailSheet.Cells(detailRow, 9).Hyperlinks(1).Follow
    Require Application.ActiveWorkbook.FullName = mComparePath, "실제 하이퍼링크 이벤트가 비교 원본을 열지 않았습니다."
    Require Application.ActiveSheet.Name = "검사'표" And Application.ActiveCell.Address(False, False) = "D7", "실제 클릭이 정확한 원본으로 이동하지 않았습니다."
    Set sourceBook = FindOpenBook(mComparePath): sourceBook.Close SaveChanges:=False: Set sourceBook = Nothing
    Application.EnableEvents = priorEvents
    Require reportBook.Saved = reportSaved, "이동 과정에서 보고서 저장 상태가 바뀌었습니다."
    Exit Sub
Failed:
    issue = Err.Description
    On Error Resume Next
    Application.AutomationSecurity = priorSecurity: Application.EnableEvents = priorEvents
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    On Error GoTo 0
    NxRaiseContractError issue
End Sub

Public Sub RunFailures(ByVal reportPath As String)
    Dim reportBook As Workbook, sourceBook As Workbook, detailSheet As Worksheet, summarySheet As Worksheet
    Dim detailRow As Long, hiddenRow As Long, issue As String, oldPath As String, oldAddress As String
    Dim priorSecurity As MsoAutomationSecurity, priorEvents As Boolean, sourceSaved As Boolean, beforeCount As Long
    On Error GoTo Failed
    priorSecurity = Application.AutomationSecurity: priorEvents = Application.EnableEvents
    Set reportBook = RequirePreparedReport(reportPath)
    Set detailSheet = reportBook.Worksheets("셀차이"): Set summarySheet = reportBook.Worksheets("비교요약")
    detailRow = FindDetailRow(detailSheet, "검사'표", "D7")
    Application.EnableEvents = False
    beforeCount = Application.Workbooks.Count

    oldPath = CStr(summarySheet.Range("B2").Value2)
    summarySheet.Range("B2").Value2 = FixturePath("missing")
    Require Not NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(detailRow, 8).Hyperlinks(1), issue), "없는 파일을 허용했습니다."
    Require InStr(issue, "파일을 찾을 수 없습니다") > 0, "없는 파일 한국어 안내가 없습니다."
    Require Application.Workbooks.Count = beforeCount, "없는 파일이 임의 통합문서를 열었습니다."
    summarySheet.Range("B2").Value2 = oldPath

    ' A user-owned, unsaved workbook must remain the same open object and keep values/Saved.
    Set sourceBook = Application.Workbooks.Open(Filename:=mBasePath, UpdateLinks:=0, ReadOnly:=False)
    sourceBook.Worksheets("검사'표").Range("D7").Value2 = "memory-unsaved"
    sourceSaved = sourceBook.Saved: beforeCount = Application.Workbooks.Count
    Require Not sourceSaved, "미저장 원본 사례가 준비되지 않았습니다."
    Require NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(detailRow, 8).Hyperlinks(1), issue), issue
    Require Application.ActiveWorkbook Is sourceBook, "열린 사용자 원본을 재사용하지 않았습니다."
    Require sourceBook.Saved = sourceSaved, "사용자 원본 Saved 상태가 바뀌었습니다."
    Require sourceBook.Worksheets("검사'표").Range("D7").Value2 = "memory-unsaved", "미저장 값이 바뀌었습니다."
    Require Application.Workbooks.Count = beforeCount, "사용자 원본을 다시 열거나 닫았습니다."
    sourceBook.Worksheets("검사'표").Name = "이름변경됨"
    Require Not NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(detailRow, 8).Hyperlinks(1), issue), "이름이 바뀐 시트를 대체 시트로 처리했습니다."
    Require InStr(issue, "시트를 찾을 수 없습니다") > 0, "누락 시트 한국어 안내가 없습니다."
    Require Application.Workbooks.Count = beforeCount And Not sourceBook.Saved, "누락 시트 실패가 사용자 원본 상태를 바꿨습니다."
    sourceBook.Worksheets("이름변경됨").Name = "검사'표"
    hiddenRow = FindDetailRow(detailSheet, "숨김원본", "A1")
    Require Not NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(hiddenRow, 8).Hyperlinks(1), issue), "숨김 시트로 이동했습니다."
    Require InStr(issue, "숨김") > 0 And sourceBook.Worksheets("숨김원본").Visible = xlSheetVeryHidden, "숨김 시트가 변경되거나 안내가 없습니다."
    hiddenRow = FindDetailRow(detailSheet, "검사'표", "C9")
    Require Not NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(hiddenRow, 8).Hyperlinks(1), issue), "숨김 행으로 이동했습니다."
    Require InStr(issue, "숨김") > 0 And sourceBook.Worksheets("검사'표").Rows(9).Hidden, "숨김 행이 변경되었습니다."
    hiddenRow = FindDetailRow(detailSheet, "검사'표", "E4")
    Require Not NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(hiddenRow, 8).Hyperlinks(1), issue), "숨김 열로 이동했습니다."
    Require InStr(issue, "숨김") > 0 And sourceBook.Worksheets("검사'표").Columns(5).Hidden, "숨김 열이 변경되었습니다."
    sourceBook.Close SaveChanges:=False: Set sourceBook = Nothing

    ' A failure after this call owns an opened source must close only that source.
    beforeCount = Application.Workbooks.Count
    Require Not NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(hiddenRow, 8).Hyperlinks(1), issue), "닫힌 원본 숨김 열이 허용되었습니다."
    Require Application.Workbooks.Count = beforeCount, "실패 후 이번에 연 원본이 남아 있습니다."
    Require Application.EnableEvents = False And Application.AutomationSecurity = priorSecurity, "실패 후 Excel 상태가 복원되지 않았습니다."

    oldAddress = CStr(detailSheet.Cells(detailRow, 2).Value2)
    detailSheet.Cells(detailRow, 2).Value2 = "A1:B2"
    Require Not NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(detailRow, 8).Hyperlinks(1), issue), "셀 범위 주입을 허용했습니다."
    detailSheet.Cells(detailRow, 2).Value2 = oldAddress
    summarySheet.Range("B2").Formula = "=1+1"
    Require Not NxWorkbookCompareNavigate(detailSheet, detailSheet.Cells(detailRow, 8).Hyperlinks(1), issue), "수식 원본 필드를 허용했습니다."
    summarySheet.Range("B2").NumberFormat = "@": summarySheet.Range("B2").Value2 = oldPath
    TestPathAndCellPolicy
    Application.AutomationSecurity = priorSecurity: Application.EnableEvents = priorEvents
    Exit Sub
Failed:
    issue = Err.Description
    On Error Resume Next
    If Len(oldPath) > 0 Then summarySheet.Range("B2").NumberFormat = "@": summarySheet.Range("B2").Value2 = oldPath
    If Len(oldAddress) > 0 Then detailSheet.Cells(detailRow, 2).Value2 = oldAddress
    If Not sourceBook Is Nothing Then sourceBook.Close SaveChanges:=False
    Application.AutomationSecurity = priorSecurity: Application.EnableEvents = priorEvents
    On Error GoTo 0
    NxRaiseContractError issue
End Sub

Public Sub TestPathAndCellPolicy()
    Dim value As Variant
    ' UNC parser checks do not contact a server and are not SMB execution evidence.
    For Each value In Array("C:\자료\보고서.xlsx", "D:\기준 문서.xlsm", "\\server\share\folder\file.xlsx", "\\server\share\한글 문서.xlsm")
        Require Len(NxWorkbookCompareNavigationPathIssue(CStr(value))) = 0, "정상 절대 경로를 거부했습니다: " & CStr(value)
    Next value
    For Each value In Array("https://example.com/file.xlsx", "file:///C:/a.xlsx", "C:relative.xlsx", "..\a.xlsx", "\\?\C:\a.xlsx", "\\.\C:\a.xlsx", "C:\a.xlsx:payload.xlsm", "C:\..\a.xlsx", "C:\a.exe", "C:\a.xls", "C:\a.xlsb", "C:\NUL.xlsx", "C:\a.\b.xlsx", "\\server\file.xlsx", "C:\a.xlsx ", "\\server@SSL\DavWWWRoot\file.xlsx", "\\server\DavWWWRoot\file.xlsx")
        Require Len(NxWorkbookCompareNavigationPathIssue(CStr(value))) > 0, "금지 경로를 허용했습니다: " & CStr(value)
    Next value
    For Each value In Array("A1", "D7", "XFD1048576")
        Require Len(NxWorkbookCompareNavigationCellIssue(CStr(value))) = 0, "정상 A1 셀을 거부했습니다."
    Next value
    For Each value In Array("A0", "A01", "a1", "$A$1", "A1:B2", "Sheet1!A1", "A1048577", "XFE1", "R1C1", "=A1", "A", "1")
        Require Len(NxWorkbookCompareNavigationCellIssue(CStr(value))) > 0, "금지 셀 주소를 허용했습니다: " & CStr(value)
    Next value
End Sub

Public Sub CleanupPreparedReport(ByVal reportPath As String)
    Dim book As Workbook, filePath As Variant
    Require Len(mReportPath) > 0 And reportPath = mReportPath, "소유하지 않은 검증 자료 정리를 거부했습니다."
    For Each filePath In Array(mBasePath, mComparePath, mReportPath)
        Require Left$(CStr(filePath), Len(Environ$("TEMP") & "\LHexcel-r62-compare-")) = Environ$("TEMP") & "\LHexcel-r62-compare-", "검증 경로 범위를 벗어났습니다."
        Set book = FindOpenBook(CStr(filePath))
        If Not book Is Nothing Then book.Close SaveChanges:=False
        If Len(Dir$(CStr(filePath))) > 0 Then Kill CStr(filePath)
    Next filePath
    mBasePath = vbNullString: mComparePath = vbNullString: mReportPath = vbNullString
End Sub

Private Function RequirePreparedReport(ByVal reportPath As String) As Workbook
    Require Len(mReportPath) > 0 And reportPath = mReportPath, "준비한 보고서와 다릅니다."
    Set RequirePreparedReport = FindOpenBook(reportPath)
    Require Not RequirePreparedReport Is Nothing, "보고서를 외부 네이티브 실행기에서 다시 연 뒤 검증하세요."
End Function

Private Function FindOpenBook(ByVal filePath As String) As Workbook
    Dim candidate As Workbook
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, filePath, vbTextCompare) = 0 Then Set FindOpenBook = candidate: Exit Function
    Next candidate
End Function

Private Function FindDetailRow(ByVal sheet As Worksheet, ByVal sheetName As String, ByVal address As String) As Long
    Dim index As Long
    For index = 2 To sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
        If CStr(sheet.Cells(index, 1).Value2) = sheetName And CStr(sheet.Cells(index, 2).Value2) = address Then FindDetailRow = index: Exit Function
    Next index
    NxRaiseContractError "예상한 비교 상세 행을 찾을 수 없습니다."
End Function

Private Function FindSummaryRow(ByVal sheet As Worksheet, ByVal sheetName As String) As Long
    Dim index As Long
    For index = 6 To sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
        If CStr(sheet.Cells(index, 1).Value2) = sheetName Then FindSummaryRow = index: Exit Function
    Next index
    NxRaiseContractError "예상한 비교 요약 행을 찾을 수 없습니다."
End Function

Private Function FixturePath(ByVal label As String) As String
    FixturePath = Environ$("TEMP") & "\LHexcel-r62-compare-" & label & "-" & Replace$(NxCreateRunUuid(), "-", vbNullString) & ".xlsx"
End Function

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then NxRaiseContractError message
End Sub
