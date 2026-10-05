Attribute VB_Name = "NxWorkbookCompareResults"
Option Explicit

Public Sub NxWorkbookCompareBrowseReport(ByVal reportPath As String)
    Dim report As Workbook, candidate As Workbook, session As CNxCompareResultsSession, service As Object
    Dim priorSecurity As MsoAutomationSecurity, priorEvents As Boolean, captured As Boolean
    Dim priorWindow As Window, priorSelection As Object
    Dim sourceHandle As Integer, opened As Boolean, shown As Boolean, issue As String, detail As String, cleanupIssue As String
    On Error GoTo Failed
    Set priorWindow = Application.ActiveWindow
    Set priorSelection = Application.Selection
    Set service = NxHostCreateWorkbookCompareResults()
    If service Is Nothing Then
        MsgBox "DLL 결과 탐색창에 연결할 수 없습니다. 저장된 xlsx 보고서의 비교요약·셀차이 시트를 사용하세요.", vbInformation, "내엑셀 - 비교 결과"
        Exit Sub
    End If
    If Len(NxWorkbookCompareNavigationPathIssue(reportPath)) > 0 Or LCase$(Right$(reportPath, 5)) <> ".xlsx" Then NxRaiseContractError "xlsx 비교 보고서의 전체 경로를 선택하세요."
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, reportPath, vbTextCompare) = 0 Then Set report = candidate: Exit For
    Next candidate
    If report Is Nothing Then
        priorSecurity = Application.AutomationSecurity: priorEvents = Application.EnableEvents: captured = True
        sourceHandle = FreeFile
        Open reportPath For Binary Access Read Lock Write As #sourceHandle
        issue = NxSafeWorkbookInputIssue(reportPath)
        If Len(issue) > 0 Then NxRaiseContractError issue
        Application.AutomationSecurity = msoAutomationSecurityForceDisable
        Application.EnableEvents = False
        Set report = Application.Workbooks.Open(Filename:=reportPath, UpdateLinks:=0, ReadOnly:=True, _
            Password:=vbNullString, WriteResPassword:=vbNullString, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        opened = True
        If Not report.ReadOnly Or StrComp(report.FullName, reportPath, vbTextCompare) <> 0 Then NxRaiseContractError "선택한 보고서를 읽기 전용으로 열지 못했습니다."
        Close #sourceHandle: sourceHandle = 0
        Application.AutomationSecurity = priorSecurity: Application.EnableEvents = priorEvents: captured = False
    End If
    Set session = New CNxCompareResultsSession
    session.Configure report
    ' Return from the VBA modal before opening this modeless, Excel-owned window.
    report.Activate
    service.ShowResults session.Rows, report.Name, CStr(Application.Hwnd), session
    NxHostRetainWorkbookCompareResults service
    shown = True
CleanUp:
    On Error Resume Next
    If sourceHandle <> 0 Then Close #sourceHandle
    If captured Then
        Err.Clear: Application.AutomationSecurity = priorSecurity
        If Err.Number <> 0 Then cleanupIssue = "매크로 보안"
        Err.Clear: Application.EnableEvents = priorEvents
        If Err.Number <> 0 Then cleanupIssue = cleanupIssue & " 이벤트"
    End If
    If Not shown Then
        If Not session Is Nothing Then session.ReleaseSession
        If opened Then
            If Not report Is Nothing Then report.Close SaveChanges:=False
        End If
        If Not priorWindow Is Nothing Then
            Err.Clear: priorWindow.Activate
            If Err.Number <> 0 Then cleanupIssue = cleanupIssue & " 활성 창"
            If Not priorSelection Is Nothing Then
                Err.Clear: priorSelection.Select
                If Err.Number <> 0 Then cleanupIssue = cleanupIssue & " 선택"
            End If
        End If
    End If
    On Error GoTo 0
    If Len(detail) > 0 Or Len(cleanupIssue) > 0 Then
        MsgBox "결과 탐색창을 열지 못했습니다. 저장된 보고서는 유지됩니다." & vbCrLf & detail & _
            IIf(Len(cleanupIssue) > 0, vbCrLf & "Excel 상태 복원 확인 필요: " & cleanupIssue, vbNullString), vbExclamation, "내엑셀 - 비교 결과"
    End If
    Exit Sub
Failed:
    detail = Err.Description: Err.Clear
    Resume CleanUp
End Sub
