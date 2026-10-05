Attribute VB_Name = "T_R58SafeOpen"
Option Explicit

' Test-only probes imported into an instrumented Product.xlam copy by
' Run-R58SafeOpenSuite.ps1.  They never create, alter, or close a caller-owned
' workbook; fixture creation and source hash assertions remain in PowerShell.
Public Function NxR58SafeOpenProbeClosedSnapshot(ByVal inputPath As String, ByVal expectedSecurity As Long, _
    ByVal expectedEvents As Boolean) As String

    Dim openedBook As Workbook, inputBook As Workbook, openedByUs As Boolean, snapshotPath As String
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    Set openedBook = NxSafeWorkbookOpen(inputPath, openedByUs)
    If Not openedByUs Then NxRaiseContractError "Closed safe-open input was not opened as a command snapshot"
    snapshotPath = openedBook.FullName
    If StrComp(snapshotPath, inputPath, vbTextCompare) = 0 Then NxRaiseContractError "Safe-open reused the closed input path"
    If Not openedBook.ReadOnly Then NxRaiseContractError "Safe-open snapshot was not read-only"
    If CLng(Application.AutomationSecurity) <> expectedSecurity Then NxRaiseContractError "AutomationSecurity was not restored"
    If CBool(Application.EnableEvents) <> expectedEvents Then NxRaiseContractError "EnableEvents was not restored"
    NxSafeWorkbookClose openedBook, openedByUs
    If Len(Dir$(snapshotPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then _
        NxRaiseContractError "Safe-open snapshot remained after command close"
    Set inputBook = R58SafeOpenFindBook(inputPath)
    If Not inputBook Is Nothing Then NxRaiseContractError "Closed source workbook was left open"
    NxR58SafeOpenProbeClosedSnapshot = "PASS|" & snapshotPath
    Exit Function
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description
    On Error Resume Next
    If Not openedBook Is Nothing Then NxSafeWorkbookClose openedBook, openedByUs
    On Error GoTo 0
    NxR58SafeOpenProbeClosedSnapshot = "FAIL|" & CStr(failureNumber) & "|" & Replace(Replace(failureDescription, vbCr, " "), vbLf, " ")
End Function

Public Function NxR58SafeOpenProbeRejectedInput(ByVal inputPath As String, ByVal expectedSecurity As Long, _
    ByVal expectedEvents As Boolean) As String

    Dim openedBook As Workbook, openedByUs As Boolean, failureNumber As Long, failureDescription As String
    On Error Resume Next
    Set openedBook = NxSafeWorkbookOpen(inputPath, openedByUs)
    failureNumber = Err.Number: failureDescription = Err.Description
    Err.Clear
    On Error GoTo Failed
    If failureNumber = 0 Or Not openedBook Is Nothing Then NxRaiseContractError "Unsafe input was accepted"
    If CLng(Application.AutomationSecurity) <> expectedSecurity Then NxRaiseContractError "Rejected input changed AutomationSecurity"
    If CBool(Application.EnableEvents) <> expectedEvents Then NxRaiseContractError "Rejected input changed EnableEvents"
    NxR58SafeOpenProbeRejectedInput = "REJECTED|" & CStr(failureNumber) & "|" & Replace(Replace(failureDescription, vbCr, " "), vbLf, " ")
    Exit Function
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description
    On Error Resume Next
    If Not openedBook Is Nothing Then NxSafeWorkbookClose openedBook, openedByUs
    On Error GoTo 0
    NxR58SafeOpenProbeRejectedInput = "FAIL|" & CStr(failureNumber) & "|" & Replace(Replace(failureDescription, vbCr, " "), vbLf, " ")
End Function

Public Function NxR58SafeOpenProbeUserOwnedOpen(ByVal inputPath As String, ByVal expectedUnsavedValue As String, _
    ByVal expectedSecurity As Long, ByVal expectedEvents As Boolean) As String

    Dim expectedBook As Workbook, openedBook As Workbook, openedByUs As Boolean
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    Set expectedBook = R58SafeOpenFindBook(inputPath)
    If expectedBook Is Nothing Then NxRaiseContractError "User-owned workbook fixture is not open"
    If expectedBook.Saved Then NxRaiseContractError "User-owned workbook fixture must be unsaved"
    Set openedBook = NxSafeWorkbookOpen(inputPath, openedByUs)
    If openedByUs Then NxRaiseContractError "User-owned workbook was replaced by a command snapshot"
    If Not (openedBook Is expectedBook) Then NxRaiseContractError "User-owned workbook identity changed"
    If openedBook.Saved Then NxRaiseContractError "User-owned workbook Saved state changed"
    If CStr(openedBook.Worksheets(1).Range("A1").Value2) <> expectedUnsavedValue Then _
        NxRaiseContractError "User-owned workbook in-memory value changed"
    If CLng(Application.AutomationSecurity) <> expectedSecurity Then NxRaiseContractError "User-owned open changed AutomationSecurity"
    If CBool(Application.EnableEvents) <> expectedEvents Then NxRaiseContractError "User-owned open changed EnableEvents"
    NxR58SafeOpenProbeUserOwnedOpen = "PASS|" & expectedBook.FullName
    Exit Function
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description
    NxR58SafeOpenProbeUserOwnedOpen = "FAIL|" & CStr(failureNumber) & "|" & Replace(Replace(failureDescription, vbCr, " "), vbLf, " ")
End Function

Private Function R58SafeOpenFindBook(ByVal inputPath As String) As Workbook
    Dim book As Workbook
    For Each book In Application.Workbooks
        If StrComp(book.FullName, inputPath, vbTextCompare) = 0 Then
            Set R58SafeOpenFindBook = book
            Exit Function
        End If
    Next book
End Function
