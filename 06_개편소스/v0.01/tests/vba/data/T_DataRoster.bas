Attribute VB_Name = "T_DataRoster"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestRosterDiffHasFiveStates"
    names.Add "TestRosterCreateWritesExactMetadata"
    names.Add "TestRosterConflictBlocksUpdate"
    names.Add "TestRosterFailureRestoresCommandOwnedChanges"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestRosterDiffHasFiveStates": TestRosterDiffHasFiveStates
        Case "TestRosterCreateWritesExactMetadata": TestRosterCreateWritesExactMetadata
        Case "TestRosterConflictBlocksUpdate": TestRosterConflictBlocksUpdate
        Case "TestRosterFailureRestoresCommandOwnedChanges": TestRosterFailureRestoresCommandOwnedChanges
        Case Else: NxRaiseContractError "Unknown Data roster test"
    End Select
End Sub

Private Sub TestRosterDiffHasFiveStates()
    Dim rows As Collection
    Set rows = NxRosterBuildDiff(NxExistingRosterFixture(), NxIncomingRosterFixture(), Array(1), Array(2))
    NxTestHarness.AssertTrue NxRosterStateSet(rows) = "new,changed,same,excluded,conflict", "Roster five-state contract changed"
End Sub

Private Sub TestRosterCreateWritesExactMetadata()
    Dim fixtureBook As Workbook
    Dim result As CNxResult
    Dim roster As Worksheet
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    TraceRosterProgress "roster-01-enter"
    Set fixtureBook = Application.Workbooks.Add
    TraceRosterProgress "roster-02-workbook-added"
    Set result = NxRosterCreate(fixtureBook, NxRosterSourceFixture(), Array(1), Array(1, 2))
    TraceRosterProgress "roster-03-create-returned"
    Set roster = fixtureBook.Worksheets(fixtureBook.Worksheets.Count)
    TraceRosterProgress "roster-04-sheet-resolved"
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Roster creation failed"
    TraceRosterProgress "roster-05-result-asserted"
    NxTestHarness.AssertTrue CStr(roster.CustomProperties("LHexcelRosterSchema").Value) = "1", "Roster schema metadata changed"
    TraceRosterProgress "roster-06-schema-read"
    NxTestHarness.AssertTrue Len(CStr(roster.CustomProperties("LHexcelRosterId").Value)) = 36, "Roster ID metadata is invalid"
    TraceRosterProgress "roster-07-id-read"
    NxTestHarness.AssertTrue CStr(roster.CustomProperties("LHexcelRosterKeyColumns").Value) = "ID", "Roster key metadata changed"
    TraceRosterProgress "roster-08-key-read"
    NxTestHarness.AssertTrue Len(CStr(roster.CustomProperties("LHexcelRosterCreatedUtc").Value)) = 20, "Roster created timestamp is invalid"
    TraceRosterProgress "roster-09-created-read"
    NxTestHarness.AssertTrue Len(CStr(roster.CustomProperties("LHexcelRosterUpdatedUtc").Value)) = 20, "Roster updated timestamp is invalid"
    TraceRosterProgress "roster-10-updated-read"
    GoTo Cleanup
Failed:
    TraceRosterProgress "roster-90-failed"
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    TraceRosterProgress "roster-91-cleanup-enter"
    On Error Resume Next
    fixtureBook.Close False
    On Error GoTo 0
    TraceRosterProgress "roster-92-cleanup-exit"
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_DataRoster.TestRosterCreateWritesExactMetadata", failureDescription
End Sub

Private Sub TraceRosterProgress(ByVal marker As String)
    Dim progressPath As String
    Dim handle As Integer
    progressPath = Environ$("LHEXCEL_OWNER_CASE_PROGRESS")
    If Len(progressPath) = 0 Then Exit Sub
    On Error GoTo CleanExit
    handle = FreeFile
    Open progressPath For Append Access Write As #handle
    Print #handle, marker
CleanExit:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    On Error GoTo 0
End Sub

Private Sub TestRosterConflictBlocksUpdate()
    Dim rows As Collection
    Set rows = NxRosterBuildDiff(NxExistingRosterFixture(), NxIncomingRosterFixture(), Array(1), Array(2))
    NxTestHarness.AssertTrue NxRosterHasConflicts(rows), "Roster conflict was not detected"
End Sub

Private Sub TestRosterFailureRestoresCommandOwnedChanges()
    Dim fixtureBook As Workbook
    Dim createResult As CNxResult
    Dim updateResult As CNxResult
    Dim roster As Worksheet
    Dim rows As Collection
    Dim existing As Variant
    Dim incoming As Variant
    Dim before As String
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    TraceRosterProgress "rollback-01-enter"
    Set fixtureBook = Application.Workbooks.Add
    TraceRosterProgress "rollback-02-workbook-added"
    Set createResult = NxRosterCreate(fixtureBook, NxRosterSourceFixture(), Array(1), Array(1, 2))
    TraceRosterProgress "rollback-03-create-returned"
    Set roster = fixtureBook.Worksheets(fixtureBook.Worksheets.Count)
    TraceRosterProgress "rollback-04-sheet-resolved"
    existing = roster.UsedRange.Value
    TraceRosterProgress "rollback-05-existing-read"
    incoming = NxRosterUpdateFixture()
    Set rows = NxRosterBuildDiff(existing, incoming, Array(1), Array(2))
    TraceRosterProgress "rollback-06-diff-built"
    before = NxRosterSheetFingerprint(roster)
    TraceRosterProgress "rollback-07-before-fingerprint"
    NxRosterSetFailureAfterRows 2
    TraceRosterProgress "rollback-08-failure-armed"
    Set updateResult = NxRosterApplyUpdate(roster, rows, incoming, True)
    TraceRosterProgress "rollback-09-update-returned"
    NxRosterClearFailurePoint
    TraceRosterProgress "rollback-10-failure-cleared"
    NxTestHarness.AssertTrue updateResult.Outcome = NxEnvironmentError, "Injected roster failure did not return an environment error"
    TraceRosterProgress "rollback-11-outcome-asserted"
    NxTestHarness.AssertTrue before = NxRosterSheetFingerprint(roster), "Roster rollback left changed cells or metadata"
    TraceRosterProgress "rollback-12-fingerprint-asserted"
    GoTo Cleanup
Failed:
    TraceRosterProgress "rollback-90-failed"
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    TraceRosterProgress "rollback-91-cleanup-enter"
    On Error Resume Next
    NxRosterClearFailurePoint
    fixtureBook.Close False
    On Error GoTo 0
    TraceRosterProgress "rollback-92-cleanup-exit"
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_DataRoster.TestRosterFailureRestoresCommandOwnedChanges", failureDescription
End Sub

Private Function NxExistingRosterFixture() As Variant
    Dim values(1 To 6, 1 To 2) As Variant
    values(1, 1) = "ID": values(1, 2) = "이름"
    values(2, 1) = "A": values(2, 2) = "이전"
    values(3, 1) = "B": values(3, 2) = "동일"
    values(4, 1) = "C": values(4, 2) = "제외"
    values(5, 1) = "X": values(5, 2) = "충돌1"
    values(6, 1) = "X": values(6, 2) = "충돌2"
    NxExistingRosterFixture = values
End Function

Private Function NxIncomingRosterFixture() As Variant
    Dim values(1 To 5, 1 To 2) As Variant
    values(1, 1) = "ID": values(1, 2) = "이름"
    values(2, 1) = "D": values(2, 2) = "신규"
    values(3, 1) = "A": values(3, 2) = "변경"
    values(4, 1) = "B": values(4, 2) = "동일"
    values(5, 1) = "X": values(5, 2) = "충돌"
    NxIncomingRosterFixture = values
End Function

Private Function NxRosterSourceFixture() As Variant
    Dim values(1 To 4, 1 To 2) As Variant
    values(1, 1) = "ID": values(1, 2) = "이름"
    values(2, 1) = "A": values(2, 2) = "이전"
    values(3, 1) = "B": values(3, 2) = "동일"
    values(4, 1) = "C": values(4, 2) = "제외"
    NxRosterSourceFixture = values
End Function

Private Function NxRosterUpdateFixture() As Variant
    Dim values(1 To 5, 1 To 2) As Variant
    values(1, 1) = "ID": values(1, 2) = "이름"
    values(2, 1) = "A": values(2, 2) = "변경"
    values(3, 1) = "B": values(3, 2) = "동일"
    values(4, 1) = "C": values(4, 2) = "제외"
    values(5, 1) = "D": values(5, 2) = "신규"
    NxRosterUpdateFixture = values
End Function

Private Function NxRosterSheetFingerprint(ByVal sheet As Worksheet) As String
    Dim cell As Range
    For Each cell In sheet.UsedRange.Cells
        NxRosterSheetFingerprint = NxRosterSheetFingerprint & CStr(cell.Row) & ":" & CStr(cell.Column) & ":" & _
            CStr(cell.Formula) & ":" & CStr(cell.NumberFormat) & "|"
    Next cell
    NxRosterSheetFingerprint = NxRosterSheetFingerprint & CStr(sheet.CustomProperties("LHexcelRosterUpdatedUtc").Value)
End Function
