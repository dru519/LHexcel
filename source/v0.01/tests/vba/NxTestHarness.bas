Attribute VB_Name = "NxTestHarness"
Option Explicit

Public Function RunSuite(ByVal Suite As String, ByVal Filter As String) As String
    Dim tests As Collection
    If Suite = "Core" Then
        Set tests = T_Core.TestNames()
        If tests.Count <> NX_LIMIT_CORE_TEST_CASES Then Err.Raise vbObjectError + 715, "NxTestHarness", "Core suite inventory must be exactly 36"
    ElseIf Suite = "Frame" Then
        Set tests = Application.Run("T_" & Suite & ".TestNames")
        If tests.Count <> 28 Then Err.Raise vbObjectError + 715, "NxTestHarness", "Frame suite inventory must be exactly 28"
    ElseIf Suite = "AI" Then
        Set tests = Application.Run("T_" & Suite & ".TestNames")
        If tests.Count <> 29 Then Err.Raise vbObjectError + 715, "NxTestHarness", "AI suite inventory must be exactly 29"
    ElseIf Suite = "G005" Then
        Set tests = Application.Run("T_" & Suite & ".TestNames")
        If tests.Count <> 10 Then Err.Raise vbObjectError + 715, "NxTestHarness", "G005 suite inventory must be exactly 10"
    ElseIf Suite = "Template" Then
        Set tests = Application.Run("T_" & Suite & ".TestNames")
        If tests.Count <> 30 Then Err.Raise vbObjectError + 715, "NxTestHarness", "Template suite inventory must be exactly 30"
    ElseIf Suite = "Data" Then
        Set tests = Application.Run("T_" & Suite & ".TestNames")
        If tests.Count <> 17 Then Err.Raise vbObjectError + 715, "NxTestHarness", "Data suite inventory must be exactly 17"
    ElseIf Suite = "Draw" Then
        Set tests = Application.Run("T_" & Suite & ".TestNames")
        If tests.Count <> 10 Then Err.Raise vbObjectError + 715, "NxTestHarness", "Draw suite inventory must be exactly 10"
    ElseIf Suite = "File" Then
        Set tests = Application.Run("T_" & Suite & ".TestNames")
        If tests.Count <> 11 Then Err.Raise vbObjectError + 715, "NxTestHarness", "File suite inventory must be exactly 11"
    ElseIf Suite = "Calculator" Then
        Set tests = Application.Run("T_" & Suite & ".TestNames")
        If tests.Count <> 28 Then Err.Raise vbObjectError + 715, "NxTestHarness", "Calculator suite inventory must be exactly 28"
    ElseIf Suite = "Symbols" Then
        Set tests = Application.Run("T_" & Suite & ".TestNames")
        If tests.Count <> 24 Then Err.Raise vbObjectError + 715, "NxTestHarness", "Symbols suite inventory must be exactly 24"
    Else
        Err.Raise vbObjectError + 710, "NxTestHarness", "Unknown suite"
    End If
    Dim passed As Long, failed As Long, skipped As Long
    Dim details As String, item As Variant, failure As String, halted As Boolean
    Dim guard As CNxStateGuard
    For Each item In tests
        If halted Or (Len(Filter) > 0 And CStr(item) <> Filter) Then
            skipped = skipped + 1
            If halted Then
                AppendCase details, CStr(item), "skipped", "Harness stopped after application-state contamination"
            Else
                AppendCase details, CStr(item), "skipped", vbNullString
            End If
        Else
            Set guard = New CNxStateGuard
            On Error GoTo CaseFailed
            If Suite = "Core" Then
                T_Core.RunCase CStr(item)
            Else
                Application.Run "T_" & Suite & ".RunCase", CStr(item)
            End If
            guard.Restore
            If Len(guard.Recovery) > 0 Then Err.Raise vbObjectError + 714, "NxTestHarness", "Harness state cleanup failed: " & guard.Recovery
            passed = passed + 1
            AppendCase details, CStr(item), "passed", vbNullString
            Set guard = Nothing
            GoTo CaseDone
CaseFailed:
            failure = CStr(Err.Number) & " " & Err.Description
            Err.Clear
            On Error Resume Next
            If Not guard Is Nothing Then guard.EmergencyCleanup
            On Error GoTo 0
            failed = failed + 1
            AppendCase details, CStr(item), "failed", failure
            halted = True
            Set guard = Nothing
CaseDone:
            On Error GoTo 0
        End If
    Next item
    RunSuite = "{""total"":" & CStr(tests.Count) & ",""passed"":" & CStr(passed) & ",""failed"":" & CStr(failed) & ",""skipped"":" & CStr(skipped) & ",""tests"":[" & details & "]}"
End Function

Public Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 711, "NxTestHarness", message
End Sub

Public Sub AssertSameObject(ByVal expected As Object, ByVal actual As Object, ByVal message As String)
    If expected Is Nothing Or actual Is Nothing Then Err.Raise vbObjectError + 713, "NxTestHarness", message
    If Not expected Is actual Then Err.Raise vbObjectError + 713, "NxTestHarness", message
End Sub

Private Sub AppendCase(ByRef details As String, ByVal name As String, ByVal status As String, ByVal failure As String)
    If Len(details) > 0 Then details = details & ","
    details = details & "{""name"":""" & EscapeJson(name) & """,""status"":""" & EscapeJson(status) & """,""error"":""" & EscapeJson(failure) & """}"
End Sub

Private Function EscapeJson(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, Chr$(34), "\" & Chr$(34))
    EscapeJson = Replace(Replace(value, vbCr, "\r"), vbLf, "\n")
End Function
