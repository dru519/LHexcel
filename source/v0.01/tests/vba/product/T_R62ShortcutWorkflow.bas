Attribute VB_Name = "T_R62ShortcutWorkflow"
Option Explicit

' Draft/rejection cases do not mutate live state.
' Persistence cases require the runner-owned profile and restore their original state.
Public Function NxR62ShortcutRunCase(ByVal kind As String) As String
    Dim draft As Object, original As String, view As FNxShortcutManager
    Const firstRoute As String = "feature:NX-DATA-NORMALIZE"
    Const secondRoute As String = "feature:NX-DATA-UNIQUE-COUNT"
    On Error GoTo Failed
    Require NxShortcutsRouteIsEligible(firstRoute), "first fixture route is not shortcut eligible"
    Require NxShortcutsRouteIsEligible(secondRoute), "second fixture route is not shortcut eligible"
    original = NxShortcutsSnapshot()
    Set draft = NxShortcutsCreateDraft()
    draft.RemoveAll
    Select Case kind
        Case "draft_cancel"
            NxShortcutDraftAssign draft, firstRoute, "Ctrl+Alt+Shift", "Q"
            Require NxShortcutDraftDisplay(draft, firstRoute) = "Ctrl+Alt+Shift+Q", "draft assignment missing"
            Require NxShortcutsSnapshot() = original, "draft changed live bindings"
            Set view = New FNxShortcutManager
            Require view.Controls("lstCommands").ColumnCount = 4, "key and status columns not separated"
            Require view.Controls("lstCommands").Left + view.Controls("lstCommands").Width <= _
                view.Controls("txtDescription").Left, "v3.4 left-list/right-details layout lost"
            Require view.Controls("lblFunctionHeaderStatus").Caption = "상태", "status heading missing"
            Require view.Controls("cmdRemove").Top < view.Controls("cmdApply").Top, "reset row order changed"
            Require view.Controls("cmdAssign").Left < view.Controls("cmdApply").Left, "preview/apply order changed"
            Require view.Controls("cmdApply").Left < view.Controls("cmdClose").Left, "apply/cancel order changed"
            Require view.Controls("cboState").ListCount = 3, "assignment filter lost"
            Unload view
            Set view = Nothing
            Set draft = Nothing
        Case "draft_remove"
            NxShortcutDraftAssign draft, firstRoute, "Ctrl+Alt+Shift", "Q"
            NxShortcutDraftRemove draft, firstRoute
            Require draft.Count = 0, "draft removal failed"
        Case "draft_clear"
            NxShortcutDraftAssign draft, firstRoute, "Ctrl+Alt+Shift", "Q"
            NxShortcutDraftAssign draft, secondRoute, "Ctrl+Alt+Shift", "W"
            draft.RemoveAll
            Require draft.Count = 0, "draft clear failed"
        Case "conflict"
            NxShortcutDraftAssign draft, firstRoute, "Ctrl+Alt+Shift", "Q"
            Require Len(NxShortcutDraftConflictReason(draft, "Ctrl+Alt+Shift", "Q", firstRoute)) = 0, "same route conflicts with itself"
            Require Len(NxShortcutDraftConflictReason(draft, "Ctrl+Alt+Shift", "Q", secondRoute)) > 0, "draft duplicate was accepted"
            Require Len(NxShortcutDraftConflictReason(draft, "Ctrl+Shift", "L", firstRoute)) > 0, "Excel reserved filter shortcut accepted"
        Case "stale_apply"
            Require ApplyRejected(draft, "stale snapshot", "변경되었습니다"), "stale dialog applied"
        Case "reserved_apply"
            draft.Add firstRoute, Array("Ctrl+Shift", "L")
            Require ApplyRejected(draft, original, "예약 단축키"), "reserved draft reached live apply"
        Case "apply_roundtrip"
            TestApplyPersistence False
        Case "save_failure_rollback"
            TestApplyPersistence True
        Case Else
            Err.Raise vbObjectError + 964, "T_R62ShortcutWorkflow", "unknown case"
    End Select
    Require NxShortcutsSnapshot() = original, "fixture changed live bindings"
    NxR62ShortcutRunCase = "PASS|" & kind
    Exit Function
Failed:
    NxR62ShortcutRunCase = "FAIL|" & kind & "|" & CStr(Err.Number) & "|" & Err.Description
    On Error Resume Next
    If Not view Is Nothing Then Unload view
    On Error GoTo 0
End Function

Private Sub TestApplyPersistence(ByVal forceSaveFailure As Boolean)
    Const firstRoute As String = "feature:NX-DATA-NORMALIZE"
    Const secondRoute As String = "feature:NX-DATA-UNIQUE-COUNT"
    Dim cfgPath As String, originalSnapshot As String, originalFile As String
    Dim baselineSnapshot As String, baselineFile As String, failureSource As String
    Dim originalDraft As Object, candidate As Object
    Dim originalFileExists As Boolean, restoreNeeded As Boolean
    Dim lockHandle As Integer, saveFailure As Long
    Dim failureNumber As Long, failureDetail As String, cleanupFailure As String
    On Error GoTo Failed
    cfgPath = FixtureSettingsPath()
    Require Not FixtureFileExists(cfgPath & ".tmp"), "pre-existing shortcut temporary file rejected"
    Require Not FixtureFileExists(cfgPath & ".bak"), "pre-existing shortcut backup file rejected"
    originalSnapshot = NxShortcutsSnapshot()
    Set originalDraft = NxShortcutsCreateDraft()
    originalFileExists = FixtureFileExists(cfgPath)
    If originalFileExists Then originalFile = ReadFixtureFileText(cfgPath)
    Set candidate = NxShortcutsCreateDraft()
    candidate.RemoveAll
    NxShortcutDraftAssign candidate, firstRoute, "Ctrl+Alt+Shift", "Q"
    If Not forceSaveFailure Then NxShortcutDraftAssign candidate, secondRoute, "Ctrl+Alt+Shift", "W"
    restoreNeeded = True
    NxShortcutsApplyDraft candidate, originalSnapshot
    Require NxShortcutsSnapshot() = ExpectedFixtureSnapshot(Not forceSaveFailure), "applied snapshot differs from fixture"
    Require ReadFixtureFileText(cfgPath) = ExpectedFixtureFile(Not forceSaveFailure), "saved cfg differs from applied bindings"
    Require NxShortcutsBindingDisplayForRoute(firstRoute) = "Ctrl+Alt+Shift+Q", "successful apply did not bind first route"
    If forceSaveFailure Then
        baselineSnapshot = NxShortcutsSnapshot()
        baselineFile = ReadFixtureFileText(cfgPath)
        NxShortcutDraftAssign candidate, secondRoute, "Ctrl+Alt+Shift", "W"
        lockHandle = FreeFile
        Open cfgPath For Binary Access Read Lock Read Write As #lockHandle
        saveFailure = ApplyFailureNumber(candidate, baselineSnapshot, failureSource)
        Require saveFailure = 55 Or saveFailure = 70 Or saveFailure = 75, "locked cfg did not raise a file-access failure"
        Require failureSource = "NxShortcuts.ApplyDraft", "failure did not enter the real apply rollback"
        Require NxShortcutsSnapshot() = baselineSnapshot, "save failure retained changed live bindings"
        Require NxShortcutsBindingDisplayForRoute(secondRoute) = vbNullString, "save failure left attempted key assigned"
        Require ReadLockedFileText(lockHandle) = baselineFile, "locked original cfg changed"
        Close #lockHandle
        lockHandle = 0
        Require ReadFixtureFileText(cfgPath) = baselineFile, "rollback did not retain the previous cfg path and text"
        Require Not FixtureFileExists(cfgPath & ".tmp"), "failed save leaked temporary file"
        Require Not FixtureFileExists(cfgPath & ".bak"), "failed save leaked backup file"
    Else
        Require NxShortcutsBindingCount() = 2, "successful apply has the wrong binding count"
        Require NxShortcutsBindingDisplayForRoute(secondRoute) = "Ctrl+Alt+Shift+W", "successful apply did not bind second route"
    End If
    GoTo CleanUp
Failed:
    failureNumber = Err.Number
    failureDetail = Err.Description
    Resume CleanUp
CleanUp:
    On Error Resume Next
    Err.Clear
    If lockHandle > 0 Then Close #lockHandle
    If Err.Number <> 0 Then cleanupFailure = "owned lock close failed: " & Err.Description
    lockHandle = 0
    Err.Clear
    If restoreNeeded Then
        ' Production apply clears the fixture OnKey chords and re-registers the original ones.
        NxShortcutsApplyDraft originalDraft, NxShortcutsSnapshot()
        If Err.Number <> 0 Then cleanupFailure = cleanupFailure & " | original bindings restore failed: " & Err.Description
        Err.Clear
        If originalFileExists Then
            WriteFixtureFileText cfgPath, originalFile
        ElseIf FixtureFileExists(cfgPath) Then
            Kill cfgPath
        End If
        If Err.Number <> 0 Then cleanupFailure = cleanupFailure & " | original cfg restore failed: " & Err.Description
        Err.Clear
        If NxShortcutsSnapshot() <> originalSnapshot Then cleanupFailure = cleanupFailure & " | original snapshot not restored"
        If originalFileExists Then
            If ReadFixtureFileText(cfgPath) <> originalFile Then cleanupFailure = cleanupFailure & " | original cfg text not restored"
        ElseIf FixtureFileExists(cfgPath) Then
            cleanupFailure = cleanupFailure & " | fixture cfg remained after restore"
        End If
        If Err.Number <> 0 Then cleanupFailure = cleanupFailure & " | restoration verification failed: " & Err.Description
    End If
    On Error GoTo 0
    If Len(cleanupFailure) > 0 Then Err.Raise vbObjectError + 966, "T_R62ShortcutWorkflow", failureDetail & cleanupFailure
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_R62ShortcutWorkflow", failureDetail
End Sub

Private Function FixtureSettingsPath() As String
    Dim profileRoot As String
    profileRoot = Environ$("LHEXCEL_PROFILE_ROOT")
    Require Len(profileRoot) > 0, "persistence test requires an explicit isolated profile"
    Require StrComp(profileRoot, NxLHexcelProfileRoot(), vbTextCompare) = 0, "profile override differs from runtime profile"
    Require StrComp(profileRoot, ThisWorkbook.Path & "\profile", vbTextCompare) = 0, "profile is not beside the runner-owned add-in"
    Require CBool((GetAttr(profileRoot) And vbDirectory) = vbDirectory), "runner-owned profile folder missing"
    FixtureSettingsPath = profileRoot & "\Settings\shortcuts-v2.cfg"
End Function

Private Function ExpectedFixtureSnapshot(ByVal includeSecond As Boolean) As String
    Dim slot As Long, row As String
    For slot = 1 To 64
        row = CStr(slot) & "|"
        If slot = 1 Then
            row = row & "feature:NX-DATA-NORMALIZE|Ctrl+Alt+Shift|Q"
        ElseIf slot = 2 And includeSecond Then
            row = row & "feature:NX-DATA-UNIQUE-COUNT|Ctrl+Alt+Shift|W"
        Else
            row = row & "||"
        End If
        ExpectedFixtureSnapshot = ExpectedFixtureSnapshot & row & vbLf
    Next slot
End Function

Private Function ExpectedFixtureFile(ByVal includeSecond As Boolean) As String
    ExpectedFixtureFile = "NXSHORT2" & vbCrLf & "1|Ctrl+Alt+Shift|Q|feature:NX-DATA-NORMALIZE" & vbCrLf
    If includeSecond Then ExpectedFixtureFile = ExpectedFixtureFile & "2|Ctrl+Alt+Shift|W|feature:NX-DATA-UNIQUE-COUNT" & vbCrLf
End Function

Private Function ApplyFailureNumber(ByVal draft As Object, ByVal snapshot As String, ByRef failureSource As String) As Long
    On Error GoTo Rejected
    NxShortcutsApplyDraft draft, snapshot
    Exit Function
Rejected:
    ApplyFailureNumber = Err.Number
    failureSource = Err.Source
    Err.Clear
End Function

Private Function FixtureFileExists(ByVal path As String) As Boolean
    FixtureFileExists = (Len(Dir$(path, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
End Function

Private Function ReadLockedFileText(ByVal handle As Integer) As String
    Dim content As String
    content = Space$(LOF(handle))
    If Len(content) > 0 Then Get #handle, 1, content
    ReadLockedFileText = content
End Function

Private Function ReadFixtureFileText(ByVal path As String) As String
    Dim handle As Integer, failureNumber As Long, failureDetail As String
    On Error GoTo Failed
    handle = FreeFile
    Open path For Binary Access Read Lock Write As #handle
    ReadFixtureFileText = ReadLockedFileText(handle)
    Close #handle
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDetail = Err.Description
    On Error Resume Next
    If handle > 0 Then Close #handle
    On Error GoTo 0
    Err.Raise failureNumber, "T_R62ShortcutWorkflow.ReadCfg", failureDetail
End Function

Private Sub WriteFixtureFileText(ByVal path As String, ByVal value As String)
    Dim handle As Integer, failureNumber As Long, failureDetail As String
    On Error GoTo Failed
    handle = FreeFile
    Open path For Output Access Write Lock Read Write As #handle
    Print #handle, value;
    Close #handle
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDetail = Err.Description
    On Error Resume Next
    If handle > 0 Then Close #handle
    On Error GoTo 0
    Err.Raise failureNumber, "T_R62ShortcutWorkflow.RestoreCfg", failureDetail
End Sub

Private Function ApplyRejected(ByVal draft As Object, ByVal snapshot As String, ByVal expected As String) As Boolean
    On Error GoTo Rejected
    NxShortcutsApplyDraft draft, snapshot
    Exit Function
Rejected:
    ApplyRejected = (InStr(1, Err.Description, expected, vbBinaryCompare) > 0)
    Err.Clear
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 965, "T_R62ShortcutWorkflow", detail
End Sub
