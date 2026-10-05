Attribute VB_Name = "T_R62AiWorkflow"
Option Explicit

' Pure synthetic sessions and an unshown form: no clipboard, files or worksheet writes.
Public Function NxR62AiRunCase(ByVal kind As String) As String
    Dim session As CNxAiInputSession, snapshot As CNxAiPrivacySnapshot
    Dim view As FNxAi, before As Long, prompt As String, feature As Variant, task As Variant
    On Error GoTo Failed
    Set session = NewSession(NxPrivacySafe)
    Select Case kind
        Case "tasks"
            For Each feature In Array("NX-AI-SUMMARY", "NX-AI-CLEAN", "NX-AI-FORMULA", "NX-AI-WRITE", "NX-AI-IMAGE")
                session.SetMode CStr(feature)
                For Each task In NxAiTaskChoices(CStr(feature))
                    session.SetTask CStr(task)
                    session.SetPurpose NxAiTaskPurpose(CStr(feature), CStr(task))
                    prompt = session.PreviewPrompt()
                    Require InStr(1, prompt, "세부 작업: " & CStr(task), vbBinaryCompare) > 0, "task missing from prompt"
                Next task
            Next feature
        Case "compatibility"
            before = session.PreviewGeneration
            session.SetOfficeVersion "Microsoft 365"
            Require session.PreviewGeneration > before, "Office change retained stale preview"
            Require InStr(1, session.PreviewPrompt(), "Microsoft 365", vbBinaryCompare) > 0, "Office context missing"
            session.SetInstruction "사용자 작성 요청 유지"
            session.SetMode "NX-AI-FORMULA"
            Require InStr(1, session.PreviewPrompt(), "사용자 작성 요청 유지", vbBinaryCompare) > 0, "mode change erased request"
            before = session.PreviewGeneration
            session.SetTargetAI "Claude"
            Require session.TargetAI = "Claude", "target AI selection lost"
            Require session.PreviewGeneration > before, "target AI change retained stale approval"
            Require NxAiTargetUrl("ChatGPT") = "https://chatgpt.com/", "ChatGPT URL differs"
            Require NxAiTargetUrl("Claude") = "https://claude.ai/", "Claude URL differs"
            Require NxAiTargetUrl("Gemini") = "https://gemini.google.com/", "Gemini URL differs"
            Require TargetRejected(session), "unlisted AI target accepted"
            Require session.TargetAI = "Claude", "invalid AI target changed valid selection"
            Require Not NxAiOpenTargetAfterCopy("Claude", NxFrameCancelled), "cancelled copy requested browser"
            Require Not NxAiOpenTargetAfterCopy("Claude", NxFrameError), "failed copy requested browser"
        Case "masked_preview"
            Set session = NewSession(NxPrivacyBlocked)
            Require session.InputPreview = "[masked fixture]", "blocked input exposed original in data preview"
            Require Not CanPreview(session), "blocked input copied without privacy decision"
            session.ChooseMaskedCopy
            prompt = session.PreviewPrompt()
            Require InStr(1, prompt, "[masked fixture]", vbBinaryCompare) > 0, "masked input missing"
            Require InStr(1, prompt, "synthetic original", vbBinaryCompare) = 0, "masked prompt retained original"
        Case "reselect"
            Set session = NewSession(NxPrivacyCaution)
            session.ChooseOriginalCopy True
            before = session.PreviewGeneration
            Set snapshot = NewSnapshot(NxPrivacyCaution)
            session.ReplaceSelection "synthetic original", "Fixture!$B$2", "1셀", snapshot
            Require session.TargetSummary = "Fixture!$B$2", "target was not replaced"
            Require session.PreviewGeneration > before, "reselect retained stale preview"
            Require session.InputPreview = "[masked fixture]", "reselect retained original approval"
            Require Not CanPreview(session), "reselect retained privacy approval"
        Case "reject_mismatch"
            before = session.PreviewGeneration
            Set snapshot = NewSnapshot(NxPrivacySafe)
            Require ReplacementRejected(session, snapshot), "mismatched input/snapshot accepted"
            Require session.TargetSummary = "Fixture!$A$1", "invalid reselect changed target"
            Require session.PreviewGeneration = before, "invalid reselect invalidated prior valid input"
        Case "incomplete"
            Set session = NewSession(NxPrivacyIncomplete)
            Require Not CanPreview(session), "uninspected picture accepted"
            Require InStr(1, session.InputPreview, "실제 이미지는 복사하지 않습니다", vbBinaryCompare) > 0, "image boundary missing"
        Case "form"
            Set view = New FNxAi
            view.BindSession session, "NX-AI-SUMMARY"
            view.Controls("txtInstruction").Value = "직접 쓴 추가 요청"
            view.Controls("cboMode").ListIndex = 2
            Require view.Controls("cboTask").ListCount = 4, "v3.4 task choices did not reload"
            Require CStr(view.Controls("txtInstruction").Value) = "직접 쓴 추가 요청", "task switch erased instructions"
            Require CBool(view.Controls("cmdExecute").Enabled), "edited request cannot execute directly"
            Require view.Controls("txtPromptPreview").Height >= 80, "preview too small"
            Require view.Controls("cboTargetAI").Top = view.Controls("cboMode").Top, "four-column option row lost"
            Require view.Controls("cboTargetAI").ListCount = 3, "target AI choices lost"
            Require view.Controls("cmdPreview").Left < view.Controls("cmdExecute").Left, "preview/copy order changed"
            Require view.Controls("cmdExecute").Left < view.Controls("cmdCancel").Left, "copy/cancel order changed"
            Require CBool(view.Controls("chkMask").Value), "masking did not default on"
            Require InStr(1, CStr(view.Controls("txtPromptPreview").Value), "대상 AI:", vbBinaryCompare) > 0, "initial summary missing target"
            view.RequestCancel
            Set view = Nothing
            Require session.IsCancelled, "form cancel did not cancel input"
        Case Else
            Err.Raise vbObjectError + 962, "T_R62AiWorkflow", "unknown case"
    End Select
    NxR62AiRunCase = "PASS|" & kind
    Exit Function
Failed:
    NxR62AiRunCase = "FAIL|" & kind & "|" & CStr(Err.Number) & "|" & Err.Description
    On Error Resume Next
    If Not view Is Nothing Then view.RequestCancel
    On Error GoTo 0
End Function

Private Function TargetRejected(ByVal session As CNxAiInputSession) As Boolean
    On Error GoTo Rejected
    session.SetTargetAI "https://example.invalid/?prompt=fixture"
    Exit Function
Rejected:
    TargetRejected = (InStr(1, Err.Description, "대상 AI", vbBinaryCompare) > 0)
    Err.Clear
End Function

Private Function NewSnapshot(ByVal state As NxPrivacyState) As CNxAiPrivacySnapshot
    Dim snapshot As New CNxAiPrivacySnapshot
    Dim masked As String
    masked = "synthetic original"
    If state = NxPrivacyCaution Or state = NxPrivacyBlocked Then masked = "[masked fixture]"
    snapshot.Configure state, "synthetic original", masked, "synthetic privacy state"
    snapshot.Seal
    Set NewSnapshot = snapshot
End Function

Private Function NewSession(ByVal state As NxPrivacyState) As CNxAiInputSession
    Dim session As New CNxAiInputSession
    session.ConfigureSelection "synthetic original", "Fixture!$A$1", "1셀", NewSnapshot(state)
    session.SetMode "NX-AI-SUMMARY"
    session.SetPurpose "fixture purpose"
    session.SetOutputFormat "간결한 요약"
    Set NewSession = session
End Function

Private Function CanPreview(ByVal session As CNxAiInputSession) As Boolean
    Dim prompt As String
    On Error GoTo Rejected
    prompt = session.PreviewPrompt()
    CanPreview = (Len(prompt) > 0)
Rejected:
    Err.Clear
End Function

Private Function ReplacementRejected(ByVal session As CNxAiInputSession, ByVal snapshot As CNxAiPrivacySnapshot) As Boolean
    On Error GoTo Rejected
    session.ReplaceSelection "different fixture", "Fixture!$B$2", "1셀", snapshot
    Exit Function
Rejected:
    ReplacementRejected = (InStr(1, Err.Description, "differ", vbBinaryCompare) > 0)
    Err.Clear
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 963, "T_R62AiWorkflow", detail
End Sub
