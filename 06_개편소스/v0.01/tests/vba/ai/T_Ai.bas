Attribute VB_Name = "T_Ai"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestAreasHaveStableOrder"
    names.Add "TestOverlappingCellsEmitOnce"
    names.Add "TestCrossWorkbookRangesRejected"
    names.Add "TestAreaLimitRejected"
    names.Add "TestCellLimitRejected"
    names.Add "TestRangeSelectionAccepted"
    names.Add "TestSinglePictureRequiresImageMode"
    names.Add "TestMixedSelectionRejected"
    names.Add "TestSummaryPromptSections"
    names.Add "TestCleanPromptSections"
    names.Add "TestFormulaPromptDistinguishesFormulaR1C1"
    names.Add "TestWritePromptSections"
    names.Add "TestImagePromptMarksPrivacyIncomplete"
    names.Add "TestPrivacyIncompleteBlocks"
    names.Add "TestBlockedRequiresMaskedCopy"
    names.Add "TestPromptRejectsUnknownMode"
    names.Add "TestClipboardSuccess"
    names.Add "TestClipboardFailure"
    names.Add "TestClipboardResultHasNoExternalCompletion"
    names.Add "TestAiCommandIsClipboardOnly"
    names.Add "TestAiRegistryHasNoTransferCapabilities"
    names.Add "TestAiFormHasRequiredControls"
    names.Add "TestAiComboChoicesAreFixed"
    names.Add "TestInputChangeInvalidatesPreview"
    names.Add "TestPrivacyIncompleteBlocksPreview"
    names.Add "TestMaskAndOriginalPolicyBindPrompt"
    names.Add "TestAiFormHasNoProviderControls"
    names.Add "TestEscapeAndCloseCancel"
    names.Add "TestForbiddenAutomationPathAbsent"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestAreasHaveStableOrder": TestAreasHaveStableOrder
        Case "TestOverlappingCellsEmitOnce": TestOverlappingCellsEmitOnce
        Case "TestCrossWorkbookRangesRejected": TestCrossWorkbookRangesRejected
        Case "TestAreaLimitRejected": TestAreaLimitRejected
        Case "TestCellLimitRejected": TestCellLimitRejected
        Case "TestRangeSelectionAccepted": TestRangeSelectionAccepted
        Case "TestSinglePictureRequiresImageMode": TestSinglePictureRequiresImageMode
        Case "TestMixedSelectionRejected": TestMixedSelectionRejected
        Case "TestSummaryPromptSections": TestSummaryPromptSections
        Case "TestCleanPromptSections": TestCleanPromptSections
        Case "TestFormulaPromptDistinguishesFormulaR1C1": TestFormulaPromptDistinguishesFormulaR1C1
        Case "TestWritePromptSections": TestWritePromptSections
        Case "TestImagePromptMarksPrivacyIncomplete": TestImagePromptMarksPrivacyIncomplete
        Case "TestPrivacyIncompleteBlocks": TestPrivacyIncompleteBlocks
        Case "TestBlockedRequiresMaskedCopy": TestBlockedRequiresMaskedCopy
        Case "TestPromptRejectsUnknownMode": TestPromptRejectsUnknownMode
        Case "TestClipboardSuccess": TestClipboardSuccess
        Case "TestClipboardFailure": TestClipboardFailure
        Case "TestClipboardResultHasNoExternalCompletion": TestClipboardResultHasNoExternalCompletion
        Case "TestAiCommandIsClipboardOnly": TestAiCommandIsClipboardOnly
        Case "TestAiRegistryHasNoTransferCapabilities": TestAiRegistryHasNoTransferCapabilities
        Case "TestAiFormHasRequiredControls": TestAiFormHasRequiredControls
        Case "TestAiComboChoicesAreFixed": TestAiComboChoicesAreFixed
        Case "TestInputChangeInvalidatesPreview": TestInputChangeInvalidatesPreview
        Case "TestPrivacyIncompleteBlocksPreview": TestPrivacyIncompleteBlocksPreview
        Case "TestMaskAndOriginalPolicyBindPrompt": TestMaskAndOriginalPolicyBindPrompt
        Case "TestAiFormHasNoProviderControls": TestAiFormHasNoProviderControls
        Case "TestEscapeAndCloseCancel": TestEscapeAndCloseCancel
        Case "TestForbiddenAutomationPathAbsent": TestForbiddenAutomationPathAbsent
        Case Else: Err.Raise vbObjectError + 881, "T_Ai", "Unknown test: " & name
    End Select
End Sub

Private Sub TestAreasHaveStableOrder()
    Dim fixtureBook As Workbook, sheet As Worksheet
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    Set fixtureBook = Application.Workbooks.Add
    Set sheet = fixtureBook.Worksheets(1)
    sheet.Range("A1:B2").Value2 = "first"
    sheet.Range("B2:C3").Value2 = "second"
    Dim ranges As New Collection
    ranges.Add sheet.Range("B2:C3")
    ranges.Add sheet.Range("A1:B2")
    Dim output As String
    output = NxAiSerializeRanges(ranges)
    NxTestHarness.AssertTrue InStr(1, output, "영역 1|", vbBinaryCompare) > 0, "Area 1 header missing"
    NxTestHarness.AssertTrue InStr(1, output, "$A$1:$B$2", vbBinaryCompare) < InStr(1, output, "$B$2:$C$3", vbBinaryCompare), "Area order is unstable"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Err.Clear
Cleanup:
    On Error Resume Next
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_Ai.TestAreasHaveStableOrder", failureDescription
End Sub

Private Sub TestOverlappingCellsEmitOnce()
    Dim fixtureBook As Workbook, sheet As Worksheet
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    Set fixtureBook = Application.Workbooks.Add
    Set sheet = fixtureBook.Worksheets(1)
    sheet.Range("B2").Value2 = "shared-value"
    Dim ranges As New Collection
    ranges.Add sheet.Range("A1:B2")
    ranges.Add sheet.Range("B2:C3")
    Dim output As String
    output = NxAiSerializeRanges(ranges)
    NxTestHarness.AssertTrue CountText(output, "shared-value") = 1, "Overlapping value was emitted more than once"
    NxTestHarness.AssertTrue CountText(output, "<겹침 제외>") = 1, "Overlap exclusion marker mismatch"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Err.Clear
Cleanup:
    On Error Resume Next
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_Ai.TestOverlappingCellsEmitOnce", failureDescription
End Sub

Private Sub TestCrossWorkbookRangesRejected()
    Dim firstBook As Workbook, secondBook As Workbook
    Dim ranges As New Collection
    Set firstBook = ActiveWorkbook
    ranges.Add firstBook.Worksheets(1).Range("A1")
    Set secondBook = Application.Workbooks.Add
    ranges.Add secondBook.Worksheets(1).Range("A1")
    Dim rejected As Boolean, detail As String
    On Error GoTo Expected
    Call NxAiSerializeRanges(ranges)
    GoTo Cleanup
Expected:
    rejected = True
    detail = Err.Description
    Err.Clear
Cleanup:
    On Error Resume Next
    secondBook.Close False
    firstBook.Activate
    On Error GoTo 0
    NxTestHarness.AssertTrue rejected And InStr(1, detail, "서로 다른 통합문서", vbBinaryCompare) > 0, "Cross-workbook input was accepted"
End Sub

Private Sub TestAreaLimitRejected()
    Dim ranges As New Collection, index As Long
    For index = 1 To 21
        ranges.Add ActiveSheet.Range("A1")
    Next index
    AssertSerializeRejected ranges, "최대 20개"
End Sub

Private Sub TestCellLimitRejected()
    AssertSerializeRejected ActiveSheet.Range("A1:A50001"), "최대 50,000개"
End Sub

Private Sub TestRangeSelectionAccepted()
    NxTestHarness.AssertTrue NxAiValidateSelectionKind(1, 0, "NX-AI-SUMMARY") = "RANGE", "Range selection was rejected"
End Sub

Private Sub TestSinglePictureRequiresImageMode()
    NxTestHarness.AssertTrue NxAiValidateSelectionKind(0, 1, "NX-AI-IMAGE") = "PICTURE", "Image mode rejected one picture"
    AssertSelectionRejected 0, 1, "NX-AI-SUMMARY", "이미지 활용 모드"
End Sub

Private Sub TestMixedSelectionRejected()
    AssertSelectionRejected 1, 1, "NX-AI-IMAGE", "동시에 선택"
End Sub

Private Sub TestSummaryPromptSections()
    AssertPromptSections BuildPrompt("NX-AI-SUMMARY", NxPrivacySafe, "SAFE", False, False)
End Sub

Private Sub TestCleanPromptSections()
    Dim prompt As String
    prompt = BuildPrompt("NX-AI-CLEAN", NxPrivacySafe, "SAFE", False, False)
    AssertPromptSections prompt
    NxTestHarness.AssertTrue InStr(1, prompt, "정규화 규칙", vbBinaryCompare) > 0, "Clean constraint missing"
End Sub

Private Sub TestFormulaPromptDistinguishesFormulaR1C1()
    Dim prompt As String
    prompt = BuildPrompt("NX-AI-FORMULA", NxPrivacySafe, "SAFE", False, False)
    AssertPromptSections prompt
    NxTestHarness.AssertTrue InStr(1, prompt, "표시값", vbBinaryCompare) > 0, "Displayed value contract missing"
    NxTestHarness.AssertTrue InStr(1, prompt, "FormulaR1C1", vbBinaryCompare) > 0, "FormulaR1C1 contract missing"
End Sub

Private Sub TestWritePromptSections()
    AssertPromptSections BuildPrompt("NX-AI-WRITE", NxPrivacySafe, "SAFE", False, False)
End Sub

Private Sub TestImagePromptMarksPrivacyIncomplete()
    Dim prompt As String
    prompt = BuildPrompt("NX-AI-IMAGE", NxPrivacyCaution, "MASK", True, False)
    AssertPromptSections prompt
    NxTestHarness.AssertTrue InStr(1, prompt, "이미지 내용 검사 미완료", vbBinaryCompare) > 0, "Image inspection warning missing"
End Sub

Private Sub TestPrivacyIncompleteBlocks()
    AssertPromptRejected "NX-AI-SUMMARY", NxPrivacyIncomplete, "MASK", True, False, "검사 미완료"
End Sub

Private Sub TestBlockedRequiresMaskedCopy()
    AssertPromptRejected "NX-AI-SUMMARY", NxPrivacyBlocked, "ORIGINAL", False, False, "차단"
    AssertPromptRejected "NX-AI-SUMMARY", NxPrivacyBlocked, "MASK", False, False, "마스킹된 입력"
    NxTestHarness.AssertTrue Len(BuildPrompt("NX-AI-SUMMARY", NxPrivacyBlocked, "MASK", True, False)) > 0, "Masked copy was blocked"
    AssertPromptRejected "NX-AI-SUMMARY", NxPrivacyBlocked, "ORIGINAL", False, True, "차단"
End Sub

Private Sub TestPromptRejectsUnknownMode()
    AssertPromptRejected "NX-AI-UNKNOWN", NxPrivacySafe, "SAFE", False, False, "지원하지 않는"
End Sub

Private Sub TestClipboardSuccess()
    NxClipboardSetTestFailure False
    Dim result As CNxResult
    Set result = NxAiCopyPrompt("합성 프롬프트", "NX-AI-SUMMARY")
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Clipboard copy did not succeed"
    NxTestHarness.AssertTrue result.Completed.Count = 1, "Clipboard completion count changed"
    NxTestHarness.AssertTrue ResultContains(result, "clipboard"), "Clipboard completion was lost"
End Sub

Private Sub TestClipboardFailure()
    NxClipboardSetTestFailure True
    Dim result As CNxResult
    Set result = NxAiCopyPrompt("합성 프롬프트", "NX-AI-SUMMARY")
    NxClipboardSetTestFailure False
    NxTestHarness.AssertTrue result.Outcome = NxEnvironmentError, "Clipboard failure outcome mismatch"
    NxTestHarness.AssertTrue result.Completed.Count = 0, "Failed clipboard recorded completion"
End Sub

Private Sub TestClipboardResultHasNoExternalCompletion()
    NxClipboardSetTestFailure False
    Dim result As CNxResult
    Set result = NxAiCopyPrompt("합성 프롬프트", "NX-AI-SUMMARY")
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Clipboard-only delivery did not succeed"
    NxTestHarness.AssertTrue ResultContains(result, "clipboard"), "Clipboard completion was lost"
    NxTestHarness.AssertTrue Not ResultContains(result, "browser"), "Browser completion must not exist"
    NxTestHarness.AssertTrue Not ResultContains(result, "network"), "Network completion must not exist"
    NxTestHarness.AssertTrue Not ResultContains(result, "process"), "Process completion must not exist"
End Sub

Private Sub TestAiCommandIsClipboardOnly()
    Dim session As CNxAiInputSession
    Dim command As CNxAiFeatureCommand
    Set session = NewAiSession(NxPrivacySafe)
    CompleteAiInputs session
    Set command = session.CreateCommand
    NxTestHarness.AssertTrue command.DeliveryMode = "CLIPBOARD_ONLY", "AI command delivery boundary changed"
    NxTestHarness.AssertTrue Len(session.PreviewPrompt) > 0, "AI command prompt was not prepared"
End Sub

Private Sub TestAiRegistryHasNoTransferCapabilities()
    Dim definition As CNxFeatureDefinition
    Set definition = NxAiFeatureDefinition("NX-AI-SUMMARY")
    NxTestHarness.AssertTrue definition.AiProcessing, "AI processing marker was lost"
    NxTestHarness.AssertTrue Not definition.ExternalDataTransfer, "AI feature declared external transfer"
    NxTestHarness.AssertTrue Not definition.NetworkTransfer, "AI feature declared network transfer"
    NxTestHarness.AssertTrue Not definition.LocalProcessTransfer, "AI feature declared process transfer"
End Sub

Private Sub TestAiFormHasRequiredControls()
    Dim view As FNxAi
    Set view = New FNxAi
    Dim required As Variant, name As Variant
    required = Array("cboMode", "cboTask", "cboOffice", "cboTargetAI", "txtInstruction", "lblLocalDelivery", _
        "lblSelection", "txtRange", "cmdPickRange", "lblPrivacy", "chkMask", "txtPromptPreview", _
        "cmdInputPreview", "cmdTogglePrompt", "cmdPreview", "cmdExecute", "cmdCancel")
    For Each name In required
        NxTestHarness.AssertTrue HasControl(view, CStr(name)), "Missing AI form control: " & CStr(name)
    Next name
    Unload view
End Sub

Private Sub TestAiComboChoicesAreFixed()
    Dim status As String
    Dim evidenceSha256 As String
    Dim evidenceRunId As String

    status = Environ$("LHEXCEL_AI_COMBO_EVIDENCE_STATUS")
    evidenceSha256 = Environ$("LHEXCEL_AI_COMBO_EVIDENCE_SHA256")
    evidenceRunId = Environ$("LHEXCEL_AI_COMBO_EVIDENCE_RUN_ID")

    NxTestHarness.AssertTrue status = "VERIFIED_EXACT_V1", "AI ComboBox build evidence was not strictly verified"
    NxTestHarness.AssertTrue Len(evidenceSha256) = 64, "AI ComboBox build evidence SHA-256 is missing"
    NxTestHarness.AssertTrue Len(evidenceRunId) = 36, "AI ComboBox build evidence run binding is missing"
End Sub

Private Sub TestInputChangeInvalidatesPreview()
    Dim session As CNxAiInputSession
    Set session = NewAiSession(NxPrivacySafe)
    session.SetPurpose "첫 목적"
    Dim beforeGeneration As Long
    beforeGeneration = session.PreviewGeneration
    session.SetPurpose "변경 목적"
    NxTestHarness.AssertTrue session.PreviewGeneration = beforeGeneration + 1, "Input edit did not invalidate preview"
    beforeGeneration = session.PreviewGeneration
    session.SetInstruction "사용자가 수정한 요청"
    NxTestHarness.AssertTrue session.PreviewGeneration = beforeGeneration + 1, "Instruction edit did not invalidate preview"
    beforeGeneration = session.PreviewGeneration
    session.SetOfficeVersion "Microsoft 365"
    NxTestHarness.AssertTrue session.PreviewGeneration = beforeGeneration + 1, "Office option did not invalidate preview"
    beforeGeneration = session.PreviewGeneration
    session.SetTargetAI "Claude"
    NxTestHarness.AssertTrue session.PreviewGeneration = beforeGeneration + 1, "Target AI change did not invalidate preview"
End Sub

Private Sub TestPrivacyIncompleteBlocksPreview()
    Dim session As CNxAiInputSession
    Dim command As CNxAiFeatureCommand
    Set session = NewAiSession(NxPrivacyIncomplete)
    CompleteAiInputs session
    Dim rejected As Boolean
    On Error GoTo Expected
    Set command = session.CreateCommand
    GoTo Finished
Expected:
    rejected = InStr(1, Err.Description, "검사 미완료", vbBinaryCompare) > 0
    Err.Clear
Finished:
    NxTestHarness.AssertTrue rejected, "Privacy-incomplete input created a command"
End Sub

Private Sub TestMaskAndOriginalPolicyBindPrompt()
    Dim masked As CNxAiInputSession
    Dim command As CNxAiFeatureCommand
    Dim prompt As String
    Set masked = NewAiSession(NxPrivacyBlocked)
    CompleteAiInputs masked
    masked.ChooseMaskedCopy
    Set command = masked.CreateCommand
    NxTestHarness.AssertTrue Not command Is Nothing, "Masked copy did not create a command"
    prompt = masked.PreviewPrompt()
    NxTestHarness.AssertTrue InStr(1, prompt, "[비밀값 마스킹]", vbBinaryCompare) > 0, "Masked prompt lost the sanitized input"
    NxTestHarness.AssertTrue InStr(1, prompt, "표시값=10", vbBinaryCompare) = 0, "Masked prompt leaked the original input"

    AssertOriginalCopyRejected NxPrivacyCaution, False, "두 번째 확인"

    Dim original As CNxAiInputSession
    Set original = NewAiSession(NxPrivacyCaution)
    CompleteAiInputs original
    original.ChooseOriginalCopy True
    Set command = original.CreateCommand
    NxTestHarness.AssertTrue Not command Is Nothing, "Confirmed original did not create a command"
    NxTestHarness.AssertTrue InStr(1, original.PreviewPrompt(), "표시값=10", vbBinaryCompare) > 0, "Confirmed original prompt lost the approved input"

    AssertOriginalCopyRejected NxPrivacyBlocked, True, "차단"
End Sub

Private Sub TestAiFormHasNoProviderControls()
    Dim view As FNxAi
    Set view = New FNxAi
    NxTestHarness.AssertTrue HasControl(view, "lblLocalDelivery"), "Local delivery notice is missing"
    ' Retain the native case name; a closed homepage target is not an API provider control.
    NxTestHarness.AssertTrue HasControl(view, "cboTargetAI"), "Closed AI homepage target selector is missing"
    NxTestHarness.AssertTrue Not HasControl(view, "cboProvider"), "AI provider control must not exist"
    NxTestHarness.AssertTrue Not HasControl(view, "cboModel"), "AI model control must not exist"
    NxTestHarness.AssertTrue Not HasControl(view, "txtApiKey") And Not HasControl(view, "cmdSendApi"), "API credentials or upload control must not exist"
    NxTestHarness.AssertTrue Not HasControl(view, "txtPurpose") And Not HasControl(view, "cboOutput"), "Removed generic workbench controls remain"
    Unload view
End Sub

Private Sub TestEscapeAndCloseCancel()
    Dim session As CNxAiInputSession, view As FNxAi
    TraceCancelProgress "01-start"
    Set session = NewAiSession(NxPrivacySafe)
    TraceCancelProgress "02-session"
    Set view = New FNxAi
    TraceCancelProgress "03-form"
    view.BindSession session
    TraceCancelProgress "04-bound"
    view.RequestCancel
    TraceCancelProgress "05-cancel-returned"
    NxTestHarness.AssertTrue session.IsCancelled, "AI form cancel was not terminal"
    TraceCancelProgress "06-asserted"
    Set view = Nothing
    TraceCancelProgress "07-released"
End Sub

Private Sub TraceCancelProgress(ByVal marker As String)
    Dim progressPath As String
    Dim handle As Integer
    progressPath = Environ$("LHEXCEL_OWNER_CASE_PROGRESS")
    If Len(progressPath) = 0 Then Exit Sub
    handle = FreeFile
    Open progressPath For Append Access Write As #handle
    Print #handle, marker
    Close #handle
End Sub

Private Sub TestForbiddenAutomationPathAbsent()
    Dim view As FNxAi
    Set view = New FNxAi
    NxTestHarness.AssertTrue Not HasControl(view, "cmdCopy"), "Direct copy control exists"
    NxTestHarness.AssertTrue Not HasControl(view, "cmdCopyOpen"), "Direct copy-open control exists"
    NxTestHarness.AssertTrue Not HasControl(view, "cmdSubmit"), "Automatic submit control exists"
    NxTestHarness.AssertTrue Not HasControl(view, "cmdAnswerImport"), "Answer import control exists"
    Unload view
End Sub

Private Function NewAiSession(ByVal state As NxPrivacyState) As CNxAiInputSession
    Dim snapshot As New CNxAiPrivacySnapshot
    Dim session As New CNxAiInputSession
    Dim masked As String
    masked = "표시값=10"
    If state = NxPrivacyCaution Or state = NxPrivacyBlocked Then masked = "[비밀값 마스킹]"
    snapshot.Configure state, "표시값=10", masked, "합성 개인정보 상태"
    snapshot.Seal
    session.ConfigureSelection "표시값=10", "Sheet1!$A$1", "1셀", snapshot
    Set NewAiSession = session
End Function

Private Sub CompleteAiInputs(ByVal session As CNxAiInputSession)
    session.SetMode "NX-AI-SUMMARY"
    session.SetPurpose "검증용 목적"
    session.SetOutputFormat "간결한 요약"
    session.SetInstruction vbNullString
End Sub

Private Sub AssertOriginalCopyRejected(ByVal privacyState As NxPrivacyState, ByVal secondConfirmation As Boolean, ByVal expectedText As String)
    Dim session As CNxAiInputSession
    Dim rejected As Boolean, detail As String
    Set session = NewAiSession(privacyState)
    CompleteAiInputs session
    On Error GoTo Expected
    session.ChooseOriginalCopy secondConfirmation
    GoTo Verify
Expected:
    rejected = True
    detail = Err.Description
    Err.Clear
    Resume Verify
Verify:
    On Error GoTo 0
    NxTestHarness.AssertTrue rejected, "Original copy was accepted"
    NxTestHarness.AssertTrue InStr(1, detail, expectedText, vbBinaryCompare) > 0, "Wrong original copy rejection"
End Sub

Private Function HasControl(ByVal view As Object, ByVal controlName As String) As Boolean
    Dim control As Object
    On Error Resume Next
    Set control = view.Controls(controlName)
    HasControl = (Err.Number = 0 And Not control Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function

Private Function ResultContains(ByVal result As CNxResult, ByVal expected As String) As Boolean
    Dim item As Variant
    For Each item In result.Completed
        If CStr(item) = expected Then ResultContains = True: Exit Function
    Next item
End Function

Private Function BuildPrompt(ByVal featureId As String, ByVal privacyState As NxPrivacyState, ByVal privacyDecision As String, ByVal inputWasMasked As Boolean, ByVal originalConfirmed As Boolean) As String
    BuildPrompt = NxAiBuildPrompt(featureId, "표시값=10|FormulaR1C1=RC[-1]", "검증용 목적", "목록", "합성 fixture만 사용", privacyState, privacyDecision, inputWasMasked, originalConfirmed)
End Function

Private Sub AssertPromptSections(ByVal prompt As String)
    Dim purposeAt As Long, inputAt As Long, constraintsAt As Long, outputAt As Long
    purposeAt = InStr(1, prompt, "[목적]", vbBinaryCompare)
    inputAt = InStr(1, prompt, "[입력]", vbBinaryCompare)
    constraintsAt = InStr(1, prompt, "[제약]", vbBinaryCompare)
    outputAt = InStr(1, prompt, "[원하는 출력]", vbBinaryCompare)
    NxTestHarness.AssertTrue purposeAt > 0 And purposeAt < inputAt, "Purpose/input section order mismatch"
    NxTestHarness.AssertTrue inputAt < constraintsAt And constraintsAt < outputAt, "Constraint/output section order mismatch"
End Sub

Private Sub AssertPromptRejected(ByVal featureId As String, ByVal privacyState As NxPrivacyState, ByVal privacyDecision As String, ByVal inputWasMasked As Boolean, ByVal originalConfirmed As Boolean, ByVal expectedText As String)
    On Error GoTo Expected
    Call BuildPrompt(featureId, privacyState, privacyDecision, inputWasMasked, originalConfirmed)
    Err.Raise vbObjectError + 885, "T_Ai", "Invalid prompt request was accepted"
Expected:
    Dim detail As String
    detail = Err.Description
    Err.Clear
    NxTestHarness.AssertTrue InStr(1, detail, expectedText, vbBinaryCompare) > 0, "Wrong prompt rejection"
End Sub

Private Sub AssertSerializeRejected(ByVal inputValue As Object, ByVal expectedText As String)
    On Error GoTo Expected
    Call NxAiSerializeRanges(inputValue)
    Err.Raise vbObjectError + 882, "T_Ai", "Invalid range input was accepted"
Expected:
    Dim detail As String
    detail = Err.Description
    Err.Clear
    NxTestHarness.AssertTrue InStr(1, detail, expectedText, vbBinaryCompare) > 0, "Wrong serializer rejection"
End Sub

Private Sub AssertSelectionRejected(ByVal rangeCount As Long, ByVal pictureCount As Long, ByVal featureId As String, ByVal expectedText As String)
    On Error GoTo Expected
    Call NxAiValidateSelectionKind(rangeCount, pictureCount, featureId)
    Err.Raise vbObjectError + 883, "T_Ai", "Invalid selection was accepted"
Expected:
    Dim detail As String
    detail = Err.Description
    Err.Clear
    NxTestHarness.AssertTrue InStr(1, detail, expectedText, vbBinaryCompare) > 0, "Wrong selection rejection"
End Sub

Private Function CountText(ByVal source As String, ByVal token As String) As Long
    If Len(token) = 0 Then Exit Function
    CountText = (Len(source) - Len(Replace(source, token, vbNullString))) \ Len(token)
End Function
