Attribute VB_Name = "NxAiUi"
Option Explicit

Private mCurrentTargetSummary As String
Private mCurrentScaleSummary As String

Public Function NxAiTargetUrl(ByVal targetName As String) As String
    Select Case targetName
        Case "ChatGPT": NxAiTargetUrl = "https://chatgpt.com/"
        Case "Claude": NxAiTargetUrl = "https://claude.ai/"
        Case "Gemini": NxAiTargetUrl = "https://gemini.google.com/"
        Case Else: NxRaiseContractError "대상 AI를 선택하세요."
    End Select
End Function

Public Function NxAiOpenTargetAfterCopy(ByVal targetName As String, ByVal copyState As NxFrameState) As Boolean
    If copyState <> NxFrameSuccess Then Exit Function
    On Error GoTo Failed
    ThisWorkbook.FollowHyperlink Address:=NxAiTargetUrl(targetName), NewWindow:=True
    NxAiOpenTargetAfterCopy = True
Failed:
    Err.Clear
End Function

Public Sub NxAiOpenInput(Optional ByVal featureId As String = vbNullString)
    Dim serialized As String
    Dim privacy As CNxAiPrivacySnapshot
    Dim session As New CNxAiInputSession
    Dim form As New FNxAi
    serialized = NxAiCaptureCurrentInput(privacy)
    session.ConfigureSelection serialized, NxAiCurrentTargetSummary(), NxAiCurrentScaleSummary(), privacy
    form.BindSession session, featureId
    form.Show
End Sub

Public Function NxAiCaptureCurrentInput(ByRef privacy As CNxAiPrivacySnapshot) As String
    mCurrentTargetSummary = vbNullString
    mCurrentScaleSummary = vbNullString
    Set privacy = Nothing

    If TypeName(Application.Selection) = "Range" Then
        NxAiCaptureCurrentInput = NxAiCaptureRangeInput(Application.Selection, privacy)
        Exit Function
    End If

    Dim shapeRange As Object
    Dim shape As Object
    Dim pictureCount As Long
    On Error Resume Next
    Set shapeRange = Application.Selection.ShapeRange
    If Not shapeRange Is Nothing Then pictureCount = shapeRange.Count
    On Error GoTo 0
    NxAiValidateSelectionKind 0, pictureCount, "NX-AI-IMAGE"
    Set shape = shapeRange.Item(1)
    If shape.Type <> 11 And shape.Type <> 13 Then NxRaiseContractError "선택 항목은 그림 형식이어야 합니다."
    NxAiCaptureCurrentInput = "그림=" & CStr(shape.Name) & "|시트=" & CStr(shape.Parent.Name)
    Dim snapshot As New CNxAiPrivacySnapshot
    snapshot.Configure NxPrivacyIncomplete, NxAiCaptureCurrentInput, NxAiCaptureCurrentInput, "이미지 내용 검사 미완료"
    snapshot.Seal
    Set privacy = snapshot
    mCurrentTargetSummary = CStr(shape.Parent.Name) & "!" & CStr(shape.Name)
    mCurrentScaleSummary = "그림 1개"
End Function

Public Function NxAiCaptureRangeInput(ByVal selectedRange As Range, ByRef privacy As CNxAiPrivacySnapshot) As String
    Dim area As Range
    Dim ranges As New Collection
    Dim cellCount As Double
    If selectedRange Is Nothing Then NxRaiseContractError "AI에 사용할 셀 범위를 선택하세요."
    For Each area In selectedRange.Areas
        ranges.Add area
        cellCount = cellCount + CDbl(area.CountLarge)
    Next area
    NxAiValidateSelectionKind ranges.Count, 0, "NX-AI-SUMMARY"
    NxAiCaptureRangeInput = NxAiSerializeRanges(ranges)
    Set privacy = NxAiScanPrivacy(NxAiCaptureRangeInput)
    mCurrentTargetSummary = selectedRange.Worksheet.Name & "!" & selectedRange.Address(True, True, xlA1, False)
    mCurrentScaleSummary = CStr(ranges.Count) & "개 영역 · " & NxInvariantUnsigned(cellCount) & "셀"
End Function

Public Function NxAiCurrentTargetSummary() As String
    If Len(mCurrentTargetSummary) = 0 Then NxRaiseContractError "AI input target was not captured"
    NxAiCurrentTargetSummary = mCurrentTargetSummary
End Function

Public Function NxAiCurrentScaleSummary() As String
    If Len(mCurrentScaleSummary) = 0 Then NxRaiseContractError "AI input scale was not captured"
    NxAiCurrentScaleSummary = mCurrentScaleSummary
End Function

Public Function NxAiShowExecutionPlan(ByVal session As CNxAiInputSession) As NxFrameState
    If session Is Nothing Then NxRaiseContractError "AI execution plan requires an input session"
    If session.IsCancelled Then NxRaiseContractError "Cancelled AI input cannot create an execution plan"

    Dim commandObject As CNxAiFeatureCommand
    Dim command As INxFeatureCommand
    Dim definition As CNxFeatureDefinition
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim result As CNxResult
    Set commandObject = session.CreateCommand()
    Set command = commandObject
    Set definition = NxAiFeatureDefinition(session.FeatureId)
    Set ticket = router.Prepare(definition, command)
    Select Case ticket.Decision.ResolvedGrade
        Case NxExecutionFast
            Set result = router.ExecuteFast(ticket)
            NxAiShowExecutionPlan = StateForResult(result)
        Case NxExecutionGuarded
            Set result = router.ExecuteUserAction(ticket)
            NxAiShowExecutionPlan = StateForResult(result)
        Case NxExecutionPlanned
            NxAiShowExecutionPlan = ShowPlannedAiDelivery(session, commandObject, command, definition, ticket, router)
        Case Else
            NxRaiseContractError "AI delivery execution grade is invalid"
    End Select
End Function

Private Function ShowPlannedAiDelivery(ByVal session As CNxAiInputSession, _
    ByVal commandObject As CNxAiFeatureCommand, ByVal command As INxFeatureCommand, _
    ByVal definition As CNxFeatureDefinition, ByVal ticket As CNxExecutionTicket, _
    ByVal router As CNxExecutionRouter) As NxFrameState
    Dim result As CNxResult
    Set result = router.ExecuteUserAction(ticket)
    If result.Outcome <> NxSuccess And result.Outcome <> NxCancelled Then NxRaiseContractError result.Recovery
    ShowPlannedAiDelivery = StateForResult(result)
End Function

Private Function StateForResult(ByVal result As CNxResult) As NxFrameState
    If result Is Nothing Or Not result.IsSealed Then NxRaiseContractError "AI apply returned no sealed result"
    Select Case result.Outcome
        Case NxSuccess: StateForResult = NxFrameSuccess
        Case NxCancelled: StateForResult = NxFrameCancelled
        Case NxInputError, NxEnvironmentError: StateForResult = NxFrameError
        Case NxPartialFailure: StateForResult = NxFramePartialFailure
        Case Else: NxRaiseContractError "AI apply result outcome is invalid"
    End Select
End Function

Private Function ActionSummaryFor(ByVal command As CNxAiFeatureCommand) As String
    ActionSummaryFor = "프롬프트를 로컬 Windows 클립보드에 복사합니다."
End Function

Private Function ResultSummaryFor(ByVal command As CNxAiFeatureCommand) As String
    ResultSummaryFor = "로컬 Windows 클립보드"
End Function

Private Function RecoverySummaryFor(ByVal command As CNxAiFeatureCommand) As String
    RecoverySummaryFor = "복사 실패 시 선택 범위와 원본 파일은 변경되지 않습니다."
End Function

Private Function CreateAiFrameTextCatalog() As CNxFrameTextCatalog
    Dim catalog As New CNxFrameTextCatalog
    catalog.Configure "성공", "취소", "입력 오류", "환경 오류", "부분 실패", _
        "없음", "완료하지 못한 작업을 확인하세요.", "원본이 변경됨", _
        "원본이 변경되지 않음", "복구 필요 없음", "결과를 확인하세요."
    Set CreateAiFrameTextCatalog = catalog
End Function
