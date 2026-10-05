VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxAiInputSession"
Option Explicit

Private mTargetSummary As String
Private mScaleSummary As String
Private mPrivacy As CNxAiPrivacySnapshot
Private mFeatureId As String
Private mPurpose As String
Private mOutputFormat As String
Private mInstruction As String
Private mTask As String
Private mOfficeVersion As String
Private mTargetAI As String
Private mPrivacyDecision As String
Private mInputWasMasked As Boolean
Private mOriginalConfirmed As Boolean
Private mPreviewGeneration As Long
Private mLastCommand As CNxAiFeatureCommand
Private mConfigured As Boolean
Private mIsCancelled As Boolean

Public Sub ConfigureSelection(ByVal serializedInput As String, ByVal targetSummary As String, _
    ByVal scaleSummary As String, ByVal privacy As CNxAiPrivacySnapshot)

    If mConfigured Then NxRaiseContractError "AI input session can be configured only once"
    If Len(Trim$(serializedInput)) = 0 Or Len(Trim$(targetSummary)) = 0 Or Len(Trim$(scaleSummary)) = 0 Then NxRaiseContractError "AI input session requires selection summaries"
    If privacy Is Nothing Or Not privacy.IsSealed Then NxRaiseContractError "AI input session requires a sealed privacy snapshot"
    mTargetSummary = targetSummary
    mScaleSummary = scaleSummary
    Set mPrivacy = privacy
    If privacy.State = NxPrivacySafe Then mPrivacyDecision = "SAFE"
    mConfigured = True
    mPreviewGeneration = 1
    mOfficeVersion = "Office 2024 이하"
    mTargetAI = "ChatGPT"
End Sub

Public Sub ReplaceSelection(ByVal serializedInput As String, ByVal targetSummary As String, _
    ByVal scaleSummary As String, ByVal privacy As CNxAiPrivacySnapshot)
    EnsureActive
    If Len(Trim$(serializedInput)) = 0 Or Len(Trim$(targetSummary)) = 0 Or Len(Trim$(scaleSummary)) = 0 Then NxRaiseContractError "AI input session requires selection summaries"
    If privacy Is Nothing Then NxRaiseContractError "AI input session requires a sealed privacy snapshot"
    If Not privacy.IsSealed Then NxRaiseContractError "AI input session requires a sealed privacy snapshot"
    If StrComp(serializedInput, privacy.OriginalCopy, vbBinaryCompare) <> 0 Then NxRaiseContractError "AI input and privacy snapshot differ"
    mTargetSummary = targetSummary
    mScaleSummary = scaleSummary
    Set mPrivacy = privacy
    mPrivacyDecision = vbNullString
    If privacy.State = NxPrivacySafe Then mPrivacyDecision = "SAFE"
    mInputWasMasked = False
    mOriginalConfirmed = False
    InvalidatePreview
End Sub

Public Sub SetMode(ByVal featureId As String)
    Dim choices As Variant
    EnsureActive
    If Not IsAiFeature(featureId) Then NxRaiseContractError "AI mode is outside the closed contract"
    If mFeatureId <> featureId Then
        choices = NxAiTaskChoices(featureId)
        mTask = CStr(choices(0))
    End If
    SetStringValue mFeatureId, featureId
End Sub

Public Sub SetTask(ByVal value As String)
    EnsureActive
    If Not NxAiTaskIsValid(mFeatureId, value) Then NxRaiseContractError "현재 모드의 세부 작업을 선택하세요."
    SetStringValue mTask, value
End Sub

Public Sub SetOfficeVersion(ByVal value As String)
    EnsureActive
    If value <> "Office 2024 이하" And value <> "Microsoft 365" Then NxRaiseContractError "지원 Excel 버전을 선택하세요."
    SetStringValue mOfficeVersion, value
End Sub

Public Sub SetTargetAI(ByVal value As String)
    Dim targetUrl As String
    EnsureActive
    targetUrl = NxAiTargetUrl(value)
    SetStringValue mTargetAI, value
End Sub

Public Property Get TaskName() As String: TaskName = mTask: End Property
Public Property Get OfficeVersion() As String: OfficeVersion = mOfficeVersion: End Property
Public Property Get TargetAI() As String: TargetAI = mTargetAI: End Property

Public Property Get InputPreview() As String
    EnsureActive
    If mPrivacy.State = NxPrivacyIncomplete Then
        InputPreview = "이미지 내용 검사 미완료: 실제 이미지는 복사하지 않습니다. 검사 가능한 셀 범위를 다시 선택하세요."
    ElseIf mPrivacy.State = NxPrivacySafe Or mPrivacyDecision = "ORIGINAL" Then
        InputPreview = mPrivacy.OriginalCopy
    Else
        InputPreview = mPrivacy.MaskedCopy
    End If
End Property

Public Sub SetPurpose(ByVal value As String)
    EnsureActive
    SetStringValue mPurpose, value
End Sub

Public Sub SetOutputFormat(ByVal value As String)
    EnsureActive
    If Not IsOutputFormat(value) Then NxRaiseContractError "AI output format is outside the closed contract"
    SetStringValue mOutputFormat, value
End Sub

Public Sub SetInstruction(ByVal value As String)
    EnsureActive
    SetStringValue mInstruction, value
End Sub

Public Sub ChooseMaskedCopy()
    EnsureActive
    If mPrivacy.State = NxPrivacySafe Or mPrivacy.State = NxPrivacyIncomplete Then NxRaiseContractError "This privacy state cannot choose a masked copy"
    If mPrivacyDecision <> "MASK" Or Not mInputWasMasked Or mOriginalConfirmed Then
        mPrivacyDecision = "MASK"
        mInputWasMasked = True
        mOriginalConfirmed = False
        InvalidatePreview
    End If
End Sub

Public Sub ChooseOriginalCopy(ByVal secondConfirmation As Boolean)
    EnsureActive
    If mPrivacy.State = NxPrivacyBlocked Then NxRaiseContractError "차단된 개인정보 후보는 원문을 사용할 수 없습니다. 마스킹 복사본을 사용하세요."
    If mPrivacy.State = NxPrivacySafe Or mPrivacy.State = NxPrivacyIncomplete Then NxRaiseContractError "This privacy state cannot choose the original copy"
    If Not secondConfirmation Then NxRaiseContractError "원문 사용은 두 번째 확인이 필요합니다."
    If mPrivacyDecision <> "ORIGINAL" Or mInputWasMasked Or Not mOriginalConfirmed Then
        mPrivacyDecision = "ORIGINAL"
        mInputWasMasked = False
        mOriginalConfirmed = True
        InvalidatePreview
    End If
End Sub

Public Function PreviewPrompt() As String
    EnsureActive
    ValidateReady
    PreviewPrompt = BuildPromptText(SelectedInputCopy)
End Function

Public Function CreateCommand() As CNxAiFeatureCommand
    EnsureActive
    ValidateReady
    If Not mLastCommand Is Nothing Then Set CreateCommand = mLastCommand: Exit Function
    Dim command As New CNxAiFeatureCommand
    command.Configure mFeatureId, BuildPromptText(SelectedInputCopy)
    Set mLastCommand = command
    Set CreateCommand = command
End Function

Public Property Get PreviewGeneration() As Long: PreviewGeneration = mPreviewGeneration: End Property
Public Property Get IsCancelled() As Boolean: IsCancelled = mIsCancelled: End Property
Public Property Get FeatureId() As String: FeatureId = mFeatureId: End Property
Public Property Get TargetSummary() As String: TargetSummary = mTargetSummary: End Property
Public Property Get ScaleSummary() As String: ScaleSummary = mScaleSummary: End Property
Public Property Get PrivacyState() As NxPrivacyState: EnsureConfigured: PrivacyState = mPrivacy.State: End Property
Public Property Get PrivacySummary() As String: EnsureConfigured: PrivacySummary = mPrivacy.Summary: End Property
Public Property Get DeliveryMode() As String: EnsureConfigured: DeliveryMode = "CLIPBOARD_ONLY": End Property
Public Property Get ExecutionPrivacyState() As NxPrivacyState
    EnsureConfigured
    If mPrivacy.State = NxPrivacySafe Or mPrivacy.State = NxPrivacyIncomplete Then
        ExecutionPrivacyState = mPrivacy.State
    ElseIf mPrivacyDecision = "MASK" Or mPrivacyDecision = "ORIGINAL" Then
        ExecutionPrivacyState = NxPrivacyCaution
    Else
        ExecutionPrivacyState = mPrivacy.State
    End If
End Property

Public Sub Cancel()
    EnsureConfigured
    If Not mIsCancelled Then
        mIsCancelled = True
        InvalidatePreview
    End If
End Sub

Private Sub ValidateReady()
    If mPrivacy.State = NxPrivacyIncomplete Then NxRaiseContractError "개인정보 검사 미완료 상태에서는 실행계획을 만들 수 없습니다."
    If Len(mFeatureId) = 0 Then NxRaiseContractError "AI 작업 모드를 선택해야 합니다."
    If Len(Trim$(mPurpose)) = 0 Then NxRaiseContractError "AI 작업 목적을 입력해야 합니다."
    If Len(mOutputFormat) = 0 Then NxRaiseContractError "AI 출력 형식을 선택해야 합니다."
    If mOutputFormat = "사용자 지정" And Len(Trim$(mInstruction)) = 0 Then NxRaiseContractError "사용자 지정 출력에는 추가 지시가 필요합니다."
    If mPrivacy.State <> NxPrivacySafe And mPrivacyDecision <> "MASK" And mPrivacyDecision <> "ORIGINAL" Then NxRaiseContractError "개인정보 후보의 복사본을 결정해야 합니다."
    If mPrivacy.State = NxPrivacyBlocked And mPrivacyDecision <> "MASK" Then NxRaiseContractError "차단된 개인정보 후보는 마스킹 복사본만 사용할 수 있습니다."
End Sub

Private Function SelectedInputCopy() As String
    If mPrivacyDecision = "MASK" Then SelectedInputCopy = mPrivacy.MaskedCopy Else SelectedInputCopy = mPrivacy.OriginalCopy
End Function

Private Function BuildPromptText(ByVal selectedCopy As String) As String
    Dim purpose As String, instruction As String
    purpose = Trim$(mPurpose) & vbLf & "세부 작업: " & mTask
    instruction = "지원 Excel: " & mOfficeVersion & ". 해당 버전에서 사용할 수 있는 수식과 대안을 구분합니다."
    If Len(Trim$(mInstruction)) > 0 Then instruction = instruction & vbLf & Trim$(mInstruction)
    BuildPromptText = NxAiBuildPrompt(mFeatureId, selectedCopy, purpose, mOutputFormat, instruction, _
        mPrivacy.State, mPrivacyDecision, mInputWasMasked, mOriginalConfirmed)
End Function

Private Sub SetStringValue(ByRef target As String, ByVal value As String)
    If StrComp(target, value, vbBinaryCompare) <> 0 Then
        target = value
        InvalidatePreview
    End If
End Sub

Private Sub InvalidatePreview()
    If mPreviewGeneration = 2147483647 Then NxRaiseContractError "AI preview generation limit reached"
    mPreviewGeneration = mPreviewGeneration + 1
    Set mLastCommand = Nothing
End Sub

Private Function IsAiFeature(ByVal featureId As String) As Boolean
    Select Case featureId
        Case "NX-AI-SUMMARY", "NX-AI-CLEAN", "NX-AI-FORMULA", "NX-AI-WRITE", "NX-AI-IMAGE"
            IsAiFeature = True
    End Select
End Function

Private Function IsOutputFormat(ByVal value As String) As Boolean
    Select Case value
        Case "간결한 요약", "표 또는 목록", "단계별 설명", "업무 문안", "사용자 지정"
            IsOutputFormat = True
    End Select
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "AI input session is not configured"
End Sub

Private Sub EnsureActive()
    EnsureConfigured
    If mIsCancelled Then NxRaiseContractError "Cancelled AI input cannot continue"
End Sub
