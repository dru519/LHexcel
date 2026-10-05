Option Explicit

Private mSession As CNxAiInputSession
Private mAllowClose As Boolean
Private mRenderedPreviewGeneration As Long
Private mLoading As Boolean
Private mSuggestedInstruction As String
Private mCurrentPrompt As String
Private mPreviewView As Long

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
    txtInstruction.ScrollBars = fmScrollBarsVertical
    txtInstruction.EnterKeyBehavior = True
    txtPromptPreview.ScrollBars = fmScrollBarsBoth
    lblPrivacy.WordWrap = True
    lblStatus.WordWrap = True
End Sub

Private Sub UserForm_Activate()
    NxFormWheelAttach Me
End Sub

Public Sub HandleWheelDelta(ByVal wheelDelta As Long)
    Dim target As Object
    Dim nextStart As Long
    On Error GoTo Finished
    Set target = Me.ActiveControl
    If TypeName(target) <> "TextBox" Then Exit Sub
    If target.Name <> "txtInstruction" And target.Name <> "txtPromptPreview" Then Exit Sub
    nextStart = CLng(target.SelStart) + IIf(wheelDelta > 0, -240, 240)
    If nextStart < 0 Then nextStart = 0
    If nextStart > Len(CStr(target.Value)) Then nextStart = Len(CStr(target.Value))
    target.SelStart = nextStart
Finished:
    Err.Clear
End Sub

Private Sub UserForm_Terminate()
    NxFormWheelDetach Me
End Sub

Public Sub BindSession(ByVal session As CNxAiInputSession, Optional ByVal featureId As String = vbNullString)
    If session Is Nothing Then NxRaiseContractError "AI form requires an input session"
    If Not mSession Is Nothing Then NxRaiseContractError "AI form session can be bound only once"
    Set mSession = session
    mLoading = True
    If Len(featureId) = 0 Then featureId = "NX-AI-SUMMARY"
    mSession.SetMode featureId
    cboMode.ListIndex = FeatureIndexForId(featureId)
    RefreshTasks featureId
    RefreshTaskInstruction
    mSession.SetTargetAI CStr(cboTargetAI.Value)
    ApplyDefaultMasking
    mLoading = False
    RenderSession
End Sub

Private Sub RefreshTasks(Optional ByVal featureId As String = vbNullString)
    Dim choice As Variant
    If Len(featureId) = 0 Then featureId = FeatureIdForIndex(cboMode.ListIndex)
    cboTask.Clear
    For Each choice In NxAiTaskChoices(featureId)
        cboTask.AddItem CStr(choice)
    Next choice
    cboTask.ListIndex = 0
End Sub

Private Sub RefreshTaskInstruction()
    Dim suggested As String
    suggested = NxAiTaskPurpose(FeatureIdForIndex(cboMode.ListIndex), CStr(cboTask.Value))
    ' Replace the prior suggestion, never a request the user has edited.
    If Len(Trim$(CStr(txtInstruction.Value))) = 0 Or _
        StrComp(CStr(txtInstruction.Value), mSuggestedInstruction, vbBinaryCompare) = 0 Then
        txtInstruction.Value = suggested
    End If
    mSuggestedInstruction = suggested
    mSession.SetPurpose suggested
    mSession.SetInstruction CStr(txtInstruction.Value)
    mSession.SetOutputFormat OutputFormatForMode()
End Sub

Private Sub cmdPickRange_Click()
    Dim selectedRange As Range
    Dim privacy As CNxAiPrivacySnapshot
    Dim serialized As String
    Dim detail As String
    EnsureSession
    NxFormWheelDetach Me
    Me.Hide
    On Error Resume Next
    Set selectedRange = Application.InputBox("AI 요청에 사용할 셀 범위를 선택하세요. 취소하면 이전 입력을 유지합니다.", _
        "내엑셀 - AI 범위 선택", Type:=8)
    Err.Clear
    On Error GoTo Failed
    If Not selectedRange Is Nothing Then
        serialized = NxAiCaptureRangeInput(selectedRange, privacy)
        mSession.ReplaceSelection serialized, NxAiCurrentTargetSummary(), NxAiCurrentScaleSummary(), privacy
        mLoading = True
        ApplyDefaultMasking
        mLoading = False
        InvalidateRenderedPreview
        RenderSession
    End If
    Me.Show
    Exit Sub
Failed:
    detail = Err.Description
    Err.Clear
    mLoading = False
    detail = NxUserErrorText(detail)
    lblStatus.Caption = detail
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - AI 범위 선택"
    Me.Show
End Sub

Public Sub RequestCancel()
    EnsureSession
    mSession.Cancel
    mAllowClose = True
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo PreviewFailed
    EnsureSession
    PushInputs
    mCurrentPrompt = mSession.PreviewPrompt()
    mRenderedPreviewGeneration = mSession.PreviewGeneration
    mPreviewView = 1
    RenderPreviewBox
    cmdExecute.Enabled = True
    cmdExecute.Default = True
    cmdPreview.Default = False
    lblStatus.Caption = "전송할 내용을 확인한 뒤 복사하세요."
    Exit Sub
PreviewFailed:
    Dim detail As String
    detail = Err.Description
    Err.Clear
    InvalidateRenderedPreview
    detail = NxUserErrorText(detail)
    lblStatus.Caption = detail
End Sub

Private Sub cmdTogglePrompt_Click()
    EnsureSession
    If mPreviewView = 1 Then
        mPreviewView = 0
        RenderPreviewBox
    ElseIf Len(mCurrentPrompt) = 0 Then
        cmdPreview_Click
    Else
        mPreviewView = 1
        RenderPreviewBox
    End If
End Sub

Private Sub cmdInputPreview_Click()
    EnsureSession
    If mPreviewView = 2 Then mPreviewView = 0 Else mPreviewView = 2
    RenderPreviewBox
End Sub

Private Sub cmdExecute_Click()
    Dim wasHidden As Boolean
    Dim copyState As NxFrameState
    On Error GoTo ExecuteFailed
    EnsureSession
    PushInputs
    If Len(mCurrentPrompt) = 0 Or mRenderedPreviewGeneration <> mSession.PreviewGeneration Then cmdPreview_Click
    If Len(mCurrentPrompt) = 0 Or mRenderedPreviewGeneration <> mSession.PreviewGeneration Then NxRaiseContractError "프롬프트를 만들지 못했습니다. 입력 내용을 확인하세요."
    cmdExecute.Enabled = False
    NxFormWheelDetach Me
    Me.Hide
    wasHidden = True
    copyState = NxAiShowExecutionPlan(mSession)
    If copyState = NxFrameSuccess Then
        If NxAiOpenTargetAfterCopy(mSession.TargetAI, copyState) Then
            lblStatus.Caption = "복사했습니다. AI에서 붙여넣고 직접 전송하세요."
        Else
            lblStatus.Caption = "복사했습니다. " & mSession.TargetAI & "를 직접 열어 붙여넣으세요."
        End If
        mAllowClose = True
        Unload Me
        Exit Sub
    ElseIf copyState = NxFrameInput Or copyState = NxFrameCancelled Then
        lblStatus.Caption = "복사를 취소했습니다."
    Else
        lblStatus.Caption = "복사하지 못했습니다. 내용을 확인하고 다시 시도하세요."
    End If
    cmdExecute.Enabled = True
    Me.Show
    Exit Sub
ExecuteFailed:
    Dim detail As String
    detail = Err.Description
    Err.Clear
    cmdExecute.Enabled = (mSession.PrivacyState <> NxPrivacyIncomplete)
    detail = NxUserErrorText(detail)
    lblStatus.Caption = detail
    If wasHidden Then Me.Show
End Sub

Private Sub ApplyDefaultMasking()
    chkMask.Value = True
    If mSession.PrivacyState = NxPrivacyCaution Or mSession.PrivacyState = NxPrivacyBlocked Then
        mSession.ChooseMaskedCopy
    End If
End Sub

Private Sub chkMask_Click()
    If mSession Is Nothing Or mLoading Then Exit Sub
    On Error GoTo Failed
    If CBool(chkMask.Value) Then
        mSession.ChooseMaskedCopy
    ElseIf ConfirmOriginalUse() Then
        mSession.ChooseOriginalCopy True
    Else
        mLoading = True
        chkMask.Value = True
        mLoading = False
        Exit Sub
    End If
    InvalidateRenderedPreview
    RenderSession
    Exit Sub
Failed:
    Dim detail As String
    detail = Err.Description
    Err.Clear
    mLoading = True
    chkMask.Value = True
    mLoading = False
    detail = NxUserErrorText(detail)
    lblStatus.Caption = detail
End Sub

Private Sub cmdCancel_Click()
    RequestCancel
End Sub

Private Sub cboMode_Change()
    If mSession Is Nothing Or mLoading Then Exit Sub
    mLoading = True
    mSession.SetMode FeatureIdForIndex(cboMode.ListIndex)
    RefreshTasks
    mSession.SetTask CStr(cboTask.Value)
    RefreshTaskInstruction
    mLoading = False
    InvalidateRenderedPreview
End Sub

Private Sub cboTask_Change()
    If mSession Is Nothing Or mLoading Then Exit Sub
    mLoading = True
    mSession.SetTask CStr(cboTask.Value)
    RefreshTaskInstruction
    mLoading = False
    InvalidateRenderedPreview
End Sub

Private Sub cboOffice_Change()
    If mSession Is Nothing Or mLoading Then Exit Sub
    mSession.SetOfficeVersion CStr(cboOffice.Value)
    InvalidateRenderedPreview
End Sub

Private Sub cboTargetAI_Change()
    If mSession Is Nothing Or mLoading Then Exit Sub
    mSession.SetTargetAI CStr(cboTargetAI.Value)
    InvalidateRenderedPreview
End Sub

Private Sub txtInstruction_Change()
    If mSession Is Nothing Or mLoading Then Exit Sub
    mSession.SetInstruction CStr(txtInstruction.Value)
    InvalidateRenderedPreview
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    NxFormWheelDetach Me
    If mAllowClose Then Exit Sub
    If mSession Is Nothing Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub PushInputs()
    EnsureSession
    mSession.SetMode FeatureIdForIndex(cboMode.ListIndex)
    mSession.SetTask CStr(cboTask.Value)
    mSession.SetOfficeVersion CStr(cboOffice.Value)
    mSession.SetTargetAI CStr(cboTargetAI.Value)
    mSession.SetOutputFormat OutputFormatForMode()
    mSession.SetPurpose NxAiTaskPurpose(mSession.FeatureId, CStr(cboTask.Value))
    mSession.SetInstruction CStr(txtInstruction.Value)
End Sub

Private Function OutputFormatForMode() As String
    Select Case cboMode.ListIndex
        Case 2: OutputFormatForMode = "단계별 설명"
        Case 3: OutputFormatForMode = "업무 문안"
        Case Else: OutputFormatForMode = "표 또는 목록"
    End Select
End Function

Private Sub RenderSession()
    EnsureSession
    txtRange.Value = mSession.TargetSummary & " · " & mSession.ScaleSummary
    lblPrivacy.Caption = mSession.PrivacySummary
    chkMask.Enabled = (mSession.PrivacyState = NxPrivacyCaution)
    cmdPreview.Enabled = (mSession.PrivacyState <> NxPrivacyIncomplete)
    cmdExecute.Enabled = cmdPreview.Enabled
    cmdExecute.Default = cmdExecute.Enabled
    cmdPreview.Default = False
    cmdTogglePrompt.Enabled = cmdPreview.Enabled
    If Not cmdPreview.Enabled Then
        InvalidateRenderedPreview
        lblStatus.Caption = "이미지는 직접 확인하고, 사용할 셀 범위를 다시 선택하세요."
    End If
    RenderPreviewBox
End Sub

Private Sub RenderPreviewBox()
    cmdTogglePrompt.Caption = "전체 프롬프트 보기"
    cmdInputPreview.Caption = "범위 내용 보기"
    Select Case mPreviewView
        Case 1
            lblWorkflow.Caption = "전체 프롬프트"
            txtPromptPreview.Value = mCurrentPrompt
            cmdTogglePrompt.Caption = "전송 요약 보기"
        Case 2
            lblWorkflow.Caption = "선택 범위 내용"
            txtPromptPreview.Value = mSession.InputPreview
            cmdInputPreview.Caption = "전송 요약 보기"
        Case Else
            lblWorkflow.Caption = "전송 요약"
            txtPromptPreview.Value = BuildTransferSummary()
    End Select
End Sub

Private Function BuildTransferSummary() As String
    Dim privacyText As String
    privacyText = mSession.PrivacySummary
    If mSession.PrivacyState = NxPrivacyCaution Or mSession.PrivacyState = NxPrivacyBlocked Then
        If CBool(chkMask.Value) Then privacyText = privacyText & " · 마스킹 적용" Else privacyText = privacyText & " · 원문 사용 확인"
    End If
    BuildTransferSummary = "선택 범위: " & mSession.TargetSummary & vbCrLf & _
        "작업 모드: " & CStr(cboMode.Value) & " / " & mSession.TaskName & vbCrLf & _
        "Office 버전: " & CStr(cboOffice.Value) & vbCrLf & _
        "대상 AI: " & CStr(cboTargetAI.Value) & vbCrLf & _
        "포함 데이터: " & mSession.ScaleSummary & vbCrLf & _
        "개인정보: " & privacyText & vbCrLf & _
        "프롬프트 길이: " & IIf(Len(mCurrentPrompt) > 0, CStr(Len(mCurrentPrompt)) & "자", "실행 시 생성")
    If mSession.FeatureId = "NX-AI-IMAGE" Then _
        BuildTransferSummary = BuildTransferSummary & vbCrLf & "이미지 첨부: 내용 확인 후 AI에서 직접 첨부"
End Function

Private Sub InvalidateRenderedPreview()
    mRenderedPreviewGeneration = 0
    mCurrentPrompt = vbNullString
    mPreviewView = 0
    cmdExecute.Enabled = (mSession.PrivacyState <> NxPrivacyIncomplete)
    cmdExecute.Default = cmdExecute.Enabled
    cmdPreview.Default = False
    RenderPreviewBox
End Sub

Private Function FeatureIdForIndex(ByVal index As Long) As String
    Select Case index
        Case 0: FeatureIdForIndex = "NX-AI-SUMMARY"
        Case 1: FeatureIdForIndex = "NX-AI-CLEAN"
        Case 2: FeatureIdForIndex = "NX-AI-FORMULA"
        Case 3: FeatureIdForIndex = "NX-AI-WRITE"
        Case 4: FeatureIdForIndex = "NX-AI-IMAGE"
        Case Else: NxRaiseContractError "AI mode selection is outside the closed contract"
    End Select
End Function

Private Function FeatureIndexForId(ByVal featureId As String) As Long
    Select Case featureId
        Case "NX-AI-SUMMARY": FeatureIndexForId = 0
        Case "NX-AI-CLEAN": FeatureIndexForId = 1
        Case "NX-AI-FORMULA": FeatureIndexForId = 2
        Case "NX-AI-WRITE": FeatureIndexForId = 3
        Case "NX-AI-IMAGE": FeatureIndexForId = 4
        Case Else: NxRaiseContractError "AI feature seed is outside the closed contract"
    End Select
End Function

Private Function ConfirmOriginalUse() As Boolean
    ConfirmOriginalUse = (MsgBox("원문이 AI용 클립보드에 포함됩니다. 원문 사용을 계속하시겠습니까?", _
        vbYesNo + vbExclamation + vbDefaultButton2, "내엑셀 - 원문 사용 확인") = vbYes)
End Function

Private Sub EnsureSession()
    If mSession Is Nothing Then NxRaiseContractError "AI form session is not bound"
End Sub
