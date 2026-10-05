Option Explicit

Private mInit As CNxFormInitializationGuard
Private mContext As CNxFeatureDialogContext
Private mRangeSession As CNxRangeSelectionSession
Private mPreviewReady As Boolean
Private mPreviewContext As CNxExecutionContext

Private Sub UserForm_Initialize()
    Set mInit = New CNxFormInitializationGuard
    mInit.BeginInitialization
    NxInitializeStaticComboChoices Me
    ApplyDefaults
    txtPreview.ScrollBars = fmScrollBarsVertical
    txtPresetDetail.ScrollBars = fmScrollBarsVertical
    txtPreview.TabStop = True
    txtPreview.BackColor = vbWhite
    mInit.CompleteInitialization
End Sub
Private Sub ApplyDefaults()
    cmdExecute.Default = True
    cmdCancel.Cancel = True
    cmdExecute.Enabled = True
    cboOutputMode.ListIndex = 0
    cboPreset.ListIndex = 0
    cboDateFormat.ListIndex = 0
    UpdateGuidance
End Sub

Public Sub BindFeatureContext(ByVal context As CNxFeatureDialogContext)
    If context Is Nothing Or Not context.IsConfigured Then NxRaiseContractError "Data normalization context is required"
    If Not mContext Is Nothing Then NxRaiseContractError "Data normalization context can be bound only once"
    If context.FeatureId <> "NX-DATA-NORMALIZE" Then NxRaiseContractError "Data normalization feature identity mismatch"
    Set mRangeSession = NxRangeSessionFromSelection()
    Set mContext = context
    txtRange.Value = mRangeSession.DisplayAddress
    UpdateGuidance
End Sub

Private Sub txtRange_Change()
    If mInit Is Nothing Then Exit Sub
    If mInit.IsInitializing Then Exit Sub
    InvalidatePreview
End Sub

Private Sub cmdPickRange_Click()
    If mInit.IsInitializing Then Exit Sub
    On Error GoTo Failed
    RequireContext
    If NxPickRange(Me, mRangeSession) Then txtRange.Value = mRangeSession.DisplayAddress
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cboPreset_Change()
    If mInit.IsInitializing Then Exit Sub
    UpdateGuidance
    InvalidatePreview
End Sub

Private Sub cboDateFormat_Change()
    If mInit.IsInitializing Then Exit Sub
    UpdateGuidance
    InvalidatePreview
End Sub

Private Sub cboOutputMode_Change()
    If mInit.IsInitializing Then Exit Sub
    UpdateGuidance
    InvalidatePreview
End Sub

Private Sub InvalidatePreview()
    If mInit Is Nothing Then Exit Sub
    If mInit.IsInitializing Then Exit Sub
    mPreviewReady = False
    Set mPreviewContext = Nothing
    txtPreview.Value = "설정이 변경되었습니다. 필요하면 미리보기를 갱신하세요. 바로 실행할 수도 있습니다."
    cmdExecute.Enabled = True
End Sub

Private Sub cmdPreview_Click()
    Dim source As Range
    If mInit.IsInitializing Then Exit Sub
    On Error GoTo Failed
    RequireContext
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set source = mRangeSession.ResolveRange()
    Set source = NxDataContentRange(source)
    source.Worksheet.Activate: source.Select
    Set mPreviewContext = NxContextFactory.CaptureCurrent()
    txtPreview.Value = NxDataNormalizationPreview(source, CStr(cboPreset.Value), _
        CStr(cboDateFormat.Value), CStr(cboOutputMode.Value))
    If Not mPreviewContext.SourceStateMatchesCurrent Then NxRaiseContractError "미리보기 중 원본이 변경되었습니다. 다시 확인하세요."
    mPreviewReady = True
    cmdExecute.Enabled = True
    Exit Sub
Failed:
    mPreviewReady = False
    cmdExecute.Enabled = True
    ShowError
End Sub

Private Sub cmdExecute_Click()
    If mInit.IsInitializing Then Exit Sub
    Dim result As CNxResult
    Dim source As Range
    On Error GoTo Failed
    RequireContext
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set source = mRangeSession.ResolveRange()
    source.Worksheet.Activate: source.Select
    Set result = NxDataRunNormalization(source, CStr(cboPreset.Value), _
        CStr(cboDateFormat.Value), CStr(cboOutputMode.Value))
    If result.Outcome = NxSuccess Then
        Unload Me
    ElseIf result.Outcome <> NxCancelled Then
        NxRaiseContractError result.Recovery
    End If
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdCancel_Click()
    If mInit.IsInitializing Then Exit Sub
    Unload Me
End Sub

Private Sub cmdResetDefaults_Click()
    If mInit.IsInitializing Then Exit Sub
    cboOutputMode.ListIndex = 0
    If cboPreset.Enabled Then
        cboPreset.ListIndex = 0
        cboDateFormat.ListIndex = 0
    Else
        cboPreset.Value = "날짜 정리"
        cboDateFormat.Value = "yyyy-mm-dd"
    End If
    txtPreview.Value = "필요하면 미리보기로 결과를 확인할 수 있습니다."
    UpdateGuidance
    InvalidatePreview
End Sub

Private Sub UpdateGuidance()
    Dim outputMode As String
    outputMode = NxDataNormalizeOutputMode(CStr(cboOutputMode.Value))
    txtPresetDetail.Value = NxDataNormalizationDetail(CStr(cboPreset.Value), _
        CStr(cboDateFormat.Value), outputMode)
    If outputMode = "원본 변경" Then
        lblContext.Caption = "선택 범위를 직접 바꿉니다. 실패하면 이전 값과 표시 형식을 복원합니다."
        lblWorkflow.Caption = "정리 방식과 결과 위치를 선택하고 실행하세요."
    Else
        lblContext.Caption = "원본을 유지하고 " & outputMode & "에 결과를 만듭니다."
        lblWorkflow.Caption = "빈 셀과 변환 불가 값은 빈칸으로 둡니다. 미리보기는 선택 사항입니다."
    End If
End Sub

Public Sub UseDateConversion()
    Me.Caption = "내엑셀 - 날짜/생년월일 변환"
    cboPreset.Value = "날짜 정리"
    cboPreset.Enabled = False
    cboDateFormat.Value = "yyyy-mm-dd"
    UpdateGuidance
End Sub

Private Sub RequireContext()
    mInit.RequireReady
    If mContext Is Nothing Then NxRaiseContractError "Data normalization context must be bound before use"
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 데이터 정규화"
End Sub

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
