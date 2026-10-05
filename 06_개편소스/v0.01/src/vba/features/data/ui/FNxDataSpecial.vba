Option Explicit

Private mInit As CNxFormInitializationGuard
Private mRangeSession As CNxRangeSelectionSession
Private mFeatureId As String
Private mPreviewContext As CNxExecutionContext
Private mOptions As CNxDataSpecialOptions

Private Sub UserForm_Initialize()
    Set mInit = New CNxFormInitializationGuard
    mInit.BeginInitialization
    NxInitializeStaticComboChoices Me
    cmdExecute.Default = True: cmdCancel.Cancel = True: cmdExecute.Enabled = True
    txtPreview.ScrollBars = fmScrollBarsVertical
    mInit.CompleteInitialization
End Sub

Public Sub BindFeature(ByVal featureId As String)
    mInit.RequireReady
    If Len(mFeatureId) > 0 Then NxRaiseContractError "데이터 기능은 한 번만 연결할 수 있습니다."
    Select Case featureId
        Case NX_FEATURE_DATA_KOREAN_MONEY
            Me.Caption = "내엑셀 - 한글 금액 변환"
            lblGuide.Caption = "정수 금액을 한글로 변환합니다. 음수·소수·인식 실패 값은 빈칸으로 둡니다."
        Case NX_FEATURE_DATA_PRIVACY_MASK
            Me.Caption = "내엑셀 - 개인정보 마스킹"
            lblGuide.Caption = "종류별로 값을 가립니다. 인식 실패는 빈칸으로 둡니다. 개인정보 유출 점검과 별도 기능입니다."
        Case Else: NxRaiseContractError "지원하지 않는 데이터 기능입니다."
    End Select
    mFeatureId = featureId
    Set mRangeSession = NxRangeSessionFromSelection()
    txtRange.Value = mRangeSession.DisplayAddress
    chkPrefix.Enabled = (featureId = NX_FEATURE_DATA_KOREAN_MONEY)
    chkSuffix.Enabled = chkPrefix.Enabled: chkMixed.Enabled = chkPrefix.Enabled: cboZeroMode.Enabled = chkPrefix.Enabled
    cboMaskType.Enabled = (featureId = NX_FEATURE_DATA_PRIVACY_MASK): cboMaskChar.Enabled = cboMaskType.Enabled
    Dim name As Variant
    For Each name In Array("lblMoney", "chkPrefix", "chkSuffix", "chkMixed", "lblZero", "cboZeroMode")
        Me.Controls(CStr(name)).Visible = (featureId = NX_FEATURE_DATA_KOREAN_MONEY)
    Next name
    For Each name In Array("lblMask", "cboMaskType", "cboMaskChar")
        Me.Controls(CStr(name)).Visible = (featureId = NX_FEATURE_DATA_PRIVACY_MASK)
    Next name
    If featureId = NX_FEATURE_DATA_PRIVACY_MASK Then
        lblMask.Top = 144: cboMaskType.Top = 138: cboMaskChar.Top = 138
    End If
    InvalidatePreview
End Sub

Private Function ReadOptions() As CNxDataSpecialOptions
    Dim options As New CNxDataSpecialOptions
    options.Configure CStr(cboOutputMode.Value), CBool(chkHeader.Value), vbNullString, _
        "숫자", CBool(chkPrefix.Value), CBool(chkSuffix.Value), CBool(chkMixed.Value), _
        CStr(cboZeroMode.Value), CStr(cboMaskType.Value), CStr(cboMaskChar.Value)
    Set ReadOptions = options
End Function

Private Sub cmdPickRange_Click()
    If Not Ready Then Exit Sub
    On Error GoTo Failed
    If NxPickRange(Me, mRangeSession) Then txtRange.Value = mRangeSession.DisplayAddress
    InvalidatePreview
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdPreview_Click()
    Dim source As Range
    If Not Ready Then Exit Sub
    On Error GoTo Failed
    InvalidatePreview
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set source = mRangeSession.ResolveRange()
    Set source = NxDataContentRange(source)
    source.Worksheet.Activate: source.Select
    Set mOptions = ReadOptions()
    Set mPreviewContext = NxContextFactory.CaptureCurrent()
    txtPreview.Value = NxDataSpecialPreview(mFeatureId, source, mOptions)
    If Not mPreviewContext.SourceStateMatchesCurrent Then NxRaiseContractError "미리보기 중 원본이 변경되었습니다."
    cmdExecute.Enabled = True
    Exit Sub
Failed:
    InvalidatePreview
    ShowError
End Sub

Private Sub cmdExecute_Click()
    Dim result As CNxResult, source As Range
    If Not Ready Then Exit Sub
    On Error GoTo Failed
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set source = mRangeSession.ResolveRange()
    Set mOptions = ReadOptions()
    Set result = NxDataSpecialRunOptions(mFeatureId, source, mOptions)
    If result.Outcome = NxSuccess Then
        Unload Me
    ElseIf result.Outcome <> NxCancelled Then
        NxRaiseContractError result.Recovery
    End If
    Exit Sub
Failed:
    InvalidatePreview
    ShowError
End Sub

Private Function Ready() As Boolean
    If mInit Is Nothing Then Exit Function
    If mInit.IsInitializing Then Exit Function
    Ready = True
End Function

Private Sub InvalidatePreview()
    If Not Ready Then Exit Sub
    Set mPreviewContext = Nothing: Set mOptions = Nothing
    cmdExecute.Enabled = True
End Sub

Private Sub txtRange_Change(): InvalidatePreview: End Sub
Private Sub cboOutputMode_Change(): InvalidatePreview: End Sub
Private Sub cboZeroMode_Change(): InvalidatePreview: End Sub
Private Sub cboMaskType_Change(): InvalidatePreview: End Sub
Private Sub cboMaskChar_Change(): InvalidatePreview: End Sub
Private Sub chkHeader_Click(): InvalidatePreview: End Sub
Private Sub chkPrefix_Click(): InvalidatePreview: End Sub
Private Sub chkSuffix_Click(): InvalidatePreview: End Sub
Private Sub chkMixed_Click(): InvalidatePreview: End Sub
Private Sub cmdCancel_Click()
    If Not Ready Then Exit Sub
    Unload Me
End Sub
Private Sub ShowError()
    Dim detail As String
    detail = Err.Description: Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, Me.Caption
End Sub

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
