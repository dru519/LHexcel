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
    If featureId <> NX_FEATURE_DATA_AGE Then NxRaiseContractError "만 나이 기능만 연결할 수 있습니다."
    mFeatureId = featureId
    Me.Caption = "내엑셀 - 만 나이 계산"
    lblGuide.Caption = "생년월일·YYMMDD·주민번호형 입력을 기준일의 만 나이로 계산합니다."
    Set mRangeSession = NxRangeSessionFromSelection()
    txtRange.Value = mRangeSession.DisplayAddress
    txtReferenceDate.Value = Format$(Date, "yyyy-mm-dd")
    InvalidatePreview
End Sub

Private Function ReadOptions() As CNxDataSpecialOptions
    Dim options As New CNxDataSpecialOptions, ageStyle As String
    ageStyle = "숫자"
    If chkAgePrefix.Value Or chkAgeSuffix.Value Then ageStyle = "사용자 지정"
    options.Configure CStr(cboOutputMode.Value), CBool(chkHeader.Value), CStr(txtReferenceDate.Value), _
        ageStyle, False, False, False, "빈칸", "자동", "*", CBool(chkAgePrefix.Value), CBool(chkAgeSuffix.Value)
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
Private Sub txtReferenceDate_Change(): InvalidatePreview: End Sub
Private Sub cboOutputMode_Change(): InvalidatePreview: End Sub
Private Sub chkAgePrefix_Click(): InvalidatePreview: End Sub
Private Sub chkAgeSuffix_Click(): InvalidatePreview: End Sub
Private Sub chkHeader_Click(): InvalidatePreview: End Sub
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
