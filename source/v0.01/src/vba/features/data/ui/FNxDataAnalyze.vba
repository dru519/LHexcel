Option Explicit

Private mInit As CNxFormInitializationGuard
Private mSource As Range
Private mRangeSession As CNxRangeSelectionSession
Private mFeatureId As String
Private mPreviewRequest As CNxDataRequest
Private mPreviewReady As Boolean
Private mAllowClose As Boolean
Private mUpdating As Boolean

Private Sub UserForm_Initialize()
    Set mInit = New CNxFormInitializationGuard
    mInit.BeginInitialization
    NxInitializeStaticComboChoices Me
    cmdExecute.Default = True
    cmdCancel.Cancel = True
    cmdExecute.Enabled = True
    txtPreview.ScrollBars = fmScrollBarsVertical
    mInit.CompleteInitialization
End Sub

Public Sub BindSelection(ByVal source As Range, ByVal featureId As String)
    If source Is Nothing Then NxRaiseContractError "데이터 범위를 확인하세요."
    If Not mSource Is Nothing Then NxRaiseContractError "분석 범위는 한 번만 연결할 수 있습니다."
    If source.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 선택하세요."
    If featureId <> "NX-DATA-UNIQUE-COUNT" And featureId <> "NX-DATA-DUPLICATE-LIST" Then NxRaiseContractError "지원하지 않는 데이터 분석입니다."
    mUpdating = True
    Set mSource = source
    Set mRangeSession = New CNxRangeSelectionSession
    mRangeSession.UpdateFromRange source
    mFeatureId = featureId
    txtRange.Value = mRangeSession.DisplayAddress
    If source.Rows.Count = 1 And source.Columns.Count > 1 Then cboDirection.ListIndex = 1 Else cboDirection.ListIndex = 0
    chkHeader.Value = (featureId = "NX-DATA-DUPLICATE-LIST")
    chkTrim.Value = (featureId = "NX-DATA-UNIQUE-COUNT")
    ApplyMode
    RefreshKeyColumns
    mUpdating = False
    InvalidatePreview
End Sub

Private Sub ApplyMode()
    Dim uniqueMode As Boolean, rowDirection As Boolean
    uniqueMode = (mFeatureId = "NX-DATA-UNIQUE-COUNT")
    rowDirection = (cboDirection.ListIndex = 1)
    If uniqueMode Then
        Me.Caption = "내엑셀 - 고유값 개수"
        lblTitle.Caption = "고유값 개수 · 값 또는 행/열 조합별 빈도 집계"
        txtGuide.Value = "집계한 값과 건수를 표로 만듭니다. 대소문자는 구분합니다. 빈 행/열 제외와 공백 정리, 비율 및 정렬을 선택하세요."
    Else
        Me.Caption = "내엑셀 - 중복 목록"
        lblTitle.Caption = "중복 목록 · 기준 목록과 대조 목록 비교"
        txtGuide.Value = "선택 범위 안의 두 열/행을 서로 대조합니다. 상대 목록에 있으면 '중복', 없으면 '고유'입니다. 열 기준은 원본 두 열과 결과 두 열, 행 기준은 결과 두 행을 만듭니다."
    End If
    chkRatio.Visible = uniqueMode
    If rowDirection Then
        chkHeader.Caption = "첫 열을 제목으로 사용": chkBlank.Caption = "빈 열 제외"
        lblLeft.Caption = "기준 행": lblRight.Caption = "대조 행"
    Else
        chkHeader.Caption = "첫 행을 제목으로 사용": chkBlank.Caption = "빈 행 제외"
        lblLeft.Caption = "기준 열": lblRight.Caption = "대조 열"
    End If
    If uniqueMode Then
        lblLeft.Caption = "집계 기준": lblRight.Caption = "정렬"
        chkTrim.Caption = "공백 정리 후 집계"
    Else
        chkBlank.Caption = "대소문자 무시"
        chkTrim.Caption = "공백·하이픈 제거"
    End If
End Sub

Private Sub RefreshKeyColumns()
    Dim index As Long, count As Long, label As String, rowDirection As Boolean
    If mSource Is Nothing Then Exit Sub
    rowDirection = (cboDirection.ListIndex = 1)
    If rowDirection Then count = mSource.Rows.Count Else count = mSource.Columns.Count
    cboKeyColumn.Clear: cboCompareColumn.Clear
    If mFeatureId = "NX-DATA-UNIQUE-COUNT" Then
        If rowDirection Then
            cboKeyColumn.AddItem "선택 범위 전체 열": cboKeyColumn.AddItem "첫 행 기준"
        Else
            cboKeyColumn.AddItem "선택 범위 전체 행": cboKeyColumn.AddItem "첫 열 기준"
        End If
        cboCompareColumn.AddItem "값 오름차순": cboCompareColumn.AddItem "값 내림차순": cboCompareColumn.AddItem "원본 순서"
        cboKeyColumn.ListIndex = 0: cboCompareColumn.ListIndex = 0
        Exit Sub
    End If
    For index = 1 To count
        If rowDirection Then
            label = CStr(mSource.Row + index - 1) & "행"
        Else
            label = Replace$(mSource.Cells(1, index).Address(False, False), CStr(mSource.Row), "") & "열"
        End If
        cboKeyColumn.AddItem label: cboCompareColumn.AddItem label
    Next index
    cboKeyColumn.ListIndex = 0
    If count > 1 Then cboCompareColumn.ListIndex = 1 Else cboCompareColumn.ListIndex = 0
End Sub

Private Sub cmdPickRange_Click()
    On Error GoTo Failed
    mInit.RequireReady
    If mRangeSession Is Nothing Then NxRaiseContractError "분석 범위가 연결되지 않았습니다."
    InvalidatePreview
    If NxPickRange(Me, mRangeSession) Then
        Set mSource = mRangeSession.ResolveRange()
        mUpdating = True
        txtRange.Value = mRangeSession.DisplayAddress
        RefreshKeyColumns
        mUpdating = False
    End If
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdPreview_Click()
    Dim request As CNxDataRequest
    On Error GoTo Failed
    mInit.RequireReady
    InvalidatePreview
    Set request = BuildAnalysisRequest()
    txtPreview.Value = NxDataAnalysisPreview(request)
    Set mPreviewRequest = request
    mPreviewReady = True
    cmdExecute.Enabled = True
    lblStatus.Caption = "결과 확인 후 실행"
    Exit Sub
Failed:
    InvalidatePreview
    ShowError
End Sub

Private Function BuildAnalysisRequest() As CNxDataRequest
    Dim request As New CNxDataRequest
    Dim sortMode As String
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set mSource = mRangeSession.ResolveRange()
    ' The user's confirmed range, including edits in the range box, becomes plan authority.
    mSource.Worksheet.Activate
    mSource.Select
    sortMode = "값 오름차순"
    If mFeatureId = "NX-DATA-UNIQUE-COUNT" Then sortMode = CStr(cboCompareColumn.Value)
    request.ConfigureAnalysis mFeatureId, mSource, CBool(chkHeader.Value), (cboDirection.ListIndex = 1), _
        (cboKeyColumn.ListIndex = 0), CBool(chkBlank.Value), CBool(chkTrim.Value), CBool(chkRatio.Value), sortMode, _
        cboKeyColumn.ListIndex + 1, cboCompareColumn.ListIndex + 1, CBool(chkBlank.Value), CBool(chkTrim.Value), _
        CBool(chkHidden.Value), CStr(cboOutputMode.Value)
    Set BuildAnalysisRequest = request
End Function

Private Sub cmdExecute_Click()
    Dim result As CNxResult
    On Error GoTo Failed
    mInit.RequireReady
    InvalidatePreview
    Set mPreviewRequest = BuildAnalysisRequest()
    mPreviewRequest.RequireFreshAnalysis
    Set result = NxDataRunAnalysis(mPreviewRequest)
    If result Is Nothing Then NxRaiseContractError "데이터 작업 결과를 확인하지 못했습니다."
    If result.Outcome = NxSuccess Then
        RequestCancel
    ElseIf result.Outcome <> NxCancelled Then
        NxRaiseContractError result.Recovery
    End If
    Exit Sub
Failed:
    InvalidatePreview
    ShowError
End Sub

Private Sub cboDirection_Change()
    If mInit Is Nothing Then Exit Sub
    If mInit.IsInitializing Then Exit Sub
    If mUpdating Then Exit Sub
    On Error GoTo Failed
    mUpdating = True
    ApplyMode
    RefreshKeyColumns
    mUpdating = False
    InvalidatePreview
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub txtRange_Change(): InvalidatePreview: End Sub
Private Sub txtRange_AfterUpdate()
    On Error GoTo Failed
    If mUpdating Then Exit Sub
    mInit.RequireReady
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set mSource = mRangeSession.ResolveRange()
    mUpdating = True
    RefreshKeyColumns
    mUpdating = False
    InvalidatePreview
    Exit Sub
Failed:
    ShowError
End Sub
Private Sub cboKeyColumn_Change(): InvalidatePreview: End Sub
Private Sub cboCompareColumn_Change(): InvalidatePreview: End Sub
Private Sub cboOutputMode_Change(): InvalidatePreview: End Sub
Private Sub chkHeader_Click(): InvalidatePreview: End Sub
Private Sub chkBlank_Click(): InvalidatePreview: End Sub
Private Sub chkTrim_Click(): InvalidatePreview: End Sub
Private Sub chkRatio_Click(): InvalidatePreview: End Sub
Private Sub chkHidden_Click(): InvalidatePreview: End Sub
Private Sub cmdCancel_Click(): RequestCancel: End Sub

Private Sub InvalidatePreview()
    If mInit Is Nothing Then Exit Sub
    If mInit.IsInitializing Then Exit Sub
    If mUpdating Then Exit Sub
    mPreviewReady = False
    Set mPreviewRequest = Nothing
    cmdExecute.Enabled = True
    lblStatus.Caption = "미리보기 선택 가능"
End Sub

Private Sub RequestCancel()
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    mUpdating = False
    If Not mInit Is Nothing Then
        If mInit.IsInitializing Then mInit.CompleteInitialization
    End If
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 데이터 분석"
End Sub

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
