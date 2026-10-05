Option Explicit

Private mSource As Range
Private mPreviewReady As Boolean
Private mAllowClose As Boolean
Private mResolvedGrade As NxExecutionGrade

Public Sub BindSelection(ByVal source As Range, Optional ByVal featureId As String = vbNullString)
    If source Is Nothing Then NxRaiseContractError "Data form selection binding is invalid"
    If Not mSource Is Nothing Then NxRaiseContractError "Data form selection binding is invalid"
    If source.Areas.Count <> 1 Then NxRaiseContractError "Data form requires one contiguous range"
    Set mSource = source
    txtRangeSummary.Value = source.Worksheet.Name & "!" & source.Address(True, True, xlA1, False) & _
        " · " & CStr(source.Rows.Count) & "행 × " & CStr(source.Columns.Count) & "열"
    RefreshKeyColumns
    If Len(featureId) > 0 Then cboFeature.ListIndex = FeatureIndexForId(featureId)
    InvalidatePreview
End Sub

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
    cmdExecute.Default = True
    cmdCancel.Cancel = True
    mResolvedGrade = NxExecutionFast
    UpdateWorkflowState False
End Sub

Public Sub RequestCancel()
    mAllowClose = True
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    Dim featureId As String
    On Error GoTo Failed
    mSource.Parent.Activate
    mSource.Select
    featureId = SelectedFeatureId()
    Select Case featureId
        Case "NX-DATA-UNIQUE-COUNT", "NX-DATA-DUPLICATE-LIST"
            txtPreview.Value = NxDataPreviewText(featureId, mSource, HasTitleRow(), SelectedKeyColumn(), _
                TrimText(), CaseSensitive(), IncludeHidden())
            mPreviewReady = True
    End Select
    ResolveWorkflowGrade
    cmdExecute.Enabled = mPreviewReady
    If mPreviewReady Then UpdateWorkflowState True
    Exit Sub
Failed:
    mPreviewReady = False
    UpdateWorkflowState False
    ShowError
End Sub

Private Sub cmdExecute_Click()
    Dim result As CNxResult
    Dim featureId As String
    On Error GoTo Failed
    cmdPreview_Click
    If Not mPreviewReady Then Exit Sub
    featureId = SelectedFeatureId()
    Select Case featureId
        Case "NX-DATA-UNIQUE-COUNT", "NX-DATA-DUPLICATE-LIST"
            Set result = NxDataRun(featureId, mSource, SelectedKeyColumn(), TrimText(), CaseSensitive(), IncludeHidden(), HasTitleRow())
    End Select
    If result.Outcome = NxSuccess Then
        RequestCancel
    ElseIf result.Outcome <> NxCancelled Then
        NxRaiseContractError result.Recovery
    End If
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdCancel_Click()
    RequestCancel
End Sub

Private Sub cboFeature_Change(): InvalidatePreview: End Sub
Private Sub cboHeader_Change(): RefreshKeyColumns: InvalidatePreview: End Sub
Private Sub cboKeyColumn_Change(): InvalidatePreview: End Sub
Private Sub cboTrim_Change(): InvalidatePreview: End Sub
Private Sub cboCase_Change(): InvalidatePreview: End Sub
Private Sub cboHidden_Change(): InvalidatePreview: End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    If CloseMode <> 0 Then Exit Sub
    mAllowClose = True
    Cancel = False
End Sub

Private Sub RefreshKeyColumns()
    Dim columnIndex As Long
    Dim label As String
    If mSource Is Nothing Then Exit Sub
    cboKeyColumn.Clear
    For columnIndex = 1 To mSource.Columns.Count
        If HasTitleRow() Then label = CStr(mSource.Cells(1, columnIndex).Value) Else label = "열 " & CStr(columnIndex)
        If Len(Trim$(label)) = 0 Then label = "열 " & CStr(columnIndex)
        cboKeyColumn.AddItem CStr(columnIndex) & " · " & label
    Next columnIndex
    cboKeyColumn.ListIndex = 0
End Sub

Private Sub InvalidatePreview()
    mPreviewReady = False
    cmdExecute.Enabled = True
    UpdateWorkflowState False
    If Not mSource Is Nothing Then txtPreview.Value = "옵션을 확인하고 미리보기를 다시 실행하세요."
End Sub

Private Sub UpdateWorkflowState(ByVal previewed As Boolean)
    If mSource Is Nothing Then
        lblWorkflow.Caption = "워크플로: 선택 범위 바인딩 대기"
    ElseIf previewed Then
        lblWorkflow.Caption = "워크플로: 입력 → 미리보기 완료 → 실행 가능"
    Else
        lblWorkflow.Caption = "워크플로: 입력 → 미리보기 선택 가능"
    End If
    lblExecutionGrade.Caption = "실행 등급: " & GradeCaption(mResolvedGrade)
    Select Case mResolvedGrade
        Case NxExecutionFast: cmdExecute.Caption = "바로 실행"
        Case NxExecutionGuarded: cmdExecute.Caption = "확인 후 실행"
        Case NxExecutionPlanned: cmdExecute.Caption = "실행계획 확인"
    End Select
End Sub

Private Function GradeCaption(ByVal grade As NxExecutionGrade) As String
    Select Case grade
        Case NxExecutionFast: GradeCaption = "[빠름]"
        Case NxExecutionGuarded: GradeCaption = "[확인]"
        Case NxExecutionPlanned: GradeCaption = "[계획]"
        Case Else: NxRaiseContractError "실행 등급이 유효하지 않습니다."
    End Select
End Function

Private Sub ResolveWorkflowGrade()
    Dim definition As CNxFeatureDefinition
    Dim decision As CNxExecutionDecision
    Dim featureId As String
    Dim resultSheetCount As Long
    featureId = SelectedFeatureId()
    Set definition = NxDataFeatureDefinition(featureId)
    If featureId = "NX-DATA-DUPLICATE-LIST" Then resultSheetCount = 1
    Set decision = NxExecutionGradePolicy.ResolveDecision(definition, CDbl(mSource.Cells.CountLarge), False, False, False, False, False, resultSheetCount, False, False)
    mResolvedGrade = decision.ResolvedGrade
End Sub

Private Function SelectedFeatureId() As String
    Select Case cboFeature.ListIndex
        Case 0: SelectedFeatureId = "NX-DATA-UNIQUE-COUNT"
        Case 1: SelectedFeatureId = "NX-DATA-DUPLICATE-LIST"
        Case Else: NxRaiseContractError "기능을 선택하세요."
    End Select
End Function

Private Function FeatureIndexForId(ByVal featureId As String) As Long
    Select Case featureId
        Case "NX-DATA-UNIQUE-COUNT": FeatureIndexForId = 0
        Case "NX-DATA-DUPLICATE-LIST": FeatureIndexForId = 1
        Case Else: NxRaiseContractError "Data feature seed is outside the closed contract"
    End Select
End Function

Private Function SelectedKeyColumn() As Long
    If cboKeyColumn.ListIndex < 0 Then NxRaiseContractError "기준 열을 선택하세요."
    SelectedKeyColumn = cboKeyColumn.ListIndex + 1
End Function

Private Function HasTitleRow() As Boolean: HasTitleRow = (cboHeader.ListIndex = 0): End Function
Private Function TrimText() As Boolean: TrimText = (cboTrim.ListIndex = 1): End Function
Private Function CaseSensitive() As Boolean: CaseSensitive = (cboCase.ListIndex = 1): End Function
Private Function IncludeHidden() As Boolean: IncludeHidden = (cboHidden.ListIndex = 1): End Function


Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 데이터/추가기능"
End Sub
