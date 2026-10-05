Option Explicit

Private mSource As Range
Private mRetained As Collection
Private mRecords As Collection
Private mTemplateIds As Collection
Private mAllowClose As Boolean
Private mInitialFeatureId As String
Private mInitialFeatureApplied As Boolean

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
End Sub

Public Sub BindSelection(ByVal source As Range, Optional ByVal featureId As String = vbNullString)
    If source Is Nothing Or Not mSource Is Nothing Then NxRaiseContractError "Template form selection binding is invalid"
    Set mSource = source
    mInitialFeatureId = featureId
    mInitialFeatureApplied = False
    Set mRetained = New Collection
    txtTargetSummary.Value = source.Worksheet.Name & "!" & source.Address(True, True, xlA1, False)
    RenderWorkflow NxExecutionPlanned, "등록은 실행계획 확인 후 진행합니다."
    RefreshTemplates
End Sub

Private Sub UserForm_Activate()
    If mInitialFeatureApplied Then Exit Sub
    mInitialFeatureApplied = True
    Select Case mInitialFeatureId
        Case "NX-TPL-REGISTER-SHEET"
            cboSourceKind.ListIndex = FeatureIndexForId(mInitialFeatureId)
        Case "NX-TPL-LIST"
            cboTemplates.SetFocus
        Case "NX-TPL-LOAD"
            If cmdLoadNewSheet.Enabled Then cmdLoadNewSheet.SetFocus
        Case "NX-TPL-RENAME"
            If cmdRename.Enabled Then cmdRename.SetFocus
        Case "NX-TPL-DELETE"
            If cmdDelete.Enabled Then cmdDelete.SetFocus
        Case vbNullString
        Case Else: NxRaiseContractError "Template feature seed is outside the closed contract"
    End Select
End Sub

Public Sub RequestCancel()
    mAllowClose = True
    Unload Me
End Sub

Private Sub cmdSelectRetained_Click()
    Dim selected As Range
    Me.Hide
    On Error Resume Next
    Set selected = Application.InputBox("저장할 상수 범위를 선택하세요.", "내엑셀 - 상수 보존 범위", Type:=8)
    On Error GoTo 0
    Me.Show
    If selected Is Nothing Then Exit Sub
    If Application.Intersect(mSource, selected) Is Nothing Then NxRaiseContractError "보존 범위는 원본 범위 안에 있어야 합니다."
    Set mRetained = New Collection
    mRetained.Add selected
    txtRetainedSummary.Value = selected.Address(True, True, xlA1, False)
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo Failed
    ApplySourceKind
    txtPreview.Value = NxTemplatePreviewText(mSource, mRetained, cboHiddenPolicy.ListIndex = 1, cboRiskConsent.ListIndex = 1)
    RenderWorkflow NxExecutionPlanned, "등록은 실행계획 확인 후 진행합니다."
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdRegister_Click()
    Dim result As CNxResult
    On Error GoTo Failed
    ApplySourceKind
    Set result = NxTemplateRunRegistrationFromForm(RegistrationFeatureId(), mSource, mRetained, CStr(txtName.Value), _
        CStr(txtDescription.Value), cboHiddenPolicy.ListIndex = 1, cboRiskConsent.ListIndex = 1, cboRiskConsent.ListIndex = 1)
    If result.Outcome = NxSuccess Then RefreshTemplates Else NxRaiseContractError result.Recovery
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdLoadNewSheet_Click()
    Dim result As CNxResult
    On Error GoTo Failed
    Set result = NxTemplateRunFromForm("NX-TPL-LOAD", SelectedTemplateId(), cboRiskConsent.ListIndex = 1)
    If result.Outcome <> NxSuccess Then NxRaiseContractError result.Recovery
    RefreshTemplates
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdRename_Click()
    Dim result As CNxResult
    On Error GoTo Failed
    Set result = NxTemplateRunFromForm("NX-TPL-RENAME", SelectedTemplateId(), True, CStr(txtName.Value), CStr(txtDescription.Value))
    If result.Outcome <> NxSuccess Then NxRaiseContractError result.Recovery
    RefreshTemplates
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdDelete_Click()
    Dim result As CNxResult
    On Error GoTo Failed
    Set result = NxTemplateRunFromForm("NX-TPL-DELETE", SelectedTemplateId(), True)
    If result.Outcome <> NxSuccess Then NxRaiseContractError result.Recovery
    RefreshTemplates
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cboTemplates_Change()
    RenderSelectedTemplate
End Sub

Private Sub cmdCancel_Click()
    RequestCancel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub ApplySourceKind()
    If cboSourceKind.ListIndex <> 0 Then NxRaiseContractError "현재 시트 사용 영역만 템플릿으로 등록할 수 있습니다."
    Set mSource = mSource.Worksheet.UsedRange
    txtTargetSummary.Value = mSource.Worksheet.Name & "!" & mSource.Address(True, True, xlA1, False)
End Sub

Private Function RegistrationFeatureId() As String
    If cboSourceKind.ListIndex = 0 Then
        RegistrationFeatureId = "NX-TPL-REGISTER-SHEET"
    Else
        NxRaiseContractError "원본 종류를 선택하세요."
    End If
End Function

Private Function FeatureIndexForId(ByVal featureId As String) As Long
    Select Case featureId
        Case "NX-TPL-REGISTER-SHEET": FeatureIndexForId = 0
        Case Else: NxRaiseContractError "Template feature seed is outside the closed contract"
    End Select
End Function

Private Sub RefreshTemplates()
    Dim item As Variant, record As CNxTemplateRecord
    Set mRecords = NxTemplateRecordsForForm()
    Set mTemplateIds = New Collection
    cboTemplates.Clear
    For Each item In mRecords
        Set record = item
        cboTemplates.AddItem record.DisplayName
        mTemplateIds.Add record.TemplateId
    Next item
    If mRecords.Count = 0 Then
        cboTemplates.AddItem "등록된 템플릿 없음"
        cboTemplates.ListIndex = 0
        txtTemplateDetails.Value = "등록된 템플릿이 없습니다."
        SetActionButtons False
    Else
        cboTemplates.ListIndex = 0
        SetActionButtons True
        RenderSelectedTemplate
    End If
End Sub

Private Sub RenderSelectedTemplate()
    Dim record As CNxTemplateRecord
    If mRecords Is Nothing Then Exit Sub
    If mRecords.Count = 0 Then Exit Sub
    If cboTemplates.ListIndex < 0 Or cboTemplates.ListIndex >= mRecords.Count Then Exit Sub
    Set record = mRecords.Item(cboTemplates.ListIndex + 1)
    txtTemplateDetails.Value = record.DisplayName & vbCrLf & record.Description & vbCrLf & _
        CStr(record.RowCount) & "행 × " & CStr(record.ColumnCount) & "열" & vbCrLf & "상태: " & record.Status
    cmdLoadNewSheet.Enabled = (record.Status = "healthy")
    RenderWorkflow NxExecutionGuarded, "불러오기는 확인 후 실행합니다."
End Sub

Private Sub RenderWorkflow(ByVal grade As NxExecutionGrade, ByVal reason As String)
    lblWorkflow.Caption = "대상 확인 > 작업 선택 > 옵션 > 미리보기/안전 > 실행"
    Select Case grade
        Case NxExecutionFast
            lblExecutionGrade.Caption = "실행 등급: [빠름] " & reason
        Case NxExecutionGuarded
            lblExecutionGrade.Caption = "실행 등급: [확인] " & reason
        Case NxExecutionPlanned
            lblExecutionGrade.Caption = "실행 등급: [계획] " & reason
        Case Else
            NxRaiseContractError "Template execution grade is invalid"
    End Select
    cmdRegister.Caption = "실행계획 확인"
    cmdLoadNewSheet.Caption = "확인 후 실행"
    cmdRename.Caption = "실행계획 확인"
    cmdDelete.Caption = "실행계획 확인"
End Sub

Private Function SelectedTemplateId() As String
    If mRecords Is Nothing Then NxRaiseContractError "템플릿을 선택하세요."
    If mRecords.Count = 0 Or cboTemplates.ListIndex < 0 Then NxRaiseContractError "템플릿을 선택하세요."
    SelectedTemplateId = CStr(mTemplateIds.Item(cboTemplates.ListIndex + 1))
End Function

Private Sub SetActionButtons(ByVal enabled As Boolean)
    cmdLoadNewSheet.Enabled = enabled
    cmdRename.Enabled = enabled
    cmdDelete.Enabled = enabled
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 나의 템플릿"
End Sub
