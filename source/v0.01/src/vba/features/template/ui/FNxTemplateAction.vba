Option Explicit

Private mFeatureId As String
Private mRecords As Collection
Private mTemplateIds As Collection
Private mAllowClose As Boolean
Private mTargetBook As Workbook

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
    cmdExecute.Default = True
    cmdCancel.Cancel = True
End Sub

Public Sub BindFeature(ByVal featureId As String)
    BindContext ActiveWorkbook, featureId
End Sub

Public Sub BindContext(ByVal targetBook As Workbook, ByVal featureId As String)
    NxTemplateValidateUserBook targetBook
    Set mTargetBook = targetBook
    Select Case featureId
        Case "NX-TPL-LIST", "NX-TPL-LOAD", "NX-TPL-RENAME", "NX-TPL-DELETE"
        Case Else: NxRaiseContractError "지원하지 않는 템플릿 기능입니다."
    End Select
    mFeatureId = featureId
    ConfigureAction
    RefreshTemplates
End Sub

Private Sub ConfigureAction()
    txtName.Enabled = (mFeatureId = "NX-TPL-RENAME")
    txtDescription.Enabled = txtName.Enabled
    Select Case mFeatureId
        Case "NX-TPL-LIST"
            Me.Caption = "내엑셀 - 템플릿 목록"
            lblTitle.Caption = "저장된 템플릿을 확인합니다."
            lblStatus.Caption = "목록은 읽기 전용입니다."
            cmdExecute.Visible = False
        Case "NX-TPL-LOAD"
            Me.Caption = "내엑셀 - 템플릿 사용"
            lblTitle.Caption = "새 시트로 불러올 템플릿을 선택하세요."
            lblStatus.Caption = "대상: " & mTargetBook.Name & " / 기존 시트 보존"
            cmdExecute.Caption = "불러오기"
        Case "NX-TPL-RENAME"
            Me.Caption = "내엑셀 - 템플릿 이름·설명 변경"
            lblTitle.Caption = "이름을 바꿀 템플릿을 선택하세요."
            cmdExecute.Caption = "이름 변경"
        Case "NX-TPL-DELETE"
            Me.Caption = "내엑셀 - 템플릿 제거"
            lblTitle.Caption = "격리 보관할 템플릿을 선택하세요."
            cmdExecute.Caption = "격리 보관"
    End Select
End Sub

Private Sub RefreshTemplates()
    Dim item As Variant
    Dim record As CNxTemplateRecord
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
        txtDetails.Value = "등록된 템플릿이 없습니다."
        cmdExecute.Enabled = False
    Else
        cboTemplates.ListIndex = 0
        cmdExecute.Enabled = (mFeatureId <> "NX-TPL-LIST")
        RenderSelected
    End If
End Sub

Private Sub cboTemplates_Change(): RenderSelected: End Sub

Private Sub RenderSelected()
    Dim record As CNxTemplateRecord
    If mRecords Is Nothing Then Exit Sub
    If mRecords.Count = 0 Then Exit Sub
    If cboTemplates.ListIndex < 0 Or cboTemplates.ListIndex >= mRecords.Count Then Exit Sub
    Set record = mRecords.Item(cboTemplates.ListIndex + 1)
    txtDetails.Value = record.DisplayName & vbCrLf & record.Description & vbCrLf & _
        CStr(CDbl(record.RowCount) * CDbl(record.ColumnCount)) & "개 사용 셀 · " & record.Status
    txtName.Value = record.DisplayName
    txtDescription.Value = record.Description
    cmdExecute.Enabled = (mFeatureId <> "NX-TPL-LIST" And (record.Status = "healthy" Or mFeatureId = "NX-TPL-DELETE"))
End Sub

Private Sub cmdExecute_Click()
    Dim result As CNxResult, action As String, templateId As String
    On Error GoTo Failed
    If mRecords Is Nothing Then NxRaiseContractError "템플릿을 선택하세요."
    If mRecords.Count = 0 Or cboTemplates.ListIndex < 0 Then NxRaiseContractError "템플릿을 선택하세요."
    If mFeatureId = "NX-TPL-RENAME" And Len(Trim$(CStr(txtName.Value))) = 0 Then NxRaiseContractError "새 이름을 입력하세요."
    templateId = CStr(mTemplateIds.Item(cboTemplates.ListIndex + 1))
    Select Case mFeatureId
        Case "NX-TPL-LOAD"
            action = "use"
            If Not NxTemplateManagerActionConsent(templateId) Then Exit Sub
        Case "NX-TPL-RENAME": action = "rename"
        Case "NX-TPL-DELETE": action = "delete"
        Case Else: NxRaiseContractError "실행할 템플릿 작업을 선택하세요."
    End Select
    Set result = NxTemplateRunManagerAction(action, mTargetBook, templateId, True, CStr(txtName.Value), CStr(txtDescription.Value))
    If result Is Nothing Then NxRaiseContractError "템플릿 작업을 완료하지 못했습니다."
    If result.Outcome = NxCancelled Then Exit Sub
    If result.Outcome <> NxSuccess Then NxRaiseContractError result.Recovery
    RequestCancel
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdCancel_Click(): RequestCancel: End Sub

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
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 템플릿"
End Sub
