Option Explicit

Private mTargetBook As Workbook
Private mSourceSheet As Worksheet
Private mSourceKind As String
Private mFilePath As String
Private mPreviewReady As Boolean
Private mBusy As Boolean
Private mResult As CNxResult
Private mPreviewDigest As String
Private mPreviewFileSha As String

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
    txtPreview.ScrollBars = fmScrollBarsVertical
    txtPreview.TabStop = True
End Sub

Public Sub BindSelection(ByVal source As Range, ByVal featureId As String)
    If source Is Nothing Then NxRaiseContractError "등록할 현재시트를 선택하세요."
    If featureId <> "NX-TPL-REGISTER-SHEET" Then NxRaiseContractError "지원하지 않는 템플릿 등록 방식입니다."
    BindSource source.Worksheet.Parent, source.Worksheet, "sheet_used_range"
End Sub

Public Sub BindSource(ByVal targetBook As Workbook, ByVal sourceSheet As Worksheet, ByVal sourceKind As String, _
    Optional ByVal filePath As String = vbNullString)
    NxTemplateValidateUserBook targetBook
    Set mTargetBook = targetBook
    Set mSourceSheet = sourceSheet
    mSourceKind = sourceKind
    mFilePath = filePath
    Me.Caption = "내엑셀 - " & NxTemplateManagerKindLabel(sourceKind) & " 등록"
    Select Case sourceKind
        Case "sheet_used_range"
            If mSourceSheet Is Nothing Then NxRaiseContractError "등록할 현재시트를 선택하세요."
            txtName.Value = mSourceSheet.Name
            txtTarget.Value = targetBook.Name & " / " & mSourceSheet.Name & "!" & mSourceSheet.UsedRange.Address
        Case "workbook"
            txtName.Value = CreateObject("Scripting.FileSystemObject").GetBaseName(targetBook.Name)
            txtTarget.Value = targetBook.Name & " / 표시 워크시트 전체"
        Case "file"
            mFilePath = NxTemplateRegistrationFilePath(mFilePath)
            txtName.Value = CreateObject("Scripting.FileSystemObject").GetBaseName(mFilePath)
            txtTarget.Value = mFilePath
        Case Else: NxRaiseContractError "지원하지 않는 템플릿 등록 원본입니다."
    End Select
    InvalidatePreview
End Sub

Public Property Get OperationResult() As CNxResult
    Set OperationResult = mResult
End Property

Private Sub cmdPreview_Click()
    Dim request As CNxTemplateRequest
    If mBusy Then Exit Sub
    On Error GoTo Failed
    InvalidatePreview
    mBusy = True
    If Len(Trim$(CStr(txtName.Value))) = 0 Then NxRaiseContractError "템플릿 이름을 입력하세요."
    If mSourceKind = "file" Then
        txtPreview.Value = NxTemplateFileRegistrationPreview(mFilePath, mTargetBook, CStr(txtName.Value), _
            CStr(txtDescription.Value), cboContents.ListIndex = 1, mPreviewFileSha)
    Else
        Set request = NxTemplateManagerRequest(mTargetBook, mSourceSheet, mSourceKind, CStr(txtName.Value), _
            CStr(txtDescription.Value), cboContents.ListIndex = 1, True)
        txtPreview.Value = NxTemplateRegistrationPreview(request, mPreviewDigest)
    End If
    txtPreview.SelStart = 0
    txtPreview.SelLength = 0
    mPreviewReady = True
    mBusy = False
    chkConfirmRisk.Enabled = True
    cmdApply.Enabled = True
    lblStatus.Caption = "스크롤로 상세 확인 후 동의하세요."
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdApply_Click()
    Dim request As CNxTemplateRequest, freshDigest As String, freshPreview As String
    If mBusy Then Exit Sub
    On Error GoTo Failed
    If Not chkConfirmRisk.Value Then NxRaiseContractError "저장 내용과 제외 항목을 확인하고 동의하세요."
    mBusy = True
    NxTemplateBeginProgress Me
    If mSourceKind = "file" Then
        If Not mPreviewReady Then mPreviewFileSha = NxTemplateFileSha256(mFilePath)
        If Len(mPreviewFileSha) <> 64 Then NxRaiseContractError "등록할 파일 상태를 확인하세요."
        Set mResult = NxTemplateRegisterFileFromManager(mFilePath, mTargetBook, CStr(txtName.Value), _
            CStr(txtDescription.Value), cboContents.ListIndex = 1, True, mPreviewFileSha, CStr(txtCategory.Value))
    Else
        Set request = NxTemplateManagerRequest(mTargetBook, mSourceSheet, mSourceKind, CStr(txtName.Value), _
            CStr(txtDescription.Value), cboContents.ListIndex = 1, True)
        request.Category = CStr(txtCategory.Value)
        If mPreviewReady Then
            freshPreview = NxTemplateRegistrationPreview(request, freshDigest)
            If Len(mPreviewDigest) = 0 Or freshDigest <> mPreviewDigest Then _
                NxRaiseContractError "미리보기 이후 원본 내용이나 시트 구성이 바뀌었습니다. 미리보기와 동의를 다시 확인하세요."
        End If
        Set mResult = NxTemplateRunRequestFromManager(request, mTargetBook)
    End If
    mBusy = False
    NxTemplateEndProgress
    If mResult Is Nothing Then NxRaiseContractError "템플릿 작업 결과가 없습니다."
    If mResult.Outcome = NxSuccess Then
        Me.Hide
    ElseIf mResult.Outcome = NxCancelled Then
        lblStatus.Caption = "등록을 취소했습니다."
    Else
        NxRaiseContractError mResult.Recovery
    End If
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cboContents_Change(): InvalidatePreview: End Sub
Private Sub txtName_Change(): InvalidatePreview: End Sub
Private Sub txtDescription_Change(): InvalidatePreview: End Sub
Private Sub txtCategory_Change(): InvalidatePreview: End Sub

Private Sub InvalidatePreview()
    mPreviewReady = False
    mPreviewDigest = vbNullString
    mPreviewFileSha = vbNullString
    chkConfirmRisk.Value = False
    chkConfirmRisk.Enabled = True
    cmdApply.Enabled = True
    txtPreview.Value = "미리보기는 선택 사항입니다. 저장 내용과 제외 항목에 동의한 뒤 바로 등록할 수 있습니다."
    lblStatus.Caption = "설정을 확인하고 등록하세요."
End Sub

Private Sub cmdCancel_Click()
    If mBusy Then
        NxTemplateRequestCancel
        lblStatus.Caption = "취소 요청 중입니다. 현재 작업의 임시 결과를 정리합니다."
        Exit Sub
    End If
    Set mResult = Nothing
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        cmdCancel_Click
    End If
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    NxTemplateEndProgress
    mBusy = False
    InvalidatePreview
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation, Me.Caption
End Sub

Public Sub UpdateTemplateProgress(ByVal stage As String, ByVal completed As Long, ByVal total As Long)
    lblStatus.Caption = NxTemplateProgressText(stage, completed, total) & " · 취소: 중단 요청"
    Me.Repaint
End Sub

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
