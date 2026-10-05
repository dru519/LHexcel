Option Explicit

Private mFeatureId As String
Private mPreviewReady As Boolean
Private mAllowClose As Boolean

Private Sub UserForm_Initialize()
    cmdExecute.Default = True
    cmdCancel.Cancel = True
    txtOutputFolder.Value = NxPdfDefaultFolder()
    txtBaseName.Value = NxPdfDefaultBaseName()
    cboQuality.Clear
    cboQuality.AddItem "표준 품질"
    cboQuality.AddItem "작은 파일"
    cboQuality.ListIndex = 0
End Sub

Public Sub BindFeature(ByVal featureId As String)
    NxPdfValidateFeature featureId
    mFeatureId = featureId
    txtBaseName.Value = NxPdfDefaultBaseName(featureId)
    Select Case featureId
        Case NX_FEATURE_FILE_PDF_CURRENT_SHEET: lblTitle.Caption = "현재 시트를 PDF 한 파일로 저장합니다."
        Case NX_FEATURE_FILE_PDF_EACH_SHEET: lblTitle.Caption = "보이는 각 시트를 개별 PDF로 저장합니다."
        Case NX_FEATURE_FILE_PDF_SELECTED_COMBINED: lblTitle.Caption = "선택한 시트를 PDF 한 파일로 합칩니다."
        Case NX_FEATURE_FILE_PDF_ALL_COMBINED: lblTitle.Caption = "보이는 시트를 탭 순서대로 PDF 한 파일로 저장합니다. 인쇄영역을 따릅니다."
        Case NX_FEATURE_FILE_PDF_SETTINGS: lblTitle.Caption = "PDF 기본 저장 폴더·이름·품질을 저장합니다."
    End Select
    InvalidatePreview
End Sub

Private Sub cmdBrowse_Click()
    Dim picker As FileDialog
    Set picker = Application.FileDialog(4)
    With picker
        .AllowMultiSelect = False
        .Title = "PDF 저장 폴더 선택"
        If .Show = -1 Then txtOutputFolder.Value = CStr(.SelectedItems(1))
    End With
End Sub

Private Sub cmdPreview_Click()
    Dim pathText As String, sheet As Worksheet, count As Long
    On Error GoTo Failed
    ValidateInputs
    If mFeatureId = NX_FEATURE_FILE_PDF_SETTINGS Then
        pathText = "기본 설정 저장: " & NxPdfSettingsPath()
    ElseIf mFeatureId = NX_FEATURE_FILE_PDF_EACH_SHEET Then
        For Each sheet In ActiveWorkbook.Worksheets
            If sheet.Visible = xlSheetVisible Then
                NxPdfRequireVacantOutput NxPdfOutputPath(CStr(txtOutputFolder.Value), CStr(txtBaseName.Value), sheet.Name)
                count = count + 1
            End If
        Next sheet
        If count = 0 Then NxRaiseContractError "PDF로 저장할 표시 워크시트가 없습니다."
        pathText = "개별 PDF " & CStr(count) & "개 · " & CStr(txtOutputFolder.Value)
    Else
        pathText = NxPdfOutputPath(CStr(txtOutputFolder.Value), CStr(txtBaseName.Value))
        NxPdfRequireVacantOutput pathText
    End If
    txtPreview.Value = lblTitle.Caption & vbCrLf & pathText & vbCrLf & "품질: " & CStr(cboQuality.Value)
    mPreviewReady = True
    cmdExecute.Enabled = True
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdExecute_Click()
    Dim options As Object, result As CNxResult, approvedTarget As String
    On Error GoTo Failed
    ValidateInputs
    Set options = CreateObject("Scripting.Dictionary")
    options.Add "output_folder", CStr(txtOutputFolder.Value)
    options.Add "base_name", CStr(txtBaseName.Value)
    options.Add "quality", CStr(cboQuality.Value)
    approvedTarget = CStr(txtOutputFolder.Value)
    If mFeatureId = NX_FEATURE_FILE_PDF_SETTINGS Then approvedTarget = NxPdfSettingsPath()
    Set result = NxFileRun(mFeatureId, approvedTarget, workflowOptions:=options)
    If result Is Nothing Then Exit Sub
    If result.Outcome = NxCancelled Then Exit Sub
    If result.Outcome <> NxSuccess Then NxRaiseContractError "PDF 작업을 완료하지 못했습니다."
    RequestCancel
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub ValidateInputs()
    If Len(mFeatureId) = 0 Then NxRaiseContractError "PDF 기능을 다시 실행하세요."
    NxPdfValidateFeature mFeatureId
    If Len(Dir$(CStr(txtOutputFolder.Value), vbDirectory)) = 0 Then NxRaiseContractError "PDF 저장 폴더를 찾을 수 없습니다."
    Call NxPdfOutputPath(CStr(txtOutputFolder.Value), CStr(txtBaseName.Value))
    Call NxPdfQualityValue(CStr(cboQuality.Value))
    If mFeatureId <> NX_FEATURE_FILE_PDF_SETTINGS And ActiveWorkbook Is Nothing Then _
        NxRaiseContractError "PDF로 저장할 통합문서가 없습니다."
End Sub

Private Sub txtOutputFolder_Change(): InvalidatePreview: End Sub
Private Sub txtBaseName_Change(): InvalidatePreview: End Sub
Private Sub cboQuality_Change(): InvalidatePreview: End Sub

Private Sub InvalidatePreview()
    mPreviewReady = False
    cmdExecute.Enabled = True
End Sub

Private Sub cmdCancel_Click(): RequestCancel: End Sub

Private Sub RequestCancel()
    mAllowClose = True
    Me.Hide
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
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - PDF"
End Sub

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
