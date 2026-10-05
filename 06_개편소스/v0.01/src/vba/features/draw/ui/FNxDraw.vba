Option Explicit

Private mSource As Range
Private mPicture As Shape
Private mPreviewReady As Boolean
Private mAllowClose As Boolean

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
End Sub

Public Sub BindSelection(ByVal source As Range, Optional ByVal featureId As String = vbNullString)
    If source Is Nothing Or source.Areas.Count <> 1 Then NxRaiseContractError "표/그리기 폼은 연속 범위가 필요합니다."
    If featureId = NX_FEATURE_DRAW_FIT_PICTURE Then NxRaiseContractError "그림 맞춤 기능은 그림 선택으로 열어야 합니다."
    If Len(featureId) > 0 And featureId <> NX_FEATURE_DRAW_TITLE_TABLE And featureId <> NX_FEATURE_DRAW_BUSINESS_TABLE And _
        featureId <> NX_FEATURE_DRAW_CLEAR_INNER Then NxRaiseContractError "Drawing feature seed is outside the closed contract"
    Set mSource = source
    Set mPicture = Nothing
    lblRange.Caption = source.Worksheet.Name & "!" & source.Address(True, True, xlA1, False)
    If Len(featureId) > 0 Then cboFeature.ListIndex = FeatureIndexForId(featureId)
    lblExecutionGrade.Caption = "[빠름] 미리보기 선택 가능"
    mPreviewReady = False: cmdExecute.Enabled = False
End Sub

Public Sub BindPictureSelection(ByVal pictureShape As Shape, ByVal target As Range, Optional ByVal featureId As String = vbNullString)
    Dim pictureSheet As Object
    If pictureShape Is Nothing Or target Is Nothing Then NxRaiseContractError "그림과 대상 셀 범위가 필요합니다."
    If target.Areas.Count <> 1 Then NxRaiseContractError "그림 대상은 연속된 한 범위여야 합니다."
    Set pictureSheet = pictureShape.Parent
    If Not pictureSheet Is target.Worksheet Then NxRaiseContractError "그림과 대상 범위는 같은 워크시트여야 합니다."
    If Len(featureId) > 0 And featureId <> NX_FEATURE_DRAW_FIT_PICTURE Then NxRaiseContractError "표 기능은 범위 선택으로 열어야 합니다."
    Set mPicture = pictureShape: Set mSource = target
    lblRange.Caption = target.Worksheet.Name & "!" & target.Address(True, True, xlA1, False) & " (그림: " & pictureShape.Name & ")"
    If Len(featureId) > 0 Then cboFeature.ListIndex = FeatureIndexForId(featureId) Else cboFeature.ListIndex = 3
    lblExecutionGrade.Caption = "[빠름] 미리보기 선택 가능"
    mPreviewReady = False: cmdExecute.Enabled = False
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo Failed
    If mSource Is Nothing Then NxRaiseContractError "선택 범위를 먼저 바인딩하세요."
    If SelectedFeatureId() = NX_FEATURE_DRAW_FIT_PICTURE Then
        If mPicture Is Nothing Then NxRaiseContractError "그림 하나를 선택한 뒤 다시 열어주세요."
        If Not IsNumeric(txtPictureMargin.Value) Then NxRaiseContractError "그림 여백은 숫자로 입력하세요."
        NxDrawValidatePictureSelection mPicture, mSource, CDbl(txtPictureMargin.Value)
    Else
        NxDrawValidateSelection mSource
    End If
    lblConflicts.Caption = "보호 상태·병합·표 객체·조건부서식 충돌 검사를 완료했습니다." & vbCrLf & "값·수식·병합 상태는 유지됩니다."
    lblExecutionGrade.Caption = "[빠름] 대상·충돌 검사 후 실행"
    cmdExecute.Caption = "바로 실행"
    mPreviewReady = True: cmdExecute.Enabled = True
    Exit Sub
Failed:
    mPreviewReady = False
    cmdExecute.Enabled = False
    ShowError
End Sub

Private Sub cmdExecute_Click()
    On Error GoTo Failed
    cmdPreview_Click
    If Not mPreviewReady Then Exit Sub
    mSource.Parent.Activate: mSource.Select
    If SelectedFeatureId() = NX_FEATURE_DRAW_FIT_PICTURE Then
        If mPicture Is Nothing Then NxRaiseContractError "그림 하나를 선택한 뒤 다시 열어주세요."
        If Not IsNumeric(txtPictureMargin.Value) Then NxRaiseContractError "그림 여백은 숫자로 입력하세요."
        NxDrawFitPictureRun mPicture, mSource, CDbl(txtPictureMargin.Value), chkMoveAndSize.Value
    Else
        NxDrawRun SelectedFeatureId(), mSource, chkTotalRow.Value, chkPreserveAlignment.Value, chkHeader.Value, chkAutoFitColumns.Value
    End If
    RequestCancel
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdCancel_Click(): RequestCancel: End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub RequestCancel()
    mAllowClose = True
    Unload Me
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 표/그리기"
End Sub

Private Function SelectedFeatureId() As String
    Select Case cboFeature.ListIndex
        Case 0: SelectedFeatureId = NX_FEATURE_DRAW_TITLE_TABLE
        Case 1: SelectedFeatureId = NX_FEATURE_DRAW_BUSINESS_TABLE
        Case 2: SelectedFeatureId = NX_FEATURE_DRAW_CLEAR_INNER
        Case 3: SelectedFeatureId = NX_FEATURE_DRAW_FIT_PICTURE
        Case Else: NxRaiseContractError "Drawing selection is outside the closed contract"
    End Select
End Function
Private Function FeatureIndexForId(ByVal featureId As String) As Long
    Select Case featureId
        Case NX_FEATURE_DRAW_TITLE_TABLE: FeatureIndexForId = 0
        Case NX_FEATURE_DRAW_BUSINESS_TABLE: FeatureIndexForId = 1
        Case NX_FEATURE_DRAW_CLEAR_INNER: FeatureIndexForId = 2
        Case NX_FEATURE_DRAW_FIT_PICTURE: FeatureIndexForId = 3
        Case Else: NxRaiseContractError "Drawing feature seed is outside the closed contract"
    End Select
End Function
