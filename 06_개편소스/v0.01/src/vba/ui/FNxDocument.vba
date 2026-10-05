Option Explicit

Private mAllowClose As Boolean
Private mPreviewReady As Boolean

Private Sub UserForm_Initialize()
    NxInitializeStaticListChoices Me
    lblSelection.Caption = "선택 범위: " & CurrentSelectionSummary()
    lblPrivacyBoundary.Caption = "안전 경계: 소유 작업대가 개인정보·출력·충돌 검사를 담당합니다."
    lblExecutionGrade.Caption = "[빠름] 미리보기"
    lstDocumentModes.ListIndex = 0
    InvalidatePreview
End Sub

Private Sub UserForm_Activate()
    NxFormWheelAttach Me
End Sub

Public Sub HandleWheelDelta(ByVal wheelDelta As Long)
    NxFormWheelScrollActiveList Me, "lstDocumentModes", wheelDelta
End Sub

Private Sub lstDocumentModes_Change()
    InvalidatePreview
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo Failed
    txtPreview.Value = PreviewSummary(SelectedModeId())
    lblExecutionGrade.Caption = GradeSummary(SelectedModeId())
    cmdOpenOwner.Enabled = True
    mPreviewReady = True
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdOpenOwner_Click()
    On Error GoTo Failed
    If Not mPreviewReady Then NxRaiseContractError "문서 작업을 먼저 미리보세요."
    ' This facade never executes owner logic. It delegates to the central router.
    Select Case SelectedModeId()
        Case "NX-AI-WRITE", "NX-AI-IMAGE"
            NxRouteFeature SelectedModeId()
        Case "NX-ENTRY-TEMPLATE"
            NxRouteCategory "NX-CAT-TEMPLATE"
        Case Else
            NxRaiseContractError "문서 작업을 선택하세요."
    End Select
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdClose_Click()
    RequestCancel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    NxFormWheelDetach Me
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub RequestCancel()
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_Terminate()
    NxFormWheelDetach Me
End Sub

Private Sub InvalidatePreview()
    mPreviewReady = False
    cmdOpenOwner.Enabled = False
    txtPreview.Value = "문서 작업을 선택하고 미리보기를 실행하세요."
End Sub

Private Function SelectedModeId() As String
    Select Case lstDocumentModes.ListIndex
        Case 0: SelectedModeId = "NX-AI-WRITE"
        Case 1: SelectedModeId = "NX-AI-IMAGE"
        Case 2: SelectedModeId = "NX-ENTRY-TEMPLATE"
        Case Else: NxRaiseContractError "문서 작업을 선택하세요."
    End Select
End Function

Private Function PreviewSummary(ByVal featureId As String) As String
    PreviewSummary = "선택 작업: " & featureId & vbCrLf & _
        "미리보기는 원본을 변경하지 않으며, 실행은 기능 소유 작업대에서 진행합니다." & vbCrLf & _
        "다음 단계: 소유 작업대 열기 → 옵션·충돌·실행등급 확인"
End Function

Private Function GradeSummary(ByVal featureId As String) As String
    If featureId = "NX-AI-WRITE" Or featureId = "NX-AI-IMAGE" Then
        GradeSummary = "[계획] 소유 작업대에서 실행계획 확인"
    Else
        GradeSummary = "[빠름] 원본 변경 없는 미리보기"
    End If
End Function

Private Function CurrentSelectionSummary() As String
    If TypeName(Application.Selection) = "Range" Then
        CurrentSelectionSummary = Application.Selection.Worksheet.Name & "!" & Application.Selection.Address(False, False)
    Else
        CurrentSelectionSummary = "범위 또는 그림을 선택하세요."
    End If
End Function

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 문서 작업대"
End Sub

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
