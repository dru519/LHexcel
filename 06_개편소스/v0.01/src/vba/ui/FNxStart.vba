Option Explicit

Private mAllowClose As Boolean

Private Sub UserForm_Initialize()
    mAllowClose = False
    LoadWorkbenches
    RefreshContextSummary
    cmdOpen.Enabled = False
End Sub

Private Sub UserForm_Activate()
    cmdClose.SetFocus
End Sub

Private Sub lstWorkbenches_Click()
    cmdOpen.Enabled = IsValidSelection()
End Sub

Private Sub cmdOpen_Click()
    Dim workbenchId As String
    On Error GoTo Failed
    If Not IsValidSelection() Then Exit Sub
    workbenchId = CStr(lstWorkbenches.List(lstWorkbenches.ListIndex, 0))
    NxRouteWorkbench workbenchId
    Exit Sub
Failed:
    lblSafety.Caption = "작업대를 열 수 없습니다. 현재 선택과 지원 상태를 확인하세요."
End Sub

Private Sub cmdClose_Click()
    RequestCancel
End Sub

Public Sub RequestCancel()
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Function IsValidSelection() As Boolean
    IsValidSelection = (lstWorkbenches.ListIndex >= 0 And Len(CStr(lstWorkbenches.List(lstWorkbenches.ListIndex, 0))) > 0)
End Function

Private Sub LoadWorkbenches()
    lstWorkbenches.Clear
    AddWorkbench "NX-WB-AI", "AI 작업대"
    AddWorkbench "NX-WB-DATA", "데이터 정리 작업대"
    AddWorkbench "NX-WB-DOC", "문서 작성 작업대"
    AddWorkbench "NX-WB-DRAW", "표·그리기 작업대"
    AddWorkbench "NX-WB-FILE", "파일 마무리 작업대"
    lstWorkbenches.ListIndex = -1
End Sub

Private Sub AddWorkbench(ByVal workbenchId As String, ByVal caption As String)
    lstWorkbenches.AddItem workbenchId
    lstWorkbenches.List(lstWorkbenches.ListCount - 1, 1) = caption
End Sub

Private Sub RefreshContextSummary()
    On Error GoTo UnknownContext
    lblWorkbook.Caption = "통합문서: " & CStr(ActiveWorkbook.Name)
    If TypeName(Application.Selection) = "Range" Then
        lblSelection.Caption = "선택 범위: " & CStr(Application.Selection.Address(False, False)) & " (" & CStr(Application.Selection.CountLarge) & "셀)"
    Else
        lblSelection.Caption = "선택 범위: 셀 범위가 선택되지 않았습니다."
    End If
    Exit Sub
UnknownContext:
    lblWorkbook.Caption = "통합문서: 확인할 수 없습니다."
    lblSelection.Caption = "선택 범위: 확인할 수 없습니다."
End Sub
