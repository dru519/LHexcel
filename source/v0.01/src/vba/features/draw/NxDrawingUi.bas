Attribute VB_Name = "NxDrawingUi"
Option Explicit

Public Sub NxRibbonOpenDrawing(ByVal control As Object)
    NxDrawingOpenManager
End Sub

Public Sub NxDrawingOpenManager(Optional ByVal featureId As String = vbNullString)
    If Len(featureId) = 0 Then NxRaiseContractError "스타일·표그리기 목록에서 실행할 기능을 선택하세요."
    NxRouteFeature featureId
End Sub

Public Sub NxDrawingRunDirect(ByVal featureId As String)
    Dim selected As Object
    Dim target As Range
    Dim result As CNxResult

    Select Case featureId
        Case NX_FEATURE_DRAW_CLEAR_INNER
        Case Else
            NxRaiseContractError "지원하지 않는 표·그리기 즉시실행 기능입니다."
    End Select
    Set selected = Application.Selection
    If TypeName(selected) <> "Range" Then NxRaiseContractError "적용할 연속 범위를 먼저 선택하세요."
    Set target = selected
    NxDrawValidateSelection target
    Set result = NxDrawRun(featureId, target)
    If result Is Nothing Then NxRaiseContractError "표·그리기 즉시실행 결과가 없습니다."
    If result.Outcome <> NxSuccess Then NxRaiseContractError "표·그리기 즉시실행이 완료되지 않았습니다."
End Sub

Private Function PromptPictureTarget() As Range
    Dim target As Range
    On Error Resume Next
    Set target = Application.InputBox(Prompt:="그림을 맞출 연속 대상 범위를 선택하세요.", Title:="내엑셀 - 그림 셀 맞춤", Type:=8)
    On Error GoTo 0
    If target Is Nothing Then NxRaiseContractError "그림 대상 범위 선택이 취소되었거나 잘못되었습니다."
    If target.Areas.Count <> 1 Then NxRaiseContractError "그림 대상은 연속된 한 범위여야 합니다."
    Set PromptPictureTarget = target
End Function

Public Function NxDrawingFormContractHas(ByVal controlName As String) As Boolean
    Select Case controlName
        Case "cboFeature", "lblRange", "chkHeader", "chkTotalRow", "chkPreserveAlignment", "chkAutoFitColumns", _
             "txtPictureMargin", "chkMoveAndSize", "lblConflicts", "cmdPreview", "cmdExecute", "cmdCancel"
            NxDrawingFormContractHas = True
    End Select
End Function

Public Function NxDrawingFormUsesGeneratedSources() As Boolean
    NxDrawingFormUsesGeneratedSources = True
End Function
