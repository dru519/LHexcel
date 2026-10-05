VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFocusSheetState"
Option Explicit

Private mSheet As Worksheet
Private mSelection As CNxFocusSelection
Private mStatus As String
Private mReason As String

Public Sub Bind(ByVal sheet As Worksheet)
    If sheet Is Nothing Then NxRaiseContractError "포커스셀 시트 상태를 연결할 수 없습니다."
    If Not mSheet Is Nothing Then NxRaiseContractError "포커스셀 시트 상태를 연결할 수 없습니다."
    Set mSheet = sheet
    mStatus = "UNSEEN"
End Sub

Public Sub MarkActive(ByVal selection As CNxFocusSelection)
    If selection Is Nothing Then NxRaiseContractError "포커스셀 선택 상태가 유효하지 않습니다."
    If Not selection.IsValid Then NxRaiseContractError "포커스셀 선택 상태가 유효하지 않습니다."
    Set mSelection = selection
    mStatus = "ACTIVE"
    mReason = vbNullString
End Sub

Public Sub MarkSkipped(ByVal reason As String)
    Set mSelection = Nothing
    mStatus = "SKIPPED"
    mReason = reason
End Sub

Public Property Get Sheet() As Worksheet: Set Sheet = mSheet: End Property
Public Property Get Selection() As CNxFocusSelection: Set Selection = mSelection: End Property
Public Property Get Status() As String: Status = mStatus: End Property
Public Property Get Reason() As String: Reason = mReason: End Property
