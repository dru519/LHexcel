VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxCalculatorHistory"
Option Explicit

Private Const NX_CALCULATOR_HISTORY_LIMIT As Long = 3
Private mItems As Collection

Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub

Public Sub Add(ByVal expressionText As String)
    If Len(Trim$(expressionText)) = 0 Then NxRaiseContractError "계산 이력은 비어 있을 수 없습니다."
    If mItems.Count = 0 Then
        mItems.Add expressionText
    Else
        mItems.Add expressionText, Before:=1
    End If
    Do While mItems.Count > NX_CALCULATOR_HISTORY_LIMIT
        mItems.Remove mItems.Count
    Loop
End Sub

Public Property Get Count() As Long
    Count = mItems.Count
End Property

Public Property Get Item(ByVal oneBasedIndex As Long) As String
    If oneBasedIndex < 1 Or oneBasedIndex > mItems.Count Then NxRaiseContractError "계산 이력 위치가 범위를 벗어났습니다."
    Item = CStr(mItems(oneBasedIndex))
End Property
