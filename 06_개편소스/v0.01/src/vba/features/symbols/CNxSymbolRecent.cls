VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxSymbolRecent"
Option Explicit

Private Const NX_SYMBOL_RECENT_LIMIT As Long = 16
Private mItems As Collection

Private Sub Class_Initialize()
    Set mItems = New Collection
End Sub

Public Sub Touch(ByVal record As CNxSymbolRecord)
    Dim index As Long, existing As CNxSymbolRecord
    If record Is Nothing Then NxRaiseContractError "최근 기호가 비어 있습니다."
    For index = mItems.Count To 1 Step -1
        Set existing = mItems(index)
        If existing.CodePoint = record.CodePoint Then mItems.Remove index
    Next index
    If mItems.Count = 0 Then
        mItems.Add record
    Else
        mItems.Add record, Before:=1
    End If
    Do While mItems.Count > NX_SYMBOL_RECENT_LIMIT
        mItems.Remove mItems.Count
    Loop
End Sub

Public Property Get Count() As Long
    Count = mItems.Count
End Property

Public Property Get Item(ByVal oneBasedIndex As Long) As CNxSymbolRecord
    Set Item = mItems(oneBasedIndex)
End Property

Public Function Snapshot() As Collection
    Dim result As New Collection, item As CNxSymbolRecord
    For Each item In mItems
        result.Add item
    Next item
    Set Snapshot = result
End Function
