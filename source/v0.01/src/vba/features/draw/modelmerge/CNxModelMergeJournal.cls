VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxModelMergeJournal"
Option Explicit

Private mSnapshots As Object

Public Function HasSnapshot(ByVal target As Range) As Boolean
    EnsureStore
    HasSnapshot = mSnapshots.Exists(KeyForRange(target))
End Function

Public Function SnapshotFor(ByVal target As Range) As CNxModelMergeSnapshot
    Dim key As String
    EnsureStore
    key = KeyForRange(target)
    If mSnapshots.Exists(key) Then Set SnapshotFor = mSnapshots(key)
End Function

Public Sub PutSnapshot(ByVal snapshot As CNxModelMergeSnapshot)
    If snapshot Is Nothing Then NxRaiseContractError "보존할 원본 상태가 필요합니다."
    EnsureStore
    If mSnapshots.Exists(snapshot.IdentityKey) Then mSnapshots.Remove snapshot.IdentityKey
    mSnapshots.Add snapshot.IdentityKey, snapshot
End Sub

Public Sub RemoveSnapshot(ByVal target As Range)
    Dim key As String
    EnsureStore
    key = KeyForRange(target)
    If mSnapshots.Exists(key) Then mSnapshots.Remove key
End Sub

Private Sub EnsureStore()
    If mSnapshots Is Nothing Then Set mSnapshots = CreateObject("Scripting.Dictionary")
End Sub

Private Function KeyForRange(ByVal target As Range) As String
    If target Is Nothing Then NxRaiseContractError "범위 식별자가 필요합니다."
    KeyForRange = target.Worksheet.Parent.FullName & "|" & target.Worksheet.CodeName & "|" & _
        target.Address(True, True, xlA1, False)
End Function
