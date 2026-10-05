VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFolderCreatePreview"
Option Explicit
Private mBaseFolder As String, mItems As Collection, mSnapshot As String, mValid As Boolean, mDirectoryPaths As Collection, mSummary As String
Public Sub Snapshot(ByVal baseFolder As String, ByVal items As Collection)
    Dim detail As String
    If Len(Trim$(baseFolder)) = 0 Or items Is Nothing Then NxRaiseContractError "Folder preview requires base folder and items"
    Set mItems = items: mBaseFolder = baseFolder
    If Not NxFolderCreatePlan(baseFolder, items, mDirectoryPaths, detail) Then NxRaiseContractError detail
    mSummary = detail: mSnapshot = BuildSnapshotIdentity(baseFolder, items): mValid = True
End Sub
Public Sub Invalidate(): mValid = False: End Sub
Public Property Get IsValid() As Boolean: IsValid = mValid And BuildSnapshotIdentity(mBaseFolder, mItems) = mSnapshot: End Property
Public Property Get BaseFolder() As String: BaseFolder = mBaseFolder: End Property
Public Property Get Items() As Collection: Set Items = mItems: End Property
Public Property Get SnapshotIdentity() As String: SnapshotIdentity = mSnapshot: End Property
Public Property Get DirectoryPaths() As Collection: Set DirectoryPaths = mDirectoryPaths: End Property
Public Property Get PlannedCount() As Long
    If Not mDirectoryPaths Is Nothing Then PlannedCount = mDirectoryPaths.Count
End Property
Public Property Get Summary() As String: Summary = mSummary: End Property
Private Function BuildSnapshotIdentity(ByVal baseFolder As String, ByVal items As Collection) As String
    Dim item As CNxFolderCreateItem, value As String: value = LCase$(Trim$(baseFolder)) & "|"
    For Each item In items: value = value & CStr(item.SelectedForCreate) & ":" & item.RelativePath & ";": Next item
    BuildSnapshotIdentity = value
End Function
