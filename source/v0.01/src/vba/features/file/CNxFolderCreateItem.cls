VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFolderCreateItem"
Option Explicit
Private mRelativePath As String, mSourceKind As String, mSelected As Boolean, mStatus As String, mDetail As String
Public Sub Configure(ByVal relativePath As String, Optional ByVal sourceKind As String = "direct input")
    If Len(Trim$(relativePath)) = 0 Then NxRaiseContractError "Folder item path is required"
    mRelativePath = Trim$(relativePath): mSourceKind = sourceKind: mSelected = True: mStatus = "pending"
End Sub
Public Property Get RelativePath() As String: RelativePath = mRelativePath: End Property
Public Property Let RelativePath(ByVal value As String)
    If Len(Trim$(value)) = 0 Then NxRaiseContractError "Folder item path is required"
    mRelativePath = value
End Property
Public Property Get SourceKind() As String: SourceKind = mSourceKind: End Property
Public Property Get SelectedForCreate() As Boolean: SelectedForCreate = mSelected: End Property
Public Property Let SelectedForCreate(ByVal value As Boolean): mSelected = value: End Property
Public Property Get StatusText() As String: StatusText = mStatus: End Property
Public Property Let StatusText(ByVal value As String): mStatus = value: End Property
Public Property Get DetailText() As String: DetailText = mDetail: End Property
Public Property Let DetailText(ByVal value As String): mDetail = value: End Property
