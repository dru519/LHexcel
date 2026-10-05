VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxCategoryDefinition"
Option Explicit

Private mId As String
Private mParentId As String
Private mLabel As String
Private mSortOrder As Long
Private mConfigured As Boolean
Private mIsSealed As Boolean

Public Sub Configure(ByVal categoryId As String, ByVal parentId As String, ByVal label As String, ByVal sortOrder As Long)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Category definition can be configured only once"
    If Len(Trim$(categoryId)) = 0 Or Len(Trim$(parentId)) = 0 Or Len(Trim$(label)) = 0 Then NxRaiseContractError "Category definition requires ID, parent, and label"
    If sortOrder <= 0 Then NxRaiseContractError "Category sort order must be positive"
    mId = categoryId
    mParentId = parentId
    mLabel = label
    mSortOrder = sortOrder
    mConfigured = True
End Sub

Public Sub Seal()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Category definition is not configured"
    mIsSealed = True
End Sub

Public Property Get Id() As String: EnsureReadable: Id = mId: End Property
Public Property Get ParentId() As String: EnsureReadable: ParentId = mParentId: End Property
Public Property Get Label() As String: EnsureReadable: Label = mLabel: End Property
Public Property Get SortOrder() As Long: EnsureReadable: SortOrder = mSortOrder: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed definition cannot be changed"
End Sub

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Category definition must be sealed before use"
End Sub
