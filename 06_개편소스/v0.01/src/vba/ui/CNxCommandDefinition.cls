VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxCommandDefinition"
Option Explicit

Private mCommandId As String
Private mLabelKo As String
Private mCategoryId As String
Private mGroupId As String
Private mHandler As String
Private mArgumentKey As String
Private mMutatesDocument As Boolean
Private mShortcutEligible As Boolean
Private mRecoveryPolicy As String
Private mSealed As Boolean

Public Sub Configure(ByVal commandId As String, ByVal labelKo As String, ByVal categoryId As String, _
    ByVal groupId As String, ByVal handler As String, ByVal argumentKey As String, _
    ByVal mutatesDocument As Boolean, ByVal shortcutEligible As Boolean, _
    Optional ByVal recoveryPolicy As String = "none")
    If mSealed Then NxRaiseContractError "Command definition cannot be changed"
    If Len(Trim$(commandId)) = 0 Or Len(Trim$(labelKo)) = 0 Then NxRaiseContractError "Command definition requires id and label"
    If Len(Trim$(categoryId)) = 0 Or Len(Trim$(groupId)) = 0 Or Len(Trim$(handler)) = 0 Or Len(Trim$(argumentKey)) = 0 Then NxRaiseContractError "Command definition requires category, group, handler and argument"
    mCommandId = commandId
    mLabelKo = labelKo
    mCategoryId = categoryId
    mGroupId = groupId
    mHandler = handler
    mArgumentKey = argumentKey
    mMutatesDocument = mutatesDocument
    mShortcutEligible = shortcutEligible
    If mutatesDocument And Len(Trim$(recoveryPolicy)) = 0 Then NxRaiseContractError "Mutation command recovery policy is required"
    If mutatesDocument Then mRecoveryPolicy = recoveryPolicy Else mRecoveryPolicy = "none"
End Sub

Public Sub Seal()
    If Len(mCommandId) = 0 Then NxRaiseContractError "Command definition is not configured"
    mSealed = True
End Sub

Public Property Get CommandId() As String
    EnsureReadable
    CommandId = mCommandId
End Property

Public Property Get GroupId() As String
    EnsureReadable
    GroupId = mGroupId
End Property

Public Property Get LabelKo() As String
    EnsureReadable
    LabelKo = mLabelKo
End Property

Public Property Get ArgumentKey() As String
    EnsureReadable
    ArgumentKey = mArgumentKey
End Property

Public Property Get CategoryId() As String
    EnsureReadable
    CategoryId = mCategoryId
End Property

Public Property Get Handler() As String
    EnsureReadable
    Handler = mHandler
End Property

Public Property Get MutatesDocument() As Boolean
    EnsureReadable
    MutatesDocument = mMutatesDocument
End Property

Public Property Get ShortcutEligible() As Boolean
    EnsureReadable
    ShortcutEligible = mShortcutEligible
End Property

Public Property Get IsSealed() As Boolean
    IsSealed = mSealed
End Property

Public Property Get RecoveryPolicy() As String
    EnsureReadable
    RecoveryPolicy = mRecoveryPolicy
End Property

Private Sub EnsureReadable()
    If Not mSealed Then NxRaiseContractError "Command definition must be sealed before use"
End Sub
