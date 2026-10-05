VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxRibbonTarget"
Option Explicit

Private mKind As String
Private mIdentifier As String
Private mSealed As Boolean

Public Sub Configure(ByVal kind As String, ByVal identifier As String)
    If mSealed Then NxRaiseContractError "Ribbon target cannot be changed"
    If Len(Trim$(kind)) = 0 Or Len(Trim$(identifier)) = 0 Then NxRaiseContractError "Ribbon target requires kind and identifier"
    mKind = kind
    mIdentifier = identifier
End Sub

Public Sub Seal()
    If Len(mKind) = 0 Or Len(mIdentifier) = 0 Then NxRaiseContractError "Ribbon target is not configured"
    mSealed = True
End Sub

Public Property Get Kind() As String
    EnsureReadable
    Kind = mKind
End Property

Public Property Get Identifier() As String
    EnsureReadable
    Identifier = mIdentifier
End Property

Public Property Get TargetKind() As String
    TargetKind = Kind
End Property

Public Property Get TargetId() As String
    TargetId = Identifier
End Property

Public Property Get IsSealed() As Boolean
    IsSealed = mSealed
End Property

Private Sub EnsureReadable()
    If Not mSealed Then NxRaiseContractError "Ribbon target must be sealed before use"
End Sub
