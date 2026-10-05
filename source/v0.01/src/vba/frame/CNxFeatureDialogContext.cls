VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFeatureDialogContext"
Option Explicit

Private mFeatureId As String
Private mDialogId As String
Private mDialogVariant As String
Private mIsWorkbench As Boolean
Private mConfigured As Boolean

Public Sub Configure(ByVal featureId As String, ByVal dialogId As String, _
    ByVal dialogVariant As String, ByVal isWorkbench As Boolean)

    If mConfigured Then NxRaiseContractError "Dialog context can be configured only once"
    If Len(featureId) = 0 And Not isWorkbench Then NxRaiseContractError "Feature dialog requires a feature ID"
    If Len(dialogId) = 0 Or Len(dialogVariant) = 0 Then NxRaiseContractError "Dialog context metadata is incomplete"
    mFeatureId = featureId
    mDialogId = dialogId
    mDialogVariant = dialogVariant
    mIsWorkbench = isWorkbench
    mConfigured = True
End Sub
Public Property Get FeatureId() As String: EnsureConfigured: FeatureId = mFeatureId: End Property
Public Property Get DialogId() As String: EnsureConfigured: DialogId = mDialogId: End Property
Public Property Get DialogVariant() As String: EnsureConfigured: DialogVariant = mDialogVariant: End Property
Public Property Get IsWorkbench() As Boolean: EnsureConfigured: IsWorkbench = mIsWorkbench: End Property
Public Property Get IsConfigured() As Boolean: IsConfigured = mConfigured: End Property

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Dialog context is not configured"
End Sub
