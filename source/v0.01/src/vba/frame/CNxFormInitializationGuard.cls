VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFormInitializationGuard"
Option Explicit

Private Const NX_FORM_INIT_NEW As Long = 0
Private Const NX_FORM_INIT_RUNNING As Long = 1
Private Const NX_FORM_INIT_READY As Long = 2

Private mState As Long

Public Sub BeginInitialization()
    If mState <> NX_FORM_INIT_NEW Then NxRaiseContractError "Form initialization has already started"
    mState = NX_FORM_INIT_RUNNING
End Sub

Public Sub CompleteInitialization()
    If mState <> NX_FORM_INIT_RUNNING Then NxRaiseContractError "Form initialization did not start"
    mState = NX_FORM_INIT_READY
End Sub

Public Property Get IsInitializing() As Boolean
    IsInitializing = (mState = NX_FORM_INIT_RUNNING)
End Property

Public Sub RequireReady()
    If mState <> NX_FORM_INIT_READY Then NxRaiseContractError "Form initialization is not ready"
End Sub
