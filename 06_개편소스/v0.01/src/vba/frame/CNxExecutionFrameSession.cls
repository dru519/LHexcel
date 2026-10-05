VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxExecutionFrameSession"
Option Explicit

Private mState As NxFrameState

Private Sub Class_Initialize()
    mState = NxFrameInput
End Sub

Public Sub BeginChecking()
    RequireState NxFrameInput
    mState = NxFrameChecking
End Sub

Public Sub MarkReady()
    RequireState NxFrameChecking
    mState = NxFrameReady
End Sub

Public Sub MarkBlocked()
    RequireState NxFrameChecking
    mState = NxFrameBlocked
End Sub

Public Sub FailChecking()
    RequireState NxFrameChecking
    mState = NxFrameError
End Sub

Public Sub BeginRunning()
    RequireState NxFrameReady
    mState = NxFrameRunning
End Sub

Public Sub CompleteFromResult(ByVal result As CNxResult)
    RequireState NxFrameRunning
    If result Is Nothing Then NxRaiseContractError "Frame completion requires a result"
    If Not result.IsSealed Then NxRaiseContractError "Frame completion requires a sealed result"
    Select Case result.Outcome
        Case NxSuccess
            mState = NxFrameSuccess
        Case NxCancelled
            mState = NxFrameCancelled
        Case NxPartialFailure
            mState = NxFramePartialFailure
        Case NxInputError, NxEnvironmentError
            mState = NxFrameError
        Case Else
            NxRaiseContractError "Frame result outcome is outside the closed contract"
    End Select
End Sub

Public Sub ReturnToInput()
    Select Case mState
        Case NxFrameReady, NxFrameBlocked
            mState = NxFrameInput
        Case Else
            NxRaiseContractError "Frame transition is not allowed"
    End Select
End Sub

Public Sub Cancel()
    Select Case mState
        Case NxFrameInput, NxFrameChecking, NxFrameReady, NxFrameBlocked
            mState = NxFrameCancelled
        Case Else
            NxRaiseContractError "Frame transition is not allowed"
    End Select
End Sub

Public Property Get State() As NxFrameState
    State = mState
End Property

Public Property Get IsTerminal() As Boolean
    Select Case mState
        Case NxFrameSuccess, NxFrameCancelled, NxFrameError, NxFramePartialFailure
            IsTerminal = True
    End Select
End Property

Private Sub RequireState(ByVal expected As NxFrameState)
    If mState <> expected Then NxRaiseContractError "Frame transition is not allowed"
End Sub
