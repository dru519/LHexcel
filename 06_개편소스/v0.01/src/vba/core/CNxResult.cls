VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxResult"
Option Explicit

Private mFeatureId As String
Private mOutcome As NxOutcome
Private mStage As String
Private mTarget As String
Private mCompleted As Collection
Private mSourceChanged As Boolean
Private mRecovery As String
Private mMessageKey As String
Private mIsSealed As Boolean

Private Sub Class_Initialize()
    Set mCompleted = New Collection
End Sub

Friend Sub Configure(ByVal featureId As String, ByVal outcome As NxOutcome, ByVal stage As String, ByVal target As String, ByVal sourceChanged As Boolean, ByVal recovery As String, ByVal messageKey As String)
    EnsureMutable
    If Len(Trim$(featureId)) = 0 Then NxRaiseContractError "Result feature is required"
    If outcome < NxSuccess Or outcome > NxPartialFailure Then NxRaiseContractError "Result outcome is outside the closed contract"
    If Len(Trim$(stage)) = 0 Or Len(Trim$(target)) = 0 Then NxRaiseContractError "Result requires stage and target"
    If outcome = NxSuccess And (stage <> "complete" Or Len(recovery) > 0) Then NxRaiseContractError "Success result must be final complete without recovery"
    mFeatureId = featureId
    mOutcome = outcome
    mStage = stage
    mTarget = target
    mSourceChanged = sourceChanged
    mRecovery = recovery
    mMessageKey = messageKey
End Sub

Friend Sub AddCompleted(ByVal item As String)
    EnsureMutable
    If Len(Trim$(item)) = 0 Then NxRaiseContractError "Completed item is required"
    mCompleted.Add item
End Sub

Friend Sub MergeRestoreFailures(ByVal failures As String)
    EnsureMutable
    If Len(failures) = 0 Then Exit Sub
    If Len(mStage) = 0 Then mStage = "restore" Else mStage = mStage & ";restore"
    If Len(mRecovery) = 0 Then mRecovery = failures Else mRecovery = mRecovery & " | " & failures
    ' Any failed rollback, compensator, restore, or cleanup leaves recovery work.
    mOutcome = NxPartialFailure
End Sub

Friend Sub Seal()
    EnsureMutable
    If Len(Trim$(mFeatureId)) = 0 Or Len(Trim$(mStage)) = 0 Or Len(Trim$(mMessageKey)) = 0 Then NxRaiseContractError "Result requires feature, stage, and message key"
    mIsSealed = True
End Sub

Friend Sub ValidateFor(ByVal expectedFeatureId As String, ByVal expectedTarget As String)
    If Not mIsSealed Then NxRaiseContractError "Result must be sealed before validation"
    If mFeatureId <> expectedFeatureId Then NxRaiseContractError "Result feature does not match approved feature"
    If mTarget <> expectedTarget Then NxRaiseContractError "Result target does not match approved target"
    If mOutcome = NxSuccess And (mStage <> "complete" Or Len(mRecovery) > 0) Then NxRaiseContractError "Success result must be complete without recovery"
    ' A valid cancellation is a terminal Execute result; runner reconciliation remeasures state.
End Sub

Friend Property Get IsSealed() As Boolean
    IsSealed = mIsSealed
End Property

Public Property Get FeatureId() As String: FeatureId = mFeatureId: End Property
Public Property Get Outcome() As NxOutcome: Outcome = mOutcome: End Property
Public Property Get Stage() As String: Stage = mStage: End Property
Public Property Get Target() As String: Target = mTarget: End Property
Public Property Get SourceChanged() As Boolean: SourceChanged = mSourceChanged: End Property
Public Property Get Recovery() As String: Recovery = mRecovery: End Property
Public Property Get MessageKey() As String: MessageKey = mMessageKey: End Property

Public Property Get Completed() As Collection
    Dim snapshot As New Collection
    Dim item As Variant
    For Each item In mCompleted
        snapshot.Add item
    Next item
    Set Completed = snapshot
End Property

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed result cannot be changed"
End Sub
