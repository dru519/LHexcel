VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxResultCardViewModel"
Option Explicit

Private mFeatureId As String, mTarget As String, mMessageKey As String
Private mOutcomeLabel As String, mCompletedSummary As String, mIncompleteSummary As String
Private mSourceChangedLabel As String, mResultRecovery As String, mRecoveryDisplaySummary As String
Private mResultPath As String, mNextAction As String, mTargetSummary As String
Private mOutcome As NxOutcome, mSourceChanged As Boolean
Private mConfigured As Boolean, mIsSealed As Boolean

Friend Sub Configure(ByVal result As CNxResult, ByVal outcomeLabel As String, _
    ByVal completedSummary As String, ByVal incompleteSummary As String, _
    ByVal sourceChangedLabel As String, ByVal recoverySummary As String, _
    ByVal resultPath As String, ByVal nextAction As String, _
    Optional ByVal targetSummary As String = vbNullString)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Result card view model can be configured only once"
    If result Is Nothing Then NxRaiseContractError "Result card view model requires a result"
    If Not result.IsSealed Then NxRaiseContractError "Result card view model requires a sealed result"
    If Len(Trim$(outcomeLabel)) = 0 Or Len(Trim$(completedSummary)) = 0 Or _
        Len(Trim$(incompleteSummary)) = 0 Or Len(Trim$(sourceChangedLabel)) = 0 Or _
        Len(Trim$(recoverySummary)) = 0 Or Len(Trim$(nextAction)) = 0 Then
        NxRaiseContractError "Result card view model requires complete display text"
    End If
    mFeatureId = result.FeatureId
    mTarget = result.Target
    mMessageKey = result.MessageKey
    mOutcome = result.Outcome
    mSourceChanged = result.SourceChanged
    mOutcomeLabel = outcomeLabel
    mCompletedSummary = completedSummary
    mIncompleteSummary = incompleteSummary
    mSourceChangedLabel = sourceChangedLabel
    ' RecoverySummary remains the raw result value because planned feature
    ' adapters use it to reconstruct the sealed result. UI text is separate.
    mResultRecovery = result.Recovery
    mRecoveryDisplaySummary = recoverySummary
    mResultPath = resultPath
    mNextAction = nextAction
    mTargetSummary = targetSummary
    If Len(Trim$(mTargetSummary)) = 0 Then mTargetSummary = "실행 계획의 대상"
    mConfigured = True
End Sub

Public Sub Seal()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Result card view model is not configured"
    mIsSealed = True
End Sub

Public Property Get FeatureId() As String: EnsureReadable: FeatureId = mFeatureId: End Property
Public Property Get Target() As String: EnsureReadable: Target = mTarget: End Property
Public Property Get MessageKey() As String: EnsureReadable: MessageKey = mMessageKey: End Property
Public Property Get Outcome() As NxOutcome: EnsureReadable: Outcome = mOutcome: End Property
Public Property Get SourceChanged() As Boolean: EnsureReadable: SourceChanged = mSourceChanged: End Property
Public Property Get OutcomeLabel() As String: EnsureReadable: OutcomeLabel = mOutcomeLabel: End Property
Public Property Get CompletedSummary() As String: EnsureReadable: CompletedSummary = mCompletedSummary: End Property
Public Property Get IncompleteSummary() As String: EnsureReadable: IncompleteSummary = mIncompleteSummary: End Property
Public Property Get SourceChangedLabel() As String: EnsureReadable: SourceChangedLabel = mSourceChangedLabel: End Property
Public Property Get RecoverySummary() As String: EnsureReadable: RecoverySummary = mResultRecovery: End Property
Public Property Get RecoveryDisplaySummary() As String: EnsureReadable: RecoveryDisplaySummary = mRecoveryDisplaySummary: End Property
Public Property Get ResultPath() As String: EnsureReadable: ResultPath = mResultPath: End Property
Public Property Get TargetSummary() As String: EnsureReadable: TargetSummary = mTargetSummary: End Property
Public Property Get NextAction() As String: EnsureReadable: NextAction = mNextAction: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed result card view model cannot be changed"
End Sub

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Result card view model must be sealed before use"
End Sub
