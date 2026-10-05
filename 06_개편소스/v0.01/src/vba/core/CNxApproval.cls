VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxApproval"
Option Explicit

Private mRunId As String
Private mCommand As INxFeatureCommand
Private mPlan As CNxPlan
Private mContext As CNxExecutionContext
Private mContextFingerprint As String
Private mPlanSealDigest As String
Private mSourceDigest As String
Private mContextSourceDigest As String
Private mSideEffectSnapshots As Collection
Private mConsumed As Boolean
Private mConfigured As Boolean
Private mIssuerAuthority As Collection

Private Sub Class_Initialize()
    Set mSideEffectSnapshots = New Collection
End Sub

Friend Sub Configure(ByVal issuerAuthority As Collection, ByVal exactCommand As INxFeatureCommand, ByVal exactPlan As CNxPlan, ByVal exactContext As CNxExecutionContext)
    Dim effect As Variant, snapshot As New Collection
    If mConfigured Then NxRaiseContractError "Approval can be configured only once"
    If issuerAuthority Is Nothing Then NxRaiseContractError "Approval issuer authority is required"
    If exactCommand Is Nothing Or exactPlan Is Nothing Or exactContext Is Nothing Then NxRaiseContractError "Approval requires exact command, plan, and context"
    If Not exactPlan.IsSealed Or exactPlan.ContextFingerprint <> exactContext.Fingerprint Then NxRaiseContractError "Approval requires current sealed plan"
    Set mIssuerAuthority = issuerAuthority
    mRunId = NxCreateRunUuid
    Set mCommand = exactCommand: Set mPlan = exactPlan: Set mContext = exactContext
    mContextFingerprint = exactContext.Fingerprint: mPlanSealDigest = exactPlan.SealDigest: mSourceDigest = exactPlan.SourceDigest: mContextSourceDigest = exactContext.SourceStateDigest
    For Each effect In exactPlan.SideEffects
        snapshot.Add effect
        snapshot.Add effect.StableIdentity
    Next effect
    Set mSideEffectSnapshots = snapshot
    mConfigured = True
End Sub

Friend Function ConsumeFor(ByVal issuerAuthority As Collection, ByVal exactCommand As INxFeatureCommand, ByVal exactPlan As CNxPlan, ByVal currentContext As CNxExecutionContext) As Boolean
    Dim index As Long, effect As CNxSideEffect, snapshotEffect As CNxSideEffect
    If mConsumed Then
        ' Approval capability was already consumed.
        Exit Function
    End If
    mConsumed = True
    If issuerAuthority Is Nothing Or mIssuerAuthority Is Nothing Then Exit Function
    If Not (issuerAuthority Is mIssuerAuthority) Then Exit Function
    If exactCommand Is Nothing Or exactPlan Is Nothing Or currentContext Is Nothing Then Exit Function
    If Not exactPlan.IsSealed Then Exit Function
    If Not (mCommand Is exactCommand) Then Exit Function
    If Not (mPlan Is exactPlan) Then Exit Function
    If Not mContext.MatchesRuntimeIdentity(currentContext) Then Exit Function
    If mContextFingerprint <> currentContext.Fingerprint Then Exit Function
    If mPlanSealDigest <> exactPlan.SealDigest Or mSourceDigest <> exactPlan.SourceDigest Then Exit Function
    ' Run has just captured the same exact selection. Compare that full fresh
    ' snapshot once rather than scanning every source property a second time.
    If mContextSourceDigest <> currentContext.SourceStateDigest Then Exit Function
    If mSideEffectSnapshots.Count <> exactPlan.SideEffects.Count * 2 Then Exit Function
    For index = 1 To mSideEffectSnapshots.Count Step 2
        Set effect = exactPlan.SideEffects.Item((index + 1) \ 2)
        Set snapshotEffect = mSideEffectSnapshots.Item(index)
        If Not (snapshotEffect Is effect) Then Exit Function
        If CStr(mSideEffectSnapshots.Item(index + 1)) <> effect.StableIdentity Then Exit Function
    Next index
    ConsumeFor = True
End Function

Friend Property Get Plan() As CNxPlan: Set Plan = mPlan: End Property
Friend Property Get Context() As CNxExecutionContext: Set Context = mContext: End Property
Public Property Get RunId() As String: RunId = mRunId: End Property
Public Property Get IsConsumed() As Boolean: IsConsumed = mConsumed: End Property
