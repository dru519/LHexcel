VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxPlan"
Option Explicit

Private mFeatureId As String
Private mPlanIdentity As String
Private mContextFingerprint As String
Private mSealDigest As String
Private mOwner As String
Private mSourceDigest As String
Private mLimitDecision As Long
Private mDeclaredCapabilities As Long
Private mObservedCellCount As Double
Private mHasObservedCellCount As Boolean
Private mSideEffects As Collection
Private mIsSealed As Boolean
Private mConfigured As Boolean

Private Sub Class_Initialize()
    Set mSideEffects = New Collection
End Sub

Friend Sub DeclareCapability(ByVal capability As NxCapability)
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Plan must be configured before declaring capabilities"
    If capability <> NxCapabilitySensitiveOutput And capability <> NxCapabilityTemplateQuarantine And _
        capability <> NxCapabilityRiskyFormula Then NxRaiseContractError "Plan marker capability is outside the closed contract"
    If (mDeclaredCapabilities And CLng(capability)) <> 0 Then NxRaiseContractError "Plan rejects duplicate declared capability"
    mDeclaredCapabilities = mDeclaredCapabilities Or CLng(capability)
End Sub

Friend Sub DeclareObservedCells(ByVal actualCellCount As Double)
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Plan must be configured before declaring observed cells"
    If mHasObservedCellCount Then NxRaiseContractError "Plan rejects duplicate observed cell declarations"
    If actualCellCount < 0 Or actualCellCount <> Fix(actualCellCount) Then NxRaiseContractError "Observed cell count must be an unsigned integer"
    If actualCellCount > CDbl(mLimitDecision) Then NxRaiseContractError "Observed cell count exceeds the exact feature limit"
    mObservedCellCount = actualCellCount
    mHasObservedCellCount = True
End Sub

Friend Sub Configure(ByVal featureId As String, ByVal contextFingerprint As String, ByVal SourceStateDigest As String)
    Dim policy As New CNxLimitsPolicy
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Plan can be configured only once"
    If Len(Trim$(featureId)) = 0 Or Len(Trim$(contextFingerprint)) = 0 Then NxRaiseContractError "Plan requires feature and context"
    mFeatureId = featureId
    mContextFingerprint = contextFingerprint
    mOwner = policy.ExactOwnerFor(featureId)
    mLimitDecision = policy.MaximumCellsFor(featureId)
    If Len(SourceStateDigest) = 0 Then NxRaiseContractError "Plan requires approved context source state"
    mSourceDigest = policy.SourceDigest & "|" & SourceStateDigest
    mPlanIdentity = "NXPLAN1|" & EncodeField(featureId) & EncodeField(contextFingerprint)
    mConfigured = True
End Sub

Public Sub AddSideEffect(ByVal effect As CNxSideEffect)
    Dim existing As Variant
    EnsureMutable
    If effect Is Nothing Then NxRaiseContractError "Typed side effect is required"
    If Not effect.IsSealed Then NxRaiseContractError "Plan owns only sealed side effects"
    If effect.Owner <> mOwner Or effect.LimitDecision <> mLimitDecision Then NxRaiseContractError "Side effect owner or limit decision does not match plan"
    For Each existing In mSideEffects
        If existing.Target = effect.Target Then NxRaiseContractError "Plan rejects duplicate approved cell targets"
    Next existing
    mSideEffects.Add effect
End Sub

Public Sub Seal()
    Dim snapshot As New Collection
    Dim effect As Variant
    Dim policy As New CNxLimitsPolicy
    EnsureMutable
    If Len(mPlanIdentity) = 0 Then NxRaiseContractError "Plan is not configured"
    If mDeclaredCapabilities <> 0 Then mPlanIdentity = mPlanIdentity & EncodeField("C" & NxInvariantUnsigned(CDbl(mDeclaredCapabilities)))
    If mHasObservedCellCount Then mPlanIdentity = mPlanIdentity & EncodeField("O" & NxInvariantUnsigned(mObservedCellCount))
    For Each effect In mSideEffects
        If Not effect.IsSealed Then NxRaiseContractError "Plan owns only sealed side effects"
        policy.ValidateEffectFor mFeatureId, effect
        If effect.Owner <> mOwner Or effect.LimitDecision <> mLimitDecision Then NxRaiseContractError "Plan seal revalidation rejected effect owner or limit"
        If Len(effect.Target) = 0 Then NxRaiseContractError "Plan seal revalidation rejected empty target"
        snapshot.Add effect
        mPlanIdentity = mPlanIdentity & EncodeField(effect.StableIdentity)
    Next effect
    Set mSideEffects = snapshot
    mSealDigest = "NXSEAL1|" & EncodeField(mPlanIdentity) & EncodeField(mOwner) & EncodeField(NxInvariantUnsigned(CDbl(mLimitDecision))) & EncodeField(mSourceDigest)
    mIsSealed = True
End Sub

Public Property Get FeatureId() As String: FeatureId = mFeatureId: End Property
Public Property Get PlanIdentity() As String: PlanIdentity = mPlanIdentity: End Property
Public Property Get ContextFingerprint() As String: ContextFingerprint = mContextFingerprint: End Property
Public Property Get SealDigest() As String: SealDigest = mSealDigest: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property
Public Property Get Owner() As String: Owner = mOwner: End Property
Public Property Get SourceDigest() As String: SourceDigest = mSourceDigest: End Property
Public Property Get LimitDecision() As Long: LimitDecision = mLimitDecision: End Property
Public Property Get DeclaredCapabilities() As Long: DeclaredCapabilities = mDeclaredCapabilities: End Property
Public Property Get HasObservedCellCount() As Boolean: HasObservedCellCount = mHasObservedCellCount: End Property
Public Property Get ObservedCellCount() As Double: ObservedCellCount = mObservedCellCount: End Property

Public Property Get SideEffects() As Collection
    Dim snapshot As New Collection
    Dim effect As Variant
    For Each effect In mSideEffects
        snapshot.Add effect
    Next effect
    Set SideEffects = snapshot
End Property

Private Function EncodeField(ByVal value As String) As String
    EncodeField = NxInvariantUnsigned(CDbl(Len(value))) & ":" & value
End Function

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed plan cannot be changed"
End Sub
