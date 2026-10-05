VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxExecutionPreview"
Option Explicit

Private mFeatureId As String, mCategoryId As String, mLabelKey As String
Private mTargetSummary As String, mActionSummary As String, mResultSummary As String
Private mScaleSummary As String, mConflictSummary As String, mRecoverySummary As String
Private mPrivacyState As NxPrivacyState, mRisk As NxRiskLevel
Private mCanExecute As Boolean, mApprovalRequired As Boolean, mBlockReason As String
Private mConfigured As Boolean, mIsSealed As Boolean

Friend Sub Configure(ByVal featureId As String, ByVal categoryId As String, ByVal labelKey As String, _
    ByVal targetSummary As String, ByVal actionSummary As String, ByVal resultSummary As String, _
    ByVal scaleSummary As String, ByVal privacyState As NxPrivacyState, ByVal conflictSummary As String, _
    ByVal recoverySummary As String, ByVal risk As NxRiskLevel, ByVal canExecute As Boolean, _
    ByVal approvalRequired As Boolean, ByVal blockReason As String)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Execution preview can be configured only once"
    If Len(featureId) = 0 Or Len(categoryId) = 0 Or Len(labelKey) = 0 Then NxRaiseContractError "Execution preview requires identity"
    If Len(targetSummary) = 0 Or Len(actionSummary) = 0 Or Len(resultSummary) = 0 Or Len(scaleSummary) = 0 Or Len(conflictSummary) = 0 Or Len(recoverySummary) = 0 Then NxRaiseContractError "Execution preview requires all summaries"
    If privacyState < NxPrivacySafe Or privacyState > NxPrivacyIncomplete Then NxRaiseContractError "Execution preview privacy state is outside the closed contract"
    If risk < NxRiskL0 Or risk > NxRiskL4 Then NxRaiseContractError "Execution preview risk is outside the closed contract"
    mFeatureId = featureId: mCategoryId = categoryId: mLabelKey = labelKey
    mTargetSummary = targetSummary: mActionSummary = actionSummary: mResultSummary = resultSummary
    mScaleSummary = scaleSummary: mPrivacyState = privacyState: mConflictSummary = conflictSummary
    mRecoverySummary = recoverySummary: mRisk = risk: mCanExecute = canExecute
    mApprovalRequired = approvalRequired: mBlockReason = blockReason: mConfigured = True
End Sub

Public Sub Seal()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Execution preview is not configured"
    mIsSealed = True
End Sub
Public Property Get FeatureId() As String: EnsureReadable: FeatureId = mFeatureId: End Property
Public Property Get CategoryId() As String: EnsureReadable: CategoryId = mCategoryId: End Property
Public Property Get LabelKey() As String: EnsureReadable: LabelKey = mLabelKey: End Property
Public Property Get TargetSummary() As String: EnsureReadable: TargetSummary = mTargetSummary: End Property
Public Property Get ActionSummary() As String: EnsureReadable: ActionSummary = mActionSummary: End Property
Public Property Get ResultSummary() As String: EnsureReadable: ResultSummary = mResultSummary: End Property
Public Property Get ScaleSummary() As String: EnsureReadable: ScaleSummary = mScaleSummary: End Property
Public Property Get ConflictSummary() As String: EnsureReadable: ConflictSummary = mConflictSummary: End Property
Public Property Get RecoverySummary() As String: EnsureReadable: RecoverySummary = mRecoverySummary: End Property
Public Property Get Risk() As NxRiskLevel: EnsureReadable: Risk = mRisk: End Property
Public Property Get PrivacyState() As NxPrivacyState: EnsureReadable: PrivacyState = mPrivacyState: End Property
Public Property Get CanExecute() As Boolean: EnsureReadable: CanExecute = mCanExecute: End Property
Public Property Get ApprovalRequired() As Boolean: EnsureReadable: ApprovalRequired = mApprovalRequired: End Property
Public Property Get BlockReason() As String: EnsureReadable: BlockReason = mBlockReason: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property
Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed execution preview cannot be changed"
End Sub

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Execution preview must be sealed before use"
End Sub
