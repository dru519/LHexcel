VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxExecutionPlanViewModel"
Option Explicit

Private mFeatureId As String, mCategoryId As String, mLabelKey As String
Private mTargetSummary As String, mActionSummary As String, mResultSummary As String
Private mScaleSummary As String, mPrivacyLabel As String, mRiskLabel As String
Private mConflictSummary As String, mRecoverySummary As String
Private mExecuteCaption As String, mBlockReason As String
Private mRisk As NxRiskLevel, mPrivacyState As NxPrivacyState
Private mCanExecute As Boolean, mApprovalRequired As Boolean
Private mConfigured As Boolean, mIsSealed As Boolean

Friend Sub Configure(ByVal preview As CNxExecutionPreview, ByVal riskLabel As String, _
    ByVal privacyLabel As String, ByVal executeCaption As String)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Execution plan view model can be configured only once"
    If preview Is Nothing Then NxRaiseContractError "Execution plan view model requires a preview"
    If Not preview.IsSealed Then NxRaiseContractError "Execution plan view model requires a sealed preview"
    If Len(Trim$(riskLabel)) = 0 Or Len(Trim$(privacyLabel)) = 0 Or Len(Trim$(executeCaption)) = 0 Then
        NxRaiseContractError "Execution plan view model requires display labels"
    End If
    mFeatureId = preview.FeatureId
    mCategoryId = preview.CategoryId
    mLabelKey = preview.LabelKey
    mTargetSummary = preview.TargetSummary
    mActionSummary = preview.ActionSummary
    mResultSummary = preview.ResultSummary
    mScaleSummary = preview.ScaleSummary
    mPrivacyLabel = privacyLabel
    mRiskLabel = riskLabel
    mConflictSummary = preview.ConflictSummary
    mRecoverySummary = preview.RecoverySummary
    mExecuteCaption = executeCaption
    mBlockReason = preview.BlockReason
    mRisk = preview.Risk
    mPrivacyState = preview.PrivacyState
    mCanExecute = preview.CanExecute And preview.Risk <> NxRiskL4
    mApprovalRequired = preview.ApprovalRequired
    If Not mCanExecute And Len(Trim$(mBlockReason)) = 0 Then
        NxRaiseContractError "Blocked execution plan requires a reason"
    End If
    mConfigured = True
End Sub

Public Sub Seal()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Execution plan view model is not configured"
    mIsSealed = True
End Sub

Public Property Get FeatureId() As String: EnsureReadable: FeatureId = mFeatureId: End Property
Public Property Get CategoryId() As String: EnsureReadable: CategoryId = mCategoryId: End Property
Public Property Get LabelKey() As String: EnsureReadable: LabelKey = mLabelKey: End Property
Public Property Get TargetSummary() As String: EnsureReadable: TargetSummary = mTargetSummary: End Property
Public Property Get ActionSummary() As String: EnsureReadable: ActionSummary = mActionSummary: End Property
Public Property Get ResultSummary() As String: EnsureReadable: ResultSummary = mResultSummary: End Property
Public Property Get ScaleSummary() As String: EnsureReadable: ScaleSummary = mScaleSummary: End Property
Public Property Get PrivacyLabel() As String: EnsureReadable: PrivacyLabel = mPrivacyLabel: End Property
Public Property Get RiskLabel() As String: EnsureReadable: RiskLabel = mRiskLabel: End Property
Public Property Get ConflictSummary() As String: EnsureReadable: ConflictSummary = mConflictSummary: End Property
Public Property Get RecoverySummary() As String: EnsureReadable: RecoverySummary = mRecoverySummary: End Property
Public Property Get ExecuteCaption() As String: EnsureReadable: ExecuteCaption = mExecuteCaption: End Property
Public Property Get BlockReason() As String: EnsureReadable: BlockReason = mBlockReason: End Property
Public Property Get Risk() As NxRiskLevel: EnsureReadable: Risk = mRisk: End Property
Public Property Get PrivacyState() As NxPrivacyState: EnsureReadable: PrivacyState = mPrivacyState: End Property
Public Property Get CanExecute() As Boolean: EnsureReadable: CanExecute = mCanExecute: End Property
Public Property Get ApprovalRequired() As Boolean: EnsureReadable: ApprovalRequired = mApprovalRequired: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property

Public Property Get DisplayCategoryLabel() As String
    Dim registry As CNxFeatureRegistry, label As String
    EnsureReadable
    DisplayCategoryLabel = "내엑셀"
    On Error GoTo MissingLabel
    Set registry = NxRegistryFactory.CreateBaseRegistry()
    label = registry.CategoryById(mCategoryId).Label
    If Len(Trim$(label)) > 0 Then DisplayCategoryLabel = label
    Exit Property
MissingLabel:
    Err.Clear
End Property

Public Property Get DisplayFeatureName() As String
    Dim label As String
    EnsureReadable
    DisplayFeatureName = "선택한 기능"
    On Error GoTo MissingLabel
    label = NxGeneratedFeatureLabel(mFeatureId)
    If Len(Trim$(label)) > 0 Then DisplayFeatureName = label
    Exit Property
MissingLabel:
    Err.Clear
End Property

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed execution plan view model cannot be changed"
End Sub

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Execution plan view model must be sealed before use"
End Sub
