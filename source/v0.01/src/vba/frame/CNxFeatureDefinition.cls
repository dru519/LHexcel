VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFeatureDefinition"
Option Explicit

Private Const NX_KNOWN_CAPABILITIES As Long = 262143
Private Const NX_KNOWN_OFFICE_CAPABILITIES As Long = 63

Private mFeatureId As String
Private mCategoryId As String
Private mLabelKey As String
Private mDescriptionKey As String
Private mCapabilityMask As Long
Private mBaseRisk As NxRiskLevel
Private mMaximumRisk As NxRiskLevel
Private mResultContract As String
Private mSortOrder As Long
Private mRequiredOfficeCapabilities As Long
Private mFallbackStrategy As String
Private mUnsupportedBehavior As String
Private mExecutionGrade As NxExecutionGrade
Private mAiProcessing As Boolean
Private mExternalDataTransfer As Boolean
Private mNetworkTransfer As Boolean
Private mLocalProcessTransfer As Boolean
Private mLocalFileCreation As Boolean
Private mPlannedCellThreshold As Long
Private mLaunchSurface As String
Private mDialogId As String
Private mDialogTemplate As String
Private mDialogVariant As String
Private mStateResolver As String
Private mSourceAssetId As String
Private mReuseMode As String
Private mConfigured As Boolean
Private mLaunchConfigured As Boolean
Private mIsSealed As Boolean

Public Sub Configure( _
    ByVal featureId As String, _
    ByVal categoryId As String, _
    ByVal labelKey As String, _
    ByVal descriptionKey As String, _
    ByVal capabilityMask As Long, _
    ByVal baseRisk As NxRiskLevel, _
    ByVal maximumRisk As NxRiskLevel, _
    ByVal resultContract As String, _
    ByVal sortOrder As Long, _
    ByVal requiredOfficeCapabilities As Long, _
    ByVal fallbackStrategy As String, _
    ByVal unsupportedBehavior As String, _
    ByVal executionGrade As NxExecutionGrade, _
    ByVal aiProcessing As Boolean, _
    ByVal externalDataTransfer As Boolean, _
    ByVal plannedCellThreshold As Long, _
    Optional ByVal networkTransfer As Boolean = False, _
    Optional ByVal localProcessTransfer As Boolean = False, _
    Optional ByVal localFileCreation As Boolean = False)

    EnsureMutable
    If mConfigured Then NxRaiseContractError "Feature definition can be configured only once"
    If Len(Trim$(featureId)) = 0 Or Len(Trim$(categoryId)) = 0 Then NxRaiseContractError "Feature definition requires feature and category IDs"
    If Len(Trim$(labelKey)) = 0 Or Len(Trim$(descriptionKey)) = 0 Then NxRaiseContractError "Feature definition requires label and description keys"
    If Len(Trim$(resultContract)) = 0 Then NxRaiseContractError "Feature definition requires a result contract"
    If sortOrder <= 0 Then NxRaiseContractError "Feature sort order must be positive"
    ValidateCapabilities capabilityMask
    ValidateRisk baseRisk, maximumRisk
    ValidateOfficeCapabilities requiredOfficeCapabilities
    ValidateExecutionMetadata executionGrade, plannedCellThreshold
    If unsupportedBehavior <> "disable_feature" Then NxRaiseContractError "Unsupported Office behavior must disable only the feature"

    mFeatureId = featureId
    mCategoryId = categoryId
    mLabelKey = labelKey
    mDescriptionKey = descriptionKey
    mCapabilityMask = capabilityMask
    mBaseRisk = baseRisk
    mMaximumRisk = maximumRisk
    mResultContract = resultContract
    mSortOrder = sortOrder
    mRequiredOfficeCapabilities = requiredOfficeCapabilities
    mFallbackStrategy = fallbackStrategy
    mUnsupportedBehavior = unsupportedBehavior
    mExecutionGrade = executionGrade
    mAiProcessing = aiProcessing
    mExternalDataTransfer = externalDataTransfer
    mNetworkTransfer = externalDataTransfer Or networkTransfer
    mLocalProcessTransfer = localProcessTransfer
    mLocalFileCreation = localFileCreation
    mPlannedCellThreshold = plannedCellThreshold
    mConfigured = True
End Sub

Public Sub ConfigureLaunch( _
    ByVal launchSurface As String, _
    ByVal dialogId As String, _
    ByVal dialogTemplate As String, _
    ByVal dialogVariant As String, _
    ByVal stateResolver As String, _
    ByVal sourceAssetId As String, _
    ByVal reuseMode As String)

    EnsureMutable
    If mLaunchConfigured Then NxRaiseContractError "Feature launch metadata can be configured only once"
    ValidateLaunchMetadata launchSurface, dialogId, dialogTemplate, dialogVariant, stateResolver
    ValidateReuseMode reuseMode
    mLaunchSurface = launchSurface
    mDialogId = dialogId
    mDialogTemplate = dialogTemplate
    mDialogVariant = dialogVariant
    mStateResolver = stateResolver
    mSourceAssetId = sourceAssetId
    mReuseMode = reuseMode
    mLaunchConfigured = True
End Sub

Public Sub Seal()
    EnsureMutable
    If Not (mConfigured And mLaunchConfigured) Then NxRaiseContractError "Feature definition is not fully configured"
    mIsSealed = True
End Sub

Public Property Get FeatureId() As String: EnsureReadable: FeatureId = mFeatureId: End Property
Public Property Get CategoryId() As String: EnsureReadable: CategoryId = mCategoryId: End Property
Public Property Get LabelKey() As String: EnsureReadable: LabelKey = mLabelKey: End Property
Public Property Get DescriptionKey() As String: EnsureReadable: DescriptionKey = mDescriptionKey: End Property
Public Property Get CapabilityMask() As Long: EnsureReadable: CapabilityMask = mCapabilityMask: End Property
Public Property Get BaseRisk() As NxRiskLevel: EnsureReadable: BaseRisk = mBaseRisk: End Property
Public Property Get MaximumRisk() As NxRiskLevel: EnsureReadable: MaximumRisk = mMaximumRisk: End Property
Public Property Get ResultContract() As String: EnsureReadable: ResultContract = mResultContract: End Property
Public Property Get SortOrder() As Long: EnsureReadable: SortOrder = mSortOrder: End Property
Public Property Get RequiredOfficeCapabilities() As Long: EnsureReadable: RequiredOfficeCapabilities = mRequiredOfficeCapabilities: End Property
Public Property Get FallbackStrategy() As String: EnsureReadable: FallbackStrategy = mFallbackStrategy: End Property
Public Property Get UnsupportedBehavior() As String: EnsureReadable: UnsupportedBehavior = mUnsupportedBehavior: End Property
Public Property Get ExecutionGrade() As NxExecutionGrade: EnsureReadable: ExecutionGrade = mExecutionGrade: End Property
Public Property Get AiProcessing() As Boolean: EnsureReadable: AiProcessing = mAiProcessing: End Property
Public Property Get ExternalDataTransfer() As Boolean: EnsureReadable: ExternalDataTransfer = mExternalDataTransfer: End Property
Public Property Get NetworkTransfer() As Boolean: EnsureReadable: NetworkTransfer = mNetworkTransfer: End Property
Public Property Get LocalProcessTransfer() As Boolean: EnsureReadable: LocalProcessTransfer = mLocalProcessTransfer: End Property
Public Property Get LocalFileCreation() As Boolean: EnsureReadable: LocalFileCreation = mLocalFileCreation: End Property
Public Property Get RequiresPrivacyInspection() As Boolean: EnsureReadable: RequiresPrivacyInspection = mAiProcessing Or mNetworkTransfer: End Property
Public Property Get PlannedCellThreshold() As Long: EnsureReadable: PlannedCellThreshold = mPlannedCellThreshold: End Property
Public Property Get LaunchSurface() As String: EnsureReadable: LaunchSurface = mLaunchSurface: End Property
Public Property Get DialogId() As String: EnsureReadable: DialogId = mDialogId: End Property
Public Property Get DialogTemplate() As String: EnsureReadable: DialogTemplate = mDialogTemplate: End Property
Public Property Get DialogVariant() As String: EnsureReadable: DialogVariant = mDialogVariant: End Property
Public Property Get StateResolver() As String: EnsureReadable: StateResolver = mStateResolver: End Property
Public Property Get SourceAssetId() As String: EnsureReadable: SourceAssetId = mSourceAssetId: End Property
Public Property Get ReuseMode() As String: EnsureReadable: ReuseMode = mReuseMode: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property

Private Sub ValidateCapabilities(ByVal capabilityMask As Long)
    If capabilityMask < 0 Or (capabilityMask And Not NX_KNOWN_CAPABILITIES) <> 0 Then NxRaiseContractError "Unknown feature capability"
End Sub

Private Sub ValidateOfficeCapabilities(ByVal capabilityMask As Long)
    If capabilityMask < 0 Or (capabilityMask And Not NX_KNOWN_OFFICE_CAPABILITIES) <> 0 Then NxRaiseContractError "Unknown Office capability"
End Sub

Private Sub ValidateRisk(ByVal baseRisk As NxRiskLevel, ByVal maximumRisk As NxRiskLevel)
    If baseRisk < NxRiskL0 Or baseRisk > NxRiskL4 Then NxRaiseContractError "Feature base risk is outside the closed contract"
    If maximumRisk < NxRiskL0 Or maximumRisk > NxRiskL4 Then NxRaiseContractError "Feature maximum risk is outside the closed contract"
    If baseRisk > maximumRisk Then NxRaiseContractError "Feature base risk exceeds maximum risk"
End Sub

Private Sub ValidateExecutionMetadata(ByVal executionGrade As NxExecutionGrade, ByVal plannedCellThreshold As Long)
    If executionGrade < NxExecutionFast Or executionGrade > NxExecutionPlanned Then NxRaiseContractError "Execution grade is outside the closed contract"
    If plannedCellThreshold <= 0 Or plannedCellThreshold > 10000 Then NxRaiseContractError "Planned cell threshold must be between 1 and 10000"
End Sub

Private Sub ValidateLaunchMetadata(ByVal launchSurface As String, ByVal dialogId As String, _
    ByVal dialogTemplate As String, ByVal dialogVariant As String, ByVal stateResolver As String)

    Select Case launchSurface
        Case "direct"
            If Len(dialogId) > 0 Or Len(dialogTemplate) > 0 Or Len(dialogVariant) > 0 Or Len(stateResolver) > 0 Then
                NxRaiseContractError "Direct feature cannot declare dialog metadata"
            End If
        Case "dialog"
            If Len(dialogId) = 0 Or Len(dialogTemplate) = 0 Or Len(dialogVariant) = 0 Or Len(stateResolver) > 0 Then
                NxRaiseContractError "Dialog feature metadata is incomplete"
            End If
        Case "hybrid"
            If Len(dialogId) = 0 Or Len(dialogTemplate) = 0 Or Len(dialogVariant) = 0 Or Len(stateResolver) = 0 Then
                NxRaiseContractError "Hybrid feature metadata is incomplete"
            End If
        Case "workbench"
            If Len(dialogId) = 0 Or Len(dialogTemplate) = 0 Or Len(dialogVariant) = 0 Or Len(stateResolver) > 0 Then
                NxRaiseContractError "Workbench feature metadata is incomplete"
            End If
        Case Else
            NxRaiseContractError "Unknown feature launch surface"
    End Select
End Sub

Private Sub ValidateReuseMode(ByVal reuseMode As String)
    Select Case reuseMode
        Case "direct", "adapt", "spec-only", "clean-room"
        Case Else: NxRaiseContractError "Unknown active feature reuse mode"
    End Select
End Sub

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed definition cannot be changed"
End Sub

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Feature definition must be sealed before use"
End Sub
