VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxOfficeCompatibility"
Option Explicit

Private Const NX_UNSUPPORTED_BEHAVIOR As String = "disable_feature"

Private mFeatureId As String
Private mState As NxCompatibilityState
Private mMissingCapabilities As Long
Private mFallbackStrategy As String
Private mBlockReason As String
Private mObservedOfficeVersion As String
Private mObservedBitness As Long
Private mConfigured As Boolean
Private mIsSealed As Boolean

Public Sub Configure( _
    ByVal featureId As String, _
    ByVal state As NxCompatibilityState, _
    ByVal missingCapabilities As Long, _
    ByVal fallbackStrategy As String, _
    ByVal blockReason As String, _
    ByVal observedOfficeVersion As String, _
    ByVal observedBitness As Long)

    EnsureMutable
    If mConfigured Then NxRaiseContractError "Office compatibility can be configured only once"
    If Len(Trim$(featureId)) = 0 Then NxRaiseContractError "Office compatibility requires a feature ID"
    If state < NxCompatibilityNative Or state > NxCompatibilityIncomplete Then NxRaiseContractError "Office compatibility state is outside the closed contract"
    If missingCapabilities < 0 Then NxRaiseContractError "Missing Office capabilities cannot be negative"
    If observedBitness <> 32 And observedBitness <> 64 Then NxRaiseContractError "Office bitness must be 32 or 64"
    If state = NxCompatibilityNative And (missingCapabilities <> 0 Or Len(blockReason) <> 0) Then NxRaiseContractError "Native Office compatibility cannot carry missing capabilities"
    If state = NxCompatibilityFallback And (missingCapabilities = 0 Or StrComp(Trim$(fallbackStrategy), NX_FALLBACK_FORMULA_LEGACY, vbBinaryCompare) <> 0) Then NxRaiseContractError "Office fallback requires the implemented formula_legacy strategy"
    If (state = NxCompatibilityUnavailable Or state = NxCompatibilityIncomplete) And Len(Trim$(blockReason)) = 0 Then NxRaiseContractError "Blocked Office compatibility requires a reason"

    mFeatureId = featureId
    mState = state
    mMissingCapabilities = missingCapabilities
    mFallbackStrategy = fallbackStrategy
    mBlockReason = blockReason
    mObservedOfficeVersion = observedOfficeVersion
    mObservedBitness = observedBitness
    mConfigured = True
End Sub

Public Sub Seal()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Office compatibility is not configured"
    mIsSealed = True
End Sub

Public Property Get FeatureId() As String: EnsureReadable: FeatureId = mFeatureId: End Property
Public Property Get State() As NxCompatibilityState: EnsureReadable: State = mState: End Property
Public Property Get MissingCapabilities() As Long: EnsureReadable: MissingCapabilities = mMissingCapabilities: End Property
Public Property Get FallbackStrategy() As String: EnsureReadable: FallbackStrategy = mFallbackStrategy: End Property
Public Property Get BlockReason() As String: EnsureReadable: BlockReason = mBlockReason: End Property
Public Property Get ObservedOfficeVersion() As String: EnsureReadable: ObservedOfficeVersion = mObservedOfficeVersion: End Property
Public Property Get ObservedBitness() As Long: EnsureReadable: ObservedBitness = mObservedBitness: End Property
Public Property Get UnsupportedBehavior() As String: EnsureReadable: UnsupportedBehavior = NX_UNSUPPORTED_BEHAVIOR: End Property
Public Property Get FallbackBehavior() As String: EnsureReadable: FallbackBehavior = mFallbackStrategy: End Property
Public Property Get CanExecute() As Boolean: EnsureReadable: CanExecute = (mState = NxCompatibilityNative Or mState = NxCompatibilityFallback): End Property
Public Property Get IsImplementedFallback() As Boolean: EnsureReadable: IsImplementedFallback = (mState = NxCompatibilityFallback And StrComp(mFallbackStrategy, NX_FALLBACK_FORMULA_LEGACY, vbBinaryCompare) = 0): End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed Office compatibility cannot be changed"
End Sub

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Office compatibility must be sealed before use"
End Sub
