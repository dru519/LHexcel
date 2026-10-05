VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxAiPrivacySnapshot"
Option Explicit

Private mState As NxPrivacyState
Private mOriginalCopy As String
Private mMaskedCopy As String
Private mSummary As String
Private mConfigured As Boolean
Private mIsSealed As Boolean

Public Sub Configure(ByVal state As NxPrivacyState, ByVal originalCopy As String, _
    ByVal maskedCopy As String, ByVal summary As String)

    EnsureMutable
    If mConfigured Then NxRaiseContractError "Privacy snapshot can be configured only once"
    If state < NxPrivacySafe Or state > NxPrivacyIncomplete Then NxRaiseContractError "Privacy snapshot state is outside the closed contract"
    If Len(Trim$(originalCopy)) = 0 Then NxRaiseContractError "Privacy snapshot requires an input copy"
    If Len(Trim$(maskedCopy)) = 0 Then NxRaiseContractError "Privacy snapshot requires a protected copy"
    If Len(Trim$(summary)) = 0 Then NxRaiseContractError "Privacy snapshot requires a summary"
    If state = NxPrivacySafe And StrComp(originalCopy, maskedCopy, vbBinaryCompare) <> 0 Then NxRaiseContractError "Safe privacy copies must be identical"
    If (state = NxPrivacyCaution Or state = NxPrivacyBlocked) And StrComp(originalCopy, maskedCopy, vbBinaryCompare) = 0 Then NxRaiseContractError "Detected privacy copies must differ"

    mState = state
    mOriginalCopy = originalCopy
    mMaskedCopy = maskedCopy
    mSummary = summary
    mConfigured = True
End Sub

Public Sub Seal()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Privacy snapshot is not configured"
    mIsSealed = True
End Sub

Public Property Get State() As NxPrivacyState: EnsureReadable: State = mState: End Property
Public Property Get OriginalCopy() As String: EnsureReadable: OriginalCopy = mOriginalCopy: End Property
Public Property Get MaskedCopy() As String: EnsureReadable: MaskedCopy = mMaskedCopy: End Property
Public Property Get Summary() As String: EnsureReadable: Summary = mSummary: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed privacy snapshot cannot be changed"
End Sub

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Privacy snapshot must be sealed before use"
End Sub
