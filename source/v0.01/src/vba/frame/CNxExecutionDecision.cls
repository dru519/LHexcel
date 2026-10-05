VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxExecutionDecision"
Option Explicit

Private mBaseGrade As NxExecutionGrade
Private mResolvedGrade As NxExecutionGrade
Private mReasonCode As String
Private mActualCellCount As Double
Private mRequiresPrivacyInspection As Boolean
Private mConfigured As Boolean
Private mIsSealed As Boolean

Public Sub Configure(ByVal baseGrade As NxExecutionGrade, ByVal resolvedGrade As NxExecutionGrade, _
    ByVal reasonCode As String, ByVal actualCellCount As Double, ByVal requiresPrivacyInspection As Boolean)

    EnsureMutable
    If mConfigured Then NxRaiseContractError "Execution decision can be configured only once"
    If baseGrade < NxExecutionFast Or baseGrade > NxExecutionPlanned Then NxRaiseContractError "Execution decision base grade is invalid"
    If resolvedGrade < baseGrade Or resolvedGrade > NxExecutionPlanned Then NxRaiseContractError "Execution decision cannot downgrade"
    If Len(Trim$(reasonCode)) = 0 Then NxRaiseContractError "Execution decision requires a reason code"
    If actualCellCount < 0 Or actualCellCount <> Fix(actualCellCount) Then NxRaiseContractError "Actual cell count must be an unsigned integer"
    mBaseGrade = baseGrade
    mResolvedGrade = resolvedGrade
    mReasonCode = reasonCode
    mActualCellCount = actualCellCount
    mRequiresPrivacyInspection = requiresPrivacyInspection
    mConfigured = True
End Sub

Public Sub Seal()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Execution decision is not configured"
    mIsSealed = True
End Sub

Public Property Get BaseGrade() As NxExecutionGrade: EnsureReadable: BaseGrade = mBaseGrade: End Property
Public Property Get ResolvedGrade() As NxExecutionGrade: EnsureReadable: ResolvedGrade = mResolvedGrade: End Property
Public Property Get ReasonCode() As String: EnsureReadable: ReasonCode = mReasonCode: End Property
Public Property Get ActualCellCount() As Double: EnsureReadable: ActualCellCount = mActualCellCount: End Property
Public Property Get RequiresPrivacyInspection() As Boolean: EnsureReadable: RequiresPrivacyInspection = mRequiresPrivacyInspection: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed execution decision cannot be changed"
End Sub

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Execution decision must be sealed before use"
End Sub
