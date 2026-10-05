VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxExecutionTicket"
Option Explicit

Private mDefinition As CNxFeatureDefinition
Private mCommand As INxFeatureCommand
Private mApproval As CNxApproval
Private mPlan As CNxPlan
Private mCompatibility As CNxOfficeCompatibility
Private mDecision As CNxExecutionDecision
Private mConsumed As Boolean
Private mConfigured As Boolean

Friend Sub Configure(ByVal definition As CNxFeatureDefinition, ByVal command As INxFeatureCommand, _
    ByVal approval As CNxApproval, ByVal decision As CNxExecutionDecision, _
    ByVal plan As CNxPlan, ByVal compatibility As CNxOfficeCompatibility)

    If mConfigured Then NxRaiseContractError "Execution ticket can be configured only once"
    If definition Is Nothing Or command Is Nothing Or approval Is Nothing Or decision Is Nothing Or plan Is Nothing Or compatibility Is Nothing Then NxRaiseContractError "Execution ticket requires every dependency"
    If Not definition.IsSealed Or Not decision.IsSealed Or Not plan.IsSealed Or Not compatibility.IsSealed Then NxRaiseContractError "Execution ticket requires sealed metadata"
    If definition.FeatureId <> command.FeatureId Then NxRaiseContractError "Execution ticket feature identity mismatch"
    If definition.FeatureId <> plan.FeatureId Or definition.FeatureId <> compatibility.FeatureId Then NxRaiseContractError "Execution ticket feature identity mismatch"
    If approval.Plan Is Nothing Or Not (approval.Plan Is plan) Then NxRaiseContractError "Execution ticket plan binding mismatch"
    If StrComp(definition.FallbackStrategy, compatibility.FallbackStrategy, vbBinaryCompare) <> 0 Then NxRaiseContractError "Execution ticket Office fallback binding mismatch"
    Set mDefinition = definition
    Set mCommand = command
    Set mApproval = approval
    Set mDecision = decision
    Set mPlan = plan
    Set mCompatibility = compatibility
    mConfigured = True
End Sub

Public Function Consume(ByVal expectedGrade As NxExecutionGrade, ByVal userConfirmed As Boolean) As CNxApproval
    EnsureConfigured
    If mConsumed Then NxRaiseContractError "Ticket can be consumed only once"
    ValidateExecutionBinding
    If mDecision.ResolvedGrade <> expectedGrade Then NxRaiseContractError "Execution ticket grade mismatch"
    If expectedGrade <> NxExecutionFast And Not userConfirmed Then NxRaiseContractError "Protected execution requires explicit confirmation"
    mConsumed = True
    Set Consume = mApproval
End Function

Public Property Get Definition() As CNxFeatureDefinition: EnsureConfigured: Set Definition = mDefinition: End Property
Public Property Get Command() As INxFeatureCommand: EnsureConfigured: Set Command = mCommand: End Property
Friend Property Get Approval() As CNxApproval: EnsureConfigured: Set Approval = mApproval: End Property
Public Property Get Decision() As CNxExecutionDecision: EnsureConfigured: Set Decision = mDecision: End Property
Public Property Get Plan() As CNxPlan: EnsureConfigured: Set Plan = mPlan: End Property
Public Property Get Compatibility() As CNxOfficeCompatibility: EnsureConfigured: Set Compatibility = mCompatibility: End Property
Public Property Get IsConsumed() As Boolean: IsConsumed = mConsumed: End Property

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Execution ticket is not configured"
End Sub

Private Sub ValidateExecutionBinding()
    If mDefinition.FeatureId <> mCommand.FeatureId Then NxRaiseContractError "Execution ticket feature identity mismatch"
    If mDefinition.FeatureId <> mPlan.FeatureId Or mDefinition.FeatureId <> mCompatibility.FeatureId Then NxRaiseContractError "Execution ticket feature identity mismatch"
    If Not mPlan.IsSealed Or Not mCompatibility.IsSealed Then NxRaiseContractError "Execution ticket metadata is no longer sealed"
    If Not (mApproval.Plan Is mPlan) Then NxRaiseContractError "Execution ticket plan binding mismatch"
    If StrComp(mDefinition.FallbackStrategy, mCompatibility.FallbackStrategy, vbBinaryCompare) <> 0 Then NxRaiseContractError "Execution ticket Office fallback binding mismatch"
    Select Case mCompatibility.State
        Case NxCompatibilityNative
            If mCompatibility.MissingCapabilities <> 0 Then NxRaiseContractError "Native Office compatibility has missing capabilities"
        Case NxCompatibilityFallback
            If Not mCompatibility.IsImplementedFallback Then NxRaiseContractError "Office fallback strategy is not implemented"
        Case NxCompatibilityUnavailable, NxCompatibilityIncomplete
            NxRaiseContractError "Office compatibility blocks execution"
        Case Else
            NxRaiseContractError "Office compatibility state is outside the closed contract"
    End Select
End Sub
