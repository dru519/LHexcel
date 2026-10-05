Attribute VB_Name = "NxExecutionPreviewFactory"
Option Explicit

Public Function NxCreateExecutionPreview(ByVal definition As CNxFeatureDefinition, ByVal plan As CNxPlan, _
    ByVal targetSummary As String, ByVal actionSummary As String, ByVal resultSummary As String, _
    ByVal scaleSummary As String, ByVal privacyState As NxPrivacyState, ByVal conflictSummary As String, _
    ByVal recoverySummary As String) As CNxExecutionPreview
    Set NxCreateExecutionPreview = NxCreateExecutionPreviewForCompatibility(definition, plan, targetSummary, actionSummary, _
        resultSummary, scaleSummary, privacyState, conflictSummary, recoverySummary, NxProbeOfficeCompatibility(definition))
End Function

Public Function NxCreateExecutionPreviewForCompatibility(ByVal definition As CNxFeatureDefinition, ByVal plan As CNxPlan, _
    ByVal targetSummary As String, ByVal actionSummary As String, ByVal resultSummary As String, _
    ByVal scaleSummary As String, ByVal privacyState As NxPrivacyState, ByVal conflictSummary As String, _
    ByVal recoverySummary As String, ByVal compatibility As CNxOfficeCompatibility) As CNxExecutionPreview
    Dim preview As New CNxExecutionPreview, risk As NxRiskLevel, capabilities As Long
    Dim canExecute As Boolean, approvalRequired As Boolean, blockReason As String
    If definition Is Nothing Then NxRaiseContractError "Preview requires a feature definition"
    If plan Is Nothing Then NxRaiseContractError "Preview requires a plan"
    If compatibility Is Nothing Then NxRaiseContractError "Preview requires Office compatibility"
    If Not definition.IsSealed Or Not plan.IsSealed Or Not compatibility.IsSealed Then NxRaiseContractError "Preview inputs must be sealed"
    If definition.FeatureId <> plan.FeatureId Or definition.FeatureId <> compatibility.FeatureId Then NxRaiseContractError "Preview identity mismatch"
    If privacyState < NxPrivacySafe Or privacyState > NxPrivacyIncomplete Then NxRaiseContractError "Privacy state is outside the closed contract"
    capabilities = CapabilityForPlan(plan)
    If (capabilities And Not definition.CapabilityMask) <> 0 Then NxRaiseContractError "Plan capability exceeds feature declaration"
    risk = RiskForPlan(plan): If definition.BaseRisk > risk Then risk = definition.BaseRisk
    If privacyState = NxPrivacyCaution And (capabilities And (NxCapabilityExternalProcess Or NxCapabilityBrowserOpen Or NxCapabilitySensitiveOutput)) <> 0 Then
        If risk < NxRiskL3 Then risk = NxRiskL3
    End If
    If risk > definition.MaximumRisk Then NxRaiseContractError "Computed risk exceeds feature maximum risk"
    canExecute = True
    If privacyState = NxPrivacyBlocked Then canExecute = False: blockReason = "Privacy blocked"
    If privacyState = NxPrivacyIncomplete Then canExecute = False: blockReason = "Privacy scan incomplete"
    If compatibility.State = NxCompatibilityUnavailable Then canExecute = False: blockReason = compatibility.BlockReason
    If compatibility.State = NxCompatibilityIncomplete Then canExecute = False: blockReason = compatibility.BlockReason
    If risk = NxRiskL4 Then canExecute = False: blockReason = "L4 is blocked in v0.01"
    approvalRequired = (risk <> NxRiskL0)
    preview.Configure definition.FeatureId, definition.CategoryId, definition.LabelKey, targetSummary, actionSummary, resultSummary, _
        scaleSummary, privacyState, conflictSummary, recoverySummary, risk, canExecute, approvalRequired, blockReason
    preview.Seal
    Set NxCreateExecutionPreviewForCompatibility = preview
End Function
