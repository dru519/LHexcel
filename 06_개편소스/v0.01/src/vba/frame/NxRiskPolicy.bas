Attribute VB_Name = "NxRiskPolicy"
Option Explicit

Public Function RiskForPlan(ByVal plan As CNxPlan) As NxRiskLevel
    Dim effect As CNxSideEffect
    Dim current As NxRiskLevel
    If plan Is Nothing Then NxRaiseContractError "Risk policy requires a plan"
    If Not plan.IsSealed Then NxRaiseContractError "Risk policy requires a sealed plan"
    If (plan.DeclaredCapabilities And NxCapabilityTemplateQuarantine) <> 0 Then RiskForPlan = NxRiskL2
    If (plan.DeclaredCapabilities And NxCapabilityRiskyFormula) <> 0 Then RiskForPlan = NxRiskL2
    For Each effect In plan.SideEffects
        current = RiskForEffect(effect.Kind)
        If current > RiskForPlan Then RiskForPlan = current
    Next effect
End Function

Public Function CapabilityForPlan(ByVal plan As CNxPlan) As Long
    Dim effect As CNxSideEffect
    If plan Is Nothing Then NxRaiseContractError "Capability policy requires a plan"
    If Not plan.IsSealed Then NxRaiseContractError "Capability policy requires a sealed plan"
    CapabilityForPlan = plan.DeclaredCapabilities
    For Each effect In plan.SideEffects
        CapabilityForPlan = CapabilityForPlan Or CapabilityForEffect(effect.Kind)
    Next effect
    If plan.SideEffects.count = 0 And plan.DeclaredCapabilities = 0 Then CapabilityForPlan = NxCapabilityRead
End Function

Public Function RiskForEffect(ByVal kind As NxSideEffectKind) As NxRiskLevel
    Select Case kind
        Case NxCellMutation, NxColumnInsert: RiskForEffect = NxRiskL1
        Case NxWorksheetCreate: RiskForEffect = NxRiskL1
        Case NxFileCreate, NxClipboardWrite, NxShapeInsert, NxShapeMutation, NxDirectoryBatchCreate: RiskForEffect = NxRiskL2
        Case NxBinaryPackageWrite: RiskForEffect = NxRiskL2
        Case NxExternalProcess, NxBrowserOpen: RiskForEffect = NxRiskL3
        Case NxFileReplace: RiskForEffect = NxRiskL4
        Case Else: NxRaiseContractError "Unknown side effect risk"
    End Select
End Function

Private Function CapabilityForEffect(ByVal kind As NxSideEffectKind) As Long
    Select Case kind
        Case NxCellMutation, NxColumnInsert: CapabilityForEffect = NxCapabilityCellMutation
        Case NxFileCreate: CapabilityForEffect = NxCapabilityFileCreate
        Case NxFileReplace: CapabilityForEffect = NxCapabilityFileReplace
        Case NxClipboardWrite: CapabilityForEffect = NxCapabilityClipboardWrite
        Case NxExternalProcess: CapabilityForEffect = NxCapabilityExternalProcess
        Case NxBrowserOpen: CapabilityForEffect = NxCapabilityBrowserOpen
        Case NxShapeInsert: CapabilityForEffect = NxCapabilityShapeInsert
        Case NxShapeMutation: CapabilityForEffect = NxCapabilityShapeInsert
        Case NxBinaryPackageWrite: CapabilityForEffect = NxCapabilityBinaryPackageWrite
        Case NxWorksheetCreate: CapabilityForEffect = NxCapabilityWorksheetCreate
        Case NxDirectoryBatchCreate: CapabilityForEffect = NxCapabilityDirectoryCreate
        Case Else: NxRaiseContractError "Unknown side effect capability"
    End Select
End Function
