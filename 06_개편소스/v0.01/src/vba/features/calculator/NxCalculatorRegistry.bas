Attribute VB_Name = "NxCalculatorRegistry"
Option Explicit

Public Sub NxCalculatorRegisterFeatures(ByVal registry As CNxFeatureRegistry)
    Dim definition As New CNxFeatureDefinition
    If registry Is Nothing Then NxRaiseContractError "계산기 Registry가 필요합니다."
    definition.Configure "NX-UTIL-CALCULATOR", "NX-CAT-CALCULATOR", "calculator", "calculator_description", _
        NxCapabilityRead Or NxCapabilityCellMutation, NxRiskL0, NxRiskL2, "calculator_result", 70, _
        NxOfficeCapabilityNone, "disable_feature", "disable_feature", NxExecutionFast, _
        False, False, 1, False, False, False
    NxConfigureGeneratedLaunch definition, "NX-UTIL-CALCULATOR"
    definition.Seal
    registry.RegisterFeature definition
End Sub
