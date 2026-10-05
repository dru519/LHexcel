Attribute VB_Name = "NxSymbolsRegistry"
Option Explicit

Public Sub NxSymbolsRegisterFeatures(ByVal registry As CNxFeatureRegistry)
    Dim definition As New CNxFeatureDefinition
    If registry Is Nothing Then NxRaiseContractError "기호표 Registry가 필요합니다."
    definition.Configure "NX-UTIL-SYMBOLS", "NX-CAT-SYMBOLS", "symbols", "symbols_description", _
        NxCapabilityRead Or NxCapabilityCellMutation, NxRiskL0, NxRiskL1, "symbols_result", 80, _
        NxOfficeCapabilityNone, "disable_feature", "disable_feature", NxExecutionFast, _
        False, False, 1, False, False, False
    NxConfigureGeneratedLaunch definition, "NX-UTIL-SYMBOLS"
    definition.Seal
    registry.RegisterFeature definition
End Sub
