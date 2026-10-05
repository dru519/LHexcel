Attribute VB_Name = "NxNavigatorRegistry"
Option Explicit

Public Sub NxNavigatorRegisterFeatures(ByVal registry As CNxFeatureRegistry)
    Dim definition As New CNxFeatureDefinition
    Dim documentDefinition As New CNxFeatureDefinition
    If registry Is Nothing Then NxRaiseContractError "탐색창 Registry가 필요합니다."
    definition.Configure "NX-UTIL-NAVIGATOR", "NX-CAT-UTIL", "navigator", "navigator_description", _
        NxCapabilityRead, NxRiskL0, NxRiskL0, "navigator_result", 90, _
        NxOfficeCapabilityNone, "disable_feature", "disable_feature", NxExecutionFast, _
        False, False, 1, False, False, False
    NxConfigureGeneratedLaunch definition, "NX-UTIL-NAVIGATOR"
    definition.Seal
    registry.RegisterFeature definition
    documentDefinition.Configure "NX-UTIL-DOCUMENT-NAVIGATOR", "NX-CAT-UTIL", "document_navigator", "document_navigator_description", _
        NxCapabilityRead, NxRiskL0, NxRiskL0, "document_navigator_result", 91, _
        NxOfficeCapabilityNone, "disable_feature", "disable_feature", NxExecutionFast, _
        False, False, 1, False, False, False
    NxConfigureGeneratedLaunch documentDefinition, "NX-UTIL-DOCUMENT-NAVIGATOR"
    documentDefinition.Seal
    registry.RegisterFeature documentDefinition
End Sub
