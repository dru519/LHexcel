Attribute VB_Name = "NxTemplateRegistry"
Option Explicit

Public Sub NxTemplateRegisterFeatures(ByVal registry As CNxFeatureRegistry)
    If registry Is Nothing Then NxRaiseContractError "Template registry is required"
    RegisterTemplateFeature registry, "NX-TPL-REGISTER-SHEET", NxExecutionPlanned, NxCapabilityRead Or NxCapabilityBinaryPackageWrite Or NxCapabilityRiskyFormula, 20
    RegisterTemplateFeature registry, "NX-TPL-LIST", NxExecutionFast, NxCapabilityRead, 30
    RegisterTemplateFeature registry, "NX-TPL-LOAD", NxExecutionGuarded, NxCapabilityRead Or NxCapabilityWorksheetCreate Or NxCapabilityRiskyFormula Or NxCapabilityBinaryPackageWrite, 40
    RegisterTemplateFeature registry, "NX-TPL-RENAME", NxExecutionPlanned, NxCapabilityBinaryPackageWrite, 50
    RegisterTemplateFeature registry, "NX-TPL-DELETE", NxExecutionPlanned, NxCapabilityBinaryPackageWrite Or NxCapabilityTemplateQuarantine, 60
End Sub

Public Function NxTemplateFeatureDefinition(ByVal featureId As String) As CNxFeatureDefinition
    Dim registry As CNxFeatureRegistry
    Set registry = CreateBaseRegistry()
    NxTemplateRegisterFeatures registry
    Set NxTemplateFeatureDefinition = registry.FeatureById(featureId)
End Function

Private Sub RegisterTemplateFeature(ByVal registry As CNxFeatureRegistry, ByVal featureId As String, _
    ByVal grade As NxExecutionGrade, ByVal capabilities As Long, ByVal sortOrder As Long)

    Dim definition As New CNxFeatureDefinition, labelKey As String
    labelKey = LCase$(Replace$(Replace$(featureId, "NX-TPL-", "template_"), "-", "_"))
    definition.Configure featureId, "NX-CAT-TEMPLATE", labelKey, labelKey & "_description", capabilities, _
        NxRiskL1, NxRiskL4, "template_result", sortOrder, NxOfficeCapabilityVba7 Or NxOfficeCapabilityWin32Api, _
        "disable_feature", "disable_feature", grade, False, False, 10000
    NxConfigureGeneratedLaunch definition, featureId
    definition.Seal
    registry.RegisterFeature definition
End Sub
