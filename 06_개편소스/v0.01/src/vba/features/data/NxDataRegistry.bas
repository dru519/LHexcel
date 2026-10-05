Attribute VB_Name = "NxDataRegistry"
Option Explicit

Public Sub NxDataRegisterFeatures(ByVal registry As CNxFeatureRegistry)
    If registry Is Nothing Then NxRaiseContractError "Data registry is required"
    RegisterDataFeature registry, "NX-DATA-NORMALIZE", NxExecutionFast, NxCapabilityRead Or NxCapabilityCellMutation Or NxCapabilityWorksheetCreate, NxRiskL1, NxRiskL2, 5
    RegisterDataFeature registry, "NX-DATA-UNIQUE-COUNT", NxExecutionFast, NxCapabilityRead Or NxCapabilityWorksheetCreate, NxRiskL0, NxRiskL1, 10
    RegisterDataFeature registry, "NX-DATA-DUPLICATE-LIST", NxExecutionFast, NxCapabilityRead Or NxCapabilityWorksheetCreate, NxRiskL1, NxRiskL2, 20
    RegisterDataFeature registry, NX_FEATURE_DATA_AGE, NxExecutionGuarded, NxCapabilityRead Or NxCapabilityCellMutation Or NxCapabilityWorksheetCreate, NxRiskL1, NxRiskL2, 30
    RegisterDataFeature registry, NX_FEATURE_DATA_KOREAN_MONEY, NxExecutionGuarded, NxCapabilityRead Or NxCapabilityCellMutation Or NxCapabilityWorksheetCreate, NxRiskL1, NxRiskL2, 40
    RegisterDataFeature registry, NX_FEATURE_DATA_PRIVACY_MASK, NxExecutionGuarded, NxCapabilityRead Or NxCapabilityCellMutation Or NxCapabilityWorksheetCreate Or NxCapabilitySensitiveOutput, NxRiskL2, NxRiskL3, 50
    RegisterDataFeature registry, NX_FEATURE_DATA_PRIVACY_SCAN, NxExecutionGuarded, NxCapabilityRead Or NxCapabilityWorksheetCreate Or NxCapabilitySensitiveOutput, NxRiskL2, NxRiskL3, 60
    RegisterDataFeature registry, "NX-DATA-FOCUS-CELL", NxExecutionFast, NxCapabilityRead, NxRiskL0, NxRiskL1, 70
    RegisterDataFeature registry, "NX-DATA-COPY-VISIBLE", NxExecutionFast, NxCapabilityRead Or NxCapabilityClipboardWrite, NxRiskL0, NxRiskL1, 80
    RegisterDataFeature registry, "NX-DATA-PASTE-VISIBLE-VALUES", NxExecutionPlanned, NxCapabilityRead Or NxCapabilityCellMutation, NxRiskL2, NxRiskL4, 90
End Sub

Public Function NxDataFeatureDefinition(ByVal featureId As String) As CNxFeatureDefinition
    Dim registry As CNxFeatureRegistry
    Set registry = CreateBaseRegistry()
    NxDataRegisterFeatures registry
    Set NxDataFeatureDefinition = registry.FeatureById(featureId)
End Function

Private Sub RegisterDataFeature(ByVal registry As CNxFeatureRegistry, ByVal featureId As String, _
    ByVal grade As NxExecutionGrade, ByVal capabilities As Long, ByVal baseRisk As NxRiskLevel, _
    ByVal maximumRisk As NxRiskLevel, ByVal sortOrder As Long)

    Dim definition As New CNxFeatureDefinition, labelKey As String
    labelKey = LCase$(Replace$(Replace$(featureId, "NX-DATA-", "data_"), "-", "_"))
    definition.Configure featureId, "NX-CAT-DATA", labelKey, labelKey & "_description", capabilities, _
        baseRisk, maximumRisk, "data_result", sortOrder, NxOfficeCapabilityNone, _
        "disable_feature", "disable_feature", grade, False, False, 10000
    NxConfigureGeneratedLaunch definition, featureId
    definition.Seal
    registry.RegisterFeature definition
End Sub
