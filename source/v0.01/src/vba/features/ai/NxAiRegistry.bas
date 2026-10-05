Attribute VB_Name = "NxAiRegistry"
Option Explicit

Public Sub NxAiRegisterFeatures(ByVal registry As CNxFeatureRegistry)
    If registry Is Nothing Then NxRaiseContractError "AI registry is required"
    RegisterAiFeature registry, "NX-AI-SUMMARY", "ai_summary", 10
    RegisterAiFeature registry, "NX-AI-CLEAN", "ai_clean", 20
    RegisterAiFeature registry, "NX-AI-FORMULA", "ai_formula", 30
    RegisterAiFeature registry, "NX-AI-WRITE", "ai_write", 40
    RegisterAiFeature registry, "NX-AI-IMAGE", "ai_image", 50
End Sub

Public Function NxAiFeatureDefinition(ByVal featureId As String) As CNxFeatureDefinition
    Dim registry As CNxFeatureRegistry
    Set registry = CreateBaseRegistry()
    NxAiRegisterFeatures registry
    Set NxAiFeatureDefinition = registry.FeatureById(featureId)
End Function

Private Sub RegisterAiFeature(ByVal registry As CNxFeatureRegistry, ByVal featureId As String, _
    ByVal labelKey As String, ByVal sortOrder As Long)

    Dim definition As New CNxFeatureDefinition
    definition.Configure featureId, "NX-CAT-AI", labelKey, labelKey & "_description", _
        NxCapabilityRead Or NxCapabilityClipboardWrite Or NxCapabilityAiProcessing, _
        NxRiskL2, NxRiskL2, "ai_prompt_clipboard", sortOrder, NxOfficeCapabilityNone, "local_clipboard", _
        "disable_feature", NxExecutionGuarded, True, False, 10000
    NxConfigureGeneratedLaunch definition, featureId
    definition.Seal
    registry.RegisterFeature definition
End Sub
