Attribute VB_Name = "NxDrawingRegistry"
Option Explicit

Public Sub NxDrawingRegisterFeatures(ByVal registry As CNxFeatureRegistry)
    If registry Is Nothing Then NxRaiseContractError "표·그리기 레지스트리가 필요합니다."
    RegisterDrawingFeature registry, NX_FEATURE_DRAW_TITLE_TABLE, "draw_title_table", 10
    RegisterDrawingFeature registry, NX_FEATURE_DRAW_BUSINESS_TABLE, "draw_business_table", 20
    RegisterDrawingFeature registry, NX_FEATURE_DRAW_ROLE_STYLE, "draw_role_style", 21
    RegisterDrawingFeature registry, NX_FEATURE_DRAW_CLEAR_INNER, "draw_clear_inner", 30
    RegisterDrawingFeature registry, NX_FEATURE_DRAW_FIT_PICTURE, "draw_fit_picture", 50
    RegisterDrawingFeature registry, NX_FEATURE_DRAW_INSERT_PICTURE, "draw_insert_picture", 51
End Sub

Public Function NxDrawingFeatureDefinition(ByVal featureId As String) As CNxFeatureDefinition
    Dim registry As CNxFeatureRegistry
    Set registry = CreateBaseRegistry(): NxDrawingRegisterFeatures registry
    Set NxDrawingFeatureDefinition = registry.FeatureById(featureId)
End Function

Private Sub RegisterDrawingFeature(ByVal registry As CNxFeatureRegistry, ByVal featureId As String, _
    ByVal labelKey As String, ByVal sortOrder As Long)

    Dim definition As New CNxFeatureDefinition
    If featureId = NX_FEATURE_DRAW_FIT_PICTURE Or featureId = NX_FEATURE_DRAW_INSERT_PICTURE Then
        definition.Configure featureId, "NX-CAT-DRAW", labelKey, labelKey & "_description", _
            NxCapabilityRead Or NxCapabilityShapeInsert, NxRiskL1, NxRiskL2, "draw_result", sortOrder, _
            NxOfficeCapabilityNone, "disable_feature", "disable_feature", NxExecutionPlanned, False, False, 10000
    Else
        definition.Configure featureId, "NX-CAT-DRAW", labelKey, labelKey & "_description", _
            NxCapabilityRead Or NxCapabilityCellMutation, NxRiskL1, NxRiskL2, "draw_result", sortOrder, _
            NxOfficeCapabilityNone, "disable_feature", "disable_feature", NxExecutionFast, False, False, 10000
    End If
    NxConfigureGeneratedLaunch definition, featureId
    definition.Seal
    registry.RegisterFeature definition
End Sub
