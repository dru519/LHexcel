Attribute VB_Name = "NxHangulRegistry"
Option Explicit

Public Sub NxHangulRegisterFeatures(ByVal registry As CNxFeatureRegistry)
    Dim definition As New CNxFeatureDefinition
    If registry Is Nothing Then NxRaiseContractError "Hangul registry is required"

    definition.Configure NX_FEATURE_HANGUL_TABLE_SEND, "NX-CAT-DRAW", _
        "hangul_table_send", "hangul_table_send_description", _
        NxCapabilityRead Or NxCapabilityBinaryPackageWrite Or NxCapabilityFileCreate Or NxCapabilityExternalProcess, _
        NxRiskL2, NxRiskL2, "hangul_result", 30, NxOfficeCapabilityNone, _
        "disable_feature", "disable_feature", NxExecutionGuarded, _
        False, False, 10000, False, False, True
    NxConfigureGeneratedLaunch definition, NX_FEATURE_HANGUL_TABLE_SEND
    definition.Seal
    registry.RegisterFeature definition
    Set definition = New CNxFeatureDefinition
    definition.Configure NX_FEATURE_HANGUL_PICTURE_SEND, "NX-CAT-DRAW", _
        "hangul_picture_send", "hangul_picture_send_description", _
        NxCapabilityRead Or NxCapabilityBinaryPackageWrite Or NxCapabilityFileCreate Or NxCapabilityExternalProcess, _
        NxRiskL2, NxRiskL2, "hangul_result", 31, NxOfficeCapabilityNone, _
        "disable_feature", "disable_feature", NxExecutionGuarded, _
        False, False, 10000, False, True, True
    NxConfigureGeneratedLaunch definition, NX_FEATURE_HANGUL_PICTURE_SEND
    definition.Seal
    registry.RegisterFeature definition
End Sub

Public Function NxHangulFeatureDefinition(ByVal featureId As String) As CNxFeatureDefinition
    Dim registry As CNxFeatureRegistry
    Set registry = CreateBaseRegistry()
    NxHangulRegisterFeatures registry
    Set NxHangulFeatureDefinition = registry.FeatureById(featureId)
End Function
