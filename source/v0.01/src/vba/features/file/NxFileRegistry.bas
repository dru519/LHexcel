Attribute VB_Name = "NxFileRegistry"
Option Explicit

Public Sub NxFileRegisterFeatures(ByVal registry As CNxFeatureRegistry)
    If registry Is Nothing Then NxRaiseContractError "File registry is required"
    RegisterFileFeature registry, "NX-FILE-CONSOLIDATE", NxCapabilityRead Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 10
    RegisterMannerSaveFeature registry
    RegisterFileFeature registry, "NX-FILE-RANGE-PNG", NxCapabilityRead Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 30
    RegisterFileFeature registry, "NX-FILE-CHART-PNG", NxCapabilityRead Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 40
    RegisterFileFeature registry, "NX-FILE-FOLDER-CREATE", NxCapabilityRead Or NxCapabilityDirectoryCreate, 50
    RegisterFileFeature registry, "NX-FILE-SHEET-COPY-SAVE", NxCapabilityRead Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 60
    RegisterFileFeature registry, "NX-FILE-RANGE-COPY-SAVE", NxCapabilityRead Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 70
    RegisterFileFeature registry, "NX-FILE-WORKBOOK-COMPARE", NxCapabilityRead Or NxCapabilityWorksheetCreate Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 80
    RegisterFileFeature registry, "NX-FILE-SHEET-COMPARE", NxCapabilityRead Or NxCapabilityWorksheetCreate, 81
    RegisterFileFeature registry, "NX-FILE-FILE-COMPARE", NxCapabilityRead Or NxCapabilityWorksheetCreate Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 82
    RegisterFileFeature registry, "NX-FILE-BATCH-RENAME", NxCapabilityRead Or NxCapabilityFileReplace Or NxCapabilityBinaryPackageWrite, 90
    RegisterFileFeature registry, "NX-FILE-SHEET-BATCH-RENAME", NxCapabilityRead Or NxCapabilityBinaryPackageWrite, 100
    RegisterFileFeature registry, "NX-FILE-PDF-CURRENT-SHEET", NxCapabilityRead Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 110
    RegisterFileFeature registry, "NX-FILE-PDF-EACH-SHEET", NxCapabilityRead Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 120
    RegisterFileFeature registry, "NX-FILE-PDF-SELECTED-COMBINED", NxCapabilityRead Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 130
    RegisterFileFeature registry, "NX-FILE-PDF-ALL-COMBINED", NxCapabilityRead Or NxCapabilityFileCreate Or NxCapabilityBinaryPackageWrite, 140
    RegisterFileFeature registry, "NX-FILE-PDF-SETTINGS", NxCapabilityRead Or NxCapabilityFileCreate Or NxCapabilityFileReplace Or NxCapabilityBinaryPackageWrite, 150
End Sub

Public Function NxFileFeatureDefinition(ByVal featureId As String) As CNxFeatureDefinition
    Dim registry As CNxFeatureRegistry
    Set registry = CreateBaseRegistry()
    NxFileRegisterFeatures registry
    Set NxFileFeatureDefinition = registry.FeatureById(featureId)
End Function

Private Sub RegisterMannerSaveFeature(ByVal registry As CNxFeatureRegistry)
    Dim definition As New CNxFeatureDefinition
    ' Native in-place Save is explicitly requested by clicking the direct command.
    ' It must never enter new-output rollback, which can delete its output path.
    definition.Configure "NX-FILE-MANNER-SAVE", "NX-CAT-FILE", "file_manner_save", _
        "file_manner_save_description", NxCapabilityRead Or NxCapabilityFileReplace, _
        NxRiskL2, NxRiskL4, "file_result", 20, NxOfficeCapabilityNone, _
        "disable_feature", "disable_feature", NxExecutionGuarded, False, False, 10000, False, False, True
    NxConfigureGeneratedLaunch definition, NX_FEATURE_FILE_MANNER_SAVE
    definition.Seal
    registry.RegisterFeature definition
End Sub

Private Sub RegisterFileFeature(ByVal registry As CNxFeatureRegistry, ByVal featureId As String, _
    ByVal capabilities As Long, ByVal sortOrder As Long)
    Dim definition As New CNxFeatureDefinition, labelKey As String
    labelKey = LCase$(Replace$(Replace$(featureId, "NX-FILE-", "file_"), "-", "_"))
    definition.Configure featureId, "NX-CAT-FILE", labelKey, labelKey & "_description", capabilities, _
        NxRiskL2, NxRiskL4, "file_result", sortOrder, NxOfficeCapabilityNone, _
        "disable_feature", "disable_feature", NxExecutionPlanned, False, False, 10000, False, False, True
    NxConfigureGeneratedLaunch definition, featureId
    definition.Seal
    registry.RegisterFeature definition
End Sub
