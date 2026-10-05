Attribute VB_Name = "T_TemplateUi"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestTemplateRegistryLocksSixGrades"
    names.Add "TestTemplateFormHasRequiredControls"
    names.Add "TestRibbonEntryOpensTemplateForm"
    names.Add "TestBinaryFormsOverwriteAndForcedLoadAreAbsent"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestTemplateRegistryLocksSixGrades": TestTemplateRegistryLocksSixGrades
        Case "TestTemplateFormHasRequiredControls": TestTemplateFormHasRequiredControls
        Case "TestRibbonEntryOpensTemplateForm": TestRibbonEntryOpensTemplateForm
        Case "TestBinaryFormsOverwriteAndForcedLoadAreAbsent": TestBinaryFormsOverwriteAndForcedLoadAreAbsent
        Case Else: NxRaiseContractError "Unknown Template UI test"
    End Select
End Sub

Private Sub TestTemplateRegistryLocksSixGrades()
    Dim registry As CNxFeatureRegistry, definition As CNxFeatureDefinition
    Set registry = CreateBaseRegistry()
    NxTemplateRegisterFeatures registry
    NxTestHarness.AssertTrue registry.FeaturesForCategory("NX-CAT-TEMPLATE").Count = 6, "Template registry must contain six features"
    Set definition = registry.FeatureById("NX-TPL-LIST")
    NxTestHarness.AssertTrue definition.ExecutionGrade = NxExecutionFast, "List must be Fast"
    Set definition = registry.FeatureById("NX-TPL-LOAD")
    NxTestHarness.AssertTrue definition.ExecutionGrade = NxExecutionGuarded, "Load base grade must be Guarded"
    NxTestHarness.AssertTrue Not definition.AiProcessing And Not definition.ExternalDataTransfer, "Local template invoked privacy gate"
    NxTestHarness.AssertTrue registry.FeatureById("NX-TPL-REGISTER-RANGE").ExecutionGrade = NxExecutionPlanned, "Registration must be Planned"
    NxTestHarness.AssertTrue registry.FeatureById("NX-TPL-DELETE").ExecutionGrade = NxExecutionPlanned, "Quarantine must be Planned"
End Sub

Private Sub TestTemplateFormHasRequiredControls()
    Dim controlName As Variant
    For Each controlName In Array("cboSourceKind", "txtTargetSummary", "cmdSelectRetained", "txtRetainedSummary", "txtName", _
        "txtDescription", "cboHiddenPolicy", "txtPreview", "cboRiskConsent", "cmdPreview", "cmdRegister", _
        "cboTemplates", "txtTemplateDetails", "cmdLoadNewSheet", "cmdRename", "cmdDelete", "cmdCancel")
        NxTestHarness.AssertTrue NxTemplateFormContractHas(CStr(controlName)), "Required Template form control is missing"
    Next controlName
End Sub

Private Sub TestRibbonEntryOpensTemplateForm()
    NxTestHarness.AssertTrue NxTemplateRibbonCallbackTarget() = "NxTemplateOpenManager", "Template ribbon callback target is invalid"
    NxFeatureRouteStubReset
    NxTemplateOpenManager "NX-TPL-LIST"
    NxTestHarness.AssertTrue NxFeatureRouteStubLastFeatureId() = "NX-TPL-LIST", "Template compatibility wrapper did not preserve the exact feature route"
End Sub

Private Sub TestBinaryFormsOverwriteAndForcedLoadAreAbsent()
    NxTestHarness.AssertTrue NxTemplateUsesGeneratedFormSources(), "Template form must use JSON plus VBA sources"
    NxTestHarness.AssertTrue Not NxTemplateFormContractHas("cmdOverwrite"), "Overwrite control must not exist"
    NxTestHarness.AssertTrue Not NxTemplateFormContractHas("cmdForceLoad"), "Forced damaged load control must not exist"
End Sub
