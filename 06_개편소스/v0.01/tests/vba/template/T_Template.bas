Attribute VB_Name = "T_Template"
Option Explicit

' The feature-only native host does not import the product Ribbon shell.
' Keep the production NxTemplateUi callback compile-bound without weakening
' the Product.xlam path, which imports the real NxRibbonCallbacks module.
Public Sub NxRibbonExecuteTag(ByVal tag As String)
    If Len(tag) = 0 Then NxRaiseContractError "Ribbon tag is required"
End Sub

Public Function TestNames() As Collection
    Dim names As New Collection
    AppendNames names, T_TemplatePolicy.TestNames
    AppendNames names, T_TemplatePackage.TestNames
    AppendNames names, T_TemplateWorkflow.TestNames
    AppendNames names, T_TemplateUi.TestNames
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestRetainedAreasMustStayInsideSource", "TestRetainedConstantsRoundTrip", _
             "TestConstantsOutsideRetainedAreasAreBlank", "TestInternalFormulaR1C1IsAllowed", _
             "TestVolatileFormulaRequiresConsent", "TestDynamicFormulaRequiresConsent", _
             "TestEnvironmentFormulaRequiresConsent", "TestExternalAndOtherSheetFormulaIsBlocked", _
             "TestExecutableOrExternalResourceIsBlocked", "TestUnsupportedDisplayResourceRequiresExclusionAck"
            T_TemplatePolicy.RunCase name
        Case "TestPackageIsMacroFreeSingleSheet", "TestPackageMetadataUsesClosedSchema", _
             "TestClosedPackageShaMatchesMetadata", "TestLogicalHashRoundTrips", _
             "TestPackageRejectsExternalWorkbookFeatures", "TestTempPackageReopensReadOnlyWithoutLinks", _
             "TestMetadataReplacementRestoresPriorFile", "TestTamperedPackageIsListedDamaged"
            T_TemplatePackage.RunCase name
        Case "TestRegisterRangeRecordsSourceKind", "TestRegisterSheetUsesUsedRange", _
             "TestDuplicateNameNeverOverwrites", "TestListIsLocalFastWithoutPrivacyScan", _
             "TestLoadCreatesNewSheetAtA1", "TestRiskyOrFailedLoadRequiresConsentAndRollsBack", _
             "TestRenameKeepsTemplateIdAndPackageName", "TestDeleteQuarantinesAtomically"
            T_TemplateWorkflow.RunCase name
        Case "TestTemplateRegistryLocksSixGrades", "TestTemplateFormHasRequiredControls", _
             "TestRibbonEntryOpensTemplateForm", "TestBinaryFormsOverwriteAndForcedLoadAreAbsent"
            T_TemplateUi.RunCase name
        Case Else: NxRaiseContractError "Unknown Template test"
    End Select
End Sub

Private Sub AppendNames(ByVal target As Collection, ByVal source As Collection)
    Dim item As Variant
    For Each item In source
        target.Add CStr(item)
    Next item
End Sub
