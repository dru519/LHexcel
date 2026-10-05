Attribute VB_Name = "T_Frame"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    Dim uiNames As Collection, item As Variant
    names.Add "TestBaseCategories"
    names.Add "TestDuplicateCategoryRejected"
    names.Add "TestDuplicateFeatureRejected"
    names.Add "TestUnknownFeatureCategoryRejected"
    names.Add "TestFeatureLookup"
    names.Add "TestCategoryFeatureSort"
    names.Add "TestRiskL0ThroughL4"
    names.Add "TestCapabilityMismatchRejected"
    names.Add "TestPrivacyBlocked"
    names.Add "TestPrivacyIncomplete"
    names.Add "TestCautionExternalEscalatesL3"
    names.Add "TestPreviewSealedAndSingleIdentity"
    names.Add "TestOfficeCapabilityNative"
    names.Add "TestOfficeCapabilityFallback"
    names.Add "TestOfficeUnavailableDisablesOnlyFeature"
    names.Add "TestOfficeIncompleteBlocksFeature"
    Set uiNames = T_FrameUi.TestNames
    For Each item In uiNames
        names.Add CStr(item)
    Next item
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestBaseCategories": TestBaseCategories
        Case "TestDuplicateCategoryRejected": TestDuplicateCategoryRejected
        Case "TestDuplicateFeatureRejected": TestDuplicateFeatureRejected
        Case "TestUnknownFeatureCategoryRejected": TestUnknownFeatureCategoryRejected
        Case "TestFeatureLookup": TestFeatureLookup
        Case "TestCategoryFeatureSort": TestCategoryFeatureSort
        Case "TestRiskL0ThroughL4": TestRiskL0ThroughL4
        Case "TestCapabilityMismatchRejected": TestCapabilityMismatchRejected
        Case "TestPrivacyBlocked": TestPrivacyBlocked
        Case "TestPrivacyIncomplete": TestPrivacyIncomplete
        Case "TestCautionExternalEscalatesL3": TestCautionExternalEscalatesL3
        Case "TestPreviewSealedAndSingleIdentity": TestPreviewSealedAndSingleIdentity
        Case "TestOfficeCapabilityNative": TestOfficeCapabilityNative
        Case "TestOfficeCapabilityFallback": TestOfficeCapabilityFallback
        Case "TestOfficeUnavailableDisablesOnlyFeature": TestOfficeUnavailableDisablesOnlyFeature
        Case "TestOfficeIncompleteBlocksFeature": TestOfficeIncompleteBlocksFeature
        Case Else: T_FrameUi.RunCase name
    End Select
End Sub

Private Sub TestBaseCategories()
    Dim registry As CNxFeatureRegistry: Set registry = CreateBaseRegistry()
    NxTestHarness.AssertTrue registry.CategoryCount = 8, "Base registry must contain eight categories"
    NxTestHarness.AssertTrue registry.CategoryById("NX-CAT-FILE").SortOrder = 60, "File category order mismatch"
End Sub

Private Sub TestDuplicateCategoryRejected()
    Dim registry As CNxFeatureRegistry: Set registry = CreateBaseRegistry()
    Dim category As CNxCategoryDefinition: Set category = MakeCategory("NX-CAT-AI", 70)
    On Error GoTo Expected
    registry.RegisterCategory category
    Err.Raise vbObjectError + 791, "T_Frame", "Duplicate category was accepted"
Expected:
    NxTestHarness.AssertTrue InStr(Err.Description, "Duplicate category ID") > 0, "Wrong duplicate category error"
    Err.Clear
End Sub

Private Sub TestDuplicateFeatureRejected()
    Dim registry As CNxFeatureRegistry: Set registry = CreateBaseRegistry()
    registry.RegisterFeature MakeFeature("NX-DATA-UNIQUE-COUNT", "NX-CAT-DATA", 20, NxCapabilityRead, NxRiskL0, NxRiskL0, 0, vbNullString)
    On Error GoTo Expected
    registry.RegisterFeature MakeFeature("NX-DATA-UNIQUE-COUNT", "NX-CAT-DATA", 30, NxCapabilityRead, NxRiskL0, NxRiskL0, 0, vbNullString)
    Err.Raise vbObjectError + 792, "T_Frame", "Duplicate feature was accepted"
Expected:
    NxTestHarness.AssertTrue InStr(Err.Description, "Duplicate feature ID") > 0, "Wrong duplicate feature error"
    Err.Clear
End Sub

Private Sub TestUnknownFeatureCategoryRejected()
    Dim registry As CNxFeatureRegistry: Set registry = CreateBaseRegistry()
    On Error GoTo Expected
    registry.RegisterFeature MakeFeature("NX-DATA-UNIQUE-COUNT", "NX-CAT-MISSING", 10, NxCapabilityRead, NxRiskL0, NxRiskL0, 0, vbNullString)
    Err.Raise vbObjectError + 793, "T_Frame", "Unknown feature category was accepted"
Expected:
    NxTestHarness.AssertTrue InStr(Err.Description, "Unknown feature category") > 0, "Wrong unknown category error"
    Err.Clear
End Sub

Private Sub TestFeatureLookup()
    Dim registry As CNxFeatureRegistry: Set registry = CreateBaseRegistry()
    Dim feature As CNxFeatureDefinition: Set feature = MakeFeature("NX-DATA-UNIQUE-COUNT", "NX-CAT-DATA", 10, NxCapabilityRead, NxRiskL0, NxRiskL0, 0, vbNullString)
    registry.RegisterFeature feature
    NxTestHarness.AssertSameObject feature, registry.FeatureById(feature.FeatureId), "Feature lookup lost identity"
End Sub

Private Sub TestCategoryFeatureSort()
    Dim registry As CNxFeatureRegistry: Set registry = CreateBaseRegistry()
    registry.RegisterFeature MakeFeature("NX-DATA-UNIQUE-COUNT", "NX-CAT-DATA", 20, NxCapabilityRead, NxRiskL0, NxRiskL0, 0, vbNullString)
    registry.RegisterFeature MakeFeature("NX-DATA-DUPLICATE-LIST", "NX-CAT-DATA", 10, NxCapabilityRead, NxRiskL0, NxRiskL0, 0, vbNullString)
    Dim items As Collection: Set items = registry.FeaturesForCategory("NX-CAT-DATA")
    NxTestHarness.AssertTrue items.Item(1).SortOrder = 10 And items.Item(2).SortOrder = 20, "Category feature sort mismatch"
End Sub

Private Sub TestRiskL0ThroughL4()
    Dim destructive As CNxFeatureDefinition
    Dim destructivePreview As CNxExecutionPreview
    Dim quarantine As CNxFeatureDefinition
    NxTestHarness.AssertTrue NxRiskL0 = 0, "L0 mismatch"
    NxTestHarness.AssertTrue RiskForEffect(NxCellMutation) = NxRiskL1, "L1 mismatch"
    NxTestHarness.AssertTrue RiskForEffect(NxFileCreate) = NxRiskL2, "L2 mismatch"
    NxTestHarness.AssertTrue RiskForEffect(NxExternalProcess) = NxRiskL3, "L3 mismatch"
    NxTestHarness.AssertTrue RiskForEffect(NxFileReplace) = NxRiskL4, "L4 mismatch"
    Set destructive = MakeFeature("NX-FILE-MANNER-SAVE", "NX-CAT-FILE", 10, NxCapabilityFileReplace Or NxCapabilityExternalProcess, NxRiskL0, NxRiskL4, 0, vbNullString)
    Set destructivePreview = MakePreview(destructive, MakePlan(destructive.FeatureId, NxFileReplace, 0, NxExternalProcess), NxPrivacyCaution, MakeCompatibility(destructive.FeatureId, 0, True, vbNullString))
    NxTestHarness.AssertTrue destructivePreview.Risk = NxRiskL4 And Not destructivePreview.CanExecute, "Privacy caution must never downgrade blocked L4"
    Set quarantine = MakeFeature("NX-TPL-DELETE", "NX-CAT-TEMPLATE", 10, NxCapabilityTemplateQuarantine, NxRiskL2, NxRiskL2, 0, vbNullString)
    Set destructivePreview = MakePreview(quarantine, MakePlan(quarantine.FeatureId, 0, NxCapabilityTemplateQuarantine), NxPrivacySafe, MakeCompatibility(quarantine.FeatureId, 0, True, vbNullString))
    NxTestHarness.AssertTrue destructivePreview.Risk = NxRiskL2 And destructivePreview.CanExecute, "Template delete plan must remain executable L2 quarantine"
End Sub

Private Sub TestCapabilityMismatchRejected()
    Dim feature As CNxFeatureDefinition: Set feature = MakeFeature("NX-DATA-UNIQUE-COUNT", "NX-CAT-DATA", 10, NxCapabilityRead, NxRiskL0, NxRiskL4, 0, vbNullString)
    Dim plan As CNxPlan: Set plan = MakePlan("NX-DATA-UNIQUE-COUNT", NxCellMutation)
    On Error GoTo Expected
    Call MakePreview(feature, plan, NxPrivacySafe, MakeCompatibility(feature.FeatureId, 0, True, vbNullString))
    Err.Raise vbObjectError + 794, "T_Frame", "Capability mismatch was accepted"
Expected:
    NxTestHarness.AssertTrue InStr(Err.Description, "capability exceeds") > 0, "Wrong capability mismatch error"
    Err.Clear
End Sub

Private Sub TestPrivacyBlocked(): AssertPrivacyBlocked NxPrivacyBlocked: End Sub
Private Sub TestPrivacyIncomplete(): AssertPrivacyBlocked NxPrivacyIncomplete: End Sub

Private Sub TestCautionExternalEscalatesL3()
    AssertMaximumRiskRejected
    Dim feature As CNxFeatureDefinition: Set feature = MakeFeature("NX-AI-SUMMARY", "NX-CAT-AI", 10, NxCapabilityExternalProcess, NxRiskL0, NxRiskL3, 0, vbNullString)
    Dim preview As CNxExecutionPreview: Set preview = MakePreview(feature, MakePlan("NX-AI-SUMMARY", NxExternalProcess), NxPrivacyCaution, MakeCompatibility(feature.FeatureId, 0, True, vbNullString))
    NxTestHarness.AssertTrue preview.Risk = NxRiskL3, "External caution must escalate to L3"
    Set feature = MakeFeature("NX-FILE-MANNER-SAVE", "NX-CAT-FILE", 10, NxCapabilityFileCreate Or NxCapabilitySensitiveOutput, NxRiskL2, NxRiskL3, 0, vbNullString)
    Set preview = MakePreview(feature, MakePlan(feature.FeatureId, NxFileCreate, NxCapabilitySensitiveOutput), NxPrivacyCaution, MakeCompatibility(feature.FeatureId, 0, True, vbNullString))
    NxTestHarness.AssertTrue preview.Risk = NxRiskL3, "Sensitive file output must escalate to L3"
End Sub

Private Sub AssertMaximumRiskRejected()
    Dim feature As CNxFeatureDefinition
    Dim plan As CNxPlan
    Set feature = MakeFeature("NX-FILE-MANNER-SAVE", "NX-CAT-FILE", 10, NxCapabilityFileCreate Or NxCapabilitySensitiveOutput, NxRiskL2, NxRiskL2, 0, vbNullString)
    Set plan = MakePlan(feature.FeatureId, NxFileCreate, NxCapabilitySensitiveOutput)
    On Error GoTo Expected
    Call MakePreview(feature, plan, NxPrivacyCaution, MakeCompatibility(feature.FeatureId, 0, True, vbNullString))
    Err.Raise vbObjectError + 795, "T_Frame", "Maximum risk overflow was accepted"
Expected:
    NxTestHarness.AssertTrue InStr(Err.Description, "exceeds feature maximum risk") > 0, "Wrong maximum risk error"
    Err.Clear
End Sub

Private Sub TestPreviewSealedAndSingleIdentity()
    AssertPreviewSummaryRejected 1
    AssertPreviewSummaryRejected 2
    AssertPreviewSummaryRejected 3
    Dim feature As CNxFeatureDefinition: Set feature = MakeFeature("NX-DATA-UNIQUE-COUNT", "NX-CAT-DATA", 10, NxCapabilityRead, NxRiskL0, NxRiskL0, 0, vbNullString)
    Dim preview As CNxExecutionPreview: Set preview = MakePreview(feature, MakePlan(feature.FeatureId, 0), NxPrivacySafe, MakeCompatibility(feature.FeatureId, 0, True, vbNullString))
    NxTestHarness.AssertTrue preview.IsSealed And preview.FeatureId = feature.FeatureId And preview.CanExecute, "Preview identity or seal mismatch"
    NxTestHarness.AssertTrue preview.LabelKey = "label.key" And preview.TargetSummary = "target" And preview.ActionSummary = "action", "Preview input summaries are not readable"
    NxTestHarness.AssertTrue preview.ResultSummary = "result" And preview.ScaleSummary = "scale" And preview.ConflictSummary = "none" And preview.RecoverySummary = "recovery", "Preview result summaries are not readable"
End Sub

Private Sub AssertPreviewSummaryRejected(ByVal emptyField As Long)
    Dim preview As New CNxExecutionPreview
    Dim scaleSummary As String, conflictSummary As String, recoverySummary As String
    scaleSummary = "scale": conflictSummary = "conflict": recoverySummary = "recovery"
    If emptyField = 1 Then scaleSummary = vbNullString
    If emptyField = 2 Then conflictSummary = vbNullString
    If emptyField = 3 Then recoverySummary = vbNullString
    On Error GoTo Expected
    preview.Configure "feature", "category", "label", "target", "action", "result", scaleSummary, NxPrivacySafe, conflictSummary, recoverySummary, NxRiskL0, True, False, vbNullString
    Err.Raise vbObjectError + 796, "T_Frame", "Empty preview summary was accepted"
Expected:
    NxTestHarness.AssertTrue InStr(Err.Description, "requires all summaries") > 0, "Wrong empty summary error"
    Err.Clear
End Sub

Private Sub TestOfficeCapabilityNative()
    Dim value As CNxOfficeCompatibility: Set value = MakeCompatibility("NX-DATA-UNIQUE-COUNT", NxOfficeCapabilityVba7, True, vbNullString)
    NxTestHarness.AssertTrue value.State = NxCompatibilityNative And value.CanExecute, "Native capability mismatch"
End Sub

Private Sub TestOfficeCapabilityFallback()
    Dim value As CNxOfficeCompatibility
    Set value = NxEvaluateOfficeCompatibility("NX-DATA-UNIQUE-COUNT", NxOfficeCapabilityDynamicArray, NX_FALLBACK_FORMULA_LEGACY, 0, True, "16.0", 64)
    NxTestHarness.AssertTrue value.State = NxCompatibilityFallback And value.CanExecute, "Fallback capability mismatch"
End Sub

Private Sub TestOfficeUnavailableDisablesOnlyFeature()
    Dim value As CNxOfficeCompatibility: Set value = NxEvaluateOfficeCompatibility("NX-DATA-UNIQUE-COUNT", NxOfficeCapabilityDynamicArray, vbNullString, 0, True, "16.0", 64)
    NxTestHarness.AssertTrue value.State = NxCompatibilityUnavailable And Not value.CanExecute And value.UnsupportedBehavior = "disable_feature", "Unavailable feature isolation mismatch"
End Sub

Private Sub TestOfficeIncompleteBlocksFeature()
    Dim value As CNxOfficeCompatibility: Set value = NxEvaluateOfficeCompatibility("NX-DATA-UNIQUE-COUNT", 0, vbNullString, 0, False, "16.0", 64)
    NxTestHarness.AssertTrue value.State = NxCompatibilityIncomplete And Not value.CanExecute, "Incomplete probe must block"
End Sub

Private Sub AssertPrivacyBlocked(ByVal state As NxPrivacyState)
    Dim feature As CNxFeatureDefinition: Set feature = MakeFeature("NX-DATA-UNIQUE-COUNT", "NX-CAT-DATA", 10, NxCapabilityRead, NxRiskL0, NxRiskL0, 0, vbNullString)
    Dim preview As CNxExecutionPreview: Set preview = MakePreview(feature, MakePlan(feature.FeatureId, 0), state, MakeCompatibility(feature.FeatureId, 0, True, vbNullString))
    NxTestHarness.AssertTrue Not preview.CanExecute, "Privacy state must block preview"
    If state = NxPrivacyIncomplete Then
        AssertUnknownPrivacyRejected -1
        AssertUnknownPrivacyRejected 4
        AssertUnknownPrivacyRejected 99
    End If
End Sub

Private Sub AssertUnknownPrivacyRejected(ByVal invalidState As Long)
    Dim feature As CNxFeatureDefinition
    Dim plan As CNxPlan
    Set feature = MakeFeature("NX-FILE-MANNER-SAVE", "NX-CAT-FILE", 10, NxCapabilityFileCreate Or NxCapabilitySensitiveOutput, NxRiskL2, NxRiskL3, 0, vbNullString)
    Set plan = MakePlan(feature.FeatureId, NxFileCreate, NxCapabilitySensitiveOutput)
    On Error GoTo Expected
    Call MakePreview(feature, plan, invalidState, MakeCompatibility(feature.FeatureId, 0, True, vbNullString))
    Err.Raise vbObjectError + 797, "T_Frame", "Unknown privacy state was accepted"
Expected:
    NxTestHarness.AssertTrue InStr(Err.Description, "Privacy state is outside") > 0, "Wrong unknown privacy error"
    Err.Clear
End Sub

Private Function MakeCategory(ByVal categoryId As String, ByVal sortOrder As Long) As CNxCategoryDefinition
    Dim value As New CNxCategoryDefinition: value.Configure categoryId, "NX-ENTRY-ALL", "label", sortOrder: value.Seal: Set MakeCategory = value
End Function

Private Function MakeFeature(ByVal featureId As String, ByVal categoryId As String, ByVal sortOrder As Long, ByVal capabilities As Long, ByVal baseRisk As NxRiskLevel, ByVal maximumRisk As NxRiskLevel, ByVal officeCapabilities As Long, ByVal fallback As String) As CNxFeatureDefinition
    Dim value As New CNxFeatureDefinition
    value.Configure featureId, categoryId, "label.key", "description.key", capabilities, baseRisk, maximumRisk, "result", sortOrder, officeCapabilities, fallback, "disable_feature", NxExecutionFast, False, False, 10000
    value.ConfigureLaunch "direct", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "direct"
    value.Seal: Set MakeFeature = value
End Function

Private Function MakePlan(ByVal featureId As String, ByVal kind As Long, Optional ByVal declaredCapability As Long = 0, Optional ByVal secondKind As Long = 0) As CNxPlan
    Dim plan As New CNxPlan: plan.Configure featureId, "context", "source"
    If declaredCapability <> 0 Then plan.DeclareCapability declaredCapability
    If kind <> 0 Then AddEffectToPlan plan, featureId, kind, "target-1"
    If secondKind <> 0 Then AddEffectToPlan plan, featureId, secondKind, "target-2"
    plan.Seal: Set MakePlan = plan
End Function

Private Sub AddEffectToPlan(ByVal plan As CNxPlan, ByVal featureId As String, ByVal kind As Long, ByVal target As String)
    Dim effect As New CNxSideEffect, policy As New CNxLimitsPolicy
    Select Case kind
        Case NxCellMutation: effect.ConfigureCell target, NxCellValueMask, 1
        Case NxFileCreate: effect.ConfigureTextAdapter kind, target, "payload", "remove created file", vbNullString, False
        Case NxFileReplace: effect.ConfigureTextAdapter kind, target, "payload", "restore prior file", "prior", True
        Case NxClipboardWrite: effect.ConfigureTextAdapter kind, target, "payload", "restore clipboard", "prior", True
        Case NxExternalProcess: effect.ConfigureManualAdapter kind, target, NxPayloadCommandLine, "cmd /c exit 0", "manual"
        Case NxBrowserOpen: effect.ConfigureManualAdapter kind, target, NxPayloadUri, "https://example.invalid", "manual"
        Case Else: NxRaiseContractError "Unknown Frame side effect fixture"
    End Select
    effect.BindDecision policy.ExactOwnerFor(featureId), policy.MaximumCellsFor(featureId)
    effect.Seal
    plan.AddSideEffect effect
End Sub

Private Function MakeCompatibility(ByVal featureId As String, ByVal detected As Long, ByVal complete As Boolean, ByVal fallback As String) As CNxOfficeCompatibility
    Set MakeCompatibility = NxEvaluateOfficeCompatibility(featureId, detected, fallback, detected, complete, "16.0", 64)
End Function

Private Function MakePreview(ByVal feature As CNxFeatureDefinition, ByVal plan As CNxPlan, ByVal privacy As NxPrivacyState, ByVal compatibility As CNxOfficeCompatibility) As CNxExecutionPreview
    Set MakePreview = NxCreateExecutionPreviewForCompatibility(feature, plan, "target", "action", "result", "scale", privacy, "none", "recovery", compatibility)
End Function
