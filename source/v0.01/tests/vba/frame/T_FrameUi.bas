Attribute VB_Name = "T_FrameUi"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestInputCheckingReady"
    names.Add "TestBlockedCannotExecute"
    names.Add "TestBackInvalidatesReadyState"
    names.Add "TestCancelIsTerminal"
    names.Add "TestRunSuccessShowsResult"
    names.Add "TestRunCancelledShowsResult"
    names.Add "TestRunInputErrorShowsResult"
    names.Add "TestRunEnvironmentErrorShowsResult"
    names.Add "TestRunPartialFailureShowsResult"
    names.Add "TestL4NeverEnablesExecute"
    names.Add "TestCloseButtonCancels"
    names.Add "TestControllerHasNoFeatureIdBranch"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestInputCheckingReady": TestInputCheckingReady
        Case "TestBlockedCannotExecute": TestBlockedCannotExecute
        Case "TestBackInvalidatesReadyState": TestBackInvalidatesReadyState
        Case "TestCancelIsTerminal": TestCancelIsTerminal
        Case "TestRunSuccessShowsResult": TestRunSuccessShowsResult
        Case "TestRunCancelledShowsResult": TestRunCancelledShowsResult
        Case "TestRunInputErrorShowsResult": TestRunInputErrorShowsResult
        Case "TestRunEnvironmentErrorShowsResult": TestRunEnvironmentErrorShowsResult
        Case "TestRunPartialFailureShowsResult": TestRunPartialFailureShowsResult
        Case "TestL4NeverEnablesExecute": TestL4NeverEnablesExecute
        Case "TestCloseButtonCancels": TestCloseButtonCancels
        Case "TestControllerHasNoFeatureIdBranch": TestControllerHasNoFeatureIdBranch
        Case Else: Err.Raise vbObjectError + 880, "T_FrameUi", "Unknown Frame UI test"
    End Select
End Sub

Private Sub TestInputCheckingReady()
    Dim controller As CNxExecutionFrameController, planView As CFakeExecutionPlanView
    Dim resultView As CFakeExecutionResultView, command As CFakeFeatureCommand
    Arrange controller, planView, resultView, command, NxSuccess, True, NxRiskL1
    controller.ShowPlan
    NxTestHarness.AssertTrue controller.State = NxFrameReady, "Plan must become ready"
    NxTestHarness.AssertTrue planView.RenderCount = 1 And planView.ExecuteEnabled, "Ready plan render mismatch"
    NxTestHarness.AssertTrue planView.LastModel.CategoryId = "NX-CAT-DATA", "Display conversion changed the category identity"
    NxTestHarness.AssertTrue planView.LastModel.DisplayCategoryLabel = "데이터/추가기능", "Category display label mismatch"
    NxTestHarness.AssertTrue planView.LastModel.DisplayFeatureName = "선택한 기능", "Unknown feature display fallback mismatch"
End Sub

Private Sub TestBlockedCannotExecute()
    Dim controller As CNxExecutionFrameController, planView As CFakeExecutionPlanView
    Dim resultView As CFakeExecutionResultView, command As CFakeFeatureCommand, failure As String
    Arrange controller, planView, resultView, command, NxSuccess, False, NxRiskL3
    controller.ShowPlan
    NxTestHarness.AssertTrue controller.State = NxFrameBlocked, "Blocked plan state mismatch"
    NxTestHarness.AssertTrue Not planView.ExecuteEnabled And planView.LastFocus = "cancel", "Blocked plan affordance mismatch"
    On Error GoTo Expected
    controller.ExecutePlan
    Err.Raise vbObjectError + 881, "T_FrameUi", "Blocked plan executed"
Expected:
    failure = Err.Description: Err.Clear
    NxTestHarness.AssertTrue InStr(failure, "transition is not allowed") > 0, "Wrong blocked execution error"
    NxTestHarness.AssertTrue command.ExecuteCount = 0, "Blocked plan reached command execution"
End Sub

Private Sub TestBackInvalidatesReadyState()
    Dim controller As CNxExecutionFrameController, planView As CFakeExecutionPlanView
    Dim resultView As CFakeExecutionResultView, command As CFakeFeatureCommand
    Arrange controller, planView, resultView, command, NxSuccess, True, NxRiskL1
    controller.ShowPlan
    controller.Back
    NxTestHarness.AssertTrue controller.State = NxFrameInput And planView.WasClosed, "Back did not return to input"
End Sub

Private Sub TestCancelIsTerminal()
    Dim controller As CNxExecutionFrameController, planView As CFakeExecutionPlanView
    Dim resultView As CFakeExecutionResultView, command As CFakeFeatureCommand
    Arrange controller, planView, resultView, command, NxSuccess, True, NxRiskL1
    controller.Cancel
    NxTestHarness.AssertTrue controller.State = NxFrameCancelled And planView.WasClosed, "Cancel must be terminal"
End Sub

Private Sub TestRunSuccessShowsResult(): AssertRunOutcome NxSuccess, NxFrameSuccess, "성공": End Sub
Private Sub TestRunCancelledShowsResult(): AssertRunOutcome NxCancelled, NxFrameCancelled, "취소": End Sub
Private Sub TestRunInputErrorShowsResult(): AssertRunOutcome NxInputError, NxFrameError, "입력 오류": End Sub
Private Sub TestRunEnvironmentErrorShowsResult(): AssertRunOutcome NxEnvironmentError, NxFrameError, "환경 오류": End Sub
Private Sub TestRunPartialFailureShowsResult(): AssertRunOutcome NxPartialFailure, NxFramePartialFailure, "부분 실패": End Sub

Private Sub TestL4NeverEnablesExecute()
    Dim controller As CNxExecutionFrameController, planView As CFakeExecutionPlanView
    Dim resultView As CFakeExecutionResultView, command As CFakeFeatureCommand
    Arrange controller, planView, resultView, command, NxSuccess, True, NxRiskL4
    controller.ShowPlan
    NxTestHarness.AssertTrue controller.State = NxFrameBlocked, "L4 must remain blocked"
    NxTestHarness.AssertTrue Not planView.ExecuteEnabled, "L4 enabled execution"
End Sub

Private Sub TestCloseButtonCancels()
    Dim controller As CNxExecutionFrameController, planView As CFakeExecutionPlanView
    Dim resultView As CFakeExecutionResultView, command As CFakeFeatureCommand
    Arrange controller, planView, resultView, command, NxSuccess, True, NxRiskL1
    controller.ShowPlan
    controller.CloseRequested
    NxTestHarness.AssertTrue controller.State = NxFrameCancelled And planView.WasClosed, "Close button must cancel"
End Sub

Private Sub TestControllerHasNoFeatureIdBranch()
    Dim controller As CNxExecutionFrameController, planView As CFakeExecutionPlanView
    Dim resultView As CFakeExecutionResultView, command As CFakeFeatureCommand
    Arrange controller, planView, resultView, command, NxSuccess, True, NxRiskL2
    controller.ShowPlan
    NxTestHarness.AssertTrue planView.LastModel.LabelKey = "fixture.label", "Controller interpreted feature presentation"
    NxTestHarness.AssertTrue controller.State = NxFrameReady, "Generic plan did not reach ready"
End Sub

Private Sub AssertRunOutcome(ByVal outcome As NxOutcome, ByVal expectedState As NxFrameState, ByVal expectedLabel As String)
    Dim controller As CNxExecutionFrameController, planView As CFakeExecutionPlanView
    Dim resultView As CFakeExecutionResultView, command As CFakeFeatureCommand
    Dim rawResult As CNxResult
    Arrange controller, planView, resultView, command, outcome, True, NxRiskL1
    command.Completed = True
    controller.ShowPlan
    controller.ExecutePlan
    Set rawResult = controller.Result
    NxTestHarness.AssertTrue controller.State = expectedState, "Result state mismatch"
    NxTestHarness.AssertTrue Not rawResult Is Nothing And rawResult.IsSealed And rawResult.Outcome = outcome, "Controller did not preserve the raw sealed result"
    NxTestHarness.AssertTrue rawResult.Recovery = vbNullString, "Display recovery text leaked into the raw result"
    NxTestHarness.AssertTrue resultView.RenderCount = 1 And resultView.WasShown, "Result view was not shown"
    NxTestHarness.AssertTrue resultView.LastModel.OutcomeLabel = expectedLabel, "Result label mismatch"
    NxTestHarness.AssertTrue resultView.LastModel.RecoverySummary = rawResult.Recovery, "Result card changed the raw recovery contract"
    NxTestHarness.AssertTrue resultView.LastModel.RecoveryDisplaySummary = "복구 필요 없음", "Result card did not preserve display recovery guidance"
    NxTestHarness.AssertTrue resultView.LastModel.Target = rawResult.Target, "Display target changed the approved identity"
    NxTestHarness.AssertTrue resultView.LastModel.TargetSummary = "target", "Result card lost the plan target summary"
    NxTestHarness.AssertTrue resultView.LastModel.ResultPath = vbNullString, "Non-file result exposed an identity as a file path"
    NxTestHarness.AssertTrue resultView.LastModel.CompletedSummary = "fake-completed", "Generic completed receipt display changed"
    NxTestHarness.AssertTrue command.ExecuteCount = 1 And planView.WasClosed, "Execution lifecycle mismatch"
End Sub

Private Sub Arrange(ByRef controller As CNxExecutionFrameController, _
    ByRef planView As CFakeExecutionPlanView, ByRef resultView As CFakeExecutionResultView, _
    ByRef commandObject As CFakeFeatureCommand, ByVal outcome As NxOutcome, _
    ByVal canExecute As Boolean, ByVal risk As NxRiskLevel)
    Dim planPort As INxExecutionPlanView, resultPort As INxExecutionResultView
    Dim command As INxFeatureCommand, ticket As CNxExecutionTicket
    Dim router As New CNxExecutionRouter
    Dim planModel As CNxExecutionPlanViewModel, catalog As CNxFrameTextCatalog
    Set planView = New CFakeExecutionPlanView
    Set resultView = New CFakeExecutionResultView
    Set commandObject = New CFakeFeatureCommand
    commandObject.Outcome = outcome
    Set command = commandObject
    Set ticket = router.Prepare(TestFrameDefinition(), command)
    Set planModel = MakePlanModel(canExecute, risk, ticket.Plan)
    Set catalog = MakeCatalog()
    Set planPort = planView
    Set resultPort = resultView
    Set controller = New CNxExecutionFrameController
    controller.Configure planPort, resultPort, planModel, ticket, catalog
End Sub

Private Function MakePlanModel(ByVal canExecute As Boolean, ByVal risk As NxRiskLevel, ByVal plan As CNxPlan) As CNxExecutionPlanViewModel
    Dim preview As New CNxExecutionPreview, model As New CNxExecutionPlanViewModel
    Dim blockReason As String
    If Not canExecute Or risk = NxRiskL4 Then blockReason = "blocked fixture"
    preview.Configure plan.FeatureId, "NX-CAT-DATA", "fixture.label", "target", "action", _
        "result-path", "one cell", NxPrivacySafe, "no conflict", "automatic", risk, _
        canExecute, risk <> NxRiskL0, blockReason
    preview.Seal
    model.Configure preview, "risk", "privacy", "계획대로 실행"
    model.Seal
    Set MakePlanModel = model
End Function

Private Function TestFrameDefinition() As CNxFeatureDefinition
    Dim definition As New CNxFeatureDefinition
    definition.Configure "NX-TEST-CORE", "NX-CAT-DATA", "test", "test_description", NxCapabilityRead, _
        NxRiskL0, NxRiskL4, "test_result", 1, NxOfficeCapabilityNone, "none", "disable_feature", _
        NxExecutionPlanned, False, False, 10000
    definition.ConfigureLaunch "direct", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "direct"
    definition.Seal
    Set TestFrameDefinition = definition
End Function

Private Function MakeCatalog() As CNxFrameTextCatalog
    Dim catalog As New CNxFrameTextCatalog
    catalog.Configure "성공", "취소", "입력 오류", "환경 오류", "부분 실패", _
        "없음", "완료하지 못한 작업을 확인하세요.", "원본이 변경됨", _
        "원본이 변경되지 않음", "복구 필요 없음", "결과를 확인하세요."
    Set MakeCatalog = catalog
End Function
