Attribute VB_Name = "T_Core"
Option Explicit

Private mExpectedRedProbe As CNxStateGuard

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestAllApplicationStates"
    names.Add "TestStatusBarFalseAndString"
    names.Add "TestActiveIdentityAfterSwitch"
    names.Add "TestRunnerNormalReturn"
    names.Add "TestRunnerForcedExecuteError"
    names.Add "TestRunnerUserCancel"
    names.Add "TestRestoreTwice"
    names.Add "TestClassTerminateRestores"
    names.Add "TestInjectedRestoreFailuresContinue"
    names.Add "TestRollbackJournalRestoresCell"
    names.Add "TestFingerprintDriftRaisesNxInputErrorBeforeGuard"
    names.Add "TestSealedPlanRejectsMutation"
    names.Add "TestExactResultContracts"
    names.Add "TestRestoreFailureDowngradesSuccessWithoutChange"
    names.Add "TestRestoreFailureDowngradesSuccessWithChange"
    names.Add "TestRestoreFailureLeavesNonSuccess"
    names.Add "TestBuildPlanMutationIsRejectedBeforeGuard"
    names.Add "TestRunnerRejectsEquivalentButDifferentPlan"
    names.Add "TestApprovalObjectBindsFeaturePlanDigestAndContext"
    names.Add "TestRunnerRevalidatesFeatureBeforeGuard"
    names.Add "TestRunnerInvalidExecuteResultIsNotPreflight"
    names.Add "TestInvalidResultFeatureIdRollsBack"
    names.Add "TestPostChangeExecuteExceptionAndRollbackFailureIsPartial"
    names.Add "TestCompletedAndRecoveryReportActualWork"
    names.Add "TestStateGuardOppositeStateBaseline"
    names.Add "TestStateGuardStatusBarFalseAndString"
    names.Add "TestStateGuardRestoreFailureAggregates"
    names.Add "TestStateGuardTerminateEmergencyCleanup"
    names.Add "TestApprovalObjectIsSingleUse"
    names.Add "TestApprovalRejectsEquivalentObjects"
    names.Add "TestApprovalBindsRunSnapshotSourceDigest"
    names.Add "TestFinalStateResultMatrix"
    names.Add "TestFormula2FallbackOnlyWhenUnsupported"
    names.Add "TestFormula2RestoreNeverSilentlyDowngrades"
    names.Add "TestTypedNonCellDispatchAndCompensation"
    names.Add "TestLimitsOwnerAndCapsFailBeforeGuard"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestAllApplicationStates": TestAllApplicationStates
        Case "TestStatusBarFalseAndString": TestStatusBarFalseAndString
        Case "TestActiveIdentityAfterSwitch": TestActiveIdentityAfterSwitch
        Case "TestRunnerNormalReturn": TestRunnerNormalReturn
        Case "TestRunnerForcedExecuteError": TestRunnerForcedExecuteError
        Case "TestRunnerUserCancel": TestRunnerUserCancel
        Case "TestRestoreTwice": TestRestoreTwice
        Case "TestClassTerminateRestores": TestClassTerminateRestores
        Case "TestInjectedRestoreFailuresContinue": TestInjectedRestoreFailuresContinue
        Case "TestRollbackJournalRestoresCell": TestRollbackJournalRestoresCell
        Case "TestFingerprintDriftRaisesNxInputErrorBeforeGuard": TestFingerprintDriftRaisesNxInputErrorBeforeGuard
        Case "TestSealedPlanRejectsMutation": TestSealedPlanRejectsMutation
        Case "TestExactResultContracts": TestExactResultContracts
        Case "TestRestoreFailureDowngradesSuccessWithoutChange": TestRestoreFailureDowngradesSuccessWithoutChange
        Case "TestRestoreFailureDowngradesSuccessWithChange": TestRestoreFailureDowngradesSuccessWithChange
        Case "TestRestoreFailureLeavesNonSuccess": TestRestoreFailureLeavesNonSuccess
        Case "TestBuildPlanMutationIsRejectedBeforeGuard": TestBuildPlanMutationIsRejectedBeforeGuard
        Case "TestRunnerRejectsEquivalentButDifferentPlan": TestRunnerRejectsEquivalentButDifferentPlan
        Case "TestApprovalObjectBindsFeaturePlanDigestAndContext": TestApprovalObjectBindsFeaturePlanDigestAndContext
        Case "TestRunnerRevalidatesFeatureBeforeGuard": TestRunnerRevalidatesFeatureBeforeGuard
        Case "TestRunnerInvalidExecuteResultIsNotPreflight": TestRunnerInvalidExecuteResultIsNotPreflight
        Case "TestInvalidResultFeatureIdRollsBack": TestInvalidResultFeatureIdRollsBack
        Case "TestPostChangeExecuteExceptionAndRollbackFailureIsPartial": TestPostChangeExecuteExceptionAndRollbackFailureIsPartial
        Case "TestCompletedAndRecoveryReportActualWork": TestCompletedAndRecoveryReportActualWork
        Case "TestStateGuardOppositeStateBaseline": TestStateGuardOppositeStateBaseline
        Case "TestStateGuardStatusBarFalseAndString": TestStateGuardStatusBarFalseAndString
        Case "TestStateGuardRestoreFailureAggregates": TestStateGuardRestoreFailureAggregates
        Case "TestStateGuardTerminateEmergencyCleanup": TestStateGuardTerminateEmergencyCleanup
        Case "TestApprovalObjectIsSingleUse": TestApprovalObjectIsSingleUse
        Case "TestApprovalRejectsEquivalentObjects": TestApprovalRejectsEquivalentObjects
        Case "TestApprovalBindsRunSnapshotSourceDigest": TestApprovalBindsRunSnapshotSourceDigest
        Case "TestFinalStateResultMatrix": TestFinalStateResultMatrix
        Case "TestFormula2FallbackOnlyWhenUnsupported": TestFormula2FallbackOnlyWhenUnsupported
        Case "TestFormula2RestoreNeverSilentlyDowngrades": TestFormula2RestoreNeverSilentlyDowngrades
        Case "TestTypedNonCellDispatchAndCompensation": TestTypedNonCellDispatchAndCompensation
        Case "TestLimitsOwnerAndCapsFailBeforeGuard": TestLimitsOwnerAndCapsFailBeforeGuard
        Case Else: Err.Raise vbObjectError + 712, "T_Core", "Unknown test: " & name
    End Select
End Sub

Private Sub TestAllApplicationStates()
    Dim guard As CNxStateGuard
    Dim screen As Boolean, events As Boolean, alerts As Boolean, calculation As XlCalculation
    screen = Application.ScreenUpdating: events = Application.EnableEvents: alerts = Application.DisplayAlerts: calculation = Application.Calculation
    Set guard = New CNxStateGuard
    Application.ScreenUpdating = Not screen: Application.EnableEvents = Not events: Application.DisplayAlerts = Not alerts: Application.Calculation = OppositeCalculation(calculation)
    guard.Restore
    NxTestHarness.AssertTrue Application.ScreenUpdating = screen, "ScreenUpdating restore failed"
    NxTestHarness.AssertTrue Application.EnableEvents = events, "EnableEvents restore failed"
    NxTestHarness.AssertTrue Application.DisplayAlerts = alerts, "DisplayAlerts restore failed"
    NxTestHarness.AssertTrue Application.Calculation = calculation, "Calculation restore failed"
End Sub

Private Sub TestStatusBarFalseAndString()
    Dim guard As CNxStateGuard, original As Variant
    original = Application.StatusBar
    On Error GoTo Cleanup
    Application.StatusBar = False
    Set guard = New CNxStateGuard
    Application.StatusBar = "Task3 status"
    guard.Restore
    NxTestHarness.AssertTrue Application.StatusBar = False, "False StatusBar restore failed"
    Set guard = Nothing
    Application.StatusBar = "captured status"
    Set guard = New CNxStateGuard
    Application.StatusBar = False
    guard.Restore
    NxTestHarness.AssertTrue Application.StatusBar = "captured status", "String StatusBar restore failed"
Cleanup:
    Application.StatusBar = original
    If Err.Number <> 0 Then Err.Raise Err.Number, "T_Core", Err.Description
End Sub

Private Sub TestActiveIdentityAfterSwitch()
    Dim guard As CNxStateGuard, originalBook As Workbook, otherBook As Workbook, originalSheet As Object, originalSelection As Range, originalAddress As String
    Set originalBook = Application.ActiveWorkbook: Set originalSheet = Application.ActiveSheet: Set originalSelection = Application.Selection: originalAddress = originalSelection.Address(False, False, xlA1, True)
    Set guard = New CNxStateGuard
    On Error GoTo Cleanup
    Set otherBook = Application.Workbooks.Add
    otherBook.Worksheets(1).Range("B2").Select
    guard.Restore
    NxTestHarness.AssertSameObject originalBook, Application.ActiveWorkbook, "ActiveWorkbook restore failed"
    NxTestHarness.AssertSameObject originalSheet, Application.ActiveSheet, "ActiveSheet restore failed"
    NxTestHarness.AssertTrue Application.Selection.Address(False, False, xlA1, True) = originalAddress, "Selection address restore failed"
Cleanup:
    If Not otherBook Is Nothing Then otherBook.Close False
    If Err.Number <> 0 Then Err.Raise Err.Number, "T_Core", Err.Description
End Sub

Private Sub TestRunnerNormalReturn()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    Set command = fake: Set result = RunApproved(command)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess And fake.ExecuteCount = 1, "Normal runner result failed"
End Sub

Private Sub TestRunnerForcedExecuteError()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    fake.ForceExecuteError = True: Set command = fake: Set result = RunApproved(command)
    NxTestHarness.AssertTrue result.Outcome = NxEnvironmentError And InStr(result.Recovery, "Forced Execute error") > 0, "Execute error was not captured"
End Sub

Private Sub TestRunnerUserCancel()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    fake.Outcome = NxCancelled: Set command = fake: Set result = RunApproved(command)
    NxTestHarness.AssertTrue result.Outcome = NxCancelled, "User cancel changed outcome"
End Sub

Private Sub TestRestoreTwice()
    Dim guard As CNxStateGuard, original As Boolean
    original = Application.ScreenUpdating: Set guard = New CNxStateGuard: Application.ScreenUpdating = Not original
    guard.Restore: guard.Restore
    NxTestHarness.AssertTrue Application.ScreenUpdating = original And Len(guard.Recovery) = 0, "Second restore changed state"
End Sub

Private Sub TestClassTerminateRestores()
    Dim guard As CNxStateGuard, original As Boolean
    original = Application.EnableEvents: Set guard = New CNxStateGuard: Application.EnableEvents = Not original
    Set guard = Nothing
    NxTestHarness.AssertTrue Application.EnableEvents = original, "Class_Terminate did not restore"
End Sub

Private Sub TestInjectedRestoreFailuresContinue()
    Dim guard As CNxStateGuard, alerts As Boolean
    alerts = Application.DisplayAlerts: Set guard = New CNxStateGuard: Application.DisplayAlerts = Not alerts
    guard.InjectRestoreFailure "ScreenUpdating": guard.InjectRestoreFailure "EnableEvents": guard.Restore
    NxTestHarness.AssertTrue InStr(guard.Recovery, "ScreenUpdating|") > 0 And InStr(guard.Recovery, "EnableEvents|") > 0 And InStr(guard.Recovery, "Injected restore failure") > 0, "Aggregate recovery missing property error detail"
    NxTestHarness.AssertTrue Application.DisplayAlerts = alerts, "Restore did not continue after injected failure"
End Sub

Private Sub TestRollbackJournalRestoresCell()
    Dim target As Range, originalValue As Variant, originalFormula As Variant, originalFormat As String
    Set target = Application.ActiveCell: target.Value2 = 17: target.NumberFormat = "0.00"
    originalValue = target.Value2: originalFormula = target.Formula: originalFormat = target.NumberFormat
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    fake.JournalActiveCell = True: fake.SourceChanged = True: fake.Completed = True: fake.Outcome = NxCancelled
    Set command = fake: Set result = RunApproved(command)
    NxTestHarness.AssertTrue target.Value2 = originalValue, "Rollback Value2 mismatch: expected=" & CStr(originalValue) & " actual=" & CStr(target.Value2)
    NxTestHarness.AssertTrue target.Formula = originalFormula, "Rollback Formula mismatch: expected=" & CStr(originalFormula) & " actual=" & CStr(target.Formula)
    NxTestHarness.AssertTrue target.NumberFormat = originalFormat, "Rollback NumberFormat mismatch: expected=" & originalFormat & " actual=" & CStr(target.NumberFormat)
    NxTestHarness.AssertTrue result.SourceChanged = False, "Rollback final-state mismatch: outcome=" & CStr(result.Outcome) & " recovery=" & result.Recovery
End Sub

Private Sub TestFingerprintDriftRaisesNxInputErrorBeforeGuard()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult, approval As CNxApproval
    Set command = fake: Set approval = NxCommandRunner.Approve(command)
    Application.ActiveCell.Offset(0, 1).Select
    Set result = NxCommandRunner.Run(command, approval)
    NxTestHarness.AssertTrue result.Outcome = NxInputError And fake.ExecuteCount = 0, "Changed fingerprint reached Execute"
End Sub

Private Sub TestSealedPlanRejectsMutation()
    Dim plan As CNxPlan, effect As CNxSideEffect, failed As Boolean
    Set plan = NxCreatePlan("NX-TEST-CORE", "context", "test-source")
    Set effect = NxCreateCellEffect("W=0:S=0:O=0:F=0", NxCellValueMask)
    plan.AddSideEffect effect
    plan.Seal
    On Error Resume Next
    plan.AddSideEffect effect: failed = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    NxTestHarness.AssertTrue failed And plan.SideEffects.Count = 1, "Sealed plan accepted mutation"
End Sub

Private Sub TestExactResultContracts()
    Dim result As CNxResult
    Set result = NxCreateResult("NX-TEST-CORE", NxSuccess, "complete", "A1", False, vbNullString, "result_success")
    NxTestHarness.AssertTrue result.FeatureId = "NX-TEST-CORE" And result.Outcome = NxSuccess And result.Stage = "complete" And result.Target = "A1" And result.Completed.Count = 0 And result.SourceChanged = False And result.Recovery = vbNullString And result.MessageKey = "result_success", "Result fields are not exact"
End Sub

Private Sub TestRestoreFailureDowngradesSuccessWithoutChange()
    Dim result As New CNxResult
    result.Configure "NX-TEST-CORE", NxSuccess, "complete", "A1", False, vbNullString, "result_success": result.MergeRestoreFailures "ScreenUpdating|1|test": result.Seal
    NxTestHarness.AssertTrue result.Outcome = NxPartialFailure And result.SourceChanged = False, "Restore failure did not become partial failure"
End Sub

Private Sub TestRestoreFailureDowngradesSuccessWithChange()
    Dim result As New CNxResult
    result.Configure "NX-TEST-CORE", NxSuccess, "complete", "A1", True, vbNullString, "result_success": result.AddCompleted "A1": result.MergeRestoreFailures "ScreenUpdating|1|test": result.Seal
    NxTestHarness.AssertTrue result.Outcome = NxPartialFailure And result.Completed.Count = 1, "Restore failure did not downgrade changed success"
End Sub

Private Sub TestRestoreFailureLeavesNonSuccess()
    Dim result As New CNxResult
    result.Configure "NX-TEST-CORE", NxCancelled, "execute", "A1", False, vbNullString, "result_cancelled": result.MergeRestoreFailures "ScreenUpdating|1|test": result.Seal
    NxTestHarness.AssertTrue result.Outcome = NxPartialFailure, "Restore failure without final source change must be partial failure"
End Sub

Private Sub TestBuildPlanMutationIsRejectedBeforeGuard()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, failed As Boolean, approval As CNxApproval
    fake.MutateContext = True
    Set command = fake
    On Error Resume Next
    Set approval = NxCommandRunner.Approve(command)
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    NxTestHarness.AssertTrue failed And fake.ExecuteCount = 0, "BuildPlan mutation must fail before guard and execute"
End Sub

Private Sub TestRunnerRejectsEquivalentButDifferentPlan()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    Set command = fake
    Set result = RunApproved(command)
    NxTestHarness.AssertTrue fake.ReceivedApprovedPlanIsBuiltPlan, "Runner must execute the same approved plan object instance"
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Identity fixture baseline did not execute"
End Sub

Private Sub TestApprovalObjectBindsFeaturePlanDigestAndContext()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    Dim legitimate As CNxApproval, badApproval As New CNxApproval, forgedContext As CNxExecutionContext, forgedAuthority As New Collection
    Set command = fake
    Set legitimate = NxCommandRunner.Approve(command)
    Set forgedContext = NxContextFactory.CaptureCurrent()
    badApproval.Configure forgedAuthority, command, legitimate.Plan, forgedContext
    Set result = NxCommandRunner.Run(command, badApproval)
    NxTestHarness.AssertTrue result.Outcome = NxInputError And fake.ExecuteCount = 0, "Forged issuer authority must stop before guard and execute"
End Sub

Private Sub TestRunnerRevalidatesFeatureBeforeGuard()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    fake.ChangeFeatureAfterPlan = True
    Set command = fake
    Set result = RunApproved(command)
    NxTestHarness.AssertTrue result.Outcome = NxInputError And fake.ExecuteCount = 0, "Feature must be revalidated immediately before guard"
End Sub

Private Sub TestRunnerInvalidExecuteResultIsNotPreflight()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    fake.ReturnUnsealedResult = True
    Set command = fake
    Set result = RunApproved(command)
    NxTestHarness.AssertTrue result.Outcome = NxEnvironmentError And result.Stage = "exception" And fake.ExecuteCount = 1, "Invalid Execute result must not become a preflight input error"
End Sub

Private Sub TestInvalidResultFeatureIdRollsBack()
    Dim scratch As Workbook, target As Range, fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    Dim originalValue As Variant, trappedNumber As Long, trappedDescription As String
    On Error GoTo Cleanup
    Set scratch = Application.Workbooks.Add
    scratch.Activate
    Set target = scratch.Worksheets(1).Range("A1")
    target.Value2 = "before-invalid-result"
    originalValue = target.Value2
    target.Select
    fake.JournalActiveCell = True
    fake.SourceChanged = True
    fake.ReturnWrongFeatureResult = True
    Set command = fake
    Set result = RunApproved(command)
    NxTestHarness.AssertTrue result.Outcome = NxEnvironmentError And result.SourceChanged = False, "Invalid result feature must retain runner exception classification after successful rollback"
    NxTestHarness.AssertTrue target.Value2 = originalValue, "Invalid result feature did not roll back changed cell"
Cleanup:
    trappedNumber = Err.Number: trappedDescription = Err.Description: Err.Clear
    On Error Resume Next
    If Not scratch Is Nothing Then scratch.Close False
    On Error GoTo 0
    If trappedNumber <> 0 Then Err.Raise trappedNumber, "T_Core", trappedDescription
End Sub

Private Sub TestPostChangeExecuteExceptionAndRollbackFailureIsPartial()
    Dim scratch As Workbook, target As Range, fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    Dim trappedNumber As Long, trappedDescription As String
    On Error GoTo Cleanup
    Set scratch = Application.Workbooks.Add
    scratch.Activate
    Set target = scratch.Worksheets(1).Range("A1")
    target.Select
    target.Value2 = "before-post-change-error"
    fake.JournalActiveCell = True
    fake.ForcePostChangeError = True
    fake.ForceRollbackFailure = True
    Set command = fake
    Set result = RunApproved(command)
    NxTestHarness.AssertTrue result.Outcome = NxPartialFailure And result.SourceChanged, "Post-change exception plus real rollback failure must be partial failure"
    NxTestHarness.AssertTrue InStr(result.Recovery, "Forced post-change Execute error") > 0 And InStr(result.Recovery, "Value2|") > 0, "Recovery must preserve the real rollback property failure"
Cleanup:
    trappedNumber = Err.Number: trappedDescription = Err.Description: Err.Clear
    On Error Resume Next
    If Not scratch Is Nothing Then scratch.Worksheets(1).Unprotect "task3-rollback"
    If Not scratch Is Nothing Then scratch.Close False
    On Error GoTo 0
    If trappedNumber <> 0 Then Err.Raise trappedNumber, "T_Core", trappedDescription
End Sub

Private Sub TestCompletedAndRecoveryReportActualWork()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, result As CNxResult
    fake.JournalActiveCell = True
    fake.SourceChanged = True
    fake.Completed = True
    fake.Outcome = NxCancelled
    Set command = fake
    Set result = RunApproved(command)
    NxTestHarness.AssertTrue result.Outcome = NxCancelled And result.Completed.Count = 1 And result.Completed.Item(1) = "fake-completed", "Returned non-success must retain actual completed work"
    NxTestHarness.AssertTrue result.SourceChanged = False And result.Recovery = vbNullString, "Successful cancelled rollback must report no final source change"
End Sub

Private Sub TestStateGuardOppositeStateBaseline()
    TestAllApplicationStates
End Sub

Private Sub TestStateGuardStatusBarFalseAndString()
    TestStatusBarFalseAndString
End Sub

Private Sub TestStateGuardRestoreFailureAggregates()
    TestInjectedRestoreFailuresContinue
End Sub

Private Sub TestStateGuardTerminateEmergencyCleanup()
    TestClassTerminateRestores
End Sub

Private Sub TestApprovalObjectIsSingleUse()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, approval As CNxApproval, result As CNxResult, failed As Boolean
    Set command = fake
    Set approval = NxCommandRunner.Approve(command)
    Set result = NxCommandRunner.Run(command, approval)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Exact approval did not execute"
    Set result = NxCommandRunner.Run(command, approval)
    NxTestHarness.AssertTrue result.Outcome = NxInputError And fake.ExecuteCount = 1, "Approval capability was reusable"
    On Error Resume Next
    approval.Configure New Collection, command, approval.Plan, NxContextFactory.CaptureCurrent
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Approval capability was reconfigured"
End Sub

Private Sub TestApprovalRejectsEquivalentObjects()
    Dim result As CNxResult, failed As Boolean
    Set result = NxCreateResult("NX-TEST-CORE", NxSuccess, "complete", "A1", False, vbNullString, "result_success")
    On Error Resume Next
    result.ValidateFor "NX-TEST-CORE", "B1"
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Result target validation accepted a different target"
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, other As New CFakeFeatureCommand, approval As CNxApproval, runResult As CNxResult
    Set command = fake: Set approval = NxCommandRunner.Approve(command)
    Set runResult = NxCommandRunner.Run(other, approval)
    NxTestHarness.AssertTrue runResult.Outcome = NxInputError And fake.ExecuteCount = 0, "Equivalent command object bypassed approval identity"
End Sub

Private Sub TestApprovalBindsRunSnapshotSourceDigest()
    Dim fake As New CFakeFeatureCommand, command As INxFeatureCommand, approval As CNxApproval, result As CNxResult
    Set command = fake: Set approval = NxCommandRunner.Approve(command)
    Application.ActiveCell.Value2 = "source-digest-drift"
    Set result = NxCommandRunner.Run(command, approval)
    NxTestHarness.AssertTrue result.Outcome = NxInputError And result.Stage = "stale_check" And fake.ExecuteCount = 0, "Stale approval reached execution"
End Sub

Private Sub TestFinalStateResultMatrix()
    Dim result As New CNxResult
    result.Configure "NX-TEST-CORE", NxCancelled, "execute", "A1", False, vbNullString, "result_cancelled": result.Seal
    result.ValidateFor "NX-TEST-CORE", "A1"
    Set result = New CNxResult
    result.Configure "NX-TEST-CORE", NxCancelled, "execute", "A1", False, vbNullString, "result_cancelled": result.MergeRestoreFailures "restore|1|test": result.Seal
    NxTestHarness.AssertTrue result.Outcome = NxPartialFailure And InStr(result.Stage, "restore") > 0, "Every restore failure must be partial at latest stage"
End Sub

Private Sub TestFormula2FallbackOnlyWhenUnsupported()
    Dim journal As New CNxCellRollbackJournal
    NxTestHarness.AssertTrue journal.Formula2FallbackContract = "438|1004", "Formula2 fallback contract drifted"
End Sub

Private Sub TestFormula2RestoreNeverSilentlyDowngrades()
    Dim journal As New CNxCellRollbackJournal
    NxTestHarness.AssertTrue journal.Formula2FallbackContract = "438|1004", "Formula2 restore must not downgrade"
End Sub

Private Sub TestTypedNonCellDispatchAndCompensation()
    Dim dispatcher As New CNxSideEffectDispatcher
    Dim context As CNxExecutionContext, plan As CNxPlan, effect As New CNxSideEffect
    Dim authority As New Collection, target As String, recovery As String
    Dim trappedNumber As Long, trappedDescription As String
    target = Environ$("TEMP") & "\LHexcel-Task3-" & Replace(NxCreateRunUuid(), "-", vbNullString) & ".txt"
    On Error GoTo Cleanup
    Set context = NxContextFactory.CaptureCurrent()
    context.BindSelectionAuthority authority
    Set plan = NxCreatePlan("NX-TEST-CORE", context.Fingerprint, context.SourceStateDigest)
    effect.ConfigureTextAdapter NxFileCreate, target, "typed-payload", "delete-created-file", vbNullString, False
    effect.BindDecision "P01", 200000
    effect.Seal
    plan.AddSideEffect effect
    plan.Seal
    dispatcher.Configure authority
    dispatcher.Dispatch plan, context
    NxTestHarness.AssertTrue Len(Dir$(target)) > 0 And FileLen(target) = Len("typed-payload"), "Typed file-create adapter did not write the sealed payload"
    recovery = dispatcher.CompensateCompleted()
    NxTestHarness.AssertTrue Len(recovery) = 0 And Len(Dir$(target)) = 0, "Typed file-create compensator did not remove its completed effect"
    AssertCommandOwnedTemplateEffects context, authority
Cleanup:
    trappedNumber = Err.Number: trappedDescription = Err.Description: Err.Clear
    On Error Resume Next
    If Len(Dir$(target)) > 0 Then Kill target
    On Error GoTo 0
    If trappedNumber <> 0 Then Err.Raise trappedNumber, "T_Core", trappedDescription
End Sub

Private Sub AssertCommandOwnedTemplateEffects(ByVal context As CNxExecutionContext, ByVal authority As Collection)
    Dim packageEffect As CNxSideEffect, sheetEffect As CNxSideEffect
    Dim plan As CNxPlan, dispatcher As New CNxSideEffectDispatcher
    Dim packageTarget As String
    packageTarget = Environ$("TEMP") & "\LHexcel-command-owned.xlsx"
    Set packageEffect = NxCreateCommandOwnedEffect("NX-TPL-REGISTER-RANGE", NxBinaryPackageWrite, _
        packageTarget, "delete temp or restore prior package")
    Set sheetEffect = NxCreateCommandOwnedEffect("NX-TPL-LOAD", NxWorksheetCreate, _
        "workbook://active/new-sheet", "delete only the newly created sheet", 25)
    NxTestHarness.AssertTrue packageEffect.PayloadKind = NxPayloadNone And packageEffect.MaximumCells = 0, "Binary package became a text adapter"
    NxTestHarness.AssertTrue sheetEffect.RollbackStrategy = NxRollbackTypedCompensator And sheetEffect.MaximumCells = 25, "New sheet compensator or cell count mismatch"

    Set plan = NxCreatePlan("NX-TPL-LOAD", context.Fingerprint, context.SourceStateDigest)
    plan.DeclareCapability NxCapabilityRiskyFormula
    plan.AddSideEffect sheetEffect
    plan.Seal
    dispatcher.Configure authority
    dispatcher.Dispatch plan, context
    NxTestHarness.AssertTrue plan.DeclaredCapabilities = NxCapabilityRiskyFormula, "Risky formula marker was not sealed"
    NxTestHarness.AssertTrue Len(Dir$(packageTarget)) = 0, "Command-owned package was dispatched as text"
End Sub

Private Sub TestLimitsOwnerAndCapsFailBeforeGuard()
    Dim policy As New CNxLimitsPolicy
    NxTestHarness.AssertTrue policy.ExactOwnerFor("NX-TEST-CORE") = "P01" And policy.MaximumCellsFor("NX-TEST-CORE") = 200000, "Limits policy drifted"
    NxTestHarness.AssertTrue policy.MaximumAreasFor("NX-AI-SUMMARY") = 20 And policy.MaximumPromptCharsFor("NX-AI-SUMMARY") = 50000 And policy.MaximumPngCellsFor("NX-FILE-RANGE-PNG") = 10000 And policy.MaximumPngDimensionFor("NX-FILE-RANGE-PNG") = 16384, "Limits dimensional projection drifted"
End Sub

Private Function RunApproved(ByVal command As INxFeatureCommand) As CNxResult
    Dim approval As CNxApproval
    Set approval = NxCommandRunner.Approve(command)
    Set RunApproved = NxCommandRunner.Run(command, approval)
End Function

Private Function OppositeCalculation(ByVal baseline As XlCalculation) As XlCalculation
    If baseline = xlCalculationManual Then
        OppositeCalculation = xlCalculationAutomatic
    Else
        OppositeCalculation = xlCalculationManual
    End If
End Function
