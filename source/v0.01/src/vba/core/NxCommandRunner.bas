Attribute VB_Name = "NxCommandRunner"
Option Explicit

Private mActiveDispatcher As CNxSideEffectDispatcher
Private mActivePlan As CNxPlan
Private mActiveContext As CNxExecutionContext
Private mApprovalCapability As Collection
Private mSelectionCapability As Collection

Public Function NxCreatePlan(ByVal featureId As String, ByVal contextFingerprint As String, ByVal sourceStateDigest As String) As CNxPlan
    Dim plan As New CNxPlan
    If Len(sourceStateDigest) = 0 Then NxRaiseContractError "Plan requires the frozen captured source state"
    plan.Configure featureId, contextFingerprint, sourceStateDigest
    Set NxCreatePlan = plan
End Function

Public Function NxCreateCellEffect(ByVal canonicalTarget As String, ByVal mutationMask As NxCellMutationMask, Optional ByVal declaredMaxCells As Long = 200000) As CNxSideEffect
    Dim effect As New CNxSideEffect
    effect.ConfigureCell canonicalTarget, mutationMask, declaredMaxCells
    effect.BindDecision "P01", 200000
    effect.Seal
    Set NxCreateCellEffect = effect
End Function

Public Function NxCreateFeatureCellEffect(ByVal featureId As String, ByVal canonicalTarget As String, _
    ByVal mutationMask As NxCellMutationMask, ByVal declaredMaxCells As Long) As CNxSideEffect

    Dim effect As New CNxSideEffect
    Dim policy As New CNxLimitsPolicy
    If declaredMaxCells <= 0 Or declaredMaxCells > policy.MaximumCellsFor(featureId) Then NxRaiseContractError "Feature cell effect exceeds the exact limit"
    effect.ConfigureCell canonicalTarget, mutationMask, declaredMaxCells
    effect.BindDecision policy.ExactOwnerFor(featureId), policy.MaximumCellsFor(featureId)
    effect.Seal
    Set NxCreateFeatureCellEffect = effect
End Function

Public Function NxCreateFeatureClipboardEffect(ByVal featureId As String, ByVal payload As String, ByVal priorText As String) As CNxSideEffect
    Dim effect As New CNxSideEffect
    Dim policy As New CNxLimitsPolicy
    If featureId <> NX_FEATURE_DATA_COPY_VISIBLE Then NxRaiseContractError "Unicode clipboard effect is restricted to visible-cell copy"
    effect.ConfigureTextAdapter NxClipboardWrite, "clipboard://unicode-text", payload, _
        "restore prior Unicode clipboard text", priorText, True
    effect.BindDecision policy.ExactOwnerFor(featureId), policy.MaximumCellsFor(featureId)
    effect.Seal
    Set NxCreateFeatureClipboardEffect = effect
End Function

Public Function NxCreateFeatureShapeEffect(ByVal featureId As String, ByVal canonicalTarget As String) As CNxSideEffect
    Dim effect As New CNxSideEffect
    Dim policy As New CNxLimitsPolicy
    If policy.MaximumPicturesFor(featureId) <> 1 Then NxRaiseContractError "Feature is not allowed to insert a shape"
    effect.ConfigureShapeMutation canonicalTarget, "restore existing shape geometry and placement"
    effect.BindDecision policy.ExactOwnerFor(featureId), policy.MaximumCellsFor(featureId)
    effect.Seal
    Set NxCreateFeatureShapeEffect = effect
End Function

Public Function NxCreateFeaturePictureInsertBatchEffect(ByVal featureId As String, ByVal kind As NxSideEffectKind, _
    ByVal declaredItems As Long, ByVal canonicalTarget As String, ByVal compensator As String) As CNxSideEffect
    Dim effect As New CNxSideEffect
    Dim policy As New CNxLimitsPolicy
    If featureId <> "NX-DRAW-INSERT-PICTURE" Then NxRaiseContractError "Picture insert batch is restricted to NX-DRAW-INSERT-PICTURE"
    If kind <> NxShapeInsert Then NxRaiseContractError "Picture insert batch requires NxShapeInsert"
    If declaredItems < 1 Or declaredItems > policy.MaximumPicturesFor(featureId) Or declaredItems > 500 Then NxRaiseContractError "Picture insert batch must contain 1-500 items"
    effect.ConfigureShapeBatch canonicalTarget, declaredItems, compensator
    effect.BindDecision policy.ExactOwnerFor(featureId), policy.MaximumCellsFor(featureId)
    effect.Seal
    Set NxCreateFeaturePictureInsertBatchEffect = effect
End Function

Public Function NxCreateFeaturePictureMutationEffect(ByVal featureId As String, ByVal canonicalTarget As String) As CNxSideEffect
    Dim effect As New CNxSideEffect
    Dim policy As New CNxLimitsPolicy
    If featureId <> "NX-DRAW-FIT-PICTURE" Then NxRaiseContractError "Picture mutation is restricted to FIT-PICTURE"
    If policy.MaximumPicturesFor(featureId) <> 1 Then NxRaiseContractError "Feature is not allowed to mutate a picture"
    effect.ConfigureShapeMutation canonicalTarget, "restore existing shape geometry and placement"
    effect.BindDecision policy.ExactOwnerFor(featureId), policy.MaximumCellsFor(featureId)
    effect.Seal
    Set NxCreateFeaturePictureMutationEffect = effect
End Function

Public Function NxCreateDirectoryBatchEffect(ByVal featureId As String, ByVal canonicalTarget As String, ByVal items As Variant) As CNxSideEffect
    Dim effect As New CNxSideEffect
    Dim policy As New CNxLimitsPolicy
    effect.ConfigureDirectoryBatch canonicalTarget, items
    If effect.DeclaredItems > policy.MaximumDirectoryItemsFor(featureId) Then NxRaiseContractError "Directory batch exceeds the exact feature limit"
    effect.BindDecision policy.ExactOwnerFor(featureId), policy.MaximumCellsFor(featureId)
    effect.Seal
    Set NxCreateDirectoryBatchEffect = effect
End Function

Public Function NxCreateCommandOwnedEffect(ByVal featureId As String, ByVal kind As NxSideEffectKind, _
    ByVal canonicalTarget As String, ByVal compensator As String, _
    Optional ByVal declaredCells As Long = 0) As CNxSideEffect

    Dim effect As New CNxSideEffect
    Dim policy As New CNxLimitsPolicy
    effect.ConfigureCommandOwned kind, canonicalTarget, compensator, declaredCells
    effect.BindDecision policy.ExactOwnerFor(featureId), policy.MaximumCellsFor(featureId)
    effect.Seal
    Set NxCreateCommandOwnedEffect = effect
End Function

Public Function NxCreateResult(ByVal featureId As String, ByVal outcome As NxOutcome, ByVal stage As String, ByVal target As String, ByVal sourceChanged As Boolean, ByVal recovery As String, ByVal messageKey As String) As CNxResult
    Dim result As New CNxResult
    result.Configure featureId, outcome, stage, target, sourceChanged, recovery, messageKey
    result.Seal
    Set NxCreateResult = result
End Function

Public Sub NxApplyApprovedCellValue(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, ByVal value As Variant)
    If mActiveDispatcher Is Nothing Or mActivePlan Is Nothing Or mActiveContext Is Nothing Then NxRaiseContractError "No active typed dispatcher boundary"
    If Not (mActivePlan Is approvedPlan) Or Not (mActiveContext Is context) Then NxRaiseContractError "Command cannot bypass planned effects"
    mActiveDispatcher.ApplyCellValue approvedPlan, context, value
End Sub

Public Sub NxApplyApprovedCellValueVector(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByVal addresses As Variant, ByVal values As Variant)

    RequireActiveBoundary approvedPlan, context
    mActiveDispatcher.ApplyCellValueVector approvedPlan, context, addresses, values
End Sub

Public Sub NxApplyApprovedCellValues(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByVal rowOffset As Long, ByVal columnOffset As Long, ByVal rowCount As Long, ByVal columnCount As Long, _
    ByVal matrix As Variant)
    RequireActiveBoundary approvedPlan, context
    mActiveDispatcher.ApplyCellValues approvedPlan, context, rowOffset, columnOffset, rowCount, columnCount, matrix
End Sub

Public Sub NxApplyApprovedCellFormulas(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByVal rowOffset As Long, ByVal columnOffset As Long, ByVal rowCount As Long, ByVal columnCount As Long, _
    ByVal matrix As Variant)
    RequireActiveBoundary approvedPlan, context
    mActiveDispatcher.ApplyCellFormulas approvedPlan, context, rowOffset, columnOffset, rowCount, columnCount, matrix
End Sub

Public Sub NxApplyApprovedCellNumberFormats(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByVal rowOffset As Long, ByVal columnOffset As Long, ByVal rowCount As Long, ByVal columnCount As Long, _
    ByVal matrix As Variant)
    RequireActiveBoundary approvedPlan, context
    mActiveDispatcher.ApplyCellNumberFormats approvedPlan, context, rowOffset, columnOffset, rowCount, columnCount, matrix
End Sub

Public Sub NxApplyApprovedImageAsset(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByVal rowOffset As Long, ByVal columnOffset As Long, ByVal rowCount As Long, ByVal columnCount As Long, _
    ByVal assetPath As String)
    RequireActiveBoundary approvedPlan, context
    mActiveDispatcher.ApplyImageAsset approvedPlan, context, rowOffset, columnOffset, rowCount, columnCount, assetPath
End Sub

Private Sub RequireActiveBoundary(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext)
    If mActiveDispatcher Is Nothing Or mActivePlan Is Nothing Or mActiveContext Is Nothing Then NxRaiseContractError "No active typed dispatcher boundary"
    If Not (mActivePlan Is approvedPlan) Or Not (mActiveContext Is context) Then NxRaiseContractError "Command cannot bypass planned effects"
End Sub

Public Function NxRunnerOwnsTableRollback(ByVal featureId As String, ByVal target As Range) As Boolean
    Dim owned As Range, effect As Variant, requiredMask As Long
    If target Is Nothing Then Exit Function
    If mActiveDispatcher Is Nothing Or mActivePlan Is Nothing Or mActiveContext Is Nothing Then Exit Function
    If featureId <> "NX-DRAW-BUSINESS-TABLE" And featureId <> "NX-DRAW-TITLE-TABLE" Then Exit Function
    If mActivePlan.FeatureId <> featureId Then Exit Function
    Set owned = mActiveContext.InternalSelectionRange(mSelectionCapability)
    If Not target.Worksheet Is owned.Worksheet Then Exit Function
    If target.Address <> owned.Address Then Exit Function
    requiredMask = NxCellFontMask Or NxCellInteriorMask Or NxCellAlignmentMask Or NxCellBorderMask
    For Each effect In mActivePlan.SideEffects
        If effect.Kind = NxCellMutation And effect.Target = mActiveContext.SelectionTargetIdentity Then
            If (CLng(effect.MutationMask) And requiredMask) = requiredMask Then
                NxRunnerOwnsTableRollback = True
                Exit Function
            End If
        End If
    Next effect
End Function

Public Function Approve(ByVal command As INxFeatureCommand) As CNxApproval
    NxDistributionEnsureExecutable
    Dim context As CNxExecutionContext, plan As CNxPlan, approval As New CNxApproval
    If command Is Nothing Then NxRaiseContractError "Approval command is required"
    EnsureCapabilities
    Set context = NxContextFactory.CaptureForFeature(command.FeatureId)
    context.BindSelectionAuthority mSelectionCapability
    Set plan = command.BuildPlan(context)
    If plan Is Nothing Or Not plan.IsSealed Then NxRaiseContractError "Approval requires a sealed plan"
    If plan.FeatureId <> command.FeatureId Or plan.ContextFingerprint <> context.Fingerprint Then NxRaiseContractError "Approval plan does not match current context"
    If mApprovalCapability Is Nothing Then Set mApprovalCapability = New Collection
    approval.Configure mApprovalCapability, command, plan, context
    Set Approve = approval
End Function

Public Function Run(ByVal command As INxFeatureCommand, ByVal approval As CNxApproval) As CNxResult
    NxDistributionEnsureExecutable
    Dim context As CNxExecutionContext
    Dim approvedPlan As CNxPlan
    Dim guard As CNxStateGuard
    Dim journal As CNxCellRollbackJournal
    Dim result As CNxResult
    Dim featureId As String
    Dim target As String
    Dim dispatcher As CNxSideEffectDispatcher
    Dim rollbackFailures As String, residual As String
    Dim trappedError As String
    Dim sourceChanged As Boolean
    Dim executeBegan As Boolean
    On Error GoTo Failed

    If command Is Nothing Or approval Is Nothing Then GoTo InvalidInput
    EnsureCapabilities
    featureId = command.FeatureId
    If Len(Trim$(featureId)) = 0 Then GoTo InvalidInput
    Set context = NxContextFactory.CaptureForFeature(command.FeatureId)
    context.BindSelectionAuthority mSelectionCapability
    target = context.WorkbookIdentity
    Set approvedPlan = approval.Plan
    If approvedPlan Is Nothing Or Not approvedPlan.IsSealed Then GoTo InvalidInput
    If mApprovalCapability Is Nothing Then GoTo StaleInput
    If Not approval.ConsumeFor(mApprovalCapability, command, approvedPlan, context) Then GoTo StaleInput
    If Not ValidateApproval(command, approvedPlan, context, featureId) Then GoTo StaleInput
    target = NxExpectedResultTarget(context, approvedPlan)

    Set guard = New CNxStateGuard
    Set journal = New CNxCellRollbackJournal
    journal.Capture context, approvedPlan, mSelectionCapability
    Set dispatcher = New CNxSideEffectDispatcher
    dispatcher.Configure mSelectionCapability
    rollbackFailures = dispatcher.Dispatch(approvedPlan, context)
    executeBegan = True
    ' command.Execute is reachable only through the typed dispatcher boundary.
    Set mActiveDispatcher = dispatcher: Set mActivePlan = approvedPlan: Set mActiveContext = context
    Set result = dispatcher.ExecuteApproved(command, context, approvedPlan)
    If result Is Nothing Then Err.Raise NX_CONTRACT_ERROR, "NxCommandRunner", "Execute returned an invalid or unsealed result"
    If Not result.IsSealed Then Err.Raise NX_CONTRACT_ERROR, "NxCommandRunner", "Execute returned an invalid or unsealed result"
    result.ValidateFor featureId, target
    If result.Outcome <> NxSuccess Then
        rollbackFailures = journal.Rollback()
        rollbackFailures = NxJoinFailure(rollbackFailures, dispatcher.Recover())
        residual = journal.VerifyResidual
        sourceChanged = FinalSourceChanged(residual)
        If Len(residual) > 0 Then rollbackFailures = NxJoinFailure(rollbackFailures, residual)
        Set result = ReconcileResult(result, sourceChanged, rollbackFailures, False)
    Else
        ' A successful mutation is expected to differ from the capture.  Residual
        ' journal verification is rollback-only, so it must never downgrade this path.
        ' VBA Or evaluates both operands. A declared mutation already determines
        ' this result flag; preserve the full comparison for undeclared changes.
        sourceChanged = result.SourceChanged
        If Not sourceChanged Then sourceChanged = (context.SourceStateMatchesCurrent() = False)
        Set result = CompleteSuccess(result, sourceChanged, dispatcher)
    End If
    GoTo FinallyRestore

InvalidInput:
    If Len(featureId) = 0 Then featureId = "NX-UNKNOWN"
    Set result = NxCreateResult(featureId, NxInputError, "preflight", target, False, "Input or approval validation failed before guard/execute", "result_input_error")
    GoTo FinallyRestore

StaleInput:
    If Len(featureId) = 0 Then featureId = "NX-UNKNOWN"
    Set result = NxCreateResult(featureId, NxInputError, "stale_check", target, False, "Stale context or approval capability", "result_input_error")
    GoTo FinallyRestore

Failed:
    trappedError = CStr(Err.Number) & " " & Err.Description
    Err.Clear
    If Not journal Is Nothing Then
        rollbackFailures = journal.Rollback()
        If Not dispatcher Is Nothing Then rollbackFailures = NxJoinFailure(rollbackFailures, dispatcher.Recover())
        residual = journal.VerifyResidual
        sourceChanged = FinalSourceChanged(residual)
        If Len(residual) > 0 Then rollbackFailures = NxJoinFailure(rollbackFailures, residual)
    End If
    If Len(featureId) = 0 Then featureId = "NX-UNKNOWN"
    If sourceChanged Or Len(rollbackFailures) > 0 Then
        Set result = NxCreateResult(featureId, NxPartialFailure, "exception", target, sourceChanged, NxJoinFailure(NxJoinFailure(trappedError, rollbackFailures), "Preserve the source and perform the recorded manual recovery action"), "result_partial_failure")
    Else
        Set result = NxCreateResult(featureId, NxEnvironmentError, "exception", target, False, trappedError, "result_environment_error")
    End If

FinallyRestore:
    Set mActiveDispatcher = Nothing: Set mActivePlan = Nothing: Set mActiveContext = Nothing
    If Not guard Is Nothing Then
        guard.Restore
        If Len(guard.Recovery) > 0 Then Set result = MergeRecovery(result, guard.Recovery, result.SourceChanged, True)
    End If
    If result.Outcome <> NxSuccess And Not journal Is Nothing Then
        residual = journal.VerifyResidual
        If Len(residual) > 0 Then Set result = MergeRecovery(result, residual, FinalSourceChanged(residual), False)
    End If
    On Error Resume Next
    If Not journal Is Nothing Then journal.ReleaseResources
    If Err.Number <> 0 Then
        trappedError = "서식 백업 정리 실패: " & Err.Description
        Err.Clear
        Set result = MergeRecovery(result, trappedError, result.SourceChanged, True)
    End If
    On Error GoTo 0
    Set Run = result
End Function

Private Function CompleteSuccess(ByVal source As CNxResult, ByVal sourceChanged As Boolean, ByVal dispatcher As CNxSideEffectDispatcher) As CNxResult
    Dim result As New CNxResult
    Dim item As Variant
    Dim completedEffect As Variant
    result.Configure source.FeatureId, NxSuccess, "complete", source.Target, sourceChanged, source.Recovery, source.MessageKey
    For Each item In source.Completed
        result.AddCompleted CStr(item)
    Next item
    If Not dispatcher Is Nothing Then
        For Each completedEffect In dispatcher.CompletedDescriptions
            result.AddCompleted CStr(completedEffect)
        Next completedEffect
    End If
    result.Seal
    Set CompleteSuccess = result
End Function

Private Function ValidateApproval(ByVal command As INxFeatureCommand, ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, ByVal featureId As String) As Boolean
    Dim policy As New CNxLimitsPolicy
    If command Is Nothing Or approvedPlan Is Nothing Or context Is Nothing Then Exit Function
    If approvedPlan.FeatureId <> featureId Or approvedPlan.ContextFingerprint <> context.Fingerprint Then Exit Function
    If approvedPlan.Owner <> policy.ExactOwnerFor(featureId) Then Exit Function
    If approvedPlan.LimitDecision <> policy.MaximumCellsFor(featureId) Then Exit Function
    ValidateApproval = context.FingerprintMatchesCurrent()
End Function

Private Function NxExpectedResultTarget(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As String
    Dim effect As Variant
    Dim packageTarget As String
    If context Is Nothing Or approvedPlan Is Nothing Then NxRaiseContractError "Result target resolution requires context and plan"
    For Each effect In approvedPlan.SideEffects
        If effect.Kind = NxBinaryPackageWrite Then
            If Len(packageTarget) > 0 Then NxRaiseContractError "Result target resolution rejects multiple binary package outputs"
            packageTarget = effect.Target
        End If
    Next effect
    If Len(packageTarget) > 0 Then
        NxExpectedResultTarget = packageTarget
    Else
        NxExpectedResultTarget = context.WorkbookIdentity
    End If
End Function

Private Function FinalSourceChanged(ByVal residual As String) As Boolean
    FinalSourceChanged = (Len(residual) > 0)
End Function

Private Function MergeRecovery(ByVal source As CNxResult, ByVal failures As String, ByVal sourceChanged As Boolean, ByVal isRestoreFailure As Boolean) As CNxResult
    Dim stage As String
    Dim result As CNxResult
    Dim item As Variant
    stage = source.Stage
    If isRestoreFailure Then stage = stage & ";restore" Else stage = stage & ";rollback"
    Set result = New CNxResult
    result.Configure source.FeatureId, NxPartialFailure, stage, source.Target, sourceChanged, NxJoinFailure(source.Recovery, failures), source.MessageKey
    For Each item In source.Completed
        result.AddCompleted CStr(item)
    Next item
    result.Seal
    Set MergeRecovery = result
End Function

Private Sub EnsureCapabilities()
    If mApprovalCapability Is Nothing Then Set mApprovalCapability = New Collection
    If mSelectionCapability Is Nothing Then Set mSelectionCapability = New Collection
End Sub

Private Function ReconcileResult(ByVal source As CNxResult, ByVal sourceChanged As Boolean, ByVal failures As String, ByVal isRestoreFailure As Boolean) As CNxResult
    Dim result As CNxResult
    Dim item As Variant
    If Len(failures) > 0 Then
        Set ReconcileResult = MergeRecovery(source, failures, sourceChanged, isRestoreFailure)
        Exit Function
    End If
    If source.SourceChanged = sourceChanged Then
        Set ReconcileResult = source
        Exit Function
    End If
    Set result = New CNxResult
    result.Configure source.FeatureId, source.Outcome, source.Stage, source.Target, sourceChanged, source.Recovery, source.MessageKey
    For Each item In source.Completed
        result.AddCompleted CStr(item)
    Next item
    result.Seal
    Set ReconcileResult = result
End Function
