VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFakeFeatureCommand"
Option Explicit
Implements INxFeatureCommand

Private mOutcome As NxOutcome
Private mForceExecuteError As Boolean
Private mForcePostChangeError As Boolean
Private mForceRollbackFailure As Boolean
Private mExecuteCount As Long
Private mSourceChanged As Boolean
Private mCompleted As Boolean
Private mMutateContext As Boolean
Private mJournalActiveCell As Boolean
Private mReturnUnsealedPlan As Boolean
Private mReturnUnsealedResult As Boolean
Private mReturnWrongFeatureResult As Boolean
Private mBuiltPlan As CNxPlan
Private mReceivedApprovedPlanIsBuiltPlan As Boolean
Private mChangeFeatureAfterPlan As Boolean
Private mFeatureIdReads As Long
Private mSelectionDriftAfterApproval As Boolean

Private Sub Class_Initialize()
    mOutcome = NxSuccess
End Sub

Public Property Let Outcome(ByVal value As NxOutcome): mOutcome = value: End Property
Public Property Let ForceExecuteError(ByVal value As Boolean): mForceExecuteError = value: End Property
Public Property Let ForcePostChangeError(ByVal value As Boolean): mForcePostChangeError = value: End Property
Public Property Let ForceRollbackFailure(ByVal value As Boolean): mForceRollbackFailure = value: End Property
Public Property Get ExecuteCount() As Long: ExecuteCount = mExecuteCount: End Property
Public Property Let SourceChanged(ByVal value As Boolean): mSourceChanged = value: End Property
Public Property Let Completed(ByVal value As Boolean): mCompleted = value: End Property
Public Property Let MutateContext(ByVal value As Boolean): mMutateContext = value: End Property
Public Property Let JournalActiveCell(ByVal value As Boolean): mJournalActiveCell = value: End Property
Public Property Let ReturnUnsealedPlan(ByVal value As Boolean): mReturnUnsealedPlan = value: End Property
Public Property Let ReturnUnsealedResult(ByVal value As Boolean): mReturnUnsealedResult = value: End Property
Public Property Let ReturnWrongFeatureResult(ByVal value As Boolean): mReturnWrongFeatureResult = value: End Property
Public Property Let ChangeFeatureAfterPlan(ByVal value As Boolean): mChangeFeatureAfterPlan = value: End Property
Public Property Get ReceivedApprovedPlanIsBuiltPlan() As Boolean: ReceivedApprovedPlanIsBuiltPlan = mReceivedApprovedPlanIsBuiltPlan: End Property
Public Property Let SelectionDriftAfterApproval(ByVal value As Boolean): mSelectionDriftAfterApproval = value: End Property

Private Property Get INxFeatureCommand_FeatureId() As String
    mFeatureIdReads = mFeatureIdReads + 1
    If mChangeFeatureAfterPlan And mFeatureIdReads > 1 Then
        INxFeatureCommand_FeatureId = "NX-TEST-CHANGED"
    Else
        INxFeatureCommand_FeatureId = "NX-TEST-CORE"
    End If
End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan
    Dim effect As CNxSideEffect
    If mMutateContext Then
        Set plan = NxCreatePlan("NX-TEST-CORE", context.Fingerprint & "-changed", context.SourceStateDigest)
    Else
        Set plan = NxCreatePlan("NX-TEST-CORE", context.Fingerprint, context.SourceStateDigest)
    End If
    If mJournalActiveCell Or mSourceChanged Or mForcePostChangeError Then
        Set effect = NxCreateCellEffect(context.SelectionTargetIdentity, NxCellAllSupportedMask)
        plan.AddSideEffect effect
    End If
    If Not mReturnUnsealedPlan Then plan.Seal
    Set mBuiltPlan = plan
    Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    Dim result As New CNxResult
    mExecuteCount = mExecuteCount + 1
    mReceivedApprovedPlanIsBuiltPlan = (approvedPlan Is mBuiltPlan)
    If mForceExecuteError Then Err.Raise vbObjectError + 720, "CFakeFeatureCommand", "Forced Execute error"
    If mSourceChanged Then NxApplyApprovedCellValue approvedPlan, context, "fake-source-change"
    If mForcePostChangeError Then
        NxApplyApprovedCellValue approvedPlan, context, "fake-source-change-before-error"
        If mForceRollbackFailure Then Application.ActiveSheet.Protect "task3-rollback"
        Err.Raise vbObjectError + 721, "CFakeFeatureCommand", "Forced post-change Execute error"
    End If
    If mReturnWrongFeatureResult Then
        result.Configure "NX-TEST-WRONG", mOutcome, ResultStage(), context.WorkbookIdentity, mSourceChanged, vbNullString, "result_fake"
    Else
        result.Configure "NX-TEST-CORE", mOutcome, ResultStage(), context.WorkbookIdentity, mSourceChanged, vbNullString, "result_fake"
    End If
    If mCompleted Then result.AddCompleted "fake-completed"
    If Not mReturnUnsealedResult Then result.Seal
    Set INxFeatureCommand_Execute = result
End Function

Private Function ResultStage() As String
    If mOutcome = NxSuccess Then ResultStage = "complete" Else ResultStage = "execute"
End Function
