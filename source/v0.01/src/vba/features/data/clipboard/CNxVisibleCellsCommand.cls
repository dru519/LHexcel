VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxVisibleCellsCommand"
Option Explicit
Implements INxFeatureCommand

Private mFeatureId As String
Private mPayload As String
Private mPriorClipboardText As String
Private mAddresses As Variant
Private mValues As Variant
Private mVisibleCellCount As Long
Private mConfigured As Boolean
Private mPlanBuilt As Boolean

Public Sub ConfigureCopy(ByVal payload As String, ByVal priorClipboardText As String, ByVal visibleCellCount As Long)
    EnsureMutable
    If visibleCellCount <= 0 Then NxRaiseContractError "Visible-cell copy requires at least one visible cell"
    mFeatureId = NX_FEATURE_DATA_COPY_VISIBLE
    mPayload = payload
    mPriorClipboardText = priorClipboardText
    mVisibleCellCount = visibleCellCount
    mConfigured = True
End Sub

Public Sub ConfigurePaste(ByVal addresses As Variant, ByVal values As Variant, ByVal visibleCellCount As Long)
    Dim valueCount As Long
    EnsureMutable
    If visibleCellCount <= 0 Then NxRaiseContractError "Visible-cell paste requires at least one visible cell"
    If NxVisibleCellsVariantCount(addresses) <> visibleCellCount Then NxRaiseContractError "Visible-cell address count mismatch"
    valueCount = NxVisibleCellsVariantCount(values)
    If valueCount <> 1 And valueCount <> visibleCellCount Then NxRaiseContractError "Clipboard value count does not match visible cells"
    mFeatureId = NX_FEATURE_DATA_PASTE_VISIBLE_VALUES
    mAddresses = addresses
    mValues = values
    mVisibleCellCount = visibleCellCount
    mConfigured = True
End Sub

Private Property Get INxFeatureCommand_FeatureId() As String
    EnsureConfigured
    INxFeatureCommand_FeatureId = mFeatureId
End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan
    EnsureConfigured
    If mPlanBuilt Then NxRaiseContractError "Visible-cell plan can be built only once"
    If context Is Nothing Or context.SelectionType <> "Range" Then NxRaiseContractError "Visible-cell command requires a range context"
    If mVisibleCellCount > context.ActualCellCount Then NxRaiseContractError "Visible-cell count exceeds the captured selection"

    Set plan = NxCreatePlan(mFeatureId, context.Fingerprint, context.SourceStateDigest)
    plan.DeclareObservedCells mVisibleCellCount
    Select Case mFeatureId
        Case NX_FEATURE_DATA_COPY_VISIBLE
            plan.AddSideEffect NxCreateFeatureClipboardEffect(mFeatureId, mPayload, mPriorClipboardText)
        Case NX_FEATURE_DATA_PASTE_VISIBLE_VALUES
            If context.ActualCellCount <> mVisibleCellCount Then NxRaiseContractError "Paste context must contain exactly the visible target cells"
            plan.AddSideEffect NxCreateFeatureCellEffect(mFeatureId, context.SelectionTargetIdentity, NxCellValueMask, mVisibleCellCount)
        Case Else
            NxRaiseContractError "Visible-cell command feature is outside the closed contract"
    End Select
    plan.Seal
    mPlanBuilt = True
    Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    EnsureConfigured
    If Not mPlanBuilt Then NxRaiseContractError "Visible-cell command requires its sealed plan"
    If context Is Nothing Or approvedPlan Is Nothing Then NxRaiseContractError "Visible-cell execution requires approved runtime facts"
    If approvedPlan.FeatureId <> mFeatureId Then NxRaiseContractError "Visible-cell approved plan identity mismatch"

    Select Case mFeatureId
        Case NX_FEATURE_DATA_COPY_VISIBLE
            Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, NxSuccess, "complete", context.WorkbookIdentity, False, vbNullString, "result_success")
        Case NX_FEATURE_DATA_PASTE_VISIBLE_VALUES
            NxApplyApprovedCellValueVector approvedPlan, context, mAddresses, mValues
            Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, NxSuccess, "complete", context.WorkbookIdentity, True, vbNullString, "result_success")
        Case Else
            NxRaiseContractError "Visible-cell command feature is outside the closed contract"
    End Select
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Visible-cell command is not configured"
End Sub

Private Sub EnsureMutable()
    If mConfigured Then NxRaiseContractError "Visible-cell command can be configured only once"
End Sub
