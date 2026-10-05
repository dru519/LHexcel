VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDataNormalizeCommand"
Option Explicit
Implements INxFeatureCommand

Private mRequest As CNxDataNormalizeRequest
Private mConfigured As Boolean
Private mPlanBuilt As Boolean
Private mDestination As New CNxDataOutput

Public Sub CompleteOutput(ByVal result As CNxResult)
    If result Is Nothing Then NxRaiseContractError "정규화 실행 결과가 없습니다."
    If result.Outcome <> NxSuccess Then
        DiscardOutput
    Else
        mDestination.ActivateResult
    End If
End Sub

Public Sub DiscardOutput()
    mDestination.Rollback
End Sub

Public Sub Configure(ByVal request As CNxDataNormalizeRequest)
    If mConfigured Then NxRaiseContractError "Data normalization command can be configured only once"
    If request Is Nothing Then NxRaiseContractError "Data normalization command requires a request"
    Set mRequest = request
    mDestination.Configure request.SourceRange, request.OutputMode
    mConfigured = True
End Sub
Private Property Get INxFeatureCommand_FeatureId() As String
    EnsureConfigured
    INxFeatureCommand_FeatureId = "NX-DATA-NORMALIZE"
End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan
    EnsureConfigured
    If mPlanBuilt Then NxRaiseContractError "Data normalization plan can be built only once"
    mRequest.RequireContext context
    Set plan = NxCreatePlan("NX-DATA-NORMALIZE", context.Fingerprint, context.SourceStateDigest)
    plan.DeclareObservedCells mRequest.ActualCellCount
    If mRequest.OutputMode = "원본 변경" Then
        plan.AddSideEffect NxCreateFeatureCellEffect("NX-DATA-NORMALIZE", context.SelectionTargetIdentity, _
            NxCellValueMask Or NxCellNumberFormatMask, mRequest.ActualCellCount)
    ElseIf mRequest.OutputMode = "오른쪽 새 열 삽입" Then
        plan.AddSideEffect NxCreateCommandOwnedEffect("NX-DATA-NORMALIZE", NxColumnInsert, context.WorkbookIdentity, _
            "delete only columns inserted by this conversion", mRequest.ActualCellCount)
    Else
        plan.AddSideEffect NxCreateCommandOwnedEffect("NX-DATA-NORMALIZE", NxWorksheetCreate, context.WorkbookIdentity, _
            "delete only the command-created normalization result sheet", mRequest.ActualCellCount)
    End If
    plan.Seal
    mPlanBuilt = True
    Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    Dim serviceResult As CNxResult
    EnsureConfigured
    If Not mPlanBuilt Then NxRaiseContractError "Data normalization command requires its sealed plan"
    If context Is Nothing Or approvedPlan Is Nothing Then NxRaiseContractError "Data normalization requires approved runtime facts"
    If approvedPlan.FeatureId <> "NX-DATA-NORMALIZE" Then NxRaiseContractError "Data normalization approved plan identity mismatch"
    mRequest.RequireContext context
    Set serviceResult = NxDataCreateNormalizationResult(mRequest, context, approvedPlan, mDestination)
    Set INxFeatureCommand_Execute = NxCreateResult(serviceResult.FeatureId, serviceResult.Outcome, serviceResult.Stage, _
        context.WorkbookIdentity, serviceResult.SourceChanged, serviceResult.Recovery, serviceResult.MessageKey)
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Data normalization command is not configured"
End Sub
