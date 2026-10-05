VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDrawingFeatureCommand"
Option Explicit
Implements INxFeatureCommand

Private mRequest As CNxDrawingRequest
Private mPlanBuilt As Boolean, mConfigured As Boolean

Public Sub Configure(ByVal request As CNxDrawingRequest)
    If mConfigured Or request Is Nothing Then NxRaiseContractError "표/그리기 명령 구성이 잘못되었습니다."
    Set mRequest = request: mConfigured = True
End Sub

Private Property Get INxFeatureCommand_FeatureId() As String
    EnsureConfigured: INxFeatureCommand_FeatureId = mRequest.FeatureId
End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan
    EnsureConfigured: If mPlanBuilt Then NxRaiseContractError "표/그리기 계획은 한 번만 만들 수 있습니다."
    mRequest.RequireContext context
    Set plan = NxCreatePlan(mRequest.FeatureId, context.Fingerprint, context.SourceStateDigest)
    plan.DeclareObservedCells mRequest.ActualCellCount
    If mRequest.FeatureId = NX_FEATURE_DRAW_FIT_PICTURE Then
        plan.AddSideEffect NxCreateFeaturePictureMutationEffect(mRequest.FeatureId, mRequest.PictureMutationTarget(context.SelectionTargetIdentity))
    Else
        plan.AddSideEffect NxCreateFeatureCellEffect(mRequest.FeatureId, context.SelectionTargetIdentity, NxCellFontMask Or NxCellInteriorMask Or NxCellAlignmentMask Or NxCellBorderMask, mRequest.ActualCellCount)
    End If
    plan.Seal: mPlanBuilt = True
    Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    Dim snapshot As New CNxFormatSnapshot
    Dim failureNumber As Long, failureDescription As String
    EnsureConfigured: If Not mPlanBuilt Or context Is Nothing Or approvedPlan Is Nothing Then NxRaiseContractError "표/그리기 명령은 승인된 계획이 필요합니다."
    If approvedPlan.FeatureId <> mRequest.FeatureId Then NxRaiseContractError "표/그리기 승인 계획의 기능 ID가 다릅니다."
    mRequest.RequireContext context
    On Error GoTo Failed
    ' The runner owns cell rollback; table options still capture column widths.
    ' Direct table calls keep a complete local snapshot. Clear-inner keeps its backup.
    If mRequest.FeatureId = NX_FEATURE_DRAW_CLEAR_INNER Then snapshot.Capture mRequest.Target
    Select Case mRequest.FeatureId
        Case NX_FEATURE_DRAW_TITLE_TABLE, NX_FEATURE_DRAW_BUSINESS_TABLE
            NxDrawApplyTableOptions mRequest.Target, mRequest.FeatureId, mRequest.DisplayMode, mRequest.TableStyle, _
                mRequest.HeaderRowCount, mRequest.HeaderFontSize, mRequest.HeaderAlignment, _
                mRequest.BodyFontSize, mRequest.BodyAlignment, mRequest.TotalRow, mRequest.AutoFitColumns
        Case NX_FEATURE_DRAW_CLEAR_INNER: NxDrawClearInner mRequest.Target
        Case NX_FEATURE_DRAW_FIT_PICTURE: NxDrawFitPicture mRequest.Picture, mRequest.Target, mRequest.Margin, mRequest.MoveAndSize
        Case Else: NxRaiseContractError "표/그리기 기능이 닫힌 계약에 없습니다."
    End Select
    Set INxFeatureCommand_Execute = NxCreateResult(mRequest.FeatureId, NxSuccess, "complete", context.WorkbookIdentity, True, vbNullString, "result_success")
    Exit Function
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description
    If mRequest.FeatureId = NX_FEATURE_DRAW_CLEAR_INNER Then
        On Error Resume Next
        snapshot.Restore
        On Error GoTo 0
    End If
    Err.Raise failureNumber, "CNxDrawingFeatureCommand", failureDescription
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "표/그리기 명령이 구성되지 않았습니다."
End Sub
