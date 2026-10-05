VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxRoleStyleCommand"
Option Explicit
Implements INxFeatureCommand

Private mRequest As CNxRoleStyleRequest

Public Sub Configure(ByVal request As CNxRoleStyleRequest)
    If request Is Nothing Then NxRaiseContractError "role-style command request is required"
    Set mRequest = request
End Sub

Private Property Get INxFeatureCommand_FeatureId() As String: INxFeatureCommand_FeatureId = "NX-DRAW-ROLE-STYLE": End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan
    mRequest.RequireContext context
    Set plan = NxCreatePlan("NX-DRAW-ROLE-STYLE", context.Fingerprint, context.SourceStateDigest)
    plan.DeclareObservedCells mRequest.ActualCellCount
    plan.AddSideEffect NxCreateFeatureCellEffect("NX-DRAW-ROLE-STYLE", context.SelectionTargetIdentity, NxCellFontMask Or NxCellInteriorMask Or NxCellAlignmentMask Or NxCellBorderMask, mRequest.ActualCellCount)
    plan.Seal
    Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    Dim snapshot As New CNxRoleStyleJournal, result As CNxResult, errorNumber As Long, errorDescription As String
    mRequest.RequireContext context: snapshot.Capture mRequest.AffectedTarget
    On Error GoTo Failed
    NxRoleStyleApply mRequest.AffectedTarget, mRequest.RoleCode, mRequest.DisplayMode
    Set result = NxCreateResult("NX-DRAW-ROLE-STYLE", NxSuccess, "complete", context.WorkbookIdentity, True, vbNullString, "role_style_applied")
    Set INxFeatureCommand_Execute = result
    Exit Function
Failed:
    errorNumber = Err.Number: errorDescription = Err.Description
    On Error Resume Next: snapshot.Restore: On Error GoTo 0
    Err.Raise errorNumber, "CNxRoleStyleCommand", errorDescription
End Function
