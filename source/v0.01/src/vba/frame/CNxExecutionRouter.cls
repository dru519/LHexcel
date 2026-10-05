VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxExecutionRouter"
Option Explicit

Public Function Prepare(ByVal definition As CNxFeatureDefinition, ByVal command As INxFeatureCommand) As CNxExecutionTicket
    Dim approval As CNxApproval
    Dim plan As CNxPlan
    Dim context As CNxExecutionContext
    Dim decision As CNxExecutionDecision
    Dim compatibility As CNxOfficeCompatibility
    Dim ticket As New CNxExecutionTicket
    Dim fileMutation As Boolean
    Dim externalDocumentMutation As Boolean
    Dim manualRecoveryMutation As Boolean
    Dim riskyFormula As Boolean
    Dim actualCellCount As Double
    Dim resultSheetCount As Long
    Dim existingContentOverwrite As Boolean

    If definition Is Nothing Or command Is Nothing Then NxRaiseContractError "Execution router requires definition and command"
    If Not definition.IsSealed Then NxRaiseContractError "Execution router requires a sealed feature definition"
    If definition.FeatureId <> command.FeatureId Then NxRaiseContractError "Execution router feature identity mismatch"
    Set approval = NxCommandRunner.Approve(command)
    Set plan = approval.Plan
    Set context = approval.Context
    InspectPlan plan, context, actualCellCount, fileMutation, externalDocumentMutation, manualRecoveryMutation, riskyFormula, resultSheetCount, existingContentOverwrite
    Set decision = NxExecutionGradePolicy.ResolveDecision(definition, actualCellCount, fileMutation, _
        externalDocumentMutation, manualRecoveryMutation, (command.FeatureId = "NX-AI-IMAGE"), riskyFormula, resultSheetCount, existingContentOverwrite, fileMutation)
    Set compatibility = NxProbeOfficeCompatibility(definition)
    If definition.FeatureId <> plan.FeatureId Or definition.FeatureId <> compatibility.FeatureId Then NxRaiseContractError "Execution router metadata identity mismatch"
    ticket.Configure definition, command, approval, decision, plan, compatibility
    Set Prepare = ticket
End Function

' A feature's Execute button is the user's approval. Plans remain sealed and single-use.
Public Function ExecuteUserAction(ByVal ticket As CNxExecutionTicket) As CNxResult
    Dim effect As CNxSideEffect, approval As CNxApproval
    If ticket Is Nothing Then NxRaiseContractError "실행할 작업이 없습니다."
    For Each effect In ticket.Plan.SideEffects
        If effect.Kind = NxFileReplace Then
            If MsgBox("기존 파일을 바꿉니다." & vbCrLf & effect.Target & vbCrLf & "덮어쓰시겠습니까?", _
                vbYesNo + vbExclamation + vbDefaultButton2, "내엑셀 - 파일 덮어쓰기") <> vbYes Then
                Set ExecuteUserAction = NxCreateResult(ticket.Command.FeatureId, NxCancelled, "consent", _
                    ticket.Approval.Context.WorkbookIdentity, False, vbNullString, "result_cancelled")
                Exit Function
            End If
        End If
    Next effect
    If ticket.Command.FeatureId = "NX-TPL-DELETE" Then
        If MsgBox("선택한 템플릿을 삭제하시겠습니까?", vbYesNo + vbExclamation + vbDefaultButton2, "내엑셀 - 템플릿 삭제") <> vbYes Then
            Set ExecuteUserAction = NxCreateResult(ticket.Command.FeatureId, NxCancelled, "consent", _
                ticket.Approval.Context.WorkbookIdentity, False, vbNullString, "result_cancelled")
            Exit Function
        End If
    End If
    Select Case ticket.Decision.ResolvedGrade
        Case NxExecutionFast: Set ExecuteUserAction = ExecuteFast(ticket)
        Case NxExecutionGuarded: Set ExecuteUserAction = ExecuteGuarded(ticket, True)
        Case NxExecutionPlanned
            Set approval = TakePlannedApproval(ticket, True)
            Set ExecuteUserAction = RunApproved(ticket, approval)
        Case Else: NxRaiseContractError "실행 등급을 확인할 수 없습니다."
    End Select
End Function

Public Function ExecuteFast(ByVal ticket As CNxExecutionTicket) As CNxResult
    Set ExecuteFast = RunTicket(ticket, NxExecutionFast, True)
End Function

Public Function ExecuteGuarded(ByVal ticket As CNxExecutionTicket, ByVal userConfirmed As Boolean) As CNxResult
    Set ExecuteGuarded = RunTicket(ticket, NxExecutionGuarded, userConfirmed)
End Function

Public Function TakePlannedApproval(ByVal ticket As CNxExecutionTicket, ByVal userConfirmed As Boolean) As CNxApproval
    If ticket Is Nothing Then NxRaiseContractError "Execution ticket is required"
    Set TakePlannedApproval = ticket.Consume(NxExecutionPlanned, userConfirmed)
End Function

Public Function RunApproved(ByVal ticket As CNxExecutionTicket, ByVal approval As CNxApproval) As CNxResult
    If ticket Is Nothing Or approval Is Nothing Then NxRaiseContractError "Approved execution requires ticket and approval"
    If Not ticket.IsConsumed Then NxRaiseContractError "Approved execution requires a consumed ticket"
    If ticket.Decision.ResolvedGrade <> NxExecutionPlanned Then NxRaiseContractError "Approved execution requires a planned ticket"
    If Not (approval Is ticket.Approval) Then NxRaiseContractError "Approved execution approval binding mismatch"
    Set RunApproved = ExecuteApprovedCore(ticket, approval)
End Function

Private Function RunTicket(ByVal ticket As CNxExecutionTicket, ByVal expectedGrade As NxExecutionGrade, _
    ByVal userConfirmed As Boolean) As CNxResult

    Dim approval As CNxApproval
    If ticket Is Nothing Then NxRaiseContractError "Execution ticket is required"
    Set approval = ticket.Consume(expectedGrade, userConfirmed)
    Set RunTicket = ExecuteApprovedCore(ticket, approval)
End Function

Private Function ExecuteApprovedCore(ByVal ticket As CNxExecutionTicket, ByVal approval As CNxApproval) As CNxResult
    If ticket Is Nothing Or approval Is Nothing Then NxRaiseContractError "Execution core requires ticket and approval"
    Set ExecuteApprovedCore = NxCommandRunner.Run(ticket.Command, approval)
End Function

Private Sub InspectPlan(ByVal plan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByRef actualCellCount As Double, ByRef fileMutation As Boolean, _
    ByRef externalDocumentMutation As Boolean, ByRef manualRecoveryMutation As Boolean, _
    ByRef riskyFormula As Boolean, ByRef resultSheetCount As Long, ByRef existingContentOverwrite As Boolean)

    Dim effect As CNxSideEffect
    If plan Is Nothing Or context Is Nothing Or Not plan.IsSealed Then NxRaiseContractError "Execution router requires sealed runtime facts"
    If plan.HasObservedCellCount Then actualCellCount = plan.ObservedCellCount
    riskyFormula = ((plan.DeclaredCapabilities And NxCapabilityRiskyFormula) <> 0)
    For Each effect In plan.SideEffects
        Select Case effect.Kind
            Case NxCellMutation
                actualCellCount = MaximumCount(actualCellCount, context.ActualCellCount)
            Case NxFileCreate, NxFileReplace
                fileMutation = True
                If effect.Kind = NxFileReplace Then existingContentOverwrite = True
            Case NxBinaryPackageWrite
                fileMutation = True
            Case NxDirectoryBatchCreate
                fileMutation = True
            Case NxWorksheetCreate, NxColumnInsert
                resultSheetCount = resultSheetCount + 1
                actualCellCount = MaximumCount(actualCellCount, effect.MaximumCells)
            Case NxExternalProcess
                externalDocumentMutation = True
                manualRecoveryMutation = True
            Case NxShapeInsert
                manualRecoveryMutation = True
            Case NxClipboardWrite, NxBrowserOpen
                ' Local clipboard and browser opening do not imply document mutation.
            Case Else
                NxRaiseContractError "Execution router found an unknown side effect"
        End Select
    Next effect
End Sub

Private Function MaximumCount(ByVal first As Double, ByVal second As Double) As Double
    If first >= second Then MaximumCount = first Else MaximumCount = second
End Function
