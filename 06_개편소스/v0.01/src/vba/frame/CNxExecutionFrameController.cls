VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxExecutionFrameController"
Option Explicit

Private mSession As CNxExecutionFrameSession
Private mPlanView As INxExecutionPlanView
Private mResultView As INxExecutionResultView
Private mPlanModel As CNxExecutionPlanViewModel
Private mTicket As CNxExecutionTicket
Private mTextCatalog As CNxFrameTextCatalog
Private mResult As CNxResult
Private mResultModel As CNxResultCardViewModel
Private mResultActions As INxResultActionAdapter
Private mExecutionCancel As INxExecutionCancelAdapter
Private mRouter As CNxExecutionRouter
Private mConfigured As Boolean

Public Sub Configure(ByVal planView As INxExecutionPlanView, _
    ByVal resultView As INxExecutionResultView, ByVal planModel As CNxExecutionPlanViewModel, _
    ByVal ticket As CNxExecutionTicket, _
    ByVal textCatalog As CNxFrameTextCatalog)
    If mConfigured Then NxRaiseContractError "Execution frame controller can be configured only once"
    If planView Is Nothing Or resultView Is Nothing Or planModel Is Nothing Or _
        ticket Is Nothing Or textCatalog Is Nothing Then
        NxRaiseContractError "Execution frame controller requires every dependency"
    End If
    If Not planModel.IsSealed Then NxRaiseContractError "Execution frame controller requires a sealed plan model"
    If planModel.FeatureId <> ticket.Definition.FeatureId Or _
        planModel.FeatureId <> ticket.Command.FeatureId Or _
        planModel.FeatureId <> ticket.Plan.FeatureId Or _
        planModel.FeatureId <> ticket.Compatibility.FeatureId Then NxRaiseContractError "Execution frame feature identity mismatch"
    Set mSession = New CNxExecutionFrameSession
    Set mPlanView = planView
    Set mResultView = resultView
    Set mPlanModel = planModel
    Set mTicket = ticket
    Set mTextCatalog = textCatalog
    Set mRouter = New CNxExecutionRouter
    mConfigured = True
End Sub

' Single side-effect execution facade. Feature owners must hand a sealed ticket here;
' grade-specific confirmation and the sole core runner call remain centralized.
Public Function ExecuteSideEffect(ByVal ticket As CNxExecutionTicket, ByVal userConfirmed As Boolean) As CNxResult
    Dim result As CNxResult
    Dim approval As CNxApproval
    If ticket Is Nothing Then NxRaiseContractError "Execution ticket is required"
    If mRouter Is Nothing Then Set mRouter = New CNxExecutionRouter
    Select Case ticket.Decision.ResolvedGrade
        Case NxExecutionFast
            Set result = mRouter.ExecuteFast(ticket)
        Case NxExecutionGuarded
            Set result = mRouter.ExecuteGuarded(ticket, userConfirmed)
        Case NxExecutionPlanned
            Set approval = mRouter.TakePlannedApproval(ticket, userConfirmed)
            Set result = mRouter.RunApproved(ticket, approval)
        Case Else
            NxRaiseContractError "Execution ticket grade is invalid"
    End Select
    If result Is Nothing Or Not result.IsSealed Then NxRaiseContractError "Execution facade returned an unsealed result"
    Set ExecuteSideEffect = result
End Function

Public Function ExecuteTicket(ByVal ticket As CNxExecutionTicket, ByVal userConfirmed As Boolean) As CNxResult
    Set ExecuteTicket = ExecuteSideEffect(ticket, userConfirmed)
End Function

Public Sub BindResultActionAdapter(ByVal adapter As INxResultActionAdapter)
    EnsureConfigured
    If adapter Is Nothing Then NxRaiseContractError "Result action adapter is required"
    If Not mResultActions Is Nothing Then NxRaiseContractError "Result action adapter can be bound only once"
    Set mResultActions = adapter
End Sub

Public Sub BindExecutionCancelAdapter(ByVal adapter As INxExecutionCancelAdapter)
    EnsureConfigured
    If adapter Is Nothing Then NxRaiseContractError "Execution cancel adapter is required"
    If Not mExecutionCancel Is Nothing Then NxRaiseContractError "Execution cancel adapter can be bound only once"
    Set mExecutionCancel = adapter
End Sub

Public Sub ShowPlan()
    EnsureConfigured
    mSession.BeginChecking
    mPlanView.Render mPlanModel
    If mPlanModel.CanExecute Then
        mSession.MarkReady
        mPlanView.SetExecuteEnabled True
    Else
        mSession.MarkBlocked
        mPlanView.SetExecuteEnabled False
        mPlanView.FocusCancel
    End If
End Sub

Public Sub ExecutePlan()
    Dim result As CNxResult
    Dim approval As CNxApproval
    EnsureConfigured
    If mTicket.IsConsumed Then NxRaiseContractError "Planned execution ticket was already consumed"
    mSession.BeginRunning
    mPlanView.SetExecuteEnabled False
    Set approval = mRouter.TakePlannedApproval(mTicket, True)
    mPlanView.CloseView
    Set result = mRouter.RunApproved(mTicket, approval)
    If result Is Nothing Then NxRaiseContractError "Core runner returned no result"
    If Not result.IsSealed Then NxRaiseContractError "Core runner returned an unsealed result"
    Set mResult = result
    mSession.CompleteFromResult result
    Set mResultModel = mTextCatalog.BuildResultModel(result, ResultFilePath(result), mPlanModel.TargetSummary)
    mResultView.Render mResultModel
    mResultView.ShowView
End Sub

Private Function ResultFilePath(ByVal result As CNxResult) As String
    Dim effect As CNxSideEffect, candidate As String, fileSystem As Object
    If result.Outcome <> NxSuccess Then Exit Function
    On Error GoTo NoPath
    For Each effect In mTicket.Plan.SideEffects
        Select Case effect.Kind
            Case NxFileCreate, NxFileReplace, NxBinaryPackageWrite
                candidate = effect.Target
                If StrComp(candidate, result.Target, vbBinaryCompare) = 0 Then
                    If candidate Like "[A-Za-z]:\*" Or Left$(candidate, 2) = "\\" Then
                        Set fileSystem = CreateObject("Scripting.FileSystemObject")
                        If fileSystem.FileExists(candidate) Then
                            ResultFilePath = candidate
                            Exit Function
                        End If
                    End If
                End If
        End Select
    Next effect
    Exit Function
NoPath:
    Err.Clear
End Function

Public Sub Back()
    EnsureConfigured
    mSession.ReturnToInput
    mPlanView.CloseView
End Sub

Public Sub Cancel()
    EnsureConfigured
    If mSession.State = NxFrameRunning Then
        If mExecutionCancel Is Nothing Then NxRaiseContractError "Running execution cannot be cancelled by this feature"
        mExecutionCancel.RequestCancel
        Exit Sub
    End If
    mSession.Cancel
    mPlanView.CloseView
End Sub

Public Sub CloseRequested()
    Cancel
End Sub

Public Sub OpenResult()
    EnsureResultActionsReady
    mResultActions.OpenPath mResultModel.ResultPath
End Sub

Public Sub CopyResultPath()
    EnsureResultActionsReady
    mResultActions.CopyPath mResultModel.ResultPath
End Sub

Public Sub CloseResult()
    EnsureConfigured
    If Not mSession.IsTerminal Then NxRaiseContractError "Result view can close only after execution completes"
    mResultView.CloseView
End Sub

Public Property Get ResultActionsReady() As Boolean
    EnsureConfigured
    If mResultActions Is Nothing Then Exit Property
    If mResultModel Is Nothing Then Exit Property
    If Not mResultModel.IsSealed Then Exit Property
    ResultActionsReady = Len(Trim$(mResultModel.ResultPath)) > 0
End Property

Public Property Get State() As NxFrameState
    EnsureConfigured
    State = mSession.State
End Property

Public Property Get ResultModel() As CNxResultCardViewModel
    EnsureConfigured
    Set ResultModel = mResultModel
End Property

Public Property Get Result() As CNxResult
    EnsureConfigured
    Set Result = mResult
End Property

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Execution frame controller is not configured"
End Sub

Private Sub EnsureResultActionsReady()
    EnsureConfigured
    If Not ResultActionsReady Then NxRaiseContractError "Result action is not available"
End Sub
