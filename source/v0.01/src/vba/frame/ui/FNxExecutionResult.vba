Option Explicit
Implements INxExecutionResultView

Private mController As CNxExecutionFrameController
Private mAllowClose As Boolean

Public Sub BindController(ByVal controller As CNxExecutionFrameController)
    If controller Is Nothing Then NxRaiseContractError "Execution result form requires a controller"
    If Not mController Is Nothing Then NxRaiseContractError "Execution result form controller can be bound only once"
    Set mController = controller
End Sub

Private Sub INxExecutionResultView_Render(ByVal model As CNxResultCardViewModel)
    If model Is Nothing Then NxRaiseContractError "Execution result form requires a model"
    If Not model.IsSealed Then NxRaiseContractError "Execution result form requires a sealed model"
    lblOutcome.Caption = model.OutcomeLabel
    txtCompleted.Text = model.CompletedSummary
    txtIncomplete.Text = model.IncompleteSummary
    lblSourceChanged.Caption = model.SourceChangedLabel
    txtRecovery.Text = model.RecoveryDisplaySummary
    txtResultPath.Text = "대상: " & model.TargetSummary
    If Len(Trim$(model.ResultPath)) > 0 Then
        txtResultPath.Text = txtResultPath.Text & vbCrLf & "결과 파일: " & model.ResultPath
    End If
    txtNextAction.Text = model.NextAction
    SetResultActionsEnabled Len(Trim$(model.ResultPath)) > 0
End Sub

Private Sub INxExecutionResultView_ShowView()
    Me.Show
End Sub

Private Sub INxExecutionResultView_CloseView()
    mAllowClose = True
    Set mController = Nothing
    Unload Me
End Sub

Private Sub cmdClose_Click()
    EnsureController
    mController.CloseResult
End Sub

Private Sub cmdOpenResult_Click()
    EnsureController
    mController.OpenResult
End Sub

Private Sub cmdCopyPath_Click()
    EnsureController
    mController.CopyResultPath
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    EnsureController
    mController.CloseResult
End Sub

Private Sub SetResultActionsEnabled(ByVal hasPath As Boolean)
    Dim enabled As Boolean
    EnsureController
    enabled = hasPath And mController.ResultActionsReady
    cmdOpenResult.Enabled = enabled
    cmdCopyPath.Enabled = enabled
End Sub

Private Sub EnsureController()
    If mController Is Nothing Then NxRaiseContractError "Execution result form controller is not bound"
End Sub
