Option Explicit
Implements INxExecutionPlanView

Private mController As CNxExecutionFrameController
Private mAllowClose As Boolean

Public Sub BindController(ByVal controller As CNxExecutionFrameController)
    If controller Is Nothing Then NxRaiseContractError "Execution plan form requires a controller"
    If Not mController Is Nothing Then NxRaiseContractError "Execution plan form controller can be bound only once"
    Set mController = controller
End Sub

Private Sub INxExecutionPlanView_Render(ByVal model As CNxExecutionPlanViewModel)
    If model Is Nothing Then NxRaiseContractError "Execution plan form requires a model"
    If Not model.IsSealed Then NxRaiseContractError "Execution plan form requires a sealed model"
    lblCategoryPath.Caption = model.DisplayCategoryLabel
    lblFeatureName.Caption = model.DisplayFeatureName
    lblRisk.Caption = model.RiskLabel
    txtTarget.Text = model.TargetSummary
    txtAction.Text = model.ActionSummary
    txtResult.Text = model.ResultSummary
    txtScale.Text = model.ScaleSummary
    txtPrivacy.Text = model.PrivacyLabel
    txtConflict.Text = model.ConflictSummary
    txtRecovery.Text = model.RecoverySummary
    lblBlockReason.Caption = model.BlockReason
    cmdExecutePlan.Caption = model.ExecuteCaption
End Sub

Private Sub INxExecutionPlanView_SetExecuteEnabled(ByVal enabled As Boolean)
    cmdExecutePlan.Enabled = enabled
End Sub

Private Sub INxExecutionPlanView_FocusCancel()
    cmdCancel.SetFocus
End Sub

Private Sub INxExecutionPlanView_CloseView()
    mAllowClose = True
    Set mController = Nothing
    Unload Me
End Sub

Private Sub cmdBack_Click()
    EnsureController
    mController.Back
End Sub

Private Sub cmdCancel_Click()
    EnsureController
    mController.Cancel
End Sub

Private Sub cmdExecutePlan_Click()
    EnsureController
    mController.ExecutePlan
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    EnsureController
    mController.CloseRequested
End Sub

Private Sub EnsureController()
    If mController Is Nothing Then NxRaiseContractError "Execution plan form controller is not bound"
End Sub
