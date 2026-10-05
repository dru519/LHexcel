VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFakeExecutionPlanView"
Option Explicit
Implements INxExecutionPlanView

Private mRenderCount As Long, mExecuteEnabled As Boolean
Private mLastFocus As String, mWasClosed As Boolean
Private mLastModel As CNxExecutionPlanViewModel

Public Property Get RenderCount() As Long: RenderCount = mRenderCount: End Property
Public Property Get ExecuteEnabled() As Boolean: ExecuteEnabled = mExecuteEnabled: End Property
Public Property Get LastFocus() As String: LastFocus = mLastFocus: End Property
Public Property Get WasClosed() As Boolean: WasClosed = mWasClosed: End Property
Public Property Get LastModel() As CNxExecutionPlanViewModel: Set LastModel = mLastModel: End Property

Private Sub INxExecutionPlanView_Render(ByVal model As CNxExecutionPlanViewModel)
    mRenderCount = mRenderCount + 1
    Set mLastModel = model
End Sub

Private Sub INxExecutionPlanView_SetExecuteEnabled(ByVal enabled As Boolean)
    mExecuteEnabled = enabled
End Sub

Private Sub INxExecutionPlanView_FocusCancel()
    mLastFocus = "cancel"
End Sub

Private Sub INxExecutionPlanView_CloseView()
    mWasClosed = True
End Sub
