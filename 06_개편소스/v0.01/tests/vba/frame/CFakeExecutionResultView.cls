VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFakeExecutionResultView"
Option Explicit
Implements INxExecutionResultView

Private mRenderCount As Long, mWasShown As Boolean, mWasClosed As Boolean
Private mLastModel As CNxResultCardViewModel

Public Property Get RenderCount() As Long: RenderCount = mRenderCount: End Property
Public Property Get WasShown() As Boolean: WasShown = mWasShown: End Property
Public Property Get WasClosed() As Boolean: WasClosed = mWasClosed: End Property
Public Property Get LastModel() As CNxResultCardViewModel: Set LastModel = mLastModel: End Property

Private Sub INxExecutionResultView_Render(ByVal model As CNxResultCardViewModel)
    mRenderCount = mRenderCount + 1
    Set mLastModel = model
End Sub

Private Sub INxExecutionResultView_ShowView()
    mWasShown = True
End Sub

Private Sub INxExecutionResultView_CloseView()
    mWasClosed = True
End Sub
