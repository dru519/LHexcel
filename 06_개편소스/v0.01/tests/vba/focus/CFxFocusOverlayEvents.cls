VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFxFocusOverlayEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mApplication As Application
Attribute mApplication.VB_VarHelpID = -1

Public Sub StartListening(ByVal host As Object)
    Set mApplication = host
End Sub

Public Sub StopListening()
    Set mApplication = Nothing
End Sub

Private Sub mApplication_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    FxOverlayRefresh
End Sub

Private Sub mApplication_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
    FxOverlayRefresh
End Sub

Private Sub mApplication_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    FxOverlayRefresh
End Sub

Private Sub mApplication_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
    FxOverlayHideAll
End Sub

Private Sub mApplication_SheetActivate(ByVal Sh As Object)
    FxOverlayRefresh
End Sub
