VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxTemplateHeaderSizer"
Option Explicit

Private WithEvents mGrip As MSForms.Label
Private mHost As Object
Private mIndex As Long
Private mDownX As Single

Public Sub Bind(ByVal host As Object, ByVal grip As MSForms.Label, ByVal index As Long)
    Set mHost = host: Set mGrip = grip: mIndex = index
End Sub

Public Sub Unbind()
    Set mGrip = Nothing
    Set mHost = Nothing
End Sub

Private Sub mGrip_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then mDownX = X
End Sub

Private Sub mGrip_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then mHost.ResizeHeader mIndex, X - mDownX
End Sub
