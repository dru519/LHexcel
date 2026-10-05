Option Explicit

Private mAllowClose As Boolean

Public Sub BindProductInfo(ByVal versionText As String, ByVal validUntilText As String)
    lblVersion.Caption = versionText
    lblExpiry.Caption = "사용기한: " & validUntilText
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdClose_Click
End Sub

Private Sub cmdClose_Click()
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    cmdClose_Click
End Sub
