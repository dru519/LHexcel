VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxSymbolButtonHandler"
Option Explicit

Public WithEvents Button As MSForms.CommandButton
Private mHost As FNxSymbols
Private mIndex As Long

Public Sub Bind(ByVal host As FNxSymbols, ByVal buttonControl As MSForms.CommandButton, ByVal oneBasedIndex As Long)
    If host Is Nothing Or buttonControl Is Nothing Or oneBasedIndex < 1 Or oneBasedIndex > 50 Then NxRaiseContractError "기호 버튼 연결이 잘못되었습니다."
    Set mHost = host
    Set Button = buttonControl
    mIndex = oneBasedIndex
End Sub

Private Sub Button_Click()
    mHost.HandleGlyphClick mIndex
End Sub

Private Sub Button_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    mHost.HandleGlyphActivate mIndex
End Sub

Private Sub Button_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    mHost.HandleGlyphKey mIndex, KeyCode, Shift
End Sub
