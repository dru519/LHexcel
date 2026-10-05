VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxCalculatorButtonHandler"
Option Explicit

Public WithEvents Button As MSForms.CommandButton
Private mHost As Object

Public Sub Bind(ByVal control As MSForms.CommandButton, ByVal host As Object)
    If control Is Nothing Or host Is Nothing Then NxRaiseContractError "계산기 버튼 연결 대상을 확인할 수 없습니다."
    Set Button = control
    Set mHost = host
End Sub

Private Sub Button_Click()
    mHost.HandleCommand CStr(Button.Tag)
End Sub

Private Sub Button_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    mHost.HandleKey keyCode, shift
End Sub
