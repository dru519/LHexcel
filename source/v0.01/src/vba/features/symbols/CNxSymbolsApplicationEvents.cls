VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxSymbolsApplicationEvents"
Option Explicit

Public WithEvents App As Excel.Application
Private mHost As FNxSymbols

Public Sub Bind(ByVal excelApplication As Excel.Application, ByVal host As FNxSymbols)
    If excelApplication Is Nothing Or host Is Nothing Then NxRaiseContractError "기호표 선택 감시 연결이 잘못되었습니다."
    Set App = excelApplication
    Set mHost = host
End Sub

Public Sub Unbind()
    Set mHost = Nothing
    Set App = Nothing
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not mHost Is Nothing Then mHost.RefreshTarget
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    If Not mHost Is Nothing Then mHost.RefreshTarget
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    If Not mHost Is Nothing Then mHost.RefreshTarget
End Sub
