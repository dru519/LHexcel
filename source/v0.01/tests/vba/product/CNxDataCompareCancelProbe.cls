VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CNxDataCompareCancelProbe"
Option Explicit
Public Calls As Long
Public Sub UpdateCompareProgress(ByVal stage As String, ByVal done As Long, ByVal total As Long)
    Calls = Calls + 1
    If stage = "차이 표시" Then NxDataCompareCancel
End Sub

