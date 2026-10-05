VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CNxR65CancelCompare"
Option Explicit
Public Triggered As Boolean
Public Sub UpdateCompareProgress(ByVal stage As String, ByVal done As Long, ByVal total As Long)
    If done >= 2048 Then
        Triggered = True
        NxWorkbookCompareCancel
    End If
End Sub
