VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CNxR66CompareProgress"
Option Explicit
Public NativeReport As Boolean, VbaReport As Boolean, Triggered As Boolean
Public CancelEnabled As Boolean
Public Sub UpdateCompareProgress(ByVal stage As String, ByVal done As Long, ByVal total As Long)
    If stage = "보고서 작성" Then NativeReport = True
    If stage = "셀 비교" Then VbaReport = True
    If CancelEnabled And stage = "보고서 작성" And done >= 256 Then
        Triggered = True
        NxWorkbookCompareCancel
    End If
End Sub
