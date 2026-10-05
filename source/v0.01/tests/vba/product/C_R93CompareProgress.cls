VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "C_R93CompareProgress"
Option Explicit
Public StopStage As String
Public Fired As Boolean
Public Sub UpdateCompareProgress(ByVal stage As String, ByVal done As Long, ByVal total As Long)
    If stage = StopStage Then
        Fired = True
        NxWorkbookCompareCancel
    End If
End Sub

Public Sub UpdateTemplateProgress(ByVal stage As String, ByVal done As Long, ByVal total As Long)
    If stage = StopStage Then
        Fired = True
        NxTemplateRequestCancel
    End If
End Sub
