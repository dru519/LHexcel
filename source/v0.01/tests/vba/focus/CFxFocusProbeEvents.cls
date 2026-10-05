VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFxFocusProbeEvents"
Option Explicit

Public WithEvents ExcelApp As Application

Private mEditPending As Boolean

Public Sub Start(ByVal app As Application)
    If app Is Nothing Then Err.Raise vbObjectError + 2800, "CFxFocusProbeEvents", "Excel application is required"
    Set ExcelApp = app
    mEditPending = False
End Sub

Public Sub StopListening()
    mEditPending = False
    Set ExcelApp = Nothing
End Sub

Public Sub ClearEditPending()
    mEditPending = False
End Sub

Private Sub ExcelApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    mEditPending = True
End Sub

Private Sub ExcelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo DeferMutation
    If ExcelApp Is Nothing Then Exit Sub
    If ExcelApp.CutCopyMode <> False Then Exit Sub
    If mEditPending Then
        If ExcelApp.CommandBars.GetEnabledMso("Undo") Then Exit Sub
        mEditPending = False
    End If
    FxCacheSelection Target
    Sh.Calculate
    Exit Sub
DeferMutation:
    Err.Clear
End Sub
