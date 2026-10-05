VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxRangeFormatBackup"
Option Explicit

Private mWorkbook As Workbook
Private mSource As Range
Private mBackup As Range

Public Sub Capture(ByVal target As Range)
    Dim owner As Window, priorEvents As Boolean, priorAlerts As Boolean
    Dim failure As Long, detail As String
    If target Is Nothing Then NxRaiseContractError "Format backup requires a target"
    If target.Areas.Count <> 1 Then NxRaiseContractError "Format backup requires one range"
    If Not mWorkbook Is Nothing Then NxRaiseContractError "Format backup was already captured"
    Set owner = Application.ActiveWindow
    priorEvents = Application.EnableEvents
    priorAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Set mSource = target
    ' Excel copies its own style and theme records with the worksheet.
    ' No temporary file or user worksheet is created in the source workbook.
    target.Worksheet.Copy
    Set mWorkbook = Application.ActiveWorkbook
    Set mBackup = mWorkbook.Worksheets(1).Range(target.Address)
    If Not owner Is Nothing Then owner.Activate
    mWorkbook.Windows(1).Visible = False
    Application.EnableEvents = priorEvents
    Application.DisplayAlerts = priorAlerts
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description
    On Error Resume Next
    Dispose
    If Not owner Is Nothing Then owner.Activate
    Application.EnableEvents = priorEvents
    Application.DisplayAlerts = priorAlerts
    Err.Raise failure, "CNxRangeFormatBackup.Capture", detail
End Sub

Public Sub Restore()
    If mWorkbook Is Nothing Or mSource Is Nothing Then NxRaiseContractError "Format backup is unavailable"
    mBackup.Copy
    mSource.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
End Sub

Public Property Get SourceRange() As Range
    Set SourceRange = mSource
End Property

Public Property Get BackupRange() As Range
    Set BackupRange = mBackup
End Property

Public Sub Dispose()
    Dim priorEvents As Boolean, failure As Long, detail As String
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set mBackup = Nothing
    Set mSource = Nothing
    If Not mWorkbook Is Nothing Then mWorkbook.Close SaveChanges:=False
    Set mWorkbook = Nothing
    Application.EnableEvents = priorEvents
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description
    Application.EnableEvents = priorEvents
    Err.Raise failure, "CNxRangeFormatBackup.Dispose", detail
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Dispose
End Sub
