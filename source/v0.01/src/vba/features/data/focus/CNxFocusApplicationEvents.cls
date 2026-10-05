VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFocusApplicationEvents"
Option Explicit

Public WithEvents ExcelApp As Application

Private mListening As Boolean
Private mEditPending As Boolean

Public Sub Start(ByVal app As Application)
    If app Is Nothing Then NxRaiseContractError "Excel 실행 상태를 확인하세요."
    Set ExcelApp = app
    mEditPending = False
    mListening = True
End Sub

Public Sub StopListening()
    mListening = False
    mEditPending = False
    Set ExcelApp = Nothing
End Sub

Private Sub ExcelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not mListening Or Target Is Nothing Then Exit Sub
    If ExcelApp.CutCopyMode <> False Then Exit Sub
    If mEditPending Then
        ' Updating add-in cache names is itself a VBA mutation and clears the
        ' native Excel Undo stack. Keep the previous focus position until the
        ' user consumes or dismisses Undo, then resume on the next selection.
        If UndoIsAvailable() Then Exit Sub
        mEditPending = False
    End If
    NxFocusControllerSelectionChanged Target, True
End Sub

Private Sub ExcelApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If mListening Then mEditPending = True
End Sub

Private Sub ExcelApp_SheetActivate(ByVal Sh As Object)
    If mListening Then NxFocusControllerRefreshActiveSelection False
End Sub

Private Sub ExcelApp_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    If mListening Then NxFocusControllerRefreshActiveSelection False
End Sub

Private Sub ExcelApp_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
    If mListening Then NxFocusControllerRefreshActiveSelection False
End Sub

Private Sub ExcelApp_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
    If mListening Then NxFocusControllerWindowDeactivated Wb
End Sub

Private Sub ExcelApp_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If mListening Then
        If Not NxFocusControllerBeforeSave(Wb) Then Cancel = True
    End If
End Sub

Private Sub ExcelApp_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    If mListening Then NxFocusControllerAfterSave Wb, Success
End Sub

Private Sub ExcelApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    If mListening Then
        If Not NxFocusControllerBeforeClose(Wb) Then Cancel = True
    End If
End Sub

Private Function UndoIsAvailable() As Boolean
    On Error Resume Next
    UndoIsAvailable = CBool(ExcelApp.CommandBars.GetEnabledMso("Undo"))
    On Error GoTo 0
End Function
