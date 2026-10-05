VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxR59CancelSave"
Option Explicit
Private WithEvents mApp As Application
Private mBook As Workbook

Public Sub Start(ByVal book As Workbook)
    Set mBook = book
    Set mApp = Application
End Sub

Public Sub StopListening()
    Set mApp = Nothing
    Set mBook = Nothing
End Sub

Private Sub mApp_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Wb Is mBook Then
        Wb.Worksheets(1).Range("B3").Value2 = "event edit preserved"
        Cancel = True
    End If
End Sub
