VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxMannerSaveEvents"
Option Explicit

Private WithEvents mApp As Application
Private mBook As Workbook
Private mSaveSucceeded As Boolean

Public Sub Start(ByVal book As Workbook)
    Set mBook = book
    Set mApp = Application
    mSaveSucceeded = False
End Sub

Public Property Get SaveSucceeded() As Boolean
    SaveSucceeded = mSaveSucceeded
End Property

Public Sub StopListening()
    Set mApp = Nothing
    Set mBook = Nothing
End Sub

Private Sub mApp_WorkbookAfterSave(ByVal Wb As Workbook, ByVal Success As Boolean)
    If Wb Is mBook Then mSaveSucceeded = Success
End Sub
