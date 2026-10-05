VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFormulaNoteBackup"
Option Explicit

Private mWorkbook As Workbook

Public Sub Capture(ByVal source As Range, ByVal slot As Long)
    If source.Comment Is Nothing Then NxRaiseContractError "보관할 셀 메모가 없습니다."
    If mWorkbook Is Nothing Then Set mWorkbook = Workbooks.Add(xlWBATWorksheet)
    source.Copy
    mWorkbook.Worksheets(1).Cells(slot, 1).PasteSpecial Paste:=xlPasteComments
    Application.CutCopyMode = False
End Sub

Public Sub Hide()
    If mWorkbook Is Nothing Then Exit Sub
    mWorkbook.Saved = True
    mWorkbook.IsAddin = True
End Sub

Public Sub Restore(ByVal target As Range, ByVal slot As Long)
    If mWorkbook Is Nothing Then NxRaiseContractError "보관한 셀 메모를 찾지 못했습니다."
    If Not target.Comment Is Nothing Then target.Comment.Delete
    mWorkbook.Worksheets(1).Cells(slot, 1).Copy
    target.PasteSpecial Paste:=xlPasteComments
    Application.CutCopyMode = False
End Sub

Public Sub Dispose()
    Dim priorEvents As Boolean, failure As Long, detail As String
    If mWorkbook Is Nothing Then Exit Sub
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    mWorkbook.Close SaveChanges:=False
    Set mWorkbook = Nothing
    Application.EnableEvents = priorEvents
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description
    Application.EnableEvents = priorEvents
    Err.Raise failure, "CNxFormulaNoteBackup.Dispose", detail
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Dispose
End Sub
