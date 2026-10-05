Attribute VB_Name = "NxSheetRebuilder"
Option Explicit

Public Sub NxRebuildVisibleSheet(ByVal source As Worksheet, ByVal target As Worksheet, ByVal targetName As String)
    Dim sourceRange As Range, targetRange As Range, detail As String
    On Error GoTo Failed
    If source Is Nothing Or target Is Nothing Then NxRaiseContractError "Sheet rebuild requires source and target"
    If source.Visible <> xlSheetVisible Then Exit Sub
    Set sourceRange = source.UsedRange
    If sourceRange Is Nothing Then Exit Sub
    Set targetRange = target.Range("A1").Resize(sourceRange.Rows.Count, sourceRange.Columns.Count)
    targetRange.Value = sourceRange.Value
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormats
    targetRange.FormatConditions.Delete
    Application.CutCopyMode = False
    CopyDimensions source, target, sourceRange
    target.Name = targetName
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    On Error GoTo 0
    NxRaiseContractError "Sheet rebuild failed: " & detail
End Sub

Private Sub CopyDimensions(ByVal source As Worksheet, ByVal target As Worksheet, ByVal area As Range)
    Dim rowIndex As Long, columnIndex As Long
    For rowIndex = 1 To area.Rows.Count
        target.Rows(rowIndex).RowHeight = source.Rows(area.Row + rowIndex - 1).RowHeight
    Next rowIndex
    For columnIndex = 1 To area.Columns.Count
        target.Columns(columnIndex).ColumnWidth = source.Columns(area.Column + columnIndex - 1).ColumnWidth
    Next columnIndex
End Sub
