VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxRoleStyleJournal"
Option Explicit

Private mStates As Collection

Public Sub Capture(ByVal target As Range)
    Dim cell As Range, state As CNxRoleStyleCellState
    Set mStates = New Collection
    For Each cell In target.Cells
        Set state = New CNxRoleStyleCellState: state.Capture cell: mStates.Add state
    Next cell
End Sub

Public Sub Restore()
    Dim i As Long, state As CNxRoleStyleCellState
    If mStates Is Nothing Then Exit Sub
    For i = mStates.Count To 1 Step -1
        Set state = mStates(i): state.Restore
    Next i
End Sub

' The cell state restores Font.Name, Font.Size, Font.Bold, Font.Color, Interior.Pattern, Interior.Color, HorizontalAlignment, VerticalAlignment, WrapText and xlEdgeLeft/xlEdgeTop/xlEdgeRight/xlEdgeBottom/xlInsideHorizontal/xlInsideVertical borders before returning.
