VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFormatSnapshot"
Option Explicit

Private Const NX_FORMAT_SNAPSHOT_MAX_CELLS As Long = 100000

Private mTarget As Range
Private mCellStates As Collection
Private mColumnWidths As Object

Public Sub Capture(ByVal target As Range, Optional ByVal captureCellStates As Boolean = True)
    Dim cell As Range, state As CNxRoleStyleCellState
    Dim failureNumber As Long, failureDescription As String
    If target Is Nothing Then NxRaiseContractError "Format snapshot requires a target"
    If target.Areas.Count <> 1 Then NxRaiseContractError "Format snapshot requires a contiguous target"
    If CDbl(target.CountLarge) > CDbl(NX_FORMAT_SNAPSHOT_MAX_CELLS) Then NxRaiseContractError "Format snapshot exceeds the Draw owner cell cap"
    On Error GoTo Failed
    Set mTarget = target
    If captureCellStates Then
        Set mCellStates = New Collection
        For Each cell In target.Cells
            Set state = New CNxRoleStyleCellState
            state.Capture cell
            mCellStates.Add state
        Next cell
    End If
    Set mColumnWidths = CreateObject("Scripting.Dictionary")
    CaptureColumnWidths
    Exit Sub
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description
    Err.Raise failureNumber, "CNxFormatSnapshot.Capture", failureDescription
End Sub

Public Sub Restore()
    Dim index As Long, state As CNxRoleStyleCellState
    Dim failureNumber As Long, failureDescription As String
    If mTarget Is Nothing Then Exit Sub
    On Error GoTo Failed
    If Not mCellStates Is Nothing Then
        For index = mCellStates.Count To 1 Step -1
            Set state = mCellStates(index)
            state.Restore
        Next index
    End If
    RestoreColumnWidths
    Exit Sub
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description
    Err.Raise failureNumber, "CNxFormatSnapshot.Restore", failureDescription
End Sub

Private Sub CaptureColumnWidths()
    Dim column As Range
    For Each column In mTarget.Columns
        mColumnWidths.Add CStr(column.Column), Array(column.ColumnWidth, column.EntireColumn.Hidden)
    Next column
End Sub

Private Sub RestoreColumnWidths()
    Dim key As Variant, values As Variant
    If mColumnWidths Is Nothing Then Exit Sub
    For Each key In mColumnWidths.Keys
        values = mColumnWidths(key)
        With mTarget.Worksheet.Columns(CLng(key))
            .ColumnWidth = values(0)
            .Hidden = CBool(values(1))
        End With
    Next key
End Sub
