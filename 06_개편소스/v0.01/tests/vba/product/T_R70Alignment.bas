Attribute VB_Name = "T_R70Alignment"
Option Explicit
Public Function Names() As String
    Names = "left|center|right|mixed_snapshot|snapshot_dictionary_cap"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, context As CNxExecutionContext
    Dim commandId As String, expected As Long, picture As Object, index As Long, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("B2:D5")
    target.Value2 = "unchanged": target.Cells(2, 2).Formula = "=1+2"
    target.Select
    Select Case name
        Case "left": commandId = "AlignLeft": expected = xlLeft
        Case "center": commandId = "AlignCenter": expected = xlCenter
        Case "right": commandId = "AlignRight": expected = xlRight
        Case "mixed_snapshot", "snapshot_dictionary_cap"
            If name = "snapshot_dictionary_cap" Then
                Set target = book.Worksheets(1).Range("A1:A4100")
                target.Value2 = 123
                For index = 1 To 4100
                    target.Cells(index, 1).Interior.Color = RGB(index Mod 256, index \ 256, 100)
                Next index
            Else
                target.Cells(1, 1).Font.Color = vbRed
                target.Cells(2, 2).Interior.Color = vbYellow
                target.Cells(3, 3).NumberFormat = "0.0000"
                target.Cells(4, 1).Borders(xlDiagonalUp).LineStyle = xlContinuous
            End If
            target.Select
            Set context = NxContextFactory.CaptureCurrent()
            If Not context.SourceStateMatchesCurrent Then Err.Raise 5, , "unchanged mixed snapshot"
            target.Cells(target.Rows.Count, 1).Font.Bold = True
            If context.SourceStateMatchesCurrent Then Err.Raise 5, , "last cell drift not detected"
            GoTo Passed
    End Select
    Set picture = Application.CommandBars.GetImageMso(commandId, 16, 16)
    If picture.Width <= 0 Then Err.Raise 5, , "alignment image empty"
    Application.CommandBars.ExecuteMso commandId
    If target.HorizontalAlignment <> expected Then Err.Raise 5, , "alignment result"
    If target.Cells(2, 2).Formula <> "=1+2" Then Err.Raise 5, , "formula changed"
    ' Undo must be tested from the Excel UI after this macro returns.
    ' Application.Undo inside the executing test macro is not a user undo check.
Passed:
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    Set picture = Nothing
    If Not book Is Nothing Then book.Close False
End Function
