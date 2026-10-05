Attribute VB_Name = "T_R80Table"
Option Explicit
Public Function Names() As String
    Names = "plain_90000|merge_wrap_90000|shape_90000|inflated_90000|autofit_90000"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, result As CNxResult, started As Double
    Dim oldAlerts As Boolean, oldEvents As Boolean, handle As Integer, elapsed As Double
    On Error GoTo Failed
    oldAlerts = Application.DisplayAlerts: oldEvents = Application.EnableEvents
    Application.DisplayAlerts = False: Application.EnableEvents = False
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1:AX1800")
    target.Value2 = "sample": target.Font.Size = 9
    target.Cells(2, 2).Formula = "=1+2": target.Cells(1800, 50).Value2 = "last"
    If name = "merge_wrap_90000" Then
        target.Range("A1:B1").Merge
        target.Cells(3, 1).Value2 = "첫 줄" & vbLf & "둘째 줄": target.Cells(3, 1).WrapText = True
    ElseIf name = "shape_90000" Then
        book.Worksheets(1).Shapes.AddShape 1, 900, 100, 30, 30
    ElseIf name = "inflated_90000" Then
        book.Worksheets(1).Range("A1:AX8000").Interior.Color = vbYellow
    End If
    target.Select
    started = Timer
    Set result = NxDrawRun(NX_FEATURE_DRAW_BUSINESS_TABLE, target, autoFitColumns:=(name = "autofit_90000"))
    elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
    If target.Cells(2, 2).Formula <> "=1+2" Or target.Cells(1800, 50).Value2 <> "last" Then Err.Raise 5, , "Content changed"
    If name = "shape_90000" Then
        If book.Worksheets(1).Shapes.Count <> 1 Then Err.Raise 5, , "Shape lost"
    End If
    If name = "inflated_90000" Then
        If book.Worksheets(1).Range("AX8000").Interior.Color <> vbYellow Then Err.Raise 5, , "Outside format changed"
    End If
    handle = FreeFile
    Open Environ$("LHEXCEL_PROFILE_ROOT") & "\full-table.tsv" For Append As #handle
    Print #handle, name & vbTab & Format$(elapsed, "0.000")
    Close #handle
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & "|" & Err.Description
Clean:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents
End Function
