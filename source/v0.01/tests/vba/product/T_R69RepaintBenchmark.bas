Attribute VB_Name = "T_R69RepaintBenchmark"
Option Explicit
Private mEvidence As String
Public Sub Configure(ByVal folder As String)
    mEvidence = folder
End Sub
Public Function Names() As String
    Names = "repaint_200|repaint_1000"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, result As CNxResult
    Dim cells As Long, iteration As Long, phase As Long, updating As Boolean
    Dim started As Double, elapsed As Double, handle As Integer, detail As String
    Dim oldUpdating As Boolean
    oldUpdating = Application.ScreenUpdating
    On Error GoTo Failed
    cells = CLng(Split(name, "_")(1))
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1").Resize(cells / 5, 5)
    target.Value2 = "sample"
    target.Font.Name = "맑은 고딕"
    target.Font.Size = 9
    target.Columns.ColumnWidth = 13
    target.RowHeight = 15
    target.Cells(2, 2).Formula = "=1+2"
    target.Select
    book.SaveAs mEvidence & Application.PathSeparator & name & ".xlsx", xlOpenXMLWorkbook
    NxDrawApplyBusinessTable target
    For iteration = 1 To 4
        For phase = 0 To 1
            updating = ((iteration + phase) Mod 2 = 1)
            Application.ScreenUpdating = True
            target.Select
            started = Timer
            Application.ScreenUpdating = updating
            Set result = NxDrawRun(NX_FEATURE_DRAW_BUSINESS_TABLE, target)
            If Application.ScreenUpdating <> updating Then Err.Raise 5, , "ScreenUpdating restoration changed"
            Application.ScreenUpdating = True
            elapsed = Timer - started
            If elapsed < 0 Then elapsed = elapsed + 86400#
            If result.Outcome <> NxSuccess Or Not result.SourceChanged Then Err.Raise 5, , "Table result changed"
            If target.Font.Size <> 9 Or target.Cells(2, 2).Formula <> "=1+2" Then Err.Raise 5, , "Table content changed"
            handle = FreeFile
            Open mEvidence & Application.PathSeparator & "repaint.tsv" For Append As #handle
            Print #handle, cells & vbTab & iteration & vbTab & IIf(updating, "updating_on", "updating_off") & vbTab & Format$(elapsed * 1000#, "0.000")
            Close #handle
        Next phase
    Next iteration
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    Application.ScreenUpdating = oldUpdating
End Function
