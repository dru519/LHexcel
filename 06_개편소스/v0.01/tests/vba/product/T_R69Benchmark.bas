Attribute VB_Name = "T_R69Benchmark"
Option Explicit
Private mEvidence As String
Public Sub Configure(ByVal folder As String)
    mEvidence = folder
End Sub
Public Function Names() As String
    Names = "table_20|table_200|table_1000|png_20|png_200|png_1000"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, result As CNxResult, context As CNxExecutionContext
    Dim snapshot As CNxFormatSnapshot, started As Double, count As Long, iteration As Long
    Dim priorEvents As Boolean, priorUpdating As Boolean, priorAlerts As Boolean, detail As String
    On Error GoTo Failed
    priorEvents = Application.EnableEvents
    priorUpdating = Application.ScreenUpdating
    priorAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    count = CLng(Split(name, "_")(1))
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1").Resize(count / 5, 5)
    target.Value2 = "표본 abc"
    target.Font.Name = "맑은 고딕"
    target.Font.Size = 9
    target.Interior.Color = RGB(255, 230, 128)
    target.Borders.LineStyle = xlContinuous
    target.Columns.ColumnWidth = 13
    target.RowHeight = 15
    target.Cells(2, 2).Formula = "=1+2"
    target.Select
    book.SaveAs mEvidence & Application.PathSeparator & name & ".xlsx", xlOpenXMLWorkbook
    For iteration = 1 To 3
        target.Select
        If Left$(name, 5) = "table" Then
            started = Timer
            Set context = NxContextFactory.CaptureCurrent()
            Metric "context", count, iteration, started
            started = Timer
            Set snapshot = New CNxFormatSnapshot
            snapshot.Capture target
            Metric "format_backup", count, iteration, started
            started = Timer
            NxDrawApplyBusinessTable target
            Metric "table_engine", count, iteration, started
            started = Timer
            Set result = NxDrawRun(NX_FEATURE_DRAW_BUSINESS_TABLE, target)
            Metric "table_full", count, iteration, started
            If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
            If target.Font.Size <> 9 Then Err.Raise 5, , "Table font"
        Else
            Application.ScreenUpdating = True
            started = Timer
            Set result = NxExportRangePng(target, mEvidence & Application.PathSeparator & "export_bench_" & count & "_" & iteration & ".png")
            Metric "png_engine", count, iteration, started
            If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
            started = Timer
            Set result = NxFileSaveRangeAutomatic(target)
            Metric "png_full", count, iteration, started
            If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
            If Not book.Saved Then Err.Raise 5, , "PNG changed Saved"
            If book.Worksheets(1).Shapes.Count <> 0 Then Err.Raise 5, , "PNG changed shapes"
        End If
        If target.Cells(2, 2).Formula <> "=1+2" Then Err.Raise 5, , "Formula changed"
    Next iteration
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    Application.EnableEvents = priorEvents
    Application.DisplayAlerts = priorAlerts
    Application.ScreenUpdating = priorUpdating
End Function
Private Sub Metric(ByVal stage As String, ByVal cells As Long, ByVal iteration As Long, ByVal started As Double)
    Dim elapsed As Double, handle As Integer
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400#
    handle = FreeFile
    Open mEvidence & Application.PathSeparator & "benchmark.tsv" For Append As #handle
    Print #handle, stage & vbTab & cells & vbTab & iteration & vbTab & Format$(elapsed * 1000#, "0.000")
    Close #handle
End Sub
