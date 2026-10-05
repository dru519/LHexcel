Attribute VB_Name = "T_R69RunnerBenchmark"
Option Explicit
Private mEvidence As String
Public Sub Configure(ByVal folder As String)
    mEvidence = folder
End Sub
Public Function Names() As String
    Names = "paired_200|paired_1000"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, result As CNxResult
    Dim cells As Long, iteration As Long, phase As Long, legacy As Boolean
    Dim started As Double, elapsed As Double, handle As Integer, detail As String
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
            ' Alternate AB/BA to reduce order and warm-up bias.
            legacy = ((iteration + phase) Mod 2 = 1)
            Application.Run "'" & ThisWorkbook.Name & "'!NxSetLegacyBenchmark", legacy
            target.Select
            started = Timer
            Set result = NxDrawRun(NX_FEATURE_DRAW_BUSINESS_TABLE, target)
            elapsed = Timer - started
            If elapsed < 0 Then elapsed = elapsed + 86400#
            If result.Outcome <> NxSuccess Or Not result.SourceChanged Then Err.Raise 5, , "Table result changed"
            If target.Font.Size <> 9 Or target.Cells(2, 2).Formula <> "=1+2" Then Err.Raise 5, , "Table content changed"
            handle = FreeFile
            Open mEvidence & Application.PathSeparator & "paired.tsv" For Append As #handle
            Print #handle, cells & vbTab & iteration & vbTab & IIf(legacy, "legacy", "optimized") & vbTab & Format$(elapsed * 1000#, "0.000")
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
    Application.Run "'" & ThisWorkbook.Name & "'!NxSetLegacyBenchmark", False
    If Not book Is Nothing Then book.Close False
End Function
