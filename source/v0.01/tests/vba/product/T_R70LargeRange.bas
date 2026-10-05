Attribute VB_Name = "T_R70LargeRange"
Option Explicit
Private mEvidence As String
Public Sub Configure(ByVal folder As String)
    mEvidence = folder
End Sub
Public Function Names() As String
    Names = "table_90000|copy_90000"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, result As CNxResult, context As CNxExecutionContext
    Dim started As Double, detail As String, path As String, other As Workbook
    Dim oldEvents As Boolean, oldAlerts As Boolean
    On Error GoTo Failed
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False: Application.DisplayAlerts = False
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1").Resize(1800, 50)
    target.Value2 = "benchmark"
    target.Font.Name = "맑은 고딕": target.Font.Size = 9
    target.Cells(2, 2).Formula = "=1+2"
    target.Cells(1800, 50).Value2 = "last"
    target.Columns.ColumnWidth = 12
    target.Select
    Metric name, "start", Timer
    started = Timer
    Set context = NxContextFactory.CaptureCurrent()
    Metric name, "context", started
    Set context = Nothing
    started = Timer
    If Left$(name, 5) = "table" Then
        Set result = NxDrawRun(NX_FEATURE_DRAW_BUSINESS_TABLE, target)
    Else
        path = mEvidence & Application.PathSeparator & name & ".xlsx"
        Set result = NxFileRun(NX_FEATURE_FILE_RANGE_COPY_SAVE, path, selectedTarget:=target)
    End If
    Metric name, "full", started
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
    If target.Cells(2, 2).Formula <> "=1+2" Or target.Cells(1800, 50).Value2 <> "last" Then Err.Raise 5, , "source content changed"
    If Left$(name, 4) = "copy" Then
        Set other = Workbooks.Open(path, 0, True)
        If other.Worksheets(1).Range("AX1800").Value2 <> "last" Then Err.Raise 5, , "last output lost"
        other.Close False: Set other = Nothing
    Else
        If target.Font.Size <> 9 Then Err.Raise 5, , "font mismatch"
    End If
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    If Not other Is Nothing Then other.Close False
    If Not book Is Nothing Then book.Close False
    Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
End Function
Private Sub Metric(ByVal name As String, ByVal stage As String, ByVal started As Double)
    Dim elapsed As Double, handle As Integer
    elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400#
    handle = FreeFile
    Open mEvidence & Application.PathSeparator & "large.tsv" For Append As #handle
    Print #handle, name & vbTab & stage & vbTab & Format$(elapsed * 1000#, "0.000")
    Close #handle
End Sub

