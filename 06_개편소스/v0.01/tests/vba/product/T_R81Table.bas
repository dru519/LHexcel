Attribute VB_Name = "T_R81Table"
Option Explicit
Public Function Names() As String
    Names = "chart_90000|picture_90000|inflated_1000000"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, ws As Worksheet, target As Range, result As CNxResult
    Dim chart As ChartObject, picture As Shape, started As Double, elapsed As Double
    Dim priorEvents As Boolean, priorAlerts As Boolean, handle As Integer, series As String
    On Error GoTo Failed
    priorEvents = Application.EnableEvents: priorAlerts = Application.DisplayAlerts
    Application.EnableEvents = False: Application.DisplayAlerts = False
    Set book = Workbooks.Add(xlWBATWorksheet): Set ws = book.Worksheets(1)
    Set target = ws.Range("A1:AX1800")
    target.Value2 = "sample": target.Font.Size = 9
    ws.Range("B2").Formula = "=1+2"
    If name = "chart_90000" Then
        ws.Range("BA1:BA3").Value2 = 7
        Set chart = ws.ChartObjects.Add(900, 100, 150, 100)
        chart.Placement = xlFreeFloating
        chart.Chart.SetSourceData ws.Range("BA1:BA3")
        series = chart.Chart.SeriesCollection(1).Formula
    ElseIf name = "picture_90000" Then
        ws.Range("A1:C3").CopyPicture Appearance:=xlScreen, Format:=xlPicture
        ws.Paste
        Set picture = ws.Shapes(ws.Shapes.Count)
        picture.Placement = xlFreeFloating
        picture.Left = 900: picture.Top = 100
        If picture.Type <> msoPicture Then Err.Raise 5, , "Picture fixture missing"
    Else
        ws.Range("A1:AX20000").Interior.Color = vbYellow
    End If
    Application.CutCopyMode = False
    target.Select
    started = Timer
    Set result = NxDrawRun(NX_FEATURE_DRAW_BUSINESS_TABLE, target)
    elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
    If ws.Range("B2").Formula <> "=1+2" Or ws.Range("AX1800").Value2 <> "sample" Then Err.Raise 5, , "Content changed"
    If name = "chart_90000" Then
        If ws.ChartObjects.Count <> 1 Or chart.Chart.SeriesCollection(1).Formula <> series Then Err.Raise 5, , "Chart changed"
        If chart.Left <> 900 Or chart.Top <> 100 Then Err.Raise 5, , "Chart moved"
    ElseIf name = "picture_90000" Then
        If ws.Shapes.Count <> 1 Or picture.Left <> 900 Or picture.Top <> 100 Then Err.Raise 5, , "Picture changed"
    Else
        If ws.Range("AX20000").Interior.Color <> vbYellow Then Err.Raise 5, , "Outside format changed"
    End If
    handle = FreeFile
    Open Environ$("LHEXCEL_PROFILE_ROOT") & "\full-table.tsv" For Append As #handle
    Print #handle, name & vbTab & Format$(elapsed, "0.000")
    Close #handle
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
Clean:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    Application.EnableEvents = priorEvents: Application.DisplayAlerts = priorAlerts
End Function
