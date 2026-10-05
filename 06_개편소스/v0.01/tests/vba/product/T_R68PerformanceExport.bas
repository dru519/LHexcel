Attribute VB_Name = "T_R68PerformanceExport"
Option Explicit

Private mEvidence As String
Private mMode As String

Public Sub Configure(ByVal evidenceRoot As String, ByVal mode As String)
    mEvidence = evidenceRoot
    mMode = mode
End Sub

Public Function Names() As String
    Names = "snapshot_20|snapshot_200|source_drift|export_visible|export_no_screen_update|export_automatic"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, sheet As Worksheet, target As Range
    Dim context As CNxExecutionContext, result As CNxResult
    Dim started As Double, fingerprintMs As Double, captureMs As Double, verifyMs As Double
    Dim rowCount As Long, i As Long, output As String, oldUpdating As Boolean
    Dim oldEvents As Boolean, oldAlerts As Boolean, detail As String
    Dim fileSystem As Object, file As Object, imageCount As Long
    On Error GoTo Failed
    oldUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = True
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1)
    If name = "snapshot_200" Then rowCount = 40 Else rowCount = 4
    Set target = sheet.Range("B2").Resize(rowCount, 5)
    target.Value2 = "r68 sample"
    target.Font.Name = "맑은 고딕"
    target.Font.Size = 9
    target.Interior.Color = RGB(255, 230, 128)
    target.Borders.LineStyle = xlContinuous
    target.Cells(2, 2).Formula = "=1+2"
    target.Cells(3, 3).Value2 = "한글 / export"
    target.Cells(4, 4).Value2 = 987.25
    sheet.Columns("B:F").ColumnWidth = 13
    target.Select
    book.Saved = True
    If Left$(name, 9) = "snapshot_" Then
        started = Timer
        output = NxContextFactory.CaptureCurrentFingerprint()
        fingerprintMs = ElapsedMs(started)
        started = Timer
        Set context = NxContextFactory.CaptureCurrent()
        captureMs = ElapsedMs(started)
        started = Timer
        If Not context.SourceStateMatchesCurrent() Then Err.Raise 5, , "Unchanged snapshot mismatch"
        verifyMs = ElapsedMs(started)
        If output <> context.Fingerprint Then Err.Raise 5, , "Identity fingerprint mismatch"
        target.Cells(1, 1).Font.Color = RGB(255, 0, 0)
        If context.SourceStateMatchesCurrent() Then Err.Raise 5, , "Font drift was not detected"
        WriteMetric name, CLng(target.CountLarge), fingerprintMs, captureMs, verifyMs, Len(context.SourceStateDigest)
    ElseIf name = "source_drift" Then
        CheckSourceDrift target
    ElseIf name = "export_automatic" Then
        book.SaveAs mEvidence & Application.PathSeparator & "export_automatic.xlsx", xlOpenXMLWorkbook
        target.Select
        Set result = NxFileSaveRangeAutomatic(target)
        If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
        Set result = NxFileSaveRangeAutomatic(target)
        If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        For Each file In fileSystem.GetFolder(mEvidence).Files
            If file.Name Like "export_automatic_*.png" Then imageCount = imageCount + 1
        Next file
        If imageCount <> 2 Then Err.Raise 5, , "Automatic saves must create two distinct images"
        If Not book.Saved Then Err.Raise 5, , "Automatic export changed source Saved state"
        If sheet.Shapes.Count <> 0 Then Err.Raise 5, , "Automatic export left source shapes"
    Else
        output = mEvidence & Application.PathSeparator & name & ".png"
        If name = "export_no_screen_update" Then Application.ScreenUpdating = False
        started = Timer
        Set result = NxExportRangePng(target, output)
        captureMs = ElapsedMs(started)
        If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
        If sheet.ChartObjects.Count <> 0 Or sheet.Shapes.Count <> 0 Then Err.Raise 5, , "Source shapes changed"
        If target.Cells(3, 3).Value2 <> "한글 / export" Then Err.Raise 5, , "Source value changed"
        If target.Cells(2, 2).Formula <> "=1+2" Then Err.Raise 5, , "Source formula changed"
        If Not book.Saved Then Err.Raise 5, , "Source Saved state changed"
        WriteMetric name, CLng(target.CountLarge), 0, captureMs, 0, 0
    End If
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = CStr(Err.Number) & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    Application.ScreenUpdating = oldUpdating
    Application.EnableEvents = oldEvents
    Application.DisplayAlerts = oldAlerts
End Function

Private Sub CheckSourceDrift(ByVal target As Range)
    Dim context As CNxExecutionContext, field As Long, cell As Range
    Set cell = target.Cells(1, 1)
    For field = 1 To 12
        target.Select
        Set context = NxContextFactory.CaptureCurrent()
        If Not context.SourceStateMatchesCurrent Then Err.Raise 5, , "Unchanged source failed before drift " & CStr(field)
        Select Case field
            Case 1: cell.Value2 = "FV5:ab:cd" & vbCrLf & "ST3:한글"
            Case 2: cell.Formula = "=12+34"
            Case 3: cell.NumberFormat = "0.000"
            Case 4: cell.Font.Size = 15
            Case 5: cell.Interior.Color = RGB(128, 255, 128)
            Case 6: cell.HorizontalAlignment = xlRight
            Case 7: cell.Locked = False
            Case 8: cell.Borders(xlEdgeBottom).Weight = xlThick
            Case 9: target.Rows(1).Merge
            Case 10: target.Rows(1).UnMerge
            Case 11: cell.Value2 = CVErr(xlErrNA)
            Case 12: cell.Value2 = True
        End Select
        If context.SourceStateMatchesCurrent Then Err.Raise 5, , "Source drift not detected " & CStr(field)
    Next field
End Sub

Private Function ElapsedMs(ByVal started As Double) As Double
    Dim elapsed As Double
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400#
    ElapsedMs = elapsed * 1000#
End Function

Private Sub WriteMetric(ByVal name As String, ByVal cells As Long, ByVal fingerprint As Double, ByVal capture As Double, ByVal verify As Double, ByVal length As Long)
    Dim handle As Integer
    handle = FreeFile
    Open mEvidence & Application.PathSeparator & "metrics.tsv" For Append As #handle
    Print #handle, name & vbTab & CStr(cells) & vbTab & CStr(fingerprint) & vbTab & CStr(capture) & vbTab & CStr(verify) & vbTab & CStr(length)
    Close #handle
End Sub
