Attribute VB_Name = "T_R78Files"
Option Explicit

Public Function Names() As String
    Names = "compare_values_20000|compare_formulas|compare_mixed|compare_literals|compare_errors|compare_merges|compare_formats|compare_single|compare_report|consolidate_plain|consolidate_inflated|consolidate_headers|consolidate_formula_blank|consolidate_mismatch"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim a As Workbook, b As Workbook, output As Workbook, leftRange As Range, rightRange As Range
    Dim started As Double, elapsed As Double, root As String, result As CNxResult
    Dim reasons As Variant, expected As String, i As Long, repetition As Long, handle As Integer
    Dim oldAlerts As Boolean, oldEvents As Boolean, oldStatus As Variant, p As String, q As String, outPath As String
    On Error GoTo Failed
    oldAlerts = Application.DisplayAlerts: oldEvents = Application.EnableEvents: oldStatus = Application.StatusBar
    Application.DisplayAlerts = False: Application.EnableEvents = False
    root = Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator
    Set a = Workbooks.Add(xlWBATWorksheet): Set b = Workbooks.Add(xlWBATWorksheet)
    If Left$(name, 8) = "compare_" Then
        Set leftRange = a.Worksheets(1).Range("B2").Resize(20, 20)
        Set rightRange = b.Worksheets(1).Range("B2").Resize(20, 20)
        If name = "compare_values_20000" Then
            Set leftRange = a.Worksheets(1).Range("B2").Resize(400, 50)
            Set rightRange = b.Worksheets(1).Range("B2").Resize(400, 50)
        ElseIf name = "compare_single" Then
            Set leftRange = a.Worksheets(1).Range("B2"): Set rightRange = b.Worksheets(1).Range("B2")
        End If
        leftRange.Value2 = 1: rightRange.Value2 = 1
        rightRange.Cells(1).Value2 = 2: expected = "값"
        Select Case name
            Case "compare_formulas"
                leftRange.Formula = "=1+2": rightRange.Formula = "=1+2"
                rightRange.Cells(1).Formula = "=2+1": expected = "수식"
            Case "compare_mixed"
                leftRange.Cells(1).Formula = "=1+2": rightRange.Cells(1).Value2 = 3
                leftRange.Cells(2).NumberFormat = "@": rightRange.Cells(2).NumberFormat = "@"
                leftRange.Cells(2).Value2 = "=1+2": rightRange.Cells(2).Value2 = "=1+2"
                expected = "수식/상수"
            Case "compare_literals"
                leftRange.NumberFormat = "@": rightRange.NumberFormat = "@"
                leftRange.Cells(1).Value2 = "=1+2": rightRange.Cells(1).Value2 = "=2+1"
            Case "compare_errors"
                leftRange.Cells(1).Value2 = CVErr(xlErrNA): rightRange.Cells(1).Value2 = CVErr(xlErrDiv0)
            Case "compare_merges"
                leftRange.Range("A1:B1").ClearContents: rightRange.Range("A1:B1").ClearContents
                leftRange.Range("A1:B1").Merge: expected = vbNullString
            Case "compare_formats"
                rightRange.Cells(1).Value2 = 1: rightRange.Cells(1).Font.Bold = True: expected = "서식"
        End Select
        a.Saved = True: b.Saved = True
        For repetition = 1 To 3
            started = Timer
            reasons = NxR78Reasons(leftRange, rightRange, name = "compare_formats")
            elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400
            If reasons(0) <> expected Then Err.Raise 5, , "Wrong first reason: " & reasons(0)
            For i = 1 To UBound(reasons)
                If name = "compare_merges" And i = 1 Then
                    If reasons(i) <> vbNullString Then Err.Raise 5, , "Blank merge follower should be excluded"
                ElseIf reasons(i) <> "" Then
                    Err.Raise 5, , "Unexpected difference"
                End If
            Next i
            If Not a.Saved Or Not b.Saved Then Err.Raise 5, , "Compare changed source"
            Metric name, elapsed
        Next repetition
        If name = "compare_report" Then
            a.Worksheets(1).Name = "비교'시트": b.Worksheets(1).Name = "비교'시트"
            p = root & name & "-a.xlsx": q = root & name & "-b.xlsx"
            a.SaveAs p, xlOpenXMLWorkbook: b.SaveAs q, xlOpenXMLWorkbook
            a.Close False: b.Close False: Set a = Nothing: Set b = Nothing
            Set result = NxWorkbookCompareReport(p, q, vbNullString, False)
            If result.Outcome <> NxSuccess Then Err.Raise 5, , "Report failed"
            NxWorkbookCompareFinishReport True
            Set output = ActiveWorkbook
            If Len(output.Path) <> 0 Or output.Saved Then Err.Raise 5, , "Report should be unsaved"
            If output.Worksheets("비교요약").Range("B4").Value2 <> 1 Then Err.Raise 5, , "Report difference count"
            If output.Worksheets("셀차이").Range("D2:E2").Hyperlinks.Count <> 2 Then Err.Raise 5, , "Snapshot links missing"
            If output.Worksheets("셀차이").Range("B2").Value2 <> "B2" Then Err.Raise 5, , "Wrong report address"
        End If
    Else
        a.Worksheets(1).Range("A1:AX1800").Value2 = 1
        b.Worksheets(1).Range("A1:AX1800").Value2 = 2
        If name = "consolidate_headers" Or name = "consolidate_mismatch" Then
            a.Worksheets(1).Rows(1).ClearContents: b.Worksheets(1).Rows(1).ClearContents
            a.Worksheets(1).Range("A1:AX1").Value2 = "header": b.Worksheets(1).Range("A1:AX1").Value2 = "header"
            If name = "consolidate_mismatch" Then b.Worksheets(1).Range("B1").Value2 = "other"
        End If
        If name = "consolidate_formula_blank" Then
            a.Worksheets(1).Range("AX1801").Formula = "=" & Chr$(34) & Chr$(34)
            b.Worksheets(1).Range("AX1801").Formula = "=" & Chr$(34) & Chr$(34)
        End If
        If name = "consolidate_inflated" Then
            a.Worksheets(1).Range("A1:AX8000").Interior.Color = vbYellow
            b.Worksheets(1).Range("A1:AX8000").Interior.Color = vbYellow
        End If
        p = root & name & "-a.xlsx": q = root & name & "-b.xlsx": outPath = root & name & "-result.xlsx"
        a.SaveAs p, xlOpenXMLWorkbook: b.SaveAs q, xlOpenXMLWorkbook
        a.Close False: b.Close False: Set a = Nothing: Set b = Nothing
        started = Timer
        Set result = NxRunConsolidation(Array(p, q), outPath, False, False, "ROWS", name = "consolidate_headers" Or name = "consolidate_mismatch")
        elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400
        Metric name, elapsed
        If name = "consolidate_mismatch" Then
            If result.Outcome = NxSuccess Or Len(Dir$(outPath)) > 0 Then Err.Raise 5, , "Mismatched headers published"
            RunCase = "PASS|" & name
            GoTo Clean
        End If
        If result.Outcome <> NxSuccess Then Err.Raise 5, , "Consolidation: " & result.Recovery
        Set output = Workbooks.Open(outPath)
        i = 3600: If name = "consolidate_headers" Then i = 3599
        If name = "consolidate_formula_blank" Then i = 3602
        If output.Worksheets(1).Cells.Find(What:="*", After:=output.Worksheets(1).Cells(1, 1), LookIn:=xlFormulas, _
            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False).Row <> _
            i - IIf(name = "consolidate_formula_blank", 1, 0) Then Err.Raise 5, , "Wrong data row count"
        If name = "consolidate_formula_blank" Then
            If output.Worksheets(1).Cells(1802, 50).Value2 <> 2 Then Err.Raise 5, , "Blank formula row dropped"
        Else
            If output.Worksheets(1).Cells(i, 50).Value2 <> 2 Then Err.Raise 5, , "Last value missing"
        End If
    End If
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & CStr(Err.Number) & "|" & Err.Description
Clean:
    On Error Resume Next
    If Not a Is Nothing Then a.Close False
    If Not b Is Nothing Then b.Close False
    If Not output Is Nothing Then output.Close False
    Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents: Application.StatusBar = oldStatus
End Function

Private Sub Metric(ByVal name As String, ByVal elapsed As Double)
    Dim handle As Integer
    handle = FreeFile
    Open Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator & "files-timing.tsv" For Append As #handle
    Print #handle, name & vbTab & Format$(elapsed, "0.000")
    Close #handle
End Sub
