Attribute VB_Name = "T_R77Normalization"
Option Explicit

Public Function Names() As String
    Names = "preview_plain_20000|preview_table_20000|preview_formula_blank_error|preview_merged|preview_all_formulas|execute_plain|execute_merged"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, sheet As Worksheet, target As Range
    Dim started As Double, elapsed As Double, text As String, preset As String
    Dim oldEvents As Boolean, oldAlerts As Boolean, failure As String, handle As Integer
    Dim result As CNxResult, output As Worksheet, repetition As Long
    On Error GoTo Failed
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False: Application.DisplayAlerts = False
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1)
    preset = "기본 정리"
    Select Case name
        Case "preview_plain_20000", "preview_table_20000"
            Set target = sheet.Range("B2").Resize(400, 50)
            target.Value2 = "  sample  value "
            If name = "preview_table_20000" Then preset = "표 구조 정리"
        Case "preview_formula_blank_error"
            Set target = sheet.Range("B2:C4")
            target.Cells(1, 1).Formula = "=1+2"
            target.Cells(1, 2).Formula = "=" & Chr$(34) & Chr$(34)
            target.Cells(2, 1).Value2 = CVErr(xlErrNA)
            target.Cells(2, 2).Value2 = " unchanged "
            target.Cells(3, 1).Value2 = 7
        Case "preview_merged"
            Set target = sheet.Range("B2:D4")
            target.Value2 = "x"
            sheet.Range("B2:C2").Merge
            sheet.Range("B2").Value2 = " merged "
            preset = "표 구조 정리"
        Case "preview_all_formulas"
            Set target = sheet.Range("B2:C4")
            target.Formula = "=1+2"
        Case "execute_plain", "execute_merged"
            Set target = sheet.Range("B2:C4")
            target.NumberFormat = "@"
            target.Value2 = "  sample  value "
            target.Cells(3, 2).Value2 = "=1+2"
            preset = "표 구조 정리"
            If name = "execute_merged" Then sheet.Range("B2:C2").Merge
        Case Else
            Err.Raise 5, , "Unknown normalization case"
    End Select
    target.Select: book.Saved = True
    started = Timer
    text = NxDataNormalizationPreview(target, preset, "yyyy-mm-dd")
    elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400
    If InStr(text, "preview") > 0 Or Len(text) < 20 Then Err.Raise 5, , "Missing preview"
    If name = "preview_formula_blank_error" Then
        If InStr(text, "수식→값 1개") = 0 Then Err.Raise 5, , "Formula count: " & text
        If InStr(text, "빈 셀 제외 2개") = 0 Then Err.Raise 5, , "Blank count: " & text
        If InStr(text, "빈칸 처리 1개") = 0 Then Err.Raise 5, , "Error count: " & text
        If target.Cells(1, 1).Formula <> "=1+2" Then Err.Raise 5, , "Formula changed"
    End If
    If name = "preview_merged" Then
        If InStr(text, "B2:") = 0 Then Err.Raise 5, , "Merged sample missing"
    End If
    If name = "preview_all_formulas" Then
        If InStr(text, "수식→값 6개") = 0 Then Err.Raise 5, , "All-formula count: " & text
    End If
    If Not book.Saved Then Err.Raise 5, , "Preview dirtied source"
    If Left$(name, 8) = "execute_" Then
        Set result = NxDataRunNormalizationForTest(target, preset, "yyyy-mm-dd", "새 시트 만들기")
        If result.Outcome <> NxSuccess Then Err.Raise 5, , "Execution failed"
        If book.Worksheets.Count <> 2 Then Err.Raise 5, , "Result sheet missing"
        If book.Worksheets(1) Is sheet Then Set output = book.Worksheets(2) Else Set output = book.Worksheets(1)
        If output.Range("A1").Value2 <> "sample value" Then Err.Raise 5, , "Normalized value mismatch"
        If output.Range("B3").Value2 <> "=1+2" Or output.Range("B3").HasFormula Then Err.Raise 5, , "Literal text changed"
        If target.Cells(1, 1).Value2 <> "  sample  value " Then Err.Raise 5, , "Source changed"
        If name = "execute_merged" Then
            If output.Range("B1").Value2 <> "sample value" Then Err.Raise 5, , "Merge value not expanded"
            If Not sheet.Range("B2").MergeCells Then Err.Raise 5, , "Source merge changed"
        End If
    End If
    handle = FreeFile
    Open Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator & "normalization-timing.tsv" For Append As #handle
    Print #handle, name & vbTab & Format$(elapsed, "0.000")
    Close #handle
    If InStr(name, "20000") > 0 Then
        For repetition = 2 To 4
            started = Timer
            text = NxDataNormalizationPreview(target, preset, "yyyy-mm-dd")
            elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400
            handle = FreeFile
            Open Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator & "normalization-timing.tsv" For Append As #handle
            Print #handle, name & vbTab & Format$(elapsed, "0.000")
            Close #handle
            If Not book.Saved Then Err.Raise 5, , "Repeated preview dirtied source"
        Next repetition
    End If
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    failure = CStr(Err.Number) & ":" & Err.Description
    RunCase = "FAIL|" & name & "|" & failure
Clean:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
End Function
