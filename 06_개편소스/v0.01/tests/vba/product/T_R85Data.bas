Attribute VB_Name = "T_R85Data"
Option Explicit
Public Function Names() As String
    Names = "normalize90k|date90k|age90k|mask90k|mixed_preview|ai_clipboard"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, sheet As Worksheet, source As Range, output As Worksheet
    Dim options As New CNxDataSpecialOptions, result As CNxResult
    Dim text As String, feature As String, preset As String, inputValue As String, expected As String
    Dim started As Double, h As Integer, oldEvents As Boolean, oldAlerts As Boolean
    Dim clip As Object
    On Error GoTo Failed
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False: Application.DisplayAlerts = False
    Set book = Workbooks.Add(xlWBATWorksheet): Set sheet = book.Worksheets(1)
    If name = "ai_clipboard" Then
        Set result = NxAiCopyPrompt("내엑셀 r85 합성 테스트", "NX-AI-SUMMARY")
        If result.Outcome <> NxSuccess Then Err.Raise 5, , "Clipboard copy failed"
        Set clip = New MSForms.DataObject
        clip.GetFromClipboard
        If clip.GetText(1) <> "내엑셀 r85 합성 테스트" Then Err.Raise 5, , "Clipboard roundtrip mismatch"
        If Not NxAiOpenTargetAfterCopy("ChatGPT", NxFrameSuccess) Then Err.Raise 5, , "AI browser dispatch failed"
        GoTo Passed
    End If
    options.Configure "새 시트 만들기", False, "2026-09-14", "숫자", False, False, False, "빈칸", "자동", "*"
    Select Case name
        Case "normalize90k": preset = "기본 정리": inputValue = "  sample  value ": expected = "sample value"
        Case "date90k": preset = "날짜 정리": inputValue = "240617": expected = "2024-06-17"
        Case "age90k": feature = NX_FEATURE_DATA_AGE: inputValue = "20000617": expected = "26"
        Case "mask90k", "mixed_preview": feature = NX_FEATURE_DATA_PRIVACY_MASK: inputValue = "01012345678": expected = "010-****-5678"
        Case Else: Err.Raise 5, , "Unknown case"
    End Select
    Set source = sheet.Range("B2").Resize(1800, 50)
    source.NumberFormat = "@": source.Value2 = inputValue
    If name = "mixed_preview" Then
        Set source = sheet.Range("B2:C4")
        source.ClearContents
        source.NumberFormat = "General"
        source.Cells(1, 1).Formula = "=" & Chr$(34) & "01012345678" & Chr$(34)
        source.Cells(1, 2).Formula = "=" & Chr$(34) & Chr$(34)
        source.Cells(2, 1).Value2 = CVErr(xlErrNA)
        source.Cells(3, 1).NumberFormat = "@"
        source.Cells(3, 1).Value2 = "01012345678"
    End If
    source.Select: book.Saved = True
    started = Timer
    If Len(feature) = 0 Then
        text = NxDataNormalizationPreview(source, preset, "yyyy-mm-dd")
    Else
        text = NxDataSpecialPreview(feature, source, options)
    End If
    LogTime name & "_preview", started
    If Not book.Saved Then Err.Raise 5, , "Preview dirtied source"
    If name = "mixed_preview" Then
        If InStr(text, "빈 셀 제외 3개") = 0 Or InStr(text, "빈칸 처리 1개") = 0 Or InStr(text, "수식→값 1개") = 0 Then Err.Raise 5, , text
        GoTo Passed
    End If
    started = Timer
    If Len(feature) = 0 Then
        Set result = NxDataRunNormalizationForTest(source, preset, "yyyy-mm-dd", "새 시트 만들기")
    Else
        Set result = NxDataSpecialRunOptions(feature, source, options, True)
    End If
    LogTime name & "_execute", started
    If result.Outcome <> NxSuccess Then Err.Raise 5, , "Execution failed"
    If book.Worksheets.Count <> 2 Then Err.Raise 5, , "Result missing"
    If book.Worksheets(1) Is sheet Then Set output = book.Worksheets(2) Else Set output = book.Worksheets(1)
    If name = "date90k" Then
        If Format$(output.Range("A1").Value, "yyyy-mm-dd") <> expected Then Err.Raise 5, , "Date output mismatch"
    Else
        If CStr(output.Range("A1").Value2) <> expected Or CStr(output.Cells(1800, 50).Value2) <> expected Then Err.Raise 5, , "Output mismatch"
    End If
    If CStr(source.Cells(1, 1).Value2) <> inputValue Then Err.Raise 5, , "Original changed"
Passed:
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
Clean:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
    Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
End Function
Private Sub LogTime(ByVal name As String, ByVal started As Double)
    Dim h As Integer, elapsed As Double
    elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400
    h = FreeFile
    Open Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator & "r85-data-timing.tsv" For Append As #h
    Print #h, name & vbTab & Format$(elapsed, "0.000")
    Close #h
End Sub
