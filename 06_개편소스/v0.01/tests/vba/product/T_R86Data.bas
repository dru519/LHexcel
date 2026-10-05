Attribute VB_Name = "T_R86Data"
Option Explicit
Private mName As String, mStarted As Double
Public Function Names() As String
    Names = "age_original|mask_original|age_columns|mask_columns|age_sheet|mask_sheet|age_book|mask_book"
End Function
Public Sub Mark(ByVal stage As String)
    Dim h As Integer, elapsed As Double
    elapsed = Timer - mStarted: If elapsed < 0 Then elapsed = elapsed + 86400
    h = FreeFile
    Open Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator & "r86-stages.tsv" For Append As #h
    Print #h, mName & vbTab & stage & vbTab & Format$(elapsed, "0.000")
    Close #h
End Sub
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, resultBook As Workbook, source As Range, output As Range
    Dim sheet As Worksheet, resultSheet As Worksheet, options As New CNxDataSpecialOptions
    Dim mode As String, feature As String, inputValue As String, expected As String
    Dim result As CNxResult, oldEvents As Boolean, oldAlerts As Boolean
    On Error GoTo Failed
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False: Application.DisplayAlerts = False
    Set book = Workbooks.Add(xlWBATWorksheet): Set sheet = book.Worksheets(1)
    Set source = sheet.Range("B2").Resize(1800, 50)
    If Left$(name, 3) = "age" Then
        feature = NX_FEATURE_DATA_AGE: inputValue = "20000617": expected = "26"
    Else
        feature = NX_FEATURE_DATA_PRIVACY_MASK: inputValue = "01012345678": expected = "010-****-5678"
    End If
    Select Case Split(name, "_")(1)
        Case "original": mode = "원본 변경"
        Case "columns": mode = "오른쪽 새 열 삽입"
        Case "sheet": mode = "새 시트 만들기"
        Case "book": mode = "새 통합문서"
    End Select
    source.NumberFormat = "@": source.Value2 = inputValue
    source.Rows(1).Value2 = "=literal"
    source.Cells(2, 1).ClearContents
    source.Cells(2, 2).Value2 = CVErr(xlErrNA)
    sheet.Range("AZ2").Value2 = "neighbor"
    options.Configure mode, True, "2026-09-15", "숫자", False, False, False, "빈칸", "자동", "*"
    source.Select: book.Saved = True
    mName = name: mStarted = Timer: Mark "start"
    Set result = NxDataSpecialRunOptions(feature, source, options, True)
    Mark "complete"
    If result.Outcome <> NxSuccess Then Err.Raise 5, , "Execution failed"
    Select Case mode
        Case "원본 변경": Set output = source
        Case "오른쪽 새 열 삽입": Set output = sheet.Range("AZ2").Resize(1800, 50)
        Case "새 시트 만들기"
            For Each resultSheet In book.Worksheets
                If Not resultSheet Is sheet Then Set output = resultSheet.Range("A1").Resize(1800, 50)
            Next resultSheet
        Case "새 통합문서"
            Set resultBook = ActiveWorkbook
            If resultBook Is book Then Err.Raise 5, , "New workbook missing"
            Set output = resultBook.Worksheets(1).Range("A1").Resize(1800, 50)
    End Select
    If output Is Nothing Then Err.Raise 5, , "Output missing"
    If CStr(output.Cells(1800, 50).Value2) <> expected Then Err.Raise 5, , "Output mismatch"
    If output.Cells(1, 1).Value2 <> "=literal" Or output.Cells(1, 1).HasFormula Then Err.Raise 5, , "Literal header changed"
    If Len(CStr(output.Cells(2, 1).Value2)) <> 0 Or Len(CStr(output.Cells(2, 2).Value2)) <> 0 Then Err.Raise 5, , "Blank/error mismatch"
    If mode <> "원본 변경" Then
        If source.Cells(1800, 50).Value2 <> inputValue Then Err.Raise 5, , "Source changed"
    End If
    If mode = "오른쪽 새 열 삽입" Then
        If sheet.Cells(2, 102).Value2 <> "neighbor" Then Err.Raise 5, , "Neighbor displaced incorrectly"
    Else
        If sheet.Range("AZ2").Value2 <> "neighbor" Then Err.Raise 5, , "Neighbor changed"
    End If
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
Clean:
    On Error Resume Next
    If Not resultBook Is Nothing Then
        If Not resultBook Is book Then resultBook.Close False
    End If
    If Not book Is Nothing Then book.Close False
    Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
End Function
