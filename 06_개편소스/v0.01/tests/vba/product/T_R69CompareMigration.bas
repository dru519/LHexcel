Attribute VB_Name = "T_R69CompareMigration"
Option Explicit
Public DisableHost As Boolean, UsedNative As Boolean
Private enhanced As Boolean
Public Sub Configure(ByVal value As Boolean)
    enhanced = value
End Sub
Public Function Names() As String
    Names = "values|formulas|literal_formula|all_formulas|formats|merged|errors|no_dll"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim a As Workbook, b As Workbook, leftRange As Range, rightRange As Range
    Dim reasons As Variant, fallback As Variant, service As Object, expected As String, i As Long, caught As Long, detail As String
    Dim nativeMs As Double, fallbackMs As Double, started As Double, handle As Integer
    On Error GoTo Failed
    DisableHost = False: UsedNative = False
    Set service = NxHostCreateWorkbookCompare()
    If service Is Nothing Then Err.Raise 5, , "Registered DLL comparison service unavailable"
    Set a = Workbooks.Add(xlWBATWorksheet): Set b = Workbooks.Add(xlWBATWorksheet)
    Set leftRange = a.Worksheets(1).Range("A1:T20")
    Set rightRange = b.Worksheets(1).Range("A1:T20")
    leftRange.Value2 = 1: rightRange.Value2 = 1
    expected = "값"
    Select Case name
        Case "values", "no_dll": rightRange.Cells(1, 1).Value2 = 2
        Case "formulas"
            leftRange.Cells(1, 1).Formula = "=1+2": rightRange.Cells(1, 1).Formula = "=2+1"
            expected = "수식"
        Case "literal_formula"
            leftRange.Cells(1, 1).Value2 = "'=1+2": rightRange.Cells(1, 1).Value2 = "'=2+1"
        Case "all_formulas"
            leftRange.Formula = "=1+2": rightRange.Formula = "=1+2"
            rightRange.Cells(1, 1).Formula = "=2+1": expected = "수식"
        Case "formats": rightRange.Cells(1, 1).Font.Bold = True: expected = "서식"
        Case "merged"
            leftRange.Range("A1:B1").ClearContents: rightRange.Range("A1:B1").ClearContents
            leftRange.Range("A1:B1").Merge
            expected = "병합"
        Case "errors"
            leftRange.Cells(1, 1).Value = CVErr(xlErrNA): rightRange.Cells(1, 1).Value = CVErr(xlErrDiv0)
    End Select
    Application.CalculateFull
    DoEvents
    a.SaveAs Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator & name & "-left.xlsx", xlOpenXMLWorkbook
    b.SaveAs Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator & name & "-right.xlsx", xlOpenXMLWorkbook
    a.Saved = True: b.Saved = True
    DisableHost = (name = "no_dll")
    On Error Resume Next
    started = Timer
    reasons = NxR69Reasons(leftRange, rightRange, True)
    nativeMs = (Timer - started) * 1000#
    caught = Err.Number: detail = Err.Description: Err.Clear
    On Error GoTo Failed
    If Not a.Saved Or Not b.Saved Then Err.Raise 5, , "Native capture changed saved state: " & CStr(a.Saved) & "/" & CStr(b.Saved)
    If name = "no_dll" And enhanced Then
        If caught = 0 Or InStr(detail, "DLL") = 0 Then Err.Raise 5, , "Missing DLL did not fail with guidance"
    Else
        If caught <> 0 Then Err.Raise caught, , detail
        If UsedNative = DisableHost Then Err.Raise 5, , "Wrong comparison engine"
        If reasons(0) <> expected Then Err.Raise 5, , "Unexpected first reason: " & reasons(0) & ";left=" & NxR69ErrorProbe(leftRange.Cells(1, 1).Value2) & ";right=" & NxR69ErrorProbe(rightRange.Cells(1, 1).Value2)
        For i = 1 To UBound(reasons)
            If name = "merged" And i = 1 Then
                If reasons(i) <> "병합" Then Err.Raise 5, , "Merge follower mismatch"
            Else
                If reasons(i) <> "" Then Err.Raise 5, , "Unexpected difference at " & i & ": " & reasons(i)
            End If
        Next i
        If Not enhanced And name <> "no_dll" Then
            DisableHost = True
            started = Timer
            fallback = NxR69Reasons(leftRange, rightRange, True)
            fallbackMs = (Timer - started) * 1000#
            For i = 0 To UBound(reasons)
                If fallback(i) <> reasons(i) Then Err.Raise 5, , "DLL/VBA parity mismatch at " & i
            Next i
            handle = FreeFile
            Open Environ$("LHEXCEL_PROFILE_ROOT") & Application.PathSeparator & "compare-timing.tsv" For Append As #handle
            Print #handle, name & vbTab & nativeMs & vbTab & fallbackMs
            Close #handle
        End If
    End If
    If Not a.Saved Or Not b.Saved Then Err.Raise 5, , "Comparison changed workbook saved state: " & CStr(a.Saved) & "/" & CStr(b.Saved)
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & "|" & Err.Description
CleanUp:
    On Error Resume Next
    DisableHost = False
    If Not a Is Nothing Then a.Close False
    If Not b Is Nothing Then b.Close False
End Function
