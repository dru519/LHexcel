Attribute VB_Name = "T_NxHwpxNative"
Option Explicit
Private mBook As Workbook
Private mSource As Range
Private mPayload As String

Public Sub NxHwpxNativeReloadPayload(ByVal root As String)
    ' Test-copy code replacement resets module state. Reload the immutable fixture only.
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.LoadFromFile root & "\excel-payload.json": mPayload = stream.ReadText: stream.Close
End Sub

Private Function StatusUnchanged(ByVal previous As Variant) As Boolean
    Dim current As Variant
    current = Application.StatusBar
    StatusUnchanged = (VarType(current) = VarType(previous) And CStr(current) = CStr(previous))
End Function

Public Function NxHwpxNativePrepare(ByVal root As String) As String
    Set mBook = Application.Workbooks.Add
    Set mSource = mBook.Worksheets(1).Range("A1:C4")
    mSource.Cells(1, 1).Value2 = " " & vbTab & "구분" & ChrW(160): mSource.Cells(1, 2).Value2 = "내용": mSource.Cells(1, 3).Value2 = "금액"
    mSource.Cells(2, 1).Value2 = "항목": mSource.Range("B2:C2").Merge
    mSource.Cells(2, 2).Value2 = ChrW(&H3000) & "한글 <표> &  인용" & vbLf & "두 줄" & ChrW(&H3000)
    mSource.Cells(3, 1).Value2 = "수식": mSource.Cells(3, 2).Formula = "=1234+5"
    mSource.Cells(3, 3).Value2 = 2500: mSource.Cells(3, 3).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
    mSource.Cells(4, 1).Value2 = "숨김"
    mSource.ColumnWidth = 14: mSource.RowHeight = 22
    mSource.Rows(4).EntireRow.Hidden = True
    mSource.Interior.Color = RGB(235, 242, 250)
    Dim settings As New CNxHangulSettings
    settings.Configure "중고딕", "HFT", 13, True, False, "공문형": settings.Seal
    mPayload = NxHwpxSerializeWithSettings(mSource, settings)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText mPayload: stream.SaveToFile root & "\excel-payload.json", 2: stream.Close
    mBook.SaveAs root & "\fixture.xlsx", xlOpenXMLWorkbook
    NxHwpxNativePrepare = "PASS|Prepared"
End Function

Public Function NxHwpxNativeRun(ByVal mode As String) As String
    Dim output As String, used As Boolean, cancelMode As XlEnableCancelKey, status As Variant
    Dim number As Long, detail As String, payload As String
    If mode = "custom" Then Application.StatusBar = "Native status restoration probe"
    cancelMode = Application.EnableCancelKey: status = Application.StatusBar
    payload = mPayload
    If mode = "invalid" Then payload = "{}"
    On Error GoTo Failed
    used = NxHwpxTryNative(payload, output)
    If mode = "unavailable" Then
        If used Then Err.Raise 5, , "Unexpected native connection"
        NxHwpxNativeRun = "PASS|Unavailable|" & NxHostHwpxConnectionStatus()
    Else
        If Not used Or Len(Dir$(output)) = 0 Then Err.Raise 5, , "Native generator unavailable: " & NxHostHwpxConnectionStatus()
        NxHwpxNativeRun = "PASS|" & output
    End If
    If Not mBook.Saved Then Err.Raise 5, , "Workbook Saved state changed"
    If mSource.Cells(3, 2).Formula <> "=1234+5" Then Err.Raise 5, , "Source formula changed"
    If mSource.Cells(1, 1).Value2 <> " " & vbTab & "구분" & ChrW(160) Then Err.Raise 5, , "Source whitespace changed"
    If mSource.Interior.Color <> RGB(235, 242, 250) Then Err.Raise 5, , "Source fill changed"
    If Not mSource.Range("B2:C2").MergeCells Then Err.Raise 5, , "Merge changed"
    If Application.EnableCancelKey <> cancelMode Then Err.Raise 5, , "Cancel mode not restored"
    If Not StatusUnchanged(status) Then Err.Raise 5, , "Status bar changed; before=" & TypeName(status) & ":" & CStr(status) & "; after=" & TypeName(Application.StatusBar) & ":" & CStr(Application.StatusBar)
    Exit Function
Failed:
    number = Err.Number: detail = Err.Description
    If mode = "invalid" And InStr(detail, "INVALID_INPUT_OR_TEMPLATE") > 0 And Len(output) = 0 And _
        Application.EnableCancelKey = cancelMode And StatusUnchanged(status) Then
        NxHwpxNativeRun = "PASS|InvalidRejected|StateRestored": Exit Function
    End If
    If mode = "cancel" And InStr(detail, "취소") > 0 And Len(output) = 0 And _
        Application.EnableCancelKey = cancelMode And StatusUnchanged(status) Then
        NxHwpxNativeRun = "PASS|CancelHandler|StateRestored": Exit Function
    End If
    NxHwpxNativeRun = "FAIL|" & CStr(number) & "|" & detail
End Function

Public Function NxHwpxNativeLegacy(ByVal root As String) As String
    Dim result As String
    result = NxPowerShellRunHwpx(root & "\excel-payload.json", root & "\legacy.hwpx", 60)
    If Len(Dir$(root & "\legacy.hwpx")) = 0 Then Err.Raise 5, , "Fallback output missing"
    NxHwpxNativeLegacy = "PASS|Legacy"
End Function

Public Sub NxHwpxNativeDispose()
    On Error Resume Next
    If Not mBook Is Nothing Then mBook.Close False
    Set mSource = Nothing: Set mBook = Nothing
End Sub
