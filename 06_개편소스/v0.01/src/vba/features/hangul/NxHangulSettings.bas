Attribute VB_Name = "NxHangulSettings"
Option Explicit

Private Const NX_HANGUL_SETTINGS_FILE As String = "hangul-table-v1.cfg"

Public Function NxHangulDefaultSettings() As CNxHangulSettings
    Dim settings As New CNxHangulSettings
    settings.Configure "중고딕", "HFT", 13, True, True, "공문형"
    settings.Seal
    Set NxHangulDefaultSettings = settings
End Function

Public Function NxHangulSettingsPath() As String
    NxHangulSettingsPath = NxLHexcelProfileRoot() & "\Settings\" & NX_HANGUL_SETTINGS_FILE
End Function

Public Function NxHangulLoadSettings() As CNxHangulSettings
    Dim text As String, settings As CNxHangulSettings
    On Error Resume Next
    text = NxHangulReadText(NxHangulSettingsPath())
    Set settings = NxHangulParseSettings(text)
    On Error GoTo 0
    If settings Is Nothing Then Set settings = NxHangulDefaultSettings()
    Set NxHangulLoadSettings = settings
End Function

Public Sub NxHangulSaveSettings(ByVal settings As CNxHangulSettings)
    Dim folder As String, path As String, temporary As String, backup As String
    Dim handle As Integer, payload As String, verified As CNxHangulSettings
    Dim failureNumber As Long, failureDescription As String
    If settings Is Nothing Then NxRaiseContractError "저장할 한글 표 설정이 없습니다."
    If Not settings.IsSealed Then NxRaiseContractError "저장할 한글 표 설정이 없습니다."
    folder = NxLHexcelProfileRoot() & "\Settings"
    If Len(Dir$(NxLHexcelProfileRoot(), vbDirectory)) = 0 Then MkDir NxLHexcelProfileRoot()
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    path = NxHangulSettingsPath()
    temporary = path & ".tmp"
    backup = path & ".bak"
    payload = NxHangulSettingsPayload(settings)
    On Error GoTo Failed
    If Len(Dir$(temporary)) > 0 Then Kill temporary
    handle = FreeFile
    Open temporary For Output Access Write Lock Read Write As #handle
    Print #handle, payload;
    Close #handle
    handle = 0
    Set verified = NxHangulParseSettings(NxHangulReadText(temporary))
    If verified Is Nothing Then NxRaiseContractError "한글 표 설정 확인에 실패했습니다."
    If NxHangulSettingsPayload(verified) <> payload Then NxRaiseContractError "한글 표 설정 확인에 실패했습니다."
    If Len(Dir$(backup)) > 0 Then Kill backup
    If Len(Dir$(path)) > 0 Then Name path As backup
    Name temporary As path
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If handle > 0 Then Close #handle
    If Len(Dir$(temporary)) > 0 Then Kill temporary
    If Len(Dir$(path)) = 0 And Len(Dir$(backup)) > 0 Then Name backup As path
    On Error GoTo 0
    If Len(failureDescription) = 0 Then failureDescription = "설정 파일을 쓸 수 없습니다."
    Err.Raise failureNumber, "NxHangulSettings", "한글 표 설정을 저장하지 못했습니다. " & failureDescription
End Sub

Public Sub NxHangulShowSettings()
    Dim form As New FNxHangulSettings
    form.BindSettings NxHangulLoadSettings()
    form.Show vbModal
End Sub

Public Sub NxHangulApplySettingsValues(ByVal fontName As String, ByVal fontSizePt As Long, _
    ByVal titleRow As Boolean, ByVal includeHidden As Boolean, ByVal tableStyle As String)
    Dim settings As New CNxHangulSettings
    settings.Configure fontName, "HFT", fontSizePt, titleRow, includeHidden, tableStyle
    settings.Seal
    NxHangulSaveSettings settings
End Sub

Public Function NxHangulTryApplySettingsValues(ByVal fontName As String, ByVal fontSizePt As Long, _
    ByVal titleRow As Boolean, ByVal includeHidden As Boolean, ByVal tableStyle As String) As String
    On Error GoTo Failed
    NxHangulApplySettingsValues fontName, fontSizePt, titleRow, includeHidden, tableStyle
    NxHangulTryApplySettingsValues = "PASS"
    Exit Function
Failed:
    NxHangulTryApplySettingsValues = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Private Function NxHangulSettingsPayload(ByVal settings As CNxHangulSettings) As String
    NxHangulSettingsPayload = "schema_version=1" & vbLf & _
        "font_name=" & settings.FontName & vbLf & _
        "font_type=" & settings.FontType & vbLf & _
        "font_size_pt=" & CStr(settings.FontSizePt) & vbLf & _
        "title_row=" & CStr(Abs(CLng(settings.TitleRow))) & vbLf & _
        "include_hidden=" & CStr(Abs(CLng(settings.IncludeHidden))) & vbLf & _
        "table_style=" & settings.TableStyle & vbLf
End Function

Private Function NxHangulParseSettings(ByVal payload As String) As CNxHangulSettings
    Dim values As Object, line As Variant, parts As Variant, settings As New CNxHangulSettings
    On Error GoTo Invalid
    If Len(payload) = 0 Then Exit Function
    Set values = CreateObject("Scripting.Dictionary")
    payload = Replace$(payload, vbCr, vbNullString)
    For Each line In Split(payload, vbLf)
        If Len(CStr(line)) > 0 Then
            parts = Split(CStr(line), "=", 2)
            If UBound(parts) <> 1 Or values.Exists(parts(0)) Then Exit Function
            values.Add parts(0), parts(1)
        End If
    Next line
    If values.Count <> 7 Or values("schema_version") <> "1" Then Exit Function
    settings.Configure values("font_name"), values("font_type"), CLng(values("font_size_pt")), _
        values("title_row") = "1", values("include_hidden") = "1", values("table_style")
    settings.Seal
    Set NxHangulParseSettings = settings
    Exit Function
Invalid:
    Set NxHangulParseSettings = Nothing
End Function

Private Function NxHangulReadText(ByVal path As String) As String
    Dim handle As Integer, currentLine As String, result As String
    On Error GoTo Failed
    If Len(Dir$(path)) = 0 Then Exit Function
    handle = FreeFile
    Open path For Input Access Read As #handle
    Do Until EOF(handle)
        Line Input #handle, currentLine
        result = result & currentLine & vbLf
    Loop
    NxHangulReadText = result
    Close #handle
    Exit Function
Failed:
    On Error Resume Next
    If handle > 0 Then Close #handle
End Function
