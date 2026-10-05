Attribute VB_Name = "NxHangulPictureTransfer"
Option Explicit

Public Sub NxHangulOpenPictures()
    Dim panel As New FNxHangulPictures
    panel.Show vbModal
    Unload panel
End Sub

Public Function NxHangulPicturePayload(ByVal paths As Collection, ByVal widthCm As Double, ByVal heightCm As Double, _
        ByVal columns As Long, ByVal rows As Long, ByVal titles As Boolean, ByVal sortMode As String, _
        Optional ByVal titleFileNames As Boolean = True) As String
    Dim path As Variant, files As String, captionHeight As Double
    If paths Is Nothing Then NxRaiseContractError "그림 파일을 선택하세요."
    If paths.Count < 1 Or paths.Count > 200 Then NxRaiseContractError "그림은 1~200개를 선택하세요."
    If widthCm < 0.5 Or heightCm < 0.5 Or columns < 1 Or rows < 1 Or columns > 8 Or rows > 12 Then _
        NxRaiseContractError "그림 크기와 가로·세로 개수를 확인하세요."
    If titles Then captionHeight = 0.7
    If widthCm * columns > 16.7 Or (heightCm + captionHeight) * rows > 24 Then _
        NxRaiseContractError "한 쪽의 배치는 가로 16.7cm, 세로 24cm 이내로 설정하세요. 제목행은 0.7cm입니다."
    If sortMode <> "title" And sortMode <> "taken" Then NxRaiseContractError "정렬 순서를 선택하세요."
    For Each path In paths
        If Len(files) > 0 Then files = files & ","
        files = files & """" & JsonText(CStr(path)) & """"
    Next path
    NxHangulPicturePayload = "{""schema_version"":1,""picture_layout"":{""width_cm"":" & Trim$(Str$(widthCm)) & _
        ",""height_cm"":" & Trim$(Str$(heightCm)) & ",""columns"":" & columns & ",""rows"":" & rows & _
        ",""titles"":" & IIf(titles, "true", "false") & ",""title_filenames"":" & IIf(titleFileNames, "true", "false") & _
        ",""sort"":""" & sortMode & """,""files"": [" & files & "]}}"
End Function

Public Function NxHangulSendPictures(ByVal payload As String) As String
    Dim outputPath As String, inputPath As String, stream As Object, result As String, detail As String
    Dim complete As Boolean
    On Error GoTo Failed
    NxDistributionEnsureExecutable
    If Not NxHwpxTryNative(payload, outputPath, 120) Then
        outputPath = NxHwpxNewOutputPath()
        inputPath = Left$(outputPath, Len(outputPath) - 5) & ".json"
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2: stream.Charset = "utf-8": stream.Open
        stream.WriteText payload: stream.SaveToFile inputPath, 1: stream.Close
        result = NxPowerShellRunHwpx(inputPath, outputPath, 120)
    End If
    If Len(Dir$(outputPath)) = 0 Then NxRaiseContractError "그림 표가 생성되지 않았습니다."
    complete = True
    If Len(inputPath) > 0 And Len(Dir$(inputPath)) > 0 Then Kill inputPath
    CreateObject("Shell.Application").ShellExecute outputPath, vbNullString, vbNullString, "open", 1
    NxHangulSendPictures = outputPath
    Exit Function
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    If Len(inputPath) > 0 Then Kill inputPath
    If Not complete And Len(outputPath) > 0 Then Kill outputPath
    On Error GoTo 0
    If complete Then detail = detail & vbCrLf & "생성된 파일: " & outputPath
    Err.Raise vbObjectError + 2194, "아래한글 그림 전송", detail
End Function

Private Function JsonText(ByVal value As String) As String
    value = Replace$(value, "\", "\\")
    value = Replace$(value, """", "\""")
    value = Replace$(value, vbCr, "\r")
    value = Replace$(value, vbLf, "\n")
    JsonText = Replace$(value, vbTab, "\t")
End Function
