Attribute VB_Name = "NxHwpxController"
Option Explicit

Private Const NX_HWPX_TITLE As String = "내엑셀 - 아래한글 표 전송하기"
Private Const NX_HWPX_MAX_CELLS As Long = 16384
Private mHwpxLastStage As String
Private mHwpxBusy As Boolean
#If VBA7 Then
Private Declare PtrSafe Sub NxHwpxSleep Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
#Else
Private Declare Sub NxHwpxSleep Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
#End If

Public Function NxHwpxSerializeWithSettings(ByVal source As Range, ByVal settings As CNxHangulSettings) As String
    Dim payload As String
    If source Is Nothing Or settings Is Nothing Or Not settings.IsSealed Then NxRaiseContractError "한글 표 전송 설정을 확인하세요."
    If CDbl(source.CountLarge) > NX_HWPX_MAX_CELLS Then NxRaiseContractError "아래한글 표는 최대 16,384셀입니다."
    payload = NxHwpxSerialize(source, settings.TitleRow, settings.IncludeHidden, "display")
    If Right$(payload, 1) <> "}" Then NxRaiseContractError "HWPX 전송 데이터가 올바르지 않습니다."
    payload = Left$(payload, Len(payload) - 1) & _
        ",""settings"":{""font_name"":""" & JsonEscape(settings.FontName) & """,""font_type"":""HFT"",""font_size_pt"":" & _
        CStr(settings.FontSizePt) & ",""title_row"":" & JsonBoolean(settings.TitleRow) & _
        ",""include_hidden"":" & JsonBoolean(settings.IncludeHidden) & _
        ",""table_style"":""" & JsonEscape(settings.TableStyle) & """}}"
    NxHwpxSerializeWithSettings = payload
End Function

Public Function NxHwpxSendTable(ByVal source As Range, ByVal settings As CNxHangulSettings) As String
    Dim outputPath As String, inputPath As String, resultJson As String, payload As String
    Dim validatedOutput As Boolean
    On Error GoTo Failed
    If source Is Nothing Or settings Is Nothing Or Not settings.IsSealed Then NxRaiseContractError "아래한글 표 전송 대상을 확인하세요."
    mHwpxLastStage = "cleanup"
    NxHwpxCleanupExpired
    mHwpxLastStage = "serialize"
    payload = NxHwpxSerializeWithSettings(source, settings)
    mHwpxLastStage = "dll"
    If Not NxHwpxTryNative(payload, outputPath) Then
        mHwpxLastStage = "output-path"
        outputPath = NxHwpxNewOutputPath()
        inputPath = Left$(outputPath, Len(outputPath) - 5) & ".json"
        WriteUtf8 inputPath, payload
        mHwpxLastStage = "powershell"
        resultJson = NxPowerShellRunHwpx(inputPath, outputPath, 60)
    End If
    mHwpxLastStage = "verify-output"
    If Len(Dir$(outputPath)) = 0 Then NxRaiseContractError "검증된 HWPX 결과가 생성되지 않았습니다."
    validatedOutput = True
    On Error Resume Next
    If Len(Dir$(inputPath)) > 0 Then Kill inputPath
    On Error GoTo Failed
    mHwpxLastStage = "open-hwpx"
    Dim shell As Object
    Set shell = CreateObject("Shell.Application")
    shell.ShellExecute outputPath, vbNullString, vbNullString, "open", 1
    NxHwpxSendTable = outputPath
    Exit Function
Failed:
    Dim detail As String
    detail = Err.Description
    On Error Resume Next
    If Len(inputPath) > 0 And Len(Dir$(inputPath)) > 0 Then Kill inputPath
    If Not validatedOutput And Len(outputPath) > 0 And Len(Dir$(outputPath)) > 0 Then Kill outputPath
    On Error GoTo 0
    If Len(detail) = 0 Then detail = "아래한글 표를 생성하지 못했습니다."
    detail = mHwpxLastStage & ": " & Err.Source & ": " & detail
    If validatedOutput Then detail = detail & vbCrLf & "생성된 파일은 보존했습니다: " & outputPath
    Err.Raise vbObjectError + 2191, NX_HWPX_TITLE, detail
End Function

Public Function NxHwpxTryNative(ByVal payload As String, ByRef outputPath As String, Optional ByVal timeoutSeconds As Long = 60) As Boolean
    Dim service As Object, jobId As String, status As String, started As Double, elapsed As Double
    Dim previousCancel As XlEnableCancelKey, detail As String
    If mHwpxBusy Then NxRaiseContractError "아래한글 표를 생성 중입니다. 완료 후 다시 실행하세요."
    Set service = NxHostCreateHwpxService()
    If service Is Nothing Then Exit Function
    previousCancel = Application.EnableCancelKey
    mHwpxBusy = True
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    jobId = CStr(service.Start(payload, NxHwpxEnsureEmbeddedTemplate()))
    started = Timer
    Do
        status = CStr(service.GetStatus(jobId))
        If status <> "running" Then Exit Do
        DoEvents
        NxHwpxSleep 25
        elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400#
        If elapsed >= timeoutSeconds Then NxRaiseContractError "아래한글 표 생성 시간이 초과되었습니다."
    Loop
    If status <> "succeeded" Then NxRaiseContractError "아래한글 표를 생성하지 못했습니다: " & CStr(service.GetError(jobId))
    outputPath = CStr(service.GetResult(jobId))
    NxHwpxTryNative = True
    Application.EnableCancelKey = previousCancel
    mHwpxBusy = False
    Exit Function
Failed:
    detail = Err.Description
    If Err.Number = 18 Then detail = "아래한글 표 생성을 취소했습니다."
    On Error Resume Next
    If Len(jobId) > 0 Then service.Cancel jobId
    Application.EnableCancelKey = previousCancel
    mHwpxBusy = False
    On Error GoTo 0
    ' No fallback after Start: an accepted native job must never be executed a second time.
    Err.Raise vbObjectError + 2192, NX_HWPX_TITLE, detail
End Function

Public Function NxHwpxLastStage() As String
    NxHwpxLastStage = mHwpxLastStage
End Function

Public Function NxHwpxNewOutputPath() As String
    Dim root As String
    root = NxHwpxRuntimeOutputRoot()
    NxHwpxEnsureFolder root
    NxHwpxNewOutputPath = root & "\" & NxCreateRunUuid() & ".hwpx"
End Function

Public Sub NxHwpxCleanupExpired()
    Dim root As String, fileSystem As Object, item As Object, cutoff As Date
    On Error Resume Next
    root = NxHwpxRuntimeOutputRoot()
    If Len(Dir$(root, vbDirectory)) = 0 Then Exit Sub
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    cutoff = DateAdd("h", -24, Now)
    For Each item In fileSystem.GetFolder(root).Files
        If LCase$(fileSystem.GetExtensionName(item.Name)) = "hwpx" Or _
            LCase$(fileSystem.GetExtensionName(item.Name)) = "json" Or _
            LCase$(fileSystem.GetExtensionName(item.Name)) = "log" Then
            If item.DateLastModified < cutoff Then fileSystem.DeleteFile item.Path, True
        End If
    Next item
    On Error GoTo 0
End Sub

Private Sub WriteUtf8(ByVal path As String, ByVal value As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText value
    stream.SaveToFile path, 2
    stream.Close
End Sub

Private Sub NxHwpxEnsureFolder(ByVal path As String)
    Dim fileSystem As Object, parent As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FolderExists(path) Then Exit Sub
    parent = fileSystem.GetParentFolderName(path)
    If Len(parent) > 0 And Not fileSystem.FolderExists(parent) Then NxHwpxEnsureFolder parent
    fileSystem.CreateFolder path
End Sub

Private Function JsonBoolean(ByVal value As Boolean) As String
    If value Then JsonBoolean = "true" Else JsonBoolean = "false"
End Function

Private Function JsonEscape(ByVal value As String) As String
    value = Replace$(value, "\", "\\")
    value = Replace$(value, Chr$(34), "\" & Chr$(34))
    value = Replace$(value, vbCr, "\r")
    value = Replace$(value, vbLf, "\n")
    JsonEscape = value
End Function
