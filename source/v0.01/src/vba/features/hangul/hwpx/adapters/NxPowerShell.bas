Attribute VB_Name = "NxPowerShell"
Option Explicit

Public Function NxPowerShellRunHwpx(ByVal inputPath As String, ByVal outputPath As String, Optional ByVal timeoutSeconds As Long = 60) As String
    Dim resources As Variant, command As String
    If timeoutSeconds <= 0 Then NxRaiseContractError "PowerShell timeout must be positive"
    resources = NxHwpxEnsureEmbeddedResources()
    command = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " & NxPowerShellQuoted(CStr(resources(0))) & _
        " -InputPath " & NxPowerShellQuoted(inputPath) & _
        " -OutputPath " & NxPowerShellQuoted(outputPath) & _
        " -TemplatePath " & NxPowerShellQuoted(CStr(resources(1)))
    NxPowerShellRunHwpx = RunPowerShellCommand(command, timeoutSeconds)
End Function

Public Function NxPowerShellQuoted(ByVal value As String) As String
    If Len(value) = 0 Or InStr(value, Chr$(34)) > 0 Or InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then
        NxRaiseContractError "PowerShell 인수 경로가 안전하지 않습니다."
    End If
    NxPowerShellQuoted = Chr$(34) & value & Chr$(34)
End Function

Private Function RunPowerShellCommand(ByVal command As String, ByVal timeoutSeconds As Long) As String
    Dim shell As Object, process As Object
    Dim started As Single, output As String, errorOutput As String
    On Error GoTo Failed
    Set shell = CreateObject("WScript.Shell")
    Set process = shell.Exec(command)
    started = Timer
    Do While process.Status = 0
        DoEvents
        If ElapsedSeconds(started) > timeoutSeconds Then
            process.Terminate
            NxRaiseContractError "HWPX 생성 시간이 " & CStr(timeoutSeconds) & "초를 초과했습니다."
        End If
    Loop
    output = process.StdOut.ReadAll
    errorOutput = process.StdErr.ReadAll
    If process.ExitCode <> 0 Then NxRaiseContractError "HWPX 생성기가 실패했습니다: " & SafeDiagnostic(errorOutput)
    If InStr(1, output, """status"":""PASS""", vbBinaryCompare) = 0 Then NxRaiseContractError "HWPX 생성 결과를 확인할 수 없습니다."
    RunPowerShellCommand = Trim$(output)
    Exit Function
Failed:
    Dim number As Long, description As String
    number = Err.Number: description = Err.Description
    Err.Raise number, "NxPowerShell", description
End Function

Private Function ElapsedSeconds(ByVal started As Single) As Double
    If Timer >= started Then ElapsedSeconds = Timer - started Else ElapsedSeconds = (86400# - started) + Timer
End Function

Private Function SafeDiagnostic(ByVal value As String) As String
    value = Replace$(Replace$(value, vbCr, " "), vbLf, " ")
    If Len(value) > 240 Then value = Left$(value, 240)
    SafeDiagnostic = value
End Function


