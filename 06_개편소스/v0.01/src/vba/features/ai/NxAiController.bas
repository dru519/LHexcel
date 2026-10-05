Attribute VB_Name = "NxAiController"
Option Explicit

Public Function NxAiCopyPrompt(ByVal prompt As String, Optional ByVal featureId As String = "NX-AI-SUMMARY") As CNxResult
    ValidateAiFeature featureId
    On Error GoTo ClipboardFailed
    NxClipboardWriteUnicode prompt
    Set NxAiCopyPrompt = CreateAdapterResult(featureId, NxSuccess, "complete", "clipboard", vbNullString, "result_success", True)
    Exit Function

ClipboardFailed:
    Dim detail As String
    detail = SafeAdapterError(Err.Number, Err.Description)
    Err.Clear
    Set NxAiCopyPrompt = CreateAdapterResult(featureId, NxEnvironmentError, "clipboard", "clipboard", detail, "result_environment_error", False)
End Function

Private Function CreateAdapterResult( _
    ByVal featureId As String, _
    ByVal outcome As NxOutcome, _
    ByVal stage As String, _
    ByVal target As String, _
    ByVal recovery As String, _
    ByVal messageKey As String, _
    ByVal clipboardCompleted As Boolean) As CNxResult

    Dim result As New CNxResult
    result.Configure featureId, outcome, stage, target, False, recovery, messageKey
    If clipboardCompleted Then result.AddCompleted "clipboard"
    result.Seal
    Set CreateAdapterResult = result
End Function

Private Sub ValidateAiFeature(ByVal featureId As String)
    Select Case featureId
        Case "NX-AI-SUMMARY", "NX-AI-CLEAN", "NX-AI-FORMULA", "NX-AI-WRITE", "NX-AI-IMAGE"
        Case Else
            Err.Raise vbObjectError + 888, "NxAiController", "지원하지 않는 AI 작업 모드입니다."
    End Select
End Sub

Private Function SafeAdapterError(ByVal errorNumber As Long, ByVal description As String) As String
    If Len(Trim$(description)) = 0 Then description = "어댑터 실행 실패"
    SafeAdapterError = CStr(errorNumber) & " " & description
End Function
