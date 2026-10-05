VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFrameTextCatalog"
Option Explicit

Private mSuccess As String, mCancelled As String, mInputError As String
Private mEnvironmentError As String, mPartialFailure As String
Private mSummaryNone As String, mIncompleteAttention As String
Private mSourceChangedYes As String, mSourceChangedNo As String
Private mRecoveryNone As String, mNextAction As String
Private mConfigured As Boolean

Public Sub Configure(ByVal successLabel As String, ByVal cancelledLabel As String, _
    ByVal inputErrorLabel As String, ByVal environmentErrorLabel As String, _
    ByVal partialFailureLabel As String, ByVal summaryNone As String, _
    ByVal incompleteAttention As String, ByVal sourceChangedYes As String, _
    ByVal sourceChangedNo As String, ByVal recoveryNone As String, _
    ByVal nextAction As String)
    If mConfigured Then NxRaiseContractError "Frame text catalog can be configured only once"
    If Len(Trim$(successLabel)) = 0 Or Len(Trim$(cancelledLabel)) = 0 Or _
        Len(Trim$(inputErrorLabel)) = 0 Or Len(Trim$(environmentErrorLabel)) = 0 Or _
        Len(Trim$(partialFailureLabel)) = 0 Or Len(Trim$(summaryNone)) = 0 Or _
        Len(Trim$(incompleteAttention)) = 0 Or Len(Trim$(sourceChangedYes)) = 0 Or _
        Len(Trim$(sourceChangedNo)) = 0 Or Len(Trim$(recoveryNone)) = 0 Or _
        Len(Trim$(nextAction)) = 0 Then
        NxRaiseContractError "Frame text catalog requires every display value"
    End If
    mSuccess = successLabel
    mCancelled = cancelledLabel
    mInputError = inputErrorLabel
    mEnvironmentError = environmentErrorLabel
    mPartialFailure = partialFailureLabel
    mSummaryNone = summaryNone
    mIncompleteAttention = incompleteAttention
    mSourceChangedYes = sourceChangedYes
    mSourceChangedNo = sourceChangedNo
    mRecoveryNone = recoveryNone
    mNextAction = nextAction
    mConfigured = True
End Sub

Public Function BuildResultModel(ByVal result As CNxResult, ByVal resultPath As String, _
    Optional ByVal targetSummary As String = vbNullString) As CNxResultCardViewModel
    Dim model As New CNxResultCardViewModel
    Dim completedSummary As String, incompleteSummary As String, recoverySummary As String
    EnsureConfigured
    If result Is Nothing Then NxRaiseContractError "Result text conversion requires a result"
    completedSummary = JoinCompleted(result.Completed)
    If Len(completedSummary) = 0 Then completedSummary = mSummaryNone
    If result.Outcome = NxSuccess Then
        incompleteSummary = mSummaryNone
    Else
        incompleteSummary = mIncompleteAttention
    End If
    recoverySummary = result.Recovery
    If Len(Trim$(recoverySummary)) = 0 Then recoverySummary = mRecoveryNone
    model.Configure result, OutcomeLabelFor(result.Outcome), completedSummary, incompleteSummary, _
        IIf(result.SourceChanged, mSourceChangedYes, mSourceChangedNo), recoverySummary, resultPath, mNextAction, targetSummary
    model.Seal
    Set BuildResultModel = model
End Function

Private Function OutcomeLabelFor(ByVal outcome As NxOutcome) As String
    Select Case outcome
        Case NxSuccess: OutcomeLabelFor = mSuccess
        Case NxCancelled: OutcomeLabelFor = mCancelled
        Case NxInputError: OutcomeLabelFor = mInputError
        Case NxEnvironmentError: OutcomeLabelFor = mEnvironmentError
        Case NxPartialFailure: OutcomeLabelFor = mPartialFailure
        Case Else: NxRaiseContractError "Result outcome is outside the frame text contract"
    End Select
End Function

Private Function JoinCompleted(ByVal items As Collection) As String
    Dim item As Variant, displayItem As String
    For Each item In items
        If Len(JoinCompleted) > 0 Then JoinCompleted = JoinCompleted & vbCrLf
        displayItem = CompletedDisplayItem(CStr(item))
        JoinCompleted = JoinCompleted & displayItem
    Next item
End Function

Private Function CompletedDisplayItem(ByVal item As String) As String
    Dim kind As String, separator As Long
    separator = InStr(1, item, ":", vbBinaryCompare)
    kind = item
    If separator > 0 Then kind = Left$(item, separator - 1)
    Select Case kind
        Case "clipboard", "clipboard_write": CompletedDisplayItem = "클립보드 복사 완료"
        Case "cell_mutation": CompletedDisplayItem = "셀 변경 완료"
        Case "file_create", "file_replace", "binary_package_write": CompletedDisplayItem = "파일 저장 완료"
        Case "shape_insert": CompletedDisplayItem = "그림 삽입 완료"
        Case "shape_mutation": CompletedDisplayItem = "그림 변경 완료"
        Case "worksheet_create": CompletedDisplayItem = "시트 생성 완료"
        Case "directory_batch_create": CompletedDisplayItem = "폴더 생성 완료"
        Case "external_process": CompletedDisplayItem = "외부 프로그램 실행 완료"
        Case "browser_open": CompletedDisplayItem = "브라우저 열기 완료"
        Case Else
            ' Preserve raw receipts in CNxResult, not in the user-facing summary.
            If separator > 0 Or InStr(1, item, "|", vbBinaryCompare) > 0 Then
                CompletedDisplayItem = "완료된 작업"
            Else
                CompletedDisplayItem = item
            End If
    End Select
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Frame text catalog is not configured"
End Sub
