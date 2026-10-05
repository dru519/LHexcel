Attribute VB_Name = "NxFileOperationState"
Option Explicit

Private mActive As Boolean
Private mCancelled As Boolean
Private mCurrent As Long
Private mTotal As Long
Private mOperationLabel As String
Private mPreviousStatusBar As Variant
Private mPreviousCancelKey As Long

Public Sub NxFileOperationBegin(ByVal totalItems As Long, ByVal operationLabel As String)
    If mActive Then NxRaiseContractError "A file operation is already active"
    If totalItems < 1 Then NxRaiseContractError "File operation item count must be positive"
    If Len(Trim$(operationLabel)) = 0 Then NxRaiseContractError "File operation label is required"

    mPreviousStatusBar = Application.StatusBar
    mPreviousCancelKey = Application.EnableCancelKey
    mTotal = totalItems
    mCurrent = 0
    mOperationLabel = Trim$(operationLabel)
    mCancelled = False
    mActive = True
    Application.EnableCancelKey = xlErrorHandler
    Application.StatusBar = NxFileOperationStatusText("준비 중")
End Sub

Public Sub NxFileOperationReport(ByVal currentItem As Long, ByVal detail As String)
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    If Not mActive Then NxRaiseContractError "File operation progress requires an active operation"
    If currentItem < 0 Or currentItem > mTotal Then NxRaiseContractError "File operation progress is outside the declared range"
    If mCancelled Then Exit Sub

    mCurrent = currentItem
    On Error GoTo CancelledByUser
    Application.StatusBar = NxFileOperationStatusText(detail)
    DoEvents
    Exit Sub
CancelledByUser:
    If Err.Number = 18 Then
        Err.Clear
        mCancelled = True
        On Error Resume Next
        Application.StatusBar = mOperationLabel & " · 취소 요청을 처리하고 있습니다."
        On Error GoTo 0
        Exit Sub
    End If
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Public Sub NxFileOperationRequestCancel()
    If Not mActive Then Exit Sub
    mCancelled = True
    On Error Resume Next
    Application.StatusBar = mOperationLabel & " · 취소 요청을 처리하고 있습니다."
    On Error GoTo 0
End Sub

Public Function NxFileOperationCancelled() As Boolean
    NxFileOperationCancelled = mActive And mCancelled
End Function

Public Sub NxFileOperationFinish()
    If Not mActive Then Exit Sub
    On Error Resume Next
    If IsEmpty(mPreviousStatusBar) Then
        Application.StatusBar = False
    ElseIf VarType(mPreviousStatusBar) = vbBoolean Then
        Application.StatusBar = False
    Else
        Application.StatusBar = mPreviousStatusBar
    End If
    Application.EnableCancelKey = mPreviousCancelKey
    On Error GoTo 0

    mActive = False
    mCancelled = False
    mCurrent = 0
    mTotal = 0
    mOperationLabel = vbNullString
    mPreviousStatusBar = Empty
    mPreviousCancelKey = 0
End Sub

Private Function NxFileOperationStatusText(ByVal detail As String) As String
    ' Excel rejects long StatusBar strings. Only the display is shortened;
    ' source paths and operation/recovery data always retain their full value.
    NxFileOperationStatusText = Left$(mOperationLabel & " · " & Format$(mCurrent, "0") & "/" & _
        Format$(mTotal, "0") & " · " & detail & " · Esc 키로 취소", 250)
End Function
