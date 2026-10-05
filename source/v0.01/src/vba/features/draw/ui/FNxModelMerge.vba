Option Explicit

Private mInit As CNxFormInitializationGuard
Private mRangeSession As CNxRangeSelectionSession
Private mAllowClose As Boolean

Private Sub UserForm_Initialize()
    Set mInit = New CNxFormInitializationGuard
    mInit.BeginInitialization
    NxInitializeStaticComboChoices Me
    cboAlignment.ListIndex = 1
    chkFillAdjacent.Value = False
    cmdApply.Default = True
    cmdClose.Cancel = True
    cmdRestore.Enabled = False
    mAllowClose = False
    mInit.CompleteInitialization
End Sub

Public Sub BindSelection()
    mInit.RequireReady
    Set mRangeSession = NxRangeSessionFromSelection()
    txtRange.Value = mRangeSession.DisplayAddress
    RefreshRestoreButton
End Sub

Private Sub cmdPickRange_Click()
    On Error GoTo Failed
    RequireRangeSession
    If NxPickRange(Me, mRangeSession) Then
        txtRange.Value = mRangeSession.DisplayAddress
        RefreshRestoreButton
    End If
    Exit Sub
Failed:
    ShowError NxUserErrorText(Err.Description)
End Sub

Private Sub cmdApply_Click()
    Dim target As Range
    On Error GoTo Failed
    RequireRangeSession
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set target = mRangeSession.ResolveRange()
    NxModelMergeApply target, SelectedAlignmentCode(), CBool(chkFillAdjacent.Value)
    cmdRestore.Enabled = True
    Exit Sub
Failed:
    ShowError NxUserErrorText(Err.Description)
End Sub

Private Sub cmdRestore_Click()
    Dim target As Range
    On Error GoTo Failed
    RequireRangeSession
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set target = mRangeSession.ResolveRange()
    NxModelMergeRestore target
    cmdRestore.Enabled = False
    Exit Sub
Failed:
    ShowError NxUserErrorText(Err.Description)
End Sub

Private Sub txtRange_Change()
    If mInit Is Nothing Then Exit Sub
    If mInit.IsInitializing Then Exit Sub
    RefreshRestoreButton
End Sub

Private Sub cmdClose_Click()
    RequestClose
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    mAllowClose = True
End Sub

Private Sub UserForm_Terminate()
    NxModelMergeFormClosed Me
End Sub

Private Sub RequestClose()
    mAllowClose = True
    Unload Me
End Sub

Private Sub RefreshRestoreButton()
    Dim target As Range
    On Error GoTo NotAvailable
    If mRangeSession Is Nothing Then GoTo NotAvailable
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set target = mRangeSession.ResolveRange()
    cmdRestore.Enabled = NxModelMergeCanRestore(target)
    Exit Sub
NotAvailable:
    Err.Clear
    cmdRestore.Enabled = False
End Sub

Private Sub RequireRangeSession()
    mInit.RequireReady
    If mRangeSession Is Nothing Then NxRaiseContractError "적용할 셀 범위를 선택하세요."
End Sub

Private Function SelectedAlignmentCode() As String
    Select Case cboAlignment.ListIndex
        Case 0: SelectedAlignmentCode = "LEFT"
        Case 1: SelectedAlignmentCode = "CENTER"
        Case 2: SelectedAlignmentCode = "RIGHT"
        Case Else: NxRaiseContractError "정렬 방법을 선택하세요."
    End Select
End Function

Private Sub ShowError(ByVal detail As String)
    Beep
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 모형병합"
End Sub
