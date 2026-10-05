Option Explicit

Private mContext As CNxFeatureDialogContext
Private mPreview As CNxFolderCreatePreview
Private mItems As Collection
Private mRangeSession As CNxRangeSelectionSession
Private mInputIdentity As String
Private mPreviewReady As Boolean
Private mAllowClose As Boolean

Public Sub BindContext(ByVal context As CNxFeatureDialogContext)
    If context Is Nothing Then NxRaiseContractError "폴더 생성 컨텍스트가 필요합니다."
    Set mContext = context
End Sub

Private Sub UserForm_Initialize()
    cmdCreate.Default = True
    cmdCancel.Cancel = True
    cmdCreate.Enabled = True
    If TypeName(Application.Selection) = "Range" Then
        Set mRangeSession = NxRangeSessionFromSelection()
        txtRange.Value = mRangeSession.DisplayAddress
    End If
End Sub

Private Sub txtBaseFolder_Change(): InvalidatePreview: End Sub
Private Sub txtRange_Change(): InvalidatePreview: End Sub
Private Sub txtPrefix_Change(): InvalidatePreview: End Sub
Private Sub txtSuffix_Change(): InvalidatePreview: End Sub

Private Sub cmdBrowse_Click()
    Dim picker As FileDialog
    Set picker = Application.FileDialog(4)
    With picker
        .AllowMultiSelect = False
        .Title = "내엑셀 - 기준 폴더 선택"
        If .Show <> -1 Then Exit Sub
        txtBaseFolder.Value = CStr(.SelectedItems(1))
    End With
End Sub

Private Sub cmdPickRange_Click()
    On Error GoTo Failed
    If mRangeSession Is Nothing Then Set mRangeSession = New CNxRangeSelectionSession
    If NxPickRange(Me, mRangeSession) Then txtRange.Value = mRangeSession.DisplayAddress
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub InvalidatePreview()
    Set mPreview = Nothing
    mPreviewReady = False
    cmdCreate.Enabled = True
    txtPreview.Value = "폴더 이름과 생성 위치를 확인하고 실행하세요."
    lblBatchLimit.Caption = "최대 500행 · 생성 폴더 최대 500개"
End Sub

Private Function CurrentItems() As Collection
    Dim source As Range
    If mRangeSession Is Nothing Then NxRaiseContractError "폴더 이름이 있는 엑셀 범위를 선택하세요."
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set source = mRangeSession.ResolveRange()
    Set CurrentItems = NxFolderCreateItemsFromRange(source, CStr(txtPrefix.Value), CStr(txtSuffix.Value))
End Function

Private Sub cmdPreview_Click()
    Dim path As Variant
    On Error GoTo Failed
    Set mItems = CurrentItems()
    Set mPreview = New CNxFolderCreatePreview
    mPreview.Snapshot CStr(txtBaseFolder.Value), mItems
    mInputIdentity = NxFolderCreateItemsText(mItems)
    txtPreview.Value = vbNullString
    For Each path In mPreview.DirectoryPaths
        txtPreview.Value = CStr(txtPreview.Value) & CStr(path) & vbCrLf
    Next path
    lblBatchLimit.Caption = "중복 제외 입력 " & CStr(mItems.Count) & "행 · 생성 예정 " & CStr(mPreview.PlannedCount) & "/500개"
    mPreviewReady = True
    cmdCreate.Enabled = True
    Exit Sub
Failed:
    ShowError
    InvalidatePreview
End Sub

Private Sub cmdCreate_Click()
    Dim result As CNxResult, current As Collection
    On Error GoTo Failed
    cmdPreview_Click
    If Not mPreviewReady Then Exit Sub
    Set current = CurrentItems()
    If NxFolderCreateItemsText(current) <> mInputIdentity Then NxRaiseContractError "미리보기 이후 엑셀 목록이 변경되었습니다. 미리보기를 다시 확인하세요."
    Set result = NxFolderCreateRun(CStr(txtBaseFolder.Value), mItems, mPreview)
    If result Is Nothing Then Exit Sub
    If result.Outcome <> NxSuccess Then Exit Sub
    RequestCancel
    Exit Sub
Failed:
    ShowError
    InvalidatePreview
End Sub

Private Sub cmdCancel_Click(): RequestCancel: End Sub

Private Sub RequestCancel()
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description: Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation, "내엑셀 - 폴더 생성"
End Sub

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
