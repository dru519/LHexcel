Option Explicit

Private mFeatureId As String
Private mPreview As Collection
Private mSealedPreview As Collection
Private mBoundWorkbook As Workbook
Private mPreviewReady As Boolean
Private mLoading As Boolean
Private mPaths As Collection
Private mRules As Collection
Private mOverrides As Object
Private mFileMode As Boolean

Private Sub UserForm_Initialize()
    cmdExecute.Default = True
    cmdCancel.Cancel = True
    Set mPaths = New Collection
    Set mRules = New Collection
    Set mOverrides = CreateObject("Scripting.Dictionary")
    mOverrides.CompareMode = vbTextCompare
End Sub

Private Sub UserForm_Activate()
    NxFormWheelAttach Me
End Sub

Public Sub HandleWheelDelta(ByVal wheelDelta As Long)
    NxFormWheelScrollActiveList Me, "lstSheets", wheelDelta
End Sub

Private Sub UserForm_Terminate()
    NxFormWheelDetach Me
End Sub

Public Sub BindFeature(ByVal featureId As String)
    Dim label As Variant, sheet As Worksheet
    mFeatureId = featureId
    Select Case featureId
        Case NX_FEATURE_FILE_BATCH_RENAME
            mFileMode = True
            Caption = "내엑셀 - 파일 이름 변경 작업대"
        Case NX_FEATURE_FILE_SHEET_BATCH_RENAME
            If ActiveWorkbook Is Nothing Then NxRaiseContractError "열린 통합문서가 없습니다."
            Set mBoundWorkbook = ActiveWorkbook
            Caption = "내엑셀 - 시트 이름 변경 작업대"
            For Each sheet In mBoundWorkbook.Worksheets
                mPaths.Add sheet.Name
            Next sheet
        Case Else
            NxRaiseContractError "지원하지 않는 이름 변경 기능입니다."
    End Select
    cmdAddFiles.Visible = mFileMode
    cmdAddFolder.Visible = mFileMode
    chkRecursive.Visible = mFileMode
    cmdRemoveItem.Visible = mFileMode
    cboRuleKind.Clear
    For Each label In Array("문자열 바꾸기", "앞에 문구 추가", "뒤에 문구 추가", "번호 붙이기", "일부 문자 지우기", "이름 전체 바꾸기", "영문 대소문자")
        cboRuleKind.AddItem CStr(label)
    Next label
    If mFileMode Then cboRuleKind.AddItem "확장자 변경"
    cboRuleKind.ListIndex = 0
    RefreshDraft
End Sub

Private Sub cmdAddFiles_Click()
    Dim picker As FileDialog, item As Variant, ignored As Boolean, detail As String
    On Error GoTo Failed
    Set picker = Application.FileDialog(3)
    picker.AllowMultiSelect = True
    picker.Title = "이름을 변경할 파일 선택"
    picker.Filters.Clear
    picker.Filters.Add "모든 파일", "*.*"
    If picker.Show <> -1 Then Exit Sub
    InvalidateSeal
    For Each item In picker.SelectedItems
        ignored = NxRenameAddFile(mPaths, CStr(item))
    Next item
    RefreshDraft
    Exit Sub
Failed:
    detail = Err.Description
    RefreshDraft
    ShowError detail
End Sub

Private Sub cmdAddFolder_Click()
    Dim picker As FileDialog, added As Long, detail As String
    On Error GoTo Failed
    Set picker = Application.FileDialog(4)
    picker.Title = "이름을 변경할 파일이 있는 폴더 선택"
    If picker.Show <> -1 Then Exit Sub
    InvalidateSeal
    added = NxRenameAddFolder(mPaths, CStr(picker.SelectedItems(1)), CBool(chkRecursive.Value))
    RefreshDraft
    Exit Sub
Failed:
    detail = Err.Description
    RefreshDraft
    ShowError detail
End Sub

Private Sub cmdRemoveItem_Click()
    Dim key As String
    If Not mFileMode Or lstSheets.ListIndex < 0 Then Exit Sub
    key = CStr(mPaths(lstSheets.ListIndex + 1))
    mPaths.Remove lstSheets.ListIndex + 1
    If mOverrides.Exists(key) Then mOverrides.Remove key
    RefreshDraft
End Sub

Private Sub cmdItemUp_Click(): MoveItem -1: End Sub
Private Sub cmdItemDown_Click(): MoveItem 1: End Sub

Private Sub MoveItem(ByVal offset As Long)
    Dim selected As Long
    selected = lstSheets.ListIndex
    If selected < 0 Or selected + offset < 0 Or selected + offset >= mPaths.Count Then Exit Sub
    NxRenameMoveItem mPaths, selected + 1, offset
    RefreshDraft
    lstSheets.ListIndex = selected + offset
End Sub

Private Sub cmdAddRule_Click()
    Dim rule As New CNxRenameRule, kinds As Variant
    kinds = Array("REPLACE", "PREFIX", "SUFFIX", "NUMBER", "DELETE", "FULL", "CASE", "EXTENSION")
    If cboRuleKind.ListIndex < 0 Then Exit Sub
    rule.Kind = CStr(kinds(cboRuleKind.ListIndex))
    If rule.Kind = "NUMBER" Then rule.Text = "_"
    EditRule rule, 0
End Sub

Private Sub cmdEditRule_Click()
    If lstRules.ListIndex < 0 Then Exit Sub
    EditRule mRules(lstRules.ListIndex + 1), lstRules.ListIndex + 1
End Sub

Private Sub lstRules_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdEditRule_Click
End Sub

Private Sub EditRule(ByVal source As CNxRenameRule, ByVal index As Long)
    Dim editor As New FNxRenameRuleEditor, result As CNxRenameRule, detail As String
    On Error GoTo Failed
    Set result = editor.EditRule(source, mFileMode)
    Unload editor
    If result Is Nothing Then Exit Sub
    If index = 0 Then
        mRules.Add result
    Else
        mRules.Add result, , index
        mRules.Remove index + 1
    End If
    RefreshRules
    RefreshDraft
    Exit Sub
Failed:
    detail = Err.Description
    Unload editor
    ShowError detail
End Sub

Private Sub cmdRuleUp_Click(): MoveRule -1: End Sub
Private Sub cmdRuleDown_Click(): MoveRule 1: End Sub

Private Sub MoveRule(ByVal offset As Long)
    Dim selected As Long
    selected = lstRules.ListIndex
    If selected < 0 Or selected + offset < 0 Or selected + offset >= mRules.Count Then Exit Sub
    NxRenameMoveItem mRules, selected + 1, offset
    RefreshRules
    lstRules.ListIndex = selected + offset
    RefreshDraft
End Sub

Private Sub cmdRemoveRule_Click()
    If lstRules.ListIndex < 0 Then Exit Sub
    mRules.Remove lstRules.ListIndex + 1
    RefreshRules
    RefreshDraft
End Sub

Private Sub RefreshRules()
    Dim rule As CNxRenameRule
    lstRules.Clear
    For Each rule In mRules
        lstRules.AddItem rule.Summary
    Next rule
End Sub

Private Sub RefreshDraft()
    Dim hasDuplicate As Boolean, hasCollision As Boolean, detail As String
    On Error GoTo Failed
    InvalidateSeal
    If mPaths.Count = 0 Then
        Set mPreview = Nothing
        mLoading = True
        lstSheets.Clear
        txtPlannedName.Value = vbNullString
        txtPlannedName.Enabled = False
        mLoading = False
        Exit Sub
    End If
    If mFileMode Then
        Set mPreview = NxRenameFileRulesPreview(mPaths, mRules, mOverrides, hasDuplicate, hasCollision)
    Else
        RequireBoundWorkbook
        Set mPreview = NxRenameSheetRulesPreview(mBoundWorkbook, mPaths, mRules, mOverrides, hasDuplicate, hasCollision)
    End If
    RenderSheetPreview
    lblStatus.Caption = CStr(mPaths.Count) & "개 / 규칙 " & CStr(mRules.Count) & "개 · 미리보기로 확인하세요. (되돌리기는 이 Excel 세션의 직전 작업)"
    Exit Sub
Failed:
    detail = Err.Description
    Set mPreview = Nothing
    mLoading = True
    lstSheets.Clear
    txtPlannedName.Enabled = False
    mLoading = False
    lblStatus.Caption = NxUserErrorText(detail)
End Sub

Private Sub cmdPreview_Click()
    Dim hasDuplicate As Boolean, hasCollision As Boolean, detail As String
    On Error GoTo Failed
    InvalidateSeal
    If mPreview Is Nothing Then NxRaiseContractError "이름 변경 대상을 추가하거나 규칙을 확인하세요."
    If mFileMode Then
        Set mPreview = NxBatchRenameRevalidatePreview(mPreview, hasDuplicate, hasCollision)
    Else
        RequireBoundWorkbook
        Set mPreview = NxSheetBatchRenameRevalidatePreview(mBoundWorkbook, mPreview, hasDuplicate, hasCollision)
    End If
    RenderSheetPreview
    If mFileMode Then
        NxBatchRenameRequireReady mPreview
    Else
        NxSheetBatchRenameRequireReady mBoundWorkbook, mPreview
    End If
    Set mSealedPreview = mPreview
    mPreviewReady = True
    cmdExecute.Enabled = True
    lblStatus.Caption = "미리보기 완료 · " & CStr(mPreview.Count) & "개 확인"
    Exit Sub
Failed:
    detail = Err.Description
    InvalidateSeal
    ShowError detail
End Sub

Private Sub lstSheets_Change()
    If mLoading Then Exit Sub
    mLoading = True
    txtPlannedName.Enabled = (lstSheets.ListIndex >= 0)
    If lstSheets.ListIndex >= 0 Then
        txtPlannedName.Value = CStr(lstSheets.List(lstSheets.ListIndex, 1))
        lblPathHelp.Caption = CStr(mPaths(lstSheets.ListIndex + 1))
    End If
    mLoading = False
End Sub

Private Sub txtPlannedName_Change()
    Dim hasDuplicate As Boolean, hasCollision As Boolean, row As Variant, edited As New Collection
    Dim index As Long, key As String, detail As String
    If mLoading Or lstSheets.ListIndex < 0 Then Exit Sub
    If mPreview Is Nothing Then Exit Sub
    On Error GoTo Failed
    InvalidateSeal
    key = CStr(mPaths(lstSheets.ListIndex + 1))
    mOverrides.Item(key) = CStr(txtPlannedName.Value)
    If mFileMode Then
        For Each row In mPreview
            index = index + 1
            If index = lstSheets.ListIndex + 1 Then row(1) = Left$(CStr(row(0)), InStrRev(CStr(row(0)), "\")) & CStr(txtPlannedName.Value)
            edited.Add row
        Next row
        Set mPreview = NxBatchRenameRevalidatePreview(edited, hasDuplicate, hasCollision)
    Else
        RequireBoundWorkbook
        Set mPreview = NxSheetBatchRenameEditPreview(mBoundWorkbook, mPreview, lstSheets.ListIndex + 1, CStr(txtPlannedName.Value), hasDuplicate, hasCollision)
    End If
    RenderSheetPreview False
    Exit Sub
Failed:
    detail = Err.Description
    InvalidateSeal
    ShowError detail
End Sub

Private Sub cmdResetName_Click()
    Dim key As String
    If lstSheets.ListIndex < 0 Then Exit Sub
    key = CStr(mPaths(lstSheets.ListIndex + 1))
    If mOverrides.Exists(key) Then mOverrides.Remove key
    RefreshDraft
End Sub

Private Sub RenderSheetPreview(Optional ByVal loadEditor As Boolean = True)
    Dim row As Variant, selectedIndex As Long, index As Long
    selectedIndex = lstSheets.ListIndex
    mLoading = True
    lstSheets.Clear
    For Each row In mPreview
        If mFileMode Then lstSheets.AddItem NxRenameLeafName(CStr(row(0))) Else lstSheets.AddItem CStr(row(0))
        If mFileMode Then lstSheets.List(index, 1) = NxRenameLeafName(CStr(row(1))) Else lstSheets.List(index, 1) = CStr(row(1))
        lstSheets.List(index, 2) = CStr(row(2))
        index = index + 1
    Next row
    If selectedIndex < 0 Or selectedIndex >= lstSheets.ListCount Then selectedIndex = 0
    If lstSheets.ListCount > 0 Then
        lstSheets.ListIndex = selectedIndex
        If loadEditor Then txtPlannedName.Value = CStr(lstSheets.List(selectedIndex, 1))
        txtPlannedName.Enabled = True
        lblPathHelp.Caption = CStr(mPaths(selectedIndex + 1))
    End If
    mLoading = False
End Sub

Private Sub RequireBoundWorkbook()
    If mBoundWorkbook Is Nothing Then NxRaiseContractError "기능을 다시 실행하세요."
    If ActiveWorkbook Is Nothing Then NxRaiseContractError "열린 통합문서가 없습니다."
    If Not ActiveWorkbook Is mBoundWorkbook Then NxRaiseContractError "처음 선택한 통합문서로 돌아오거나 기능을 다시 실행하세요."
End Sub

Private Sub cmdExecute_Click()
    Dim options As Object, result As CNxResult, target As String, detail As String
    On Error GoTo Failed
    cmdPreview_Click
    If Not mPreviewReady Then Exit Sub
    Set options = CreateObject("Scripting.Dictionary")
    Set options.Item("preview") = mSealedPreview
    If mFileMode Then
        NxBatchRenameRequireReady mSealedPreview
        target = NxBatchRenameRootTarget(mSealedPreview)
    Else
        RequireBoundWorkbook
        NxSheetBatchRenameRequireReady mBoundWorkbook, mSealedPreview
        target = mBoundWorkbook.FullName
    End If
    Set result = NxFileRun(mFeatureId, target, workflowOptions:=options)
    If result Is Nothing Then NxRaiseContractError "이름 변경을 완료하지 못했습니다."
    If result.Outcome <> NxSuccess Then NxRaiseContractError "이름 변경을 완료하지 못했습니다."
    ResetAfterOperation
    lblStatus.Caption = "이름 변경 완료 · 직전 작업을 되돌릴 수 있습니다."
    Exit Sub
Failed:
    detail = Err.Description
    InvalidateSeal
    ShowError detail
End Sub

Private Sub cmdUndo_Click()
    Dim count As Long, detail As String, options As Object, result As CNxResult, target As String
    On Error GoTo Failed
    If MsgBox("이 Excel 세션의 직전 이름 변경을 원래 이름으로 되돌리시겠습니까?", vbQuestion + vbYesNo + vbDefaultButton2, "내엑셀") <> vbYes Then Exit Sub
    If Not mFileMode Then RequireBoundWorkbook
    count = NxBatchRenameUndoLast(mFileMode, mBoundWorkbook, True)
    target = NxBatchRenameUndoTarget(mFileMode, mBoundWorkbook)
    Set options = CreateObject("Scripting.Dictionary")
    options.Add "rename_undo", True
    Set result = NxFileRun(mFeatureId, target, selectedTarget:=mBoundWorkbook, workflowOptions:=options)
    If result Is Nothing Then NxRaiseContractError "이름 복구 결과를 확인할 수 없습니다."
    If result.Outcome = NxCancelled Then Exit Sub
    If result.Outcome <> NxSuccess Then NxRaiseContractError "이름 복구를 완료하지 못했습니다."
    ResetAfterOperation
    lblStatus.Caption = CStr(count) & "개 이름을 복구했습니다."
    Exit Sub
Failed:
    detail = Err.Description
    InvalidateSeal
    ShowError detail
End Sub

Private Sub ResetAfterOperation()
    Dim sheet As Worksheet
    Set mPaths = New Collection
    Set mRules = New Collection
    mOverrides.RemoveAll
    If Not mFileMode Then
        For Each sheet In mBoundWorkbook.Worksheets
            mPaths.Add sheet.Name
        Next sheet
    End If
    RefreshRules
    RefreshDraft
End Sub

Private Sub InvalidateSeal()
    mPreviewReady = False
    Set mSealedPreview = Nothing
    cmdExecute.Enabled = True
    cmdUndo.Enabled = NxBatchRenameCanUndo(mFileMode, mBoundWorkbook)
    lblStatus.Caption = "옵션을 확인하고 실행하세요."
End Sub

Private Sub cmdCancel_Click(): RequestCancel: End Sub

Private Sub RequestCancel()
    NxFormWheelDetach Me
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    NxFormWheelDetach Me
End Sub

Private Sub ShowError(ByVal detail As String)
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 이름 변경 작업대"
End Sub
