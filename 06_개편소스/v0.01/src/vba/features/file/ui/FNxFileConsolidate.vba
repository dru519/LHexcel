Option Explicit

Private mPreviewReady As Boolean
Private mAllowClose As Boolean
Private mExecuting As Boolean

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
    lstInputFiles.MultiSelect = fmMultiSelectExtended
    cmdExecute.Default = True
    cmdCancel.Cancel = True
End Sub

Private Sub UserForm_Activate()
    NxFormWheelAttach Me
End Sub

Public Sub HandleWheelDelta(ByVal wheelDelta As Long)
    NxFormWheelScrollActiveList Me, "lstInputFiles", wheelDelta
End Sub

Public Sub BindFeature()
    InvalidatePreview
End Sub

Private Sub cmdAddFiles_Click()
    Dim picker As FileDialog
    Dim item As Variant
    Dim index As Long
    Dim exists As Boolean
    Set picker = Application.FileDialog(3)
    With picker
        .AllowMultiSelect = True
        .Title = "통합할 파일 선택"
        .Filters.Clear
        .Filters.Add "Excel 통합문서", "*.xlsx;*.xlsm"
        If .Show <> -1 Then Exit Sub
        For Each item In .SelectedItems
            exists = False
            For index = 0 To lstInputFiles.ListCount - 1
                If StrComp(CStr(lstInputFiles.List(index)), CStr(item), vbTextCompare) = 0 Then exists = True: Exit For
            Next index
            If Not exists Then lstInputFiles.AddItem CStr(item)
        Next item
    End With
    InvalidatePreview
End Sub

Private Sub cmdRemoveFile_Click()
    Dim index As Long
    If mExecuting Then Exit Sub
    For index = lstInputFiles.ListCount - 1 To 0 Step -1
        If lstInputFiles.Selected(index) Then lstInputFiles.RemoveItem index
    Next index
    InvalidatePreview
End Sub

Private Sub cmdSelectAll_Click()
    Dim index As Long
    If mExecuting Then Exit Sub
    For index = 0 To lstInputFiles.ListCount - 1
        lstInputFiles.Selected(index) = True
    Next index
End Sub

Private Sub lstInputFiles_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyA And (Shift And 2) <> 0 Then
        cmdSelectAll_Click
        KeyCode = 0
    End If
End Sub

Private Sub cmdMoveUp_Click(): MoveSelectedFiles -1: End Sub
Private Sub cmdMoveDown_Click(): MoveSelectedFiles 1: End Sub

Private Sub MoveSelectedFiles(ByVal direction As Long)
    Dim paths() As String, chosen() As Boolean, index As Long, other As Long
    Dim first As Long, last As Long, temp As String
    If mExecuting Or lstInputFiles.ListCount < 2 Then Exit Sub
    ReDim paths(0 To lstInputFiles.ListCount - 1)
    ReDim chosen(0 To lstInputFiles.ListCount - 1)
    For index = 0 To lstInputFiles.ListCount - 1
        paths(index) = CStr(lstInputFiles.List(index))
        chosen(index) = lstInputFiles.Selected(index)
    Next index
    If direction < 0 Then
        first = 1: last = UBound(paths)
    Else
        first = UBound(paths) - 1: last = 0
    End If
    For index = first To last Step -direction
        other = index + direction
        If chosen(index) And Not chosen(other) Then
            temp = paths(other): paths(other) = paths(index): paths(index) = temp
            chosen(other) = True: chosen(index) = False
        End If
    Next index
    For index = 0 To UBound(paths)
        lstInputFiles.List(index) = paths(index)
        lstInputFiles.Selected(index) = chosen(index)
    Next index
    InvalidatePreview
End Sub

Private Sub cmdBrowse_Click()
    Dim selectedPath As Variant
    selectedPath = Application.GetSaveAsFilename(InitialFileName:="내엑셀_통합.xlsx", _
        FileFilter:="Excel 통합문서 (*.xlsx), *.xlsx", FilterIndex:=1, Title:="저장 폴더와 파일 이름 선택")
    If VarType(selectedPath) <> vbBoolean Then txtOutputPath.Value = CStr(selectedPath)
End Sub

Private Sub cmdPreview_Click()
    Dim validInputs As Collection
    Dim excludedSummary As String
    On Error GoTo Failed
    If lstInputFiles.ListCount = 0 Then NxRaiseContractError "통합할 파일을 하나 이상 추가하세요."
    NxConsolidationValidateOutput CStr(txtOutputPath.Value)
    Set validInputs = NxConsolidationPlan(InputPaths(), chkAllowPartial.Value, excludedSummary)
    txtPreview.Value = CStr(cboMode.Value) & " · 유효 입력 " & CStr(validInputs.Count) & "개"
    If Len(excludedSummary) > 0 Then txtPreview.Value = txtPreview.Value & vbCrLf & "제외 항목" & vbCrLf & excludedSummary
    txtPreview.Value = txtPreview.Value & vbCrLf & "원본 파일은 읽기 전용으로 처리합니다."
    mPreviewReady = True
    cmdExecute.Enabled = True
    lblStep.Caption = "1. 파일 선택  2. 미리보기 완료  3. 통합"
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdExecute_Click()
    Dim result As CNxResult
    Dim secondApproval As Boolean
    Dim options As Object
    On Error GoTo Failed
    If mExecuting Then Exit Sub
    If lstInputFiles.ListCount = 0 Then NxRaiseContractError "통합할 파일을 하나 이상 추가하세요."
    NxConsolidationValidateOutput CStr(txtOutputPath.Value)
    Set options = CreateObject("Scripting.Dictionary")
    options.Add "consolidation_mode", Choose(cboMode.ListIndex + 1, "SHEETS", "ROWS", "BOTH")
    options.Add "skip_headers", CBool(chkSkipHeaders.Value)
    options.Add "create_contents", CBool(chkContents.Value) And cboMode.ListIndex <> 1
    BeginExecution
    Set result = NxFileRun("NX-FILE-CONSOLIDATE", CStr(txtOutputPath.Value), InputPaths(), chkAllowPartial.Value, workflowOptions:=options)
    If result Is Nothing Then NxRaiseContractError "파일 통합 결과를 확인할 수 없습니다."
    If result.Outcome = NxCancelled And chkAllowPartial.Value Then
        If MsgBox("실패 파일을 제외한 부분결과를 만들까요?", vbYesNo + vbQuestion + vbDefaultButton2, "내엑셀") = vbYes Then
            secondApproval = True
            Set result = NxFileRun("NX-FILE-CONSOLIDATE", CStr(txtOutputPath.Value), InputPaths(), True, secondApproval:=secondApproval, workflowOptions:=options)
        End If
    End If
    EndExecution
    If result Is Nothing Then NxRaiseContractError "파일 통합 결과를 확인할 수 없습니다."
    If result.Outcome <> NxSuccess And result.Outcome <> NxPartialFailure Then NxRaiseContractError "파일을 통합하지 못했습니다."
    RequestCancel
    Exit Sub
Failed:
    EndExecution
    ShowError
End Sub

Private Function InputPaths() As Variant
    Dim result() As String
    Dim index As Long
    ReDim result(0 To lstInputFiles.ListCount - 1)
    For index = 0 To lstInputFiles.ListCount - 1
        result(index) = CStr(lstInputFiles.List(index))
    Next index
    InputPaths = result
End Function

Private Sub BeginExecution()
    mExecuting = True
    SetInputsEnabled False
    cmdExecute.Enabled = False
    cmdPreview.Enabled = False
    cmdCancel.Caption = "중단 요청"
    lblStep.Caption = "파일을 통합하는 중입니다."
End Sub

Private Sub EndExecution()
    If Not mExecuting Then Exit Sub
    mExecuting = False
    SetInputsEnabled True
    cmdPreview.Enabled = True
    cmdExecute.Enabled = True
    cmdCancel.Caption = "취소"
End Sub

Private Sub txtOutputPath_Change(): InvalidatePreview: End Sub
Private Sub chkAllowPartial_Click(): InvalidatePreview: End Sub
Private Sub chkSkipHeaders_Click(): InvalidatePreview: End Sub
Private Sub chkContents_Click(): InvalidatePreview: End Sub
Private Sub cboMode_Change()
    chkSkipHeaders.Enabled = (cboMode.ListIndex <> 0)
    chkContents.Enabled = (cboMode.ListIndex <> 1)
    InvalidatePreview
End Sub

Private Sub SetInputsEnabled(ByVal enabled As Boolean)
    lstInputFiles.Enabled = enabled
    cmdMoveUp.Enabled = enabled: cmdMoveDown.Enabled = enabled: cmdSelectAll.Enabled = enabled
    cmdAddFiles.Enabled = enabled: cmdRemoveFile.Enabled = enabled: cmdBrowse.Enabled = enabled
    txtOutputPath.Enabled = enabled: cboMode.Enabled = enabled: chkAllowPartial.Enabled = enabled
    chkSkipHeaders.Enabled = enabled And cboMode.ListIndex <> 0
    chkContents.Enabled = enabled And cboMode.ListIndex <> 1
End Sub

Private Sub cmdCancel_Click()
    If mExecuting Then
        NxFileOperationRequestCancel
        Exit Sub
    End If
    RequestCancel
End Sub

Private Sub InvalidatePreview()
    mPreviewReady = False
    cmdExecute.Enabled = True
    lblStep.Caption = "목록 " & lstInputFiles.ListCount & "개 전체를 위에서부터 통합합니다."
End Sub

Private Sub RequestCancel()
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    NxFormWheelDetach Me
    If mExecuting Then
        Cancel = True
        NxFileOperationRequestCancel
        Exit Sub
    End If
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub UserForm_Terminate()
    NxFormWheelDetach Me
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 파일 통합"
End Sub

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
