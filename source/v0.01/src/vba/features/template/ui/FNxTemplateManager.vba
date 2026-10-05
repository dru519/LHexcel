Option Explicit

Private mTargetBook As Workbook
Private mSourceSheet As Worksheet
Private mRecords As Collection
Private mBusy As Boolean
Private mRefreshing As Boolean
Private mKeys As Collection
Private mCatalog As Collection
Private mCatalogKeys As Collection
Private mHeaderSizers As Collection

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
    lstTemplates.IntegralHeight = False
    txtSearch.ControlTipText = "이름·설명·분류 검색"
    cboSort.Clear
    cboSort.AddItem "사용자 순서": cboSort.AddItem "최근 수정순": cboSort.AddItem "최근 사용순": cboSort.AddItem "이름순"
    cboSort.ListIndex = 0
    cboOutput.Clear
    cboOutput.AddItem "새 통합문서": cboOutput.AddItem "현재 문서에 새 시트"
    cboOutput.ListIndex = 0
    cmdCancel.Cancel = True
    Set mHeaderSizers = New Collection
    Dim index As Long, sizer As CNxTemplateHeaderSizer
    For index = 1 To 3
        Set sizer = New CNxTemplateHeaderSizer
        sizer.Bind Me, Me.Controls("grip" & CStr(index)), index
        mHeaderSizers.Add sizer
    Next index
    RestoreColumnWidths
    ResizeHeader 0, 0
End Sub

Private Sub UserForm_Activate()
    NxFormWheelAttach Me
End Sub

Public Sub HandleWheelDelta(ByVal wheelDelta As Long)
    NxFormWheelScrollActiveList Me, "lstTemplates", wheelDelta
End Sub

Private Sub UserForm_Terminate()
    SaveColumnWidths
    NxFormWheelDetach Me
End Sub

Private Sub RestoreColumnWidths()
    Dim saved As String, widths As Variant, i As Long, total As Double
    On Error GoTo IgnorePreference
    saved = GetSetting("LHExcel", "TemplateLibrary", "ColumnWidths", vbNullString)
    If Len(saved) = 0 Then Exit Sub
    widths = Split(saved, ";")
    If UBound(widths) <> 4 Then Exit Sub
    For i = 0 To 3
        widths(i) = Trim$(Replace(LCase$(CStr(widths(i))), "pt", vbNullString))
        If Not IsNumeric(widths(i)) Then Exit Sub
        If CDbl(widths(i)) < 45 Then Exit Sub
        total = total + CDbl(widths(i))
    Next i
    If total > lstTemplates.Width - 12 Or total < 180 Then Exit Sub
    If Val(widths(4)) <> 0 Then Exit Sub
    lstTemplates.ColumnWidths = saved
IgnorePreference:
End Sub

Private Sub SaveColumnWidths()
    On Error Resume Next
    SaveSetting "LHExcel", "TemplateLibrary", "ColumnWidths", lstTemplates.ColumnWidths
    On Error GoTo 0
End Sub

Public Sub ResizeHeader(ByVal index As Long, ByVal delta As Single)
    Dim widths As Variant, i As Long, cursor As Single, current As Single, adjacent As Single
    widths = Split(lstTemplates.ColumnWidths, ";")
    If index > 0 Then
        current = Val(widths(index - 1)) + delta
        adjacent = Val(widths(index)) - delta
        If current < 45 Or adjacent < 45 Then Exit Sub
        widths(index - 1) = CStr(current): widths(index) = CStr(adjacent)
        lstTemplates.ColumnWidths = Join(widths, ";")
    End If
    cursor = lstTemplates.Left
    For i = 1 To 4
        Me.Controls("header" & CStr(i)).Left = cursor
        Me.Controls("header" & CStr(i)).Width = Val(widths(i - 1)) - 2
        cursor = cursor + Val(widths(i - 1))
        If i < 4 Then Me.Controls("grip" & CStr(i)).Left = cursor - 2
    Next i
End Sub

Private Sub cmdUp_Click()
    MoveSelected -1
End Sub

Private Sub cmdDown_Click()
    MoveSelected 1
End Sub

Private Sub MoveSelected(ByVal direction As Long)
    Dim result As CNxResult
    If mBusy Then Exit Sub
    If direction = -1 And Not cmdUp.Enabled Then Exit Sub
    If direction = 1 And Not cmdDown.Enabled Then Exit Sub
    On Error GoTo Failed
    mBusy = True
    Set result = NxTemplateRunManagerAction("move", mTargetBook, SelectedTemplateId(), True, moveDirection:=direction)
    cboSort.ListIndex = 0
    FinishAction result, "순서 변경"
    Exit Sub
Failed:
    ShowFailure
End Sub

Public Sub BindContext(ByVal targetBook As Workbook, Optional ByVal featureId As String = "NX-TPL-LIST")
    Select Case featureId
        Case "NX-TPL-LIST", "NX-TPL-LOAD", "NX-TPL-RENAME", "NX-TPL-DELETE"
        Case Else: NxRaiseContractError "지원하지 않는 템플릿 보관함 작업입니다."
    End Select
    NxTemplateValidateUserBook targetBook
    Set mTargetBook = targetBook
    If TypeName(targetBook.ActiveSheet) = "Worksheet" Then Set mSourceSheet = targetBook.ActiveSheet
    txtFolder.Value = NxTemplateStoreRoot()
    RefreshTemplates
End Sub

Private Sub cmdRegisterFile_Click()
    Dim filePath As String
    If mBusy Then Exit Sub
    On Error GoTo Failed
    filePath = NxTemplatePickRegistrationFile()
    If Len(filePath) = 0 Then Exit Sub
    RegisterSource "file", filePath
    Exit Sub
Failed:
    ShowFailure
End Sub

Private Sub cmdRegisterSheet_Click()
    RegisterSource "sheet_used_range"
End Sub

Private Sub RegisterSource(ByVal sourceKind As String, Optional ByVal filePath As String = vbNullString)
    Dim result As CNxResult
    If mBusy Then Exit Sub
    On Error GoTo Failed
    mBusy = True
    Set result = NxTemplateShowRegistration(mTargetBook, mSourceSheet, sourceKind, filePath)
    mBusy = False
    If result Is Nothing Then
        lblStatus.Caption = "등록을 취소했습니다. 등록된 내용은 바뀌지 않았습니다."
    Else
        FinishAction result, NxTemplateManagerKindLabel(sourceKind) & " 등록"
    End If
    Exit Sub
Failed:
    ShowFailure
End Sub

Private Sub cmdUse_Click()
    If cboOutput.ListIndex = 0 Then
        RunSelectedAction "copy", "템플릿 사용"
    Else
        RunSelectedAction "use", "템플릿 사용"
    End If
End Sub

Private Sub cmdDelete_Click()
    RunSelectedAction "delete", "템플릿 삭제"
End Sub

Private Sub cmdEdit_Click()
    RunSelectedAction "rename", "상세 저장"
End Sub

Private Sub txtSearch_Change()
    FilterChanged
End Sub

Private Sub cboFilter_Change()
    FilterChanged
End Sub

Private Sub cboSort_Change()
    FilterChanged
End Sub

Private Sub FilterChanged()
    If mTargetBook Is Nothing Or mBusy Or mRefreshing Then Exit Sub
    On Error GoTo Failed
    RefreshTemplates False
    Exit Sub
Failed:
    ShowFailure
End Sub

Private Sub cmdCatalog_Click()
    RunSelectedAction "catalog", "새 관리표 생성"
End Sub

Private Sub RunSelectedAction(ByVal action As String, ByVal actionLabel As String)
    Dim result As CNxResult, templateId As String
    If mBusy Then Exit Sub
    On Error GoTo Failed
    If action <> "catalog" Then templateId = SelectedTemplateId()
    If action = "use" Or action = "copy" Then
        If Not NxTemplateManagerActionConsent(templateId) Then
            lblStatus.Caption = actionLabel & "를 취소했습니다."
            Exit Sub
        End If
    End If
    mBusy = True
    NxTemplateBeginProgress Me
    If action = "rename" Then
        Set result = NxTemplateRunManagerAction(action, mTargetBook, templateId, True, CStr(txtName.Value), CStr(txtDescription.Value), CStr(txtCategory.Value))
    ElseIf action = "restore" Then
        Set result = NxTemplateRunManagerAction(action, mTargetBook, templateId, True, quarantineId:=CStr(mKeys(lstTemplates.ListIndex + 1)))
    Else
        Set result = NxTemplateRunManagerAction(action, mTargetBook, templateId, True)
    End If
    If Not result Is Nothing Then
        If result.Outcome = NxSuccess And (action = "use" Or action = "copy") Then
            mBusy = False
            NxTemplateEndProgress
            Unload Me
            Exit Sub
        End If
    End If
    FinishAction result, actionLabel
    Exit Sub
Failed:
    ShowFailure
End Sub

Private Sub FinishAction(ByVal result As CNxResult, ByVal actionLabel As String)
    NxTemplateEndProgress
    mBusy = False
    If result Is Nothing Then NxRaiseContractError "템플릿 작업 결과가 없습니다."
    Select Case result.Outcome
        Case NxSuccess
            RefreshTemplates
            lblStatus.Caption = actionLabel & " 완료. 원본 파일·기존 시트는 보존했습니다."
        Case NxCancelled
            lblStatus.Caption = actionLabel & " 취소. 등록된 내용은 바뀌지 않았습니다."
        Case Else
            lblStatus.Caption = actionLabel & " 실패: " & NxUserErrorText(result.Recovery)
            MsgBox NxUserErrorText(result.Recovery), vbExclamation, "내엑셀 - 템플릿 관리"
    End Select
End Sub

Private Sub cmdRefresh_Click()
    If mBusy Then Exit Sub
    On Error GoTo Failed
    RefreshTemplates
    Exit Sub
Failed:
    ShowFailure
End Sub

Private Sub cmdOpenFolder_Click()
    If mBusy Then Exit Sub
    On Error GoTo Failed
    NxTemplateManagerOpenFolder
    lblStatus.Caption = "템플릿 저장 폴더를 열었습니다. 저장소 파일을 직접 바꾸면 무결성 검사가 실패할 수 있습니다."
    Exit Sub
Failed:
    ShowFailure
End Sub

Private Sub RefreshTemplates(Optional ByVal reloadStore As Boolean = True)
    Dim record As CNxTemplateRecord, selectedId As String, index As Long, selectedIndex As Long
    Dim priorBook As Workbook, filter As String, categories As Object, item As Variant, allRecords As Collection
    On Error GoTo Failed
    Set priorBook = ActiveWorkbook
    If Not mRecords Is Nothing Then
        If lstTemplates.ListIndex >= 0 And lstTemplates.ListIndex < mRecords.Count Then selectedId = mRecords.Item(lstTemplates.ListIndex + 1).TemplateId
    End If
    mRefreshing = True
    filter = CStr(cboFilter.Value)
    If filter = "전체 분류" Then filter = vbNullString
    If reloadStore Or mCatalog Is Nothing Then _
        Set mCatalog = NxTemplateLibraryRecords("", "", 0, False, mCatalogKeys)
    Set mRecords = NxTemplateFilterRecords(mCatalog, mCatalogKeys, Trim$(CStr(txtSearch.Value)), filter, cboSort.ListIndex, mKeys)
    Set categories = CreateObject("Scripting.Dictionary")
    categories.CompareMode = vbTextCompare
    Set allRecords = mCatalog
    For Each record In allRecords
        If Len(record.Category) > 0 Then categories(record.Category) = True
    Next record
    cboFilter.Clear: cboFilter.AddItem "전체 분류"
    For Each item In categories.Keys: cboFilter.AddItem CStr(item): Next item
    If Len(filter) > 0 And Not categories.Exists(filter) Then cboFilter.AddItem filter
    If Len(filter) = 0 Then cboFilter.ListIndex = 0 Else cboFilter.Value = filter
    lstTemplates.Clear
    For Each record In mRecords
        lstTemplates.AddItem record.DisplayName
        lstTemplates.List(index, 1) = record.Category
        lstTemplates.List(index, 2) = NxTemplateManagerKindLabel(record.SourceKind)
        lstTemplates.List(index, 3) = Left$(Replace(record.ModifiedUtc, "T", " "), 16)
        lstTemplates.List(index, 4) = record.TemplateId
        If record.TemplateId = selectedId Then selectedIndex = index
        index = index + 1
    Next record
    If mRecords.Count > 0 Then lstTemplates.ListIndex = selectedIndex
    mRefreshing = False
    RenderSelected
    If Not priorBook Is Nothing Then priorBook.Activate
    If mRecords.Count = 0 Then lblStatus.Caption = "조건에 맞는 템플릿이 없습니다. 검색·분류를 바꾸거나 새로 등록하세요."
    Exit Sub
Failed:
    mRefreshing = False
    On Error Resume Next
    If Not priorBook Is Nothing Then priorBook.Activate
    On Error GoTo 0
    NxRaiseContractError "템플릿 목록을 읽지 못했습니다. 저장소 경로와 파일 상태를 확인하세요."
End Sub

Private Function PackageSizeText(ByVal templateId As String) As String
    On Error GoTo Unavailable
    PackageSizeText = Format$(CDbl(CreateObject("Scripting.FileSystemObject").GetFile(NxTemplatePackagePath(templateId)).Size) / 1024#, "0.0")
    Exit Function
Unavailable:
    PackageSizeText = "-"
End Function

Private Sub lstTemplates_Change()
    If mRefreshing Then Exit Sub
    RenderSelected
End Sub

Private Sub lstTemplates_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If cmdUse.Enabled Then cmdUse_Click
End Sub

Private Sub RenderSelected()
    Dim record As CNxTemplateRecord, selected As Boolean, healthy As Boolean
    If Not mRecords Is Nothing Then selected = (lstTemplates.ListIndex >= 0 And lstTemplates.ListIndex < mRecords.Count)
    If selected Then
        Set record = mRecords.Item(lstTemplates.ListIndex + 1)
        healthy = (record.Status = "healthy")
        txtName.Value = record.DisplayName: txtDescription.Value = record.Description: txtCategory.Value = record.Category
        lblStatus.Caption = CStr(mRecords.Count) & "개 / " & record.DisplayName & " · " & _
            NxTemplateManagerKindLabel(record.SourceKind) & " · 실행 시 무결성 검사" & vbCrLf & _
            "사용 셀 " & CStr(CDbl(record.RowCount) * CDbl(record.ColumnCount)) & _
            "개. 제거는 복구 가능한 격리 보관입니다."
        If Not healthy Then lblStatus.Caption = "손상된 항목입니다. 사용하거나 복원할 수 없습니다."
    Else
        txtName.Value = vbNullString: txtDescription.Value = vbNullString: txtCategory.Value = vbNullString
    End If
    cmdUse.Enabled = healthy
    cmdEdit.Enabled = healthy
    cmdDelete.Enabled = selected
    cmdUp.Enabled = False
    cmdDown.Enabled = False
    If selected And cboSort.ListIndex = 0 And Len(Trim$(CStr(txtSearch.Value))) = 0 And _
        (CStr(cboFilter.Value) = "전체 분류" Or Len(CStr(cboFilter.Value)) = 0) Then
        cmdUp.Enabled = (lstTemplates.ListIndex > 0)
        cmdDown.Enabled = (lstTemplates.ListIndex < lstTemplates.ListCount - 1)
    End If
    cmdUp.ControlTipText = "전체 분류 · 검색 없음 · 사용자 순서에서 위로 이동"
    cmdDown.ControlTipText = "전체 분류 · 검색 없음 · 사용자 순서에서 아래로 이동"
End Sub

Private Function SelectedTemplateId() As String
    If mRecords Is Nothing Then NxRaiseContractError "템플릿을 먼저 선택하세요."
    If lstTemplates.ListIndex < 0 Or lstTemplates.ListIndex >= mRecords.Count Then NxRaiseContractError "템플릿을 먼저 선택하세요."
    SelectedTemplateId = mRecords.Item(lstTemplates.ListIndex + 1).TemplateId
End Function

Private Sub ShowFailure()
    Dim message As String
    message = Err.Description
    NxTemplateEndProgress
    mBusy = False
    lblStatus.Caption = "작업을 완료하지 못했습니다: " & message
    message = NxUserErrorText(message)
    MsgBox message, vbExclamation, "내엑셀 - 템플릿 관리"
End Sub

Private Sub cmdCancel_Click()
    If mBusy Then
        NxTemplateRequestCancel
        lblStatus.Caption = "취소 요청 중입니다. 현재 작업의 임시 결과를 정리합니다."
        Exit Sub
    End If
    NxFormWheelDetach Me
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim sizer As CNxTemplateHeaderSizer
    If mBusy Then
        NxTemplateRequestCancel
        Cancel = True
        Exit Sub
    End If
    NxFormWheelDetach Me
    If Not mHeaderSizers Is Nothing Then
        For Each sizer In mHeaderSizers
            sizer.Unbind
        Next sizer
        Set mHeaderSizers = Nothing
    End If
End Sub

Public Sub UpdateTemplateProgress(ByVal stage As String, ByVal completed As Long, ByVal total As Long)
    lblStatus.Caption = NxTemplateProgressText(stage, completed, total) & " · 닫기: 취소"
    Me.Repaint
End Sub
