Option Explicit

Private mPreviewReady As Boolean
Private mSelectedChart As Object
Private mAllowClose As Boolean
Private mResolvedGrade As NxExecutionGrade
Private mExecuting As Boolean
Private mFeatureSeeded As Boolean

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
    cmdExecute.Default = True
    cmdCancel.Cancel = True
    mResolvedGrade = NxExecutionPlanned
    UpdateWorkflowState False
End Sub

Public Sub BindSelection(ByVal selectedRange As Range, Optional ByVal featureId As String = vbNullString)
    If selectedRange Is Nothing Then NxRaiseContractError "파일관리 선택 범위가 필요합니다."
    selectedRange.Select
    ApplyFeatureSeed featureId
End Sub

Public Sub BindFeatureSeed(ByVal featureId As String)
    ApplyFeatureSeed featureId
End Sub

Private Sub ApplyFeatureSeed(ByVal featureId As String)
    If Len(featureId) = 0 Then Exit Sub
    mFeatureSeeded = True
    cboFeature.ListIndex = FeatureIndexForId(featureId)
    cboFeature.Enabled = False
End Sub

Private Sub cmdAddFiles_Click()
    Dim picker As FileDialog, item As Variant, index As Long, exists As Boolean
    Set picker = Application.FileDialog(3)
    With picker
        .AllowMultiSelect = True
        .Title = "통합할 파일 선택"
        .Filters.Clear
        .Filters.Add "Excel 통합문서", "*.xlsx;*.xlsm;*.xlsb;*.xls"
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
    If lstInputFiles.ListIndex < 0 Then NxRaiseContractError "제거할 파일을 선택하세요."
    lstInputFiles.RemoveItem lstInputFiles.ListIndex
    InvalidatePreview
End Sub

Private Sub cmdBrowse_Click()
    Dim selectedPath As Variant, filter As String, defaultName As String
    If SelectedFeatureId() = "NX-FILE-RANGE-PNG" Or SelectedFeatureId() = "NX-FILE-CHART-PNG" Then
        filter = "PNG 파일 (*.png), *.png": defaultName = "내엑셀_결과.png"
    Else
        filter = "Excel 통합문서 (*.xlsx), *.xlsx": defaultName = "내엑셀_결과.xlsx"
    End If
    selectedPath = Application.GetSaveAsFilename(InitialFileName:=defaultName, FileFilter:=filter, _
        FilterIndex:=1, Title:="결과 경로 선택")
    If VarType(selectedPath) <> vbBoolean Then txtOutputPath.Value = CStr(selectedPath)
    InvalidatePreview
End Sub

Private Sub cmdPreview_Click()
    Dim validInputs As Collection, excludedSummary As String
    On Error GoTo Failed
    lstDecisions.Clear
    Set mSelectedChart = Nothing
    Select Case SelectedFeatureId()
        Case "NX-FILE-CONSOLIDATE"
            RequireOutputPath
            If InputPathCount() = 0 Then NxRaiseContractError "통합할 파일을 하나 이상 추가하세요."
            Set validInputs = NxConsolidationPlan(InputPaths(), chkAllowPartial.Value, excludedSummary)
            lstDecisions.AddItem "입력 한도: 최대 20개 · 파일당 100 MiB · 읽기 전용"
            lstDecisions.AddItem "유효 입력 " & CStr(validInputs.Count) & "개 · 제외 입력 " & CStr(ExcludedInputCount(excludedSummary)) & "개"
            lstDecisions.AddItem "부분결과 허용: " & IIf(chkAllowPartial.Value, "예(추가 승인 필요)", "아니오")
        Case "NX-FILE-RANGE-PNG"
            RequireOutputPath
            If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "범위 PNG는 셀 범위를 먼저 선택하세요."
            If Application.Selection.Areas.Count <> 1 Then NxRaiseContractError "범위 PNG는 연속된 한 범위만 지원합니다."
            lstDecisions.AddItem "선택 범위 " & Application.Selection.Address(False, False) & " · 최대 10,000셀"
        Case "NX-FILE-CHART-PNG"
            RequireOutputPath
            Set mSelectedChart = ResolveSelectedChart()
            lstDecisions.AddItem "선택 차트 " & CStr(mSelectedChart.Name) & " · 원본 변경 없음"
    End Select
    lblOutputSafety.Caption = "실행 등급: Planned · 출력 안전: 새 파일 생성 · 기존 파일 자동 덮어쓰기 금지"
    lstDecisions.AddItem "실행 등급 Planned · 출력 파일은 공통 실행 프레임 승인 후 생성"
    mPreviewReady = True
    ResolveWorkflowGrade
    UpdateWorkflowState True
    cmdExecute.Enabled = True
    Exit Sub
Failed:
    mPreviewReady = False
    UpdateWorkflowState False
    cmdExecute.Enabled = False
    ShowError
End Sub

Private Sub cmdExecute_Click()
    Dim result As CNxResult, featureId As String, secondApproval As Boolean, failureDescription As String
    On Error GoTo Failed
    cmdPreview_Click
    If Not mPreviewReady Then Exit Sub
    featureId = SelectedFeatureId()
    If featureId = "NX-FILE-CHART-PNG" And mSelectedChart Is Nothing Then NxRaiseContractError "차트 선택이 사라졌습니다. 미리보기를 다시 실행하세요."
    BeginExecutionUi
    Set result = NxFileRun(featureId, CStr(txtOutputPath.Value), InputPaths(), chkAllowPartial.Value, mSelectedChart, False)
    If result.Outcome = NxCancelled And featureId = "NX-FILE-CONSOLIDATE" And chkAllowPartial.Value Then
        If MsgBox("실패 파일을 제외한 부분결과와 실패 목록을 생성하시겠습니까?", vbYesNo + vbQuestion + vbDefaultButton2, "내엑셀 - 2차 승인") = vbYes Then
            secondApproval = True
            Set result = NxFileRun(featureId, CStr(txtOutputPath.Value), InputPaths(), True, mSelectedChart, secondApproval)
        End If
    End If
    If result Is Nothing Then NxRaiseContractError "파일 작업 결과를 확인할 수 없습니다."
    If result.Outcome = NxCancelled Then
        EndExecutionUi
        MsgBox "작업을 취소했습니다. 원본 파일은 변경되지 않았습니다.", vbInformation + vbOKOnly, "내엑셀 - 파일관리"
        Exit Sub
    End If
    If result.Outcome = NxPartialFailure Then
        EndExecutionUi
        MsgBox "부분결과와 제외·실패 목록을 생성했습니다." & vbCrLf & result.Target, vbInformation + vbOKOnly, "내엑셀 - 파일관리"
        RequestCancel
        Exit Sub
    End If
    If result.Outcome <> NxSuccess Then NxRaiseContractError result.Recovery
    EndExecutionUi
    RequestCancel
    Exit Sub
Failed:
    failureDescription = Err.Description
    Err.Clear
    EndExecutionUi
    MsgBox NxUserErrorText(failureDescription), vbExclamation + vbOKOnly, "내엑셀 - 파일관리"
End Sub

Private Sub cmdCancel_Click()
    If mExecuting Then
        NxFileOperationRequestCancel
        cmdCancel.Caption = "취소 요청됨"
        Exit Sub
    End If
    RequestCancel
End Sub

Private Sub cboFeature_Change(): InvalidatePreview: End Sub
Private Sub chkAllowPartial_Click(): InvalidatePreview: End Sub
Private Sub txtOutputPath_Change(): InvalidatePreview: End Sub

Private Function SelectedFeatureId() As String
    Select Case cboFeature.ListIndex
        Case 0: SelectedFeatureId = "NX-FILE-CONSOLIDATE"
        Case 1: SelectedFeatureId = "NX-FILE-RANGE-PNG"
        Case 2: SelectedFeatureId = "NX-FILE-CHART-PNG"
        Case Else: NxRaiseContractError "기능을 선택하세요."
    End Select
End Function

Private Function FeatureIndexForId(ByVal featureId As String) As Long
    Select Case featureId
        Case "NX-FILE-CONSOLIDATE": FeatureIndexForId = 0
        Case "NX-FILE-RANGE-PNG": FeatureIndexForId = 1
        Case "NX-FILE-CHART-PNG": FeatureIndexForId = 2
        Case Else: NxRaiseContractError "File feature seed is outside the closed contract"
    End Select
End Function

Private Sub RequireOutputPath()
    If Len(Trim$(CStr(txtOutputPath.Value))) = 0 Or CStr(txtOutputPath.Value) = "새 결과 경로" Then NxRaiseContractError "결과 경로를 지정하세요."
End Sub

Private Function InputPathCount() As Long
    InputPathCount = lstInputFiles.ListCount
End Function

Private Function InputPaths() As Variant
    Dim result() As String, index As Long
    If lstInputFiles.ListCount = 0 Then InputPaths = Empty: Exit Function
    ReDim result(0 To lstInputFiles.ListCount - 1)
    For index = 0 To lstInputFiles.ListCount - 1: result(index) = CStr(lstInputFiles.List(index)): Next index
    InputPaths = result
End Function

Private Function ExcludedInputCount(ByVal excludedSummary As String) As Long
    If Len(excludedSummary) = 0 Then Exit Function
    ExcludedInputCount = UBound(Split(excludedSummary, vbCrLf)) + 1
End Function

Private Function ResolveSelectedChart() As Object
    Dim selected As Object, candidate As Object, chart As Object, owner As Object
    On Error Resume Next
    Set selected = Application.Selection
    If selected Is Nothing Then On Error GoTo 0: NxRaiseContractError "차트 하나를 먼저 선택하세요."
    Select Case TypeName(selected)
        Case "Chart": Set chart = selected
        Case "ChartObject": Set ResolveSelectedChart = selected: On Error GoTo 0: Exit Function
        Case "Shape": Set chart = selected.Chart
        Case "ShapeRange"
            If selected.Count <> 1 Then On Error GoTo 0: NxRaiseContractError "차트 하나만 선택하세요."
            Set candidate = selected.Item(1): Set chart = candidate.Chart
        Case "ChartArea", "PlotArea", "Legend", "Series": Set chart = selected.Parent
    End Select
    If chart Is Nothing Then Err.Clear: On Error GoTo 0: NxRaiseContractError "선택 항목이 차트가 아닙니다."
    If Len(CStr(chart.Name)) = 0 Then Err.Clear: On Error GoTo 0: NxRaiseContractError "차트 이름을 확인할 수 없습니다."
    On Error Resume Next
    Set owner = chart.Parent
    If Not owner Is Nothing Then
        If TypeName(owner) = "ChartObject" Then Set ResolveSelectedChart = owner
    End If
    On Error GoTo 0
    If owner Is Nothing Or TypeName(owner) <> "ChartObject" Then Set ResolveSelectedChart = chart
End Function

Private Sub InvalidatePreview()
    mPreviewReady = False
    Set mSelectedChart = Nothing
    cmdExecute.Enabled = True
    UpdateWorkflowState False
End Sub

Private Sub UpdateWorkflowState(ByVal previewed As Boolean)
    If previewed Then
        lblWorkflow.Caption = "워크플로: 입력 → 미리보기 완료 → 실행 가능"
    Else
        lblWorkflow.Caption = "워크플로: 입력 → 미리보기 선택 가능"
    End If
    lblExecutionGrade.Caption = "실행 등급: " & GradeCaption(mResolvedGrade)
    If mResolvedGrade = NxExecutionPlanned Then
        lblOutputSafety.Caption = "실행 등급 [계획] · 출력 안전: 새 파일 생성 · 기존 파일 자동 덮어쓰기 금지"
    Else
        lblOutputSafety.Caption = lblExecutionGrade.Caption & " · 출력 안전: 새 파일 생성 · 기존 파일 자동 덮어쓰기 금지"
    End If
    Select Case mResolvedGrade
        Case NxExecutionFast: cmdExecute.Caption = "바로 실행"
        Case NxExecutionGuarded: cmdExecute.Caption = "확인 후 실행"
        Case NxExecutionPlanned: cmdExecute.Caption = "실행계획 확인"
    End Select
End Sub

Private Function GradeCaption(ByVal grade As NxExecutionGrade) As String
    Select Case grade
        Case NxExecutionFast: GradeCaption = "[빠름]"
        Case NxExecutionGuarded: GradeCaption = "[확인]"
        Case NxExecutionPlanned: GradeCaption = "[계획]"
        Case Else: NxRaiseContractError "실행 등급이 유효하지 않습니다."
    End Select
End Function

Private Sub ResolveWorkflowGrade()
    Dim definition As CNxFeatureDefinition
    Dim decision As CNxExecutionDecision
    Set definition = NxFileFeatureDefinition(SelectedFeatureId())
    Set decision = NxExecutionGradePolicy.ResolveDecision( _
        definition:=definition, actualCellCount:=0, fileMutation:=False, _
        externalDocumentMutation:=False, manualRecoveryMutation:=False, _
        fileCreation:=True)
    mResolvedGrade = decision.ResolvedGrade
End Sub

Private Sub RequestCancel()
    If mExecuting Then
        NxFileOperationRequestCancel
        Exit Sub
    End If
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mExecuting Then
        Cancel = True
        NxFileOperationRequestCancel
        cmdCancel.Caption = "취소 요청됨"
        Exit Sub
    End If
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub BeginExecutionUi()
    mExecuting = True
    cboFeature.Enabled = False
    lstInputFiles.Enabled = False
    cmdAddFiles.Enabled = False
    cmdRemoveFile.Enabled = False
    txtOutputPath.Enabled = False
    cmdBrowse.Enabled = False
    chkAllowPartial.Enabled = False
    cmdPreview.Enabled = False
    cmdExecute.Enabled = False
    cmdCancel.Caption = "작업 중단"
    lblWorkflow.Caption = "워크플로: 승인 후 실행 중 · Excel 상태표시줄에서 진행률 확인 · Esc로 취소"
End Sub

Private Sub EndExecutionUi()
    If Not mExecuting Then Exit Sub
    mExecuting = False
    cboFeature.Enabled = Not mFeatureSeeded
    lstInputFiles.Enabled = True
    cmdAddFiles.Enabled = True
    cmdRemoveFile.Enabled = True
    txtOutputPath.Enabled = True
    cmdBrowse.Enabled = True
    chkAllowPartial.Enabled = True
    cmdPreview.Enabled = True
    cmdExecute.Enabled = True
    cmdCancel.Caption = "닫기"
    UpdateWorkflowState mPreviewReady
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 파일관리"
End Sub
