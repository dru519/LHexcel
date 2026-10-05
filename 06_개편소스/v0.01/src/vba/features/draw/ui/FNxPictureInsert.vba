Option Explicit

Private mInit As CNxFormInitializationGuard, mContext As CNxFeatureDialogContext
Private mRequest As CNxPictureInsertRequest, mPreview As CNxPictureInsertPreview
Private mFitRequest As CNxDrawingRequest, mSelectedFiles As Collection
Private mPicture As Shape, mRangeSession As CNxRangeSelectionSession
Private mBoundBook As Workbook, mBoundSheet As Worksheet, mOriginalSelection As Object
Private mPreviewReady As Boolean, mAllowClose As Boolean, mBinding As Boolean

Public Sub BindFeatureContext(ByVal context As CNxFeatureDialogContext)
    Dim target As Range
    If context Is Nothing Then NxRaiseContractError "그림 대화상자 컨텍스트가 필요합니다."
    If Not context.IsConfigured Then NxRaiseContractError "그림 대화상자 컨텍스트가 필요합니다."
    If context.DialogId <> "NX-DLG-DRAW-INSERT-PICTURE" Then NxRaiseContractError "그림 대화상자 ID가 일치하지 않습니다."
    If Not mContext Is Nothing Then NxRaiseContractError "그림 대화상자는 한 번만 연결할 수 있습니다."
    Select Case context.FeatureId
        Case "NX-DRAW-INSERT-PICTURE"
            If context.DialogVariant <> "picture-insert" Then NxRaiseContractError "그림 삽입 유형이 일치하지 않습니다."
        Case "NX-DRAW-FIT-PICTURE"
            If context.DialogVariant <> "picture-fit" Then NxRaiseContractError "선택 그림 맞춤 유형이 일치하지 않습니다."
        Case Else: NxRaiseContractError "그림 기능 ID가 일치하지 않습니다."
    End Select
    If ActiveWorkbook Is Nothing Then NxRaiseContractError "통합문서를 먼저 여세요."
    If TypeName(ActiveSheet) <> "Worksheet" Then NxRaiseContractError "워크시트에서 그림 기능을 여세요."
    Set mBoundBook = ActiveWorkbook: Set mBoundSheet = ActiveSheet
    Set mOriginalSelection = Application.Selection
    Set mContext = context: Set mRangeSession = New CNxRangeSelectionSession
    mBinding = True
    If TypeName(mOriginalSelection) = "Range" Then
        Set target = mOriginalSelection
        If target.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위를 선택하세요."
        cboMode.ListIndex = 0
        If context.FeatureId = "NX-DRAW-FIT-PICTURE" Then cboMode.ListIndex = 1
    ElseIf context.FeatureId = "NX-DRAW-INSERT-PICTURE" And IsMultipleDrawingSelection(mOriginalSelection) Then
        ' Existing selections are not the source of a new-file insertion.
        ' Keep them intact for picker/cancel; single-picture FIT remains strict.
        Set target = ActiveCell.MergeArea
        cboMode.ListIndex = 0
    Else
        BindPictureSelection
        Set target = mPicture.TopLeftCell.MergeArea
        cboMode.ListIndex = 1
    End If
    mRangeSession.UpdateFromRange target
    txtRange.Value = mRangeSession.DisplayAddress
    mBinding = False
    ApplyMode
    InvalidatePreview
End Sub

Private Function IsMultipleDrawingSelection(ByVal selected As Object) As Boolean
    Select Case TypeName(selected)
        Case "DrawingObjects"
            IsMultipleDrawingSelection = (selected.ShapeRange.Count > 1)
        Case "ShapeRange"
            IsMultipleDrawingSelection = (selected.Count > 1)
    End Select
End Function

Private Sub UserForm_Initialize()
    Set mInit = New CNxFormInitializationGuard
    mInit.BeginInitialization
    NxInitializeStaticComboChoices Me
    cmdExecute.Default = True: cmdCancel.Cancel = True
    txtMargin.Value = "4": chkMoveAndSize.Value = True
    cmdExecute.Enabled = True
    mInit.CompleteInitialization
End Sub

Private Function ReadyForEvents() As Boolean
    If mInit Is Nothing Then Exit Function
    If mInit.IsInitializing Or mBinding Then Exit Function
    ReadyForEvents = True
End Function

Private Function IsFitMode() As Boolean
    If cboMode.ListIndex < 0 Or cboMode.ListIndex > 1 Then NxRaiseContractError "그림 작업 모드를 선택하세요."
    IsFitMode = (cboMode.ListIndex = 1)
End Function

Private Sub ApplyMode()
    Dim fitting As Boolean
    fitting = IsFitMode()
    cmdBrowse.Visible = Not fitting: cmdUseSelection.Visible = fitting
    cboSizeMode.Visible = Not fitting: lblSizeHelp.Visible = Not fitting
    lblMargin.Visible = fitting: txtMargin.Visible = fitting: chkMoveAndSize.Visible = fitting
    If fitting Then
        lblContext.Caption = "현재 시트의 그림 하나를 선택하고 대상 범위에 맞춥니다."
        lblSource.Caption = "맞출 그림 · 다른 그림을 선택한 뒤 [선택 그림 가져오기]"
        cmdExecute.Caption = "그림 맞춤"
        txtFiles.Value = "시트에서 그림 하나를 선택한 뒤 [선택 그림 가져오기]를 누르세요."
        If Not mPicture Is Nothing Then txtFiles.Value = "선택 그림: " & mPicture.Name
    Else
        lblContext.Caption = "여러 그림 파일을 선택한 범위의 보이는 셀에 순서대로 넣습니다."
        lblSource.Caption = "삽입할 그림 파일 · 여러 파일 선택 가능"
        cmdExecute.Caption = "그림 삽입"
        txtFiles.Value = JoinPictureInsertFiles(mSelectedFiles)
        If Len(txtFiles.Value) = 0 Then txtFiles.Value = "[찾아보기]에서 그림 파일을 선택하세요."
    End If
End Sub

Private Sub cmdBrowse_Click()
    If Not ReadyForEvents() Then Exit Sub
    Dim files As Collection
    On Error GoTo Failed
    InvalidatePreview
    RequireContext
    If IsFitMode() Then NxRaiseContractError "새 그림 삽입 모드에서 파일을 선택하세요."
    Set files = NxPictureInsertPickFiles()
    If files Is Nothing Then Exit Sub
    Set mSelectedFiles = files
    txtFiles.Value = JoinPictureInsertFiles(mSelectedFiles)
    Exit Sub
Failed: ShowFailure NxUserErrorText(Err.Description)
End Sub

Private Sub cmdUseSelection_Click()
    If Not ReadyForEvents() Then Exit Sub
    On Error GoTo Failed
    InvalidatePreview
    RequireContext
    If Not IsFitMode() Then NxRaiseContractError "선택 그림 맞춤 모드에서 그림을 가져오세요."
    BindPictureSelection
    ApplyMode
    Exit Sub
Failed: ShowFailure NxUserErrorText(Err.Description)
End Sub

Private Sub BindPictureSelection()
    Dim selected As Object, candidate As Shape
    Set selected = Application.Selection
    Select Case TypeName(selected)
        Case "Shape", "Picture"
            Set candidate = selected.Parent.Shapes(selected.Name)
        Case "ShapeRange"
            If selected.Count <> 1 Then NxRaiseContractError "그림은 하나만 선택하세요."
            Set candidate = selected.Item(1)
        Case Else: NxRaiseContractError "현재 시트에서 그림 하나를 선택하세요."
    End Select
    If candidate.Type <> msoPicture And candidate.Type <> msoLinkedPicture Then NxRaiseContractError "사진이나 그림만 맞출 수 있습니다. 도형과 차트는 제외됩니다."
    If Not (candidate.Parent Is mBoundSheet) Then NxRaiseContractError "처음 연 시트의 그림을 선택하세요."
    Set mPicture = candidate
End Sub

Private Sub cmdPickRange_Click()
    If Not ReadyForEvents() Then Exit Sub
    Dim target As Range
    On Error GoTo Failed
    InvalidatePreview
    RequireContext
    If Not NxPickRange(Me, mRangeSession) Then Exit Sub
    RequireContext
    Set target = mRangeSession.ResolveRange
    If Not (target.Worksheet Is mBoundSheet) Then NxRaiseContractError "처음 연 시트 안에서 범위를 선택하세요."
    txtRange.Value = mRangeSession.DisplayAddress
    Exit Sub
Failed: ShowFailure NxUserErrorText(Err.Description)
End Sub

Private Sub cmdPreview_Click()
    If Not ReadyForEvents() Then Exit Sub
    On Error GoTo Failed
    RefreshPreview
    If Not IsFitMode() Then
        If Not NxPictureInsertNativePreview(mRequest, mPreview) Then InvalidatePreview
    End If
    Exit Sub
Failed: ShowFailure NxUserErrorText(Err.Description)
End Sub

Private Sub RefreshPreview()
    Dim target As Range, margin As Double, context As CNxExecutionContext
    InvalidatePreview
    RequireContext
    Set target = ResolveTarget()
    ' The sealed drawing context is a Range context, even when entry selected a picture.
    target.Select
    If IsFitMode() Then
        If mPicture Is Nothing Then NxRaiseContractError "시트에서 그림을 선택하고 [선택 그림 가져오기]를 누르세요."
        If Not IsNumeric(txtMargin.Value) Then NxRaiseContractError "여백은 0 이상의 숫자(pt)로 입력하세요."
        margin = CDbl(txtMargin.Value)
        NxDrawValidatePictureSelection mPicture, target, margin
        If target.Width <= margin * 2 Or target.Height <= margin * 2 Then NxRaiseContractError "여백을 제외한 대상 범위의 너비와 높이가 0보다 커야 합니다."
        Set mFitRequest = New CNxDrawingRequest
        mFitRequest.ConfigurePicture mPicture, target, margin, CBool(chkMoveAndSize.Value)
        Set context = NxContextFactory.CaptureCurrent()
        mFitRequest.RequireContext context
        mPreviewReady = True
        txtPreview.Value = "그림: " & mPicture.Name & vbCrLf & "대상: " & target.Address(External:=True) & vbCrLf & "여백 " & CStr(margin) & "pt · 비율 유지 · " & IIf(chkMoveAndSize.Value, "셀과 함께 이동/크기 변경", "기존 배치 속성 유지")
    Else
        If mSelectedFiles Is Nothing Then NxRaiseContractError "[찾아보기]에서 삽입할 그림 파일을 선택하세요."
        Set mRequest = NxPictureInsertCreateRequest(target, mSelectedFiles, cboSizeMode.Value)
        Set mPreview = NxPictureInsertBuildPreview(mRequest)
        mPreviewReady = mPreview.IsSealed
        txtPreview.Value = "대상: " & mPreview.TargetAddress & vbCrLf & mPreview.Summary
    End If
    cmdExecute.Enabled = mPreviewReady
End Sub

Private Sub cmdExecute_Click()
    If Not ReadyForEvents() Then Exit Sub
    Dim result As CNxResult, context As CNxExecutionContext
    Dim commandObject As CNxDrawingFeatureCommand, command As INxFeatureCommand
    On Error GoTo Failed
    RequireContext
    RefreshPreview
    If Not mPreviewReady Then Exit Sub
    If IsFitMode() Then
        If mFitRequest Is Nothing Then NxRaiseContractError "선택 그림 미리보기가 없습니다."
        RequireSameTarget mFitRequest.Target
        Set context = NxContextFactory.CaptureCurrent()
        mFitRequest.RequireContext context
        Set commandObject = New CNxDrawingFeatureCommand
        commandObject.Configure mFitRequest
        Set command = commandObject
        Set result = NxRunDrawingFeatureCommand(command, NxDrawingTargetSummary(mFitRequest.Target))
    Else
        If mRequest Is Nothing Then NxRaiseContractError "삽입 미리보기가 없습니다."
        If mPreview Is Nothing Then NxRaiseContractError "삽입 미리보기가 없습니다."
        RequireSameTarget mRequest.Target
        If Not mPreview.Revalidate(mRequest) Then NxRaiseContractError "파일 또는 범위가 변경되었습니다. 미리보기를 다시 확인하세요."
        Set result = NxPictureInsertRunSealed(mRequest, mPreview)
    End If
    InvalidatePreview
    If result Is Nothing Then Exit Sub
    If result.Outcome <> NxSuccess Then Exit Sub
    mAllowClose = True
    Unload Me
    Exit Sub
Failed: ShowFailure NxUserErrorText(Err.Description)
End Sub

Private Sub RequireSameTarget(ByVal target As Range)
    Dim current As Range
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "미리보기 이후 선택이 변경되었습니다."
    Set current = Application.Selection
    If Not (current.Worksheet Is mBoundSheet) Then NxRaiseContractError "미리보기 이후 시트가 변경되었습니다."
    If current.Address(External:=True) <> target.Address(External:=True) Then NxRaiseContractError "미리보기 이후 선택 범위가 변경되었습니다."
End Sub

Private Function ResolveTarget() As Range
    Dim target As Range
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set target = mRangeSession.ResolveRange
    If Not (target.Worksheet Is mBoundSheet) Then NxRaiseContractError "처음 연 시트 안에서 범위를 선택하세요."
    Set ResolveTarget = target
End Function

Private Sub RequireContext()
    mInit.RequireReady
    If mContext Is Nothing Then NxRaiseContractError "그림 대화상자 컨텍스트가 필요합니다."
    If Not (ActiveWorkbook Is mBoundBook) Then NxRaiseContractError "처음 연 통합문서로 돌아오세요."
    If Not (ActiveSheet Is mBoundSheet) Then NxRaiseContractError "처음 연 시트로 돌아오세요."
End Sub

Private Sub cboMode_Change()
    If Not ReadyForEvents() Then Exit Sub
    On Error GoTo Failed
    InvalidatePreview
    ApplyMode
    Exit Sub
Failed: ShowFailure NxUserErrorText(Err.Description)
End Sub

Private Sub cboSizeMode_Change()
    If Not ReadyForEvents() Then Exit Sub
    InvalidatePreview
End Sub

Private Sub txtMargin_Change()
    If Not ReadyForEvents() Then Exit Sub
    InvalidatePreview
End Sub

Private Sub chkMoveAndSize_Click()
    If Not ReadyForEvents() Then Exit Sub
    InvalidatePreview
End Sub

Private Sub txtRange_Change()
    If Not ReadyForEvents() Then Exit Sub
    InvalidatePreview
End Sub

Private Sub InvalidatePreview()
    mPreviewReady = False
    Set mRequest = Nothing: Set mPreview = Nothing: Set mFitRequest = Nothing
    cmdExecute.Enabled = True
    txtPreview.Value = "설정과 대상을 확인하고 실행하세요. 미리보기는 선택 사항입니다."
End Sub

Private Sub ShowFailure(ByVal detail As String)
    InvalidatePreview
    txtPreview.Value = "오류: " & detail
End Sub

Private Sub cmdCancel_Click()
    If Not ReadyForEvents() Then Exit Sub
    RequestCancel
End Sub

Private Sub RequestCancel()
    InvalidatePreview
    RestoreOriginalSelection
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    InvalidatePreview
    RestoreOriginalSelection
End Sub

Private Sub RestoreOriginalSelection()
    On Error Resume Next
    If Not mOriginalSelection Is Nothing Then
        mBoundBook.Activate
        mBoundSheet.Activate
        mOriginalSelection.Select
    End If
    On Error GoTo 0
End Sub

Private Function JoinPictureInsertFiles(ByVal source As Collection) As String
    Dim item As Variant, answer As String
    If source Is Nothing Then Exit Function
    For Each item In source
        If Len(answer) > 0 Then answer = answer & vbCrLf
        answer = answer & CStr(item)
    Next item
    JoinPictureInsertFiles = answer
End Function

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
