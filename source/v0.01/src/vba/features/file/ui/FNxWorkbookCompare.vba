Option Explicit

Private mPreviewReady As Boolean
Private mAllowClose As Boolean
Private mBusy As Boolean
Private mRequestedReport As String
Private mLastReport As String
Private mDestination As Workbook

Private Sub UserForm_Initialize()
    Dim service As Object
    cmdExecute.Default = True
    cmdCancel.Cancel = True
    Set mDestination = ActiveWorkbook
    cboOutput.Clear: cboOutput.AddItem "새 창(새 통합문서)": cboOutput.AddItem "현재 파일의 새 시트"
    cboOutput.ListIndex = 0
    ResetSheets
    txtPreview.BackColor = vbWhite
    Set service = NxHostCreateWorkbookCompareResults()
    cmdResults.Visible = Not service Is Nothing
End Sub

Private Sub ResetSheets()
    cboBaseSheet.Clear: cboBaseSheet.AddItem "전체 시트": cboBaseSheet.ListIndex = 0
    cboCompareSheet.Clear: cboCompareSheet.AddItem "전체 시트": cboCompareSheet.ListIndex = 0
End Sub

Private Sub cmdLoadSheets_Click()
    Dim book As Workbook, sheet As Worksheet, owned As Boolean, guard As CNxStateGuard, detail As String
    On Error GoTo Failed
    If mBusy Then Exit Sub
    ValidateInputs
    Set guard = New CNxStateGuard
    Application.EnableEvents = False: Application.ScreenUpdating = False
    ResetSheets
    Set book = NxSafeWorkbookOpen(CStr(txtBasePath.Value), owned)
    For Each sheet In book.Worksheets: cboBaseSheet.AddItem sheet.Name: Next sheet
    NxSafeWorkbookClose book, owned: Set book = Nothing
    owned = False
    Set book = NxSafeWorkbookOpen(CStr(txtComparePath.Value), owned)
    For Each sheet In book.Worksheets: cboCompareSheet.AddItem sheet.Name: Next sheet
    NxSafeWorkbookClose book, owned: Set book = Nothing
    guard.Restore
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then NxSafeWorkbookClose book, owned
    If Not guard Is Nothing Then guard.Restore
    txtPreview.Value = NxUserErrorText(detail)
End Sub

Public Sub BindFeature()
    InvalidatePreview
End Sub

Public Property Get RequestedReport() As String
    RequestedReport = mRequestedReport
End Property

Private Sub cmdResults_Click()
    Dim selected As Variant
    If mBusy Then Exit Sub
    If Len(mLastReport) > 0 Then
        mRequestedReport = mLastReport
    Else
        selected = Application.GetOpenFilename(FileFilter:="비교 보고서 (*.xlsx), *.xlsx", Title:="내엑셀 비교 보고서 선택")
        If VarType(selected) = vbBoolean Then Exit Sub
        mRequestedReport = CStr(selected)
    End If
    mAllowClose = True
    Me.Hide
End Sub

Private Sub cmdBaseBrowse_Click()
    txtBasePath.Value = PickWorkbook(CStr(txtBasePath.Value))
End Sub

Private Sub cmdCompareBrowse_Click()
    txtComparePath.Value = PickWorkbook(CStr(txtComparePath.Value))
End Sub



Private Function PickWorkbook(ByVal initialPath As String) As String
    Dim selectedPath As Variant
    selectedPath = Application.GetOpenFilename( _
        FileFilter:="Excel 통합문서 (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", _
        Title:="비교할 통합문서 선택")
    If VarType(selectedPath) = vbBoolean Then PickWorkbook = initialPath Else PickWorkbook = CStr(selectedPath)
End Function

Private Sub cmdPreview_Click()
    On Error GoTo Failed
    ValidateInputs
    txtPreview.Value = "두 파일을 읽기 전용으로 비교해 새 보고서를 만듭니다." & vbCrLf & _
        "비교: 값·수식·병합" & IIf(chkCompareFormats.Value, "·표시 형식", "") & vbCrLf & _
        "메모·도형·매크로는 비교하지 않습니다. 원본은 변경하지 않습니다." & vbCrLf & ComparisonSample()
    mPreviewReady = True
    cmdExecute.Enabled = True
    Exit Sub
Failed:
    ShowError
End Sub

Private Function ComparisonSample() As String
    Dim leftBook As Workbook, rightBook As Workbook, leftOwned As Boolean, rightOwned As Boolean
    Dim leftSheet As Worksheet, rightSheet As Worksheet, guard As CNxStateGuard
    Dim r As Long, c As Long, detail As String, a As String, b As String
    On Error GoTo Failed
    Set guard = New CNxStateGuard
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set leftBook = NxSafeWorkbookOpen(CStr(txtBasePath.Value), leftOwned)
    Set rightBook = NxSafeWorkbookOpen(CStr(txtComparePath.Value), rightOwned)
    If cboBaseSheet.ListIndex > 0 Then
        Set leftSheet = leftBook.Worksheets(CStr(cboBaseSheet.Value))
    Else
        Set leftSheet = leftBook.Worksheets(1)
    End If
    If cboCompareSheet.ListIndex > 0 Then
        Set rightSheet = rightBook.Worksheets(CStr(cboCompareSheet.Value))
    Else
        On Error Resume Next
        Set rightSheet = rightBook.Worksheets(leftSheet.Name)
        On Error GoTo Failed
        If rightSheet Is Nothing Then Set rightSheet = rightBook.Worksheets(1)
    End If
    ComparisonSample = "비교 표본 · " & leftSheet.Name & " / " & rightSheet.Name & " · A1:D6 표시 값"
    For r = 1 To 6
        For c = 1 To 4
            a = Replace$(Replace$(CStr(leftSheet.Cells(r, c).Text), vbCr, " "), vbLf, " ")
            b = Replace$(Replace$(CStr(rightSheet.Cells(r, c).Text), vbCr, " "), vbLf, " ")
            If Len(a) > 0 Or Len(b) > 0 Then ComparisonSample = ComparisonSample & vbCrLf & _
                leftSheet.Cells(r, c).Address(False, False) & ": " & Left$(a, 60) & " → " & Left$(b, 60)
        Next c
    Next r
CleanUp:
    On Error Resume Next
    If Not rightBook Is Nothing Then NxSafeWorkbookClose rightBook, rightOwned
    If Not leftBook Is Nothing Then NxSafeWorkbookClose leftBook, leftOwned
    If Not guard Is Nothing Then guard.Restore
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Function
Failed:
    detail = Err.Description
    GoTo CleanUp
End Function

Private Sub cmdExecute_Click()
    Dim options As Object, result As CNxResult
    On Error GoTo Failed
    ValidateInputs
    If mBusy Then Exit Sub
    If (cboBaseSheet.ListIndex = 0) Xor (cboCompareSheet.ListIndex = 0) Then NxRaiseContractError "양쪽 시트를 각각 선택하거나 모두 전체 시트로 선택하세요."
    If cboOutput.ListIndex = 1 Then NxCompareRequireDestination mDestination
    mBusy = True
    cmdResults.Enabled = False
    cmdExecute.Enabled = False: cmdPreview.Enabled = False
    NxWorkbookCompareBindProgress Me
    Set options = CreateObject("Scripting.Dictionary")
    options.Add "base_path", CStr(txtBasePath.Value)
    options.Add "compare_path", CStr(txtComparePath.Value)
    options.Add "compare_formats", CBool(chkCompareFormats.Value)
    If cboBaseSheet.ListIndex > 0 Then
        options.Add "base_sheet", CStr(cboBaseSheet.Value)
        options.Add "compare_sheet", CStr(cboCompareSheet.Value)
    End If
    Set result = NxFileRun("NX-FILE-FILE-COMPARE", vbNullString, workflowOptions:=options)
    If result Is Nothing Then NxRaiseContractError "파일 비교 결과를 확인하지 못했습니다."
    If result.Outcome <> NxSuccess Then NxRaiseContractError result.Recovery
    If cboOutput.ListIndex = 1 Then NxCompareAppendResultSheets ActiveWorkbook, mDestination
    mBusy = False
    mAllowClose = True
    Me.Hide
    Exit Sub
Failed:
    mBusy = False
    cmdResults.Enabled = True
    NxWorkbookCompareBindProgress Nothing
    cmdPreview.Enabled = True: cmdExecute.Enabled = True
    ShowError
End Sub

Public Sub UpdateCompareProgress(ByVal stage As String, ByVal done As Long, ByVal total As Long)
    If Not mBusy Then Exit Sub
    txtPreview.Value = stage & " · " & CStr(done) & " / " & CStr(total) & vbCrLf & "취소하면 미완성 보고서를 남기지 않습니다."
End Sub

Private Sub ValidateInputs()
    Dim basePath As String, comparePath As String, outputPath As String
    basePath = Trim$(CStr(txtBasePath.Value))
    comparePath = Trim$(CStr(txtComparePath.Value))
    If Len(Dir$(basePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then NxRaiseContractError "기준 파일을 찾을 수 없습니다."
    If Len(Dir$(comparePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then NxRaiseContractError "비교 파일을 찾을 수 없습니다."
    If StrComp(basePath, comparePath, vbTextCompare) = 0 Then NxRaiseContractError "서로 다른 두 파일을 선택하세요."
End Sub

Private Sub txtBasePath_Change(): ResetSheets: InvalidatePreview: End Sub
Private Sub txtComparePath_Change(): ResetSheets: InvalidatePreview: End Sub
Private Sub chkCompareFormats_Click(): InvalidatePreview: End Sub

Private Sub InvalidatePreview()
    mLastReport = vbNullString
    mPreviewReady = False
    cmdExecute.Enabled = True
End Sub

Private Sub cmdCancel_Click(): RequestCancel: End Sub

Private Sub RequestCancel()
    If mBusy Then
        NxWorkbookCompareCancel
        txtPreview.Value = "취소하고 있습니다. 잠시 기다려주세요."
        Exit Sub
    End If
    mAllowClose = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 파일 비교"
End Sub

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
