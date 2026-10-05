Option Explicit
Private mFeatureId As String
Private mLeft As Range, mRight As Range
Private mSheets As Collection, mOwned As Collection
Private mBusy As Boolean, mAllowClose As Boolean
Private mDestination As Workbook

Public Sub BindFeature(ByVal featureId As String)
    mFeatureId = featureId
    Set mDestination = ActiveWorkbook
    cboOutput.Clear
    cboOutput.AddItem "새 창(새 통합문서)": cboOutput.AddItem "현재 파일의 새 시트"
    cboOutput.ListIndex = 0
    Me.Caption = "내엑셀 - " & IIf(featureId = "NX-FILE-SHEET-COMPARE", "시트 비교", "범위 비교")
    Set mOwned = New Collection
    RefreshSheets
    cboLeft.Visible = (featureId = "NX-FILE-SHEET-COMPARE")
    cboRight.Visible = cboLeft.Visible
    txtLeft.Visible = Not cboLeft.Visible: txtRight.Visible = Not cboLeft.Visible
    cmdLeft.Visible = Not cboLeft.Visible: cmdRight.Visible = Not cboLeft.Visible
    If TypeName(Application.Selection) = "Range" Then
        Set mLeft = Application.Selection
        txtLeft.Value = mLeft.Address(External:=True)
    End If
    lblHelp.Caption = IIf(cboLeft.Visible, "같은 셀 주소끼리 비교합니다.", "두 범위의 왼쪽 위 셀을 맞춰 비교합니다.") & vbCrLf & "새 창이 기본입니다. 현재 파일의 새 시트도 선택할 수 있습니다. 차이는 연한 빨강으로 표시합니다."
    If cboLeft.Visible Then lblHelp.Caption = "같은 셀 주소끼리 비교합니다. 차이 목록의 링크로 기준·비교 셀에 이동합니다." & vbCrLf & "새 창이 기본이며 현재 파일의 새 시트도 선택할 수 있습니다."
End Sub

Private Sub UserForm_Initialize()
    cmdExecute.Default = True: cmdCancel.Cancel = True
End Sub

Private Sub RefreshSheets()
    Dim book As Workbook, sheet As Worksheet
    Set mSheets = New Collection
    cboLeft.Clear: cboRight.Clear
    For Each book In Application.Workbooks
        If Not book.IsAddin Then
            For Each sheet In book.Worksheets
                mSheets.Add sheet
                cboLeft.AddItem "[" & book.Name & "] " & sheet.Name
                cboRight.AddItem "[" & book.Name & "] " & sheet.Name
            Next sheet
        End If
    Next book
    If mSheets.Count > 0 Then
        cboLeft.ListIndex = 0: cboRight.ListIndex = IIf(mSheets.Count > 1, 1, 0)
    End If
End Sub

Private Sub cmdOpen_Click()
    Dim path As Variant, book As Workbook, owned As Boolean
    On Error GoTo Failed
    If mBusy Then Exit Sub
    path = Application.GetOpenFilename("Excel 통합문서 (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", , "비교할 파일 열기")
    If VarType(path) = vbBoolean Then Exit Sub
    Set book = NxSafeWorkbookOpen(CStr(path), owned)
    If owned Then mOwned.Add book
    RefreshSheets
    Exit Sub
Failed:
    MsgBox NxUserErrorText(Err.Description), vbExclamation, Me.Caption
End Sub

Private Sub cmdLeft_Click(): PickRange True: End Sub
Private Sub cmdRight_Click(): PickRange False: End Sub

Private Sub PickRange(ByVal isLeft As Boolean)
    Dim selected As Range
    If mBusy Then Exit Sub
    Me.Hide
    On Error Resume Next
    Set selected = Application.InputBox("비교할 연속 범위를 드래그하세요.", "내엑셀 - 범위 선택", Type:=8)
    On Error GoTo 0
    If Not selected Is Nothing Then
        If isLeft Then
            Set mLeft = selected: txtLeft.Value = selected.Address(External:=True)
        Else
            Set mRight = selected: txtRight.Value = selected.Address(External:=True)
        End If
    End If
    Me.Show vbModal
End Sub

Private Function ResolveRange(ByVal text As String, ByVal fallback As Range) As Range
    If Not fallback Is Nothing Then
        If StrComp(text, fallback.Address(External:=True), vbTextCompare) = 0 Then
            Set ResolveRange = fallback: Exit Function
        End If
    End If
    On Error Resume Next
    Set ResolveRange = Application.Range(text)
    On Error GoTo 0
    If ResolveRange Is Nothing Then NxRaiseContractError "범위 주소를 확인하거나 다시 선택하세요."
End Function

Private Sub cmdExecute_Click()
    Dim options As Object, result As CNxResult, sheet As Worksheet, detail As String
    On Error GoTo Failed
    If mBusy Then Exit Sub
    If mFeatureId = "NX-FILE-SHEET-COMPARE" Then
        If cboLeft.ListIndex < 0 Or cboRight.ListIndex < 0 Then NxRaiseContractError "두 시트를 선택하세요."
        Set sheet = mSheets(cboLeft.ListIndex + 1): Set mLeft = NxDataCompareUsedRange(sheet)
        Set sheet = mSheets(cboRight.ListIndex + 1): Set mRight = NxDataCompareUsedRange(sheet)
    Else
        Set mLeft = ResolveRange(CStr(txtLeft.Value), mLeft)
        Set mRight = ResolveRange(CStr(txtRight.Value), mRight)
    End If
    If cboOutput.ListIndex = 1 Then NxCompareRequireDestination mDestination
    mBusy = True: cmdExecute.Enabled = False: cmdOpen.Enabled = False
    NxDataCompareBind Me
    Set options = CreateObject("Scripting.Dictionary")
    options.Add "data_compare", True
    options.Add "left_range", mLeft: options.Add "right_range", mRight
    options.Add "compare_formulas", CBool(chkFormulas.Value)
    Set result = NxFileRun(mFeatureId, vbNullString, workflowOptions:=options)
    If result Is Nothing Then NxRaiseContractError "비교 결과를 확인하지 못했습니다."
    If result.Outcome <> NxSuccess Then NxRaiseContractError result.Recovery
    If cboOutput.ListIndex = 1 Then NxCompareAppendResultSheets ActiveWorkbook, mDestination
    mBusy = False: mAllowClose = True
    ReleaseInputs
    Me.Hide
    Exit Sub
Failed:
    detail = Err.Description
    mBusy = False: cmdExecute.Enabled = True: cmdOpen.Enabled = True
    NxDataCompareBind Nothing
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation, Me.Caption
End Sub

Public Sub UpdateCompareProgress(ByVal stage As String, ByVal done As Long, ByVal total As Long)
    lblStatus.Caption = stage & " · " & done & " / " & total
End Sub

Private Sub cmdCancel_Click()
    If mBusy Then
        NxDataCompareCancel
    Else
        mAllowClose = True: ReleaseInputs: Me.Hide
    End If
End Sub

Private Sub ReleaseInputs()
    Dim book As Workbook
    Set mLeft = Nothing: Set mRight = Nothing: Set mSheets = Nothing
    If Not mOwned Is Nothing Then
        For Each book In mOwned
            NxSafeWorkbookClose book, True
        Next book
    End If
    Set mOwned = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True: cmdCancel_Click
End Sub
