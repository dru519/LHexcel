Option Explicit
Private mFiles As Collection
Private mLoading As Boolean, mBusy As Boolean

Private Sub UserForm_Initialize()
    mLoading = True
    NxInitializeStaticComboChoices Me
    Set mFiles = New Collection
    cmdExecute.Default = True: cmdCancel.Cancel = True
    mLoading = False
    RefreshPreview
End Sub

Private Sub cmdFiles_Click()
    Dim picker As FileDialog, path As Variant
    On Error GoTo Failed
    Set picker = Application.FileDialog(msoFileDialogFilePicker)
    With picker
        .Title = "아래한글로 보낼 그림 선택"
        .AllowMultiSelect = True
        .Filters.Clear: .Filters.Add "그림 파일", "*.jpg;*.jpeg;*.png;*.bmp"
        If .Show <> -1 Then Exit Sub
        If .SelectedItems.Count > 200 Then NxRaiseContractError "최대 200개까지 선택하세요."
        Set mFiles = New Collection
        lstFiles.Clear
        For Each path In .SelectedItems
            mFiles.Add CStr(path)
            lstFiles.AddItem Mid$(CStr(path), InStrRev(CStr(path), "\") + 1)
        Next path
    End With
    RefreshPreview
    Exit Sub
Failed:
    MsgBox NxUserErrorText(Err.Description), vbExclamation, Me.Caption
End Sub

Private Function Payload() As String
    If Not IsNumeric(txtWidth.Value) Or Not IsNumeric(txtHeight.Value) Then NxRaiseContractError "크기는 cm 단위 숫자로 입력하세요."
    If Not IsNumeric(cboColumns.Value) Or Not IsNumeric(cboRows.Value) Then NxRaiseContractError "가로·세로 개수를 선택하세요."
    Payload = NxHangulPicturePayload(mFiles, CDbl(txtWidth.Value), CDbl(txtHeight.Value), _
        CLng(cboColumns.Value), CLng(cboRows.Value), CBool(chkTitles.Value), IIf(cboSort.ListIndex = 1, "taken", "title"), CBool(chkTitleFileNames.Value))
End Function

Private Sub RefreshPreview()
    Dim r As Long, c As Long, item As Object, key As String, n As Long
    Dim columns As Long, rows As Long, cw As Single, ch As Single, titleHeight As Single, unused As String
    If mLoading Or mBusy Then Exit Sub
    chkTitleFileNames.Enabled = CBool(chkTitles.Value)
    On Error GoTo Invalid
    For Each item In Me.Controls
        If Left$(item.Name, 8) = "nxPhoto_" Then item.Visible = False
    Next item
    columns = CLng(cboColumns.Value): rows = CLng(cboRows.Value)
    If columns < 1 Or rows < 1 Then Exit Sub
    cw = 204 / columns: ch = 114 / rows
    If chkTitles.Value Then titleHeight = 12
    For r = 1 To rows
        For c = 1 To columns
            n = n + 1: key = "nxPhoto_" & n
            On Error Resume Next
            Set item = Nothing: Set item = Me.Controls(key)
            On Error GoTo Invalid
            If item Is Nothing Then Set item = Me.Controls.Add("Forms.Label.1", key, True)
            With item
                .Left = 252 + (c - 1) * cw: .Top = 100 + (r - 1) * ch
                .Width = cw - 2: .Height = ch - 2
                .BackColor = RGB(237, 245, 249): .BorderStyle = 1: .BorderColor = RGB(180, 190, 195)
                .Font.Name = "맑은 고딕": .Font.Size = 9: .TextAlign = 2
                .Caption = "그림 " & n & IIf(chkTitles.Value, vbCrLf & IIf(chkTitleFileNames.Value, "파일명", "빈 제목칸"), vbNullString)
                .Visible = True
            End With
        Next c
    Next r
    lblPreview.Caption = "한 쪽 배치 · " & txtWidth.Value & " × " & txtHeight.Value & " cm / 셀"
    lblStatus.Caption = "선택 " & mFiles.Count & "개 · " & cboSort.Value & " · 셀 배경에 문서 포함"
    If mFiles.Count > 0 Then unused = Payload()
    cmdExecute.Enabled = mFiles.Count > 0
    Exit Sub
Invalid:
    lblStatus.Caption = NxUserErrorText(Err.Description): Err.Clear
    cmdExecute.Enabled = False
End Sub

Private Sub cmdExecute_Click()
    Dim text As String, outputPath As String, detail As String
    On Error GoTo Failed
    text = Payload()
    mBusy = True: cmdExecute.Enabled = False: cmdFiles.Enabled = False: cmdCancel.Enabled = False
    lblStatus.Caption = "그림을 문서에 포함하고 있습니다."
    Me.Repaint
    outputPath = NxHangulSendPictures(text)
    mBusy = False
    Me.Hide
    Exit Sub
Failed:
    detail = Err.Description
    mBusy = False: cmdExecute.Enabled = True: cmdFiles.Enabled = True: cmdCancel.Enabled = True
    lblStatus.Caption = "설정을 확인하고 다시 실행하세요."
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation, Me.Caption
End Sub
Private Sub txtWidth_Change(): RefreshPreview: End Sub
Private Sub txtHeight_Change(): RefreshPreview: End Sub
Private Sub cboColumns_Change(): RefreshPreview: End Sub
Private Sub cboRows_Change(): RefreshPreview: End Sub
Private Sub cboSort_Change(): RefreshPreview: End Sub
Private Sub chkTitles_Click(): RefreshPreview: End Sub
Private Sub chkTitleFileNames_Click(): RefreshPreview: End Sub
Private Sub cmdCancel_Click(): If Not mBusy Then Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mBusy Then Cancel = True
End Sub
