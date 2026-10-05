Option Explicit

Private mInit As CNxFormInitializationGuard
Private mContext As CNxFeatureDialogContext
Private mRangeSession As CNxRangeSelectionSession
Private mPreviewAddress As String
Private mPreviewOptions As String
Private mPreviewReady As Boolean
Private mAllowClose As Boolean
Private mUpdating As Boolean

Private Sub UserForm_Initialize()
    Set mInit = New CNxFormInitializationGuard
    mInit.BeginInitialization
    NxInitializeStaticComboChoices Me
    ApplyDefaults
    mInit.CompleteInitialization
End Sub

Public Sub BindFeatureContext(ByVal context As CNxFeatureDialogContext)
    If context Is Nothing Then NxRaiseContractError "Draw table context is required"
    If Not context.IsConfigured Then NxRaiseContractError "Draw table context is required"
    If Not mContext Is Nothing Then NxRaiseContractError "Draw table context can be bound only once"
    Select Case context.FeatureId
        Case NX_FEATURE_DRAW_TITLE_TABLE, NX_FEATURE_DRAW_BUSINESS_TABLE
        Case Else: NxRaiseContractError "Draw table feature is outside the closed contract"
    End Select
    Set mContext = context
    Set mRangeSession = NxRangeSessionFromSelection()
    txtRange.Value = mRangeSession.DisplayAddress
    ConfigureStyleChoices
    RefreshPreview
End Sub

Private Sub ConfigureStyleChoices()
    Dim titleMode As Boolean
    mUpdating = True
    titleMode = (mContext.FeatureId = NX_FEATURE_DRAW_TITLE_TABLE)
    cboTableStyle.Clear
    Me.Caption = "내엑셀 - 표 그리기"
    If mContext.FeatureId = NX_FEATURE_DRAW_TITLE_TABLE Then
        cboTableStyle.AddItem "스타일: 제목"
        cboTableStyle.AddItem "스타일: 소제목"
    Else
        cboTableStyle.AddItem "스타일: 닫힌 표"
        cboTableStyle.AddItem "스타일: 열린 표"
    End If
    cboTableStyle.ListIndex = IIf(titleMode, 0, 1)
    cboHeaderSize.Value = IIf(titleMode, "머리글 크기: 13", "머리글 크기: 9")
    cboHeaderAlign.ListIndex = IIf(titleMode, 0, 1)
    cboHeaderRows.Visible = Not titleMode
    lblHeaderRows.Visible = Not titleMode
    cboBodySize.Visible = Not titleMode
    cboBodyAlign.Visible = Not titleMode
    lblBodySize.Visible = Not titleMode
    lblBodyAlign.Visible = Not titleMode
    cboFinishing.Visible = Not titleMode
    lblFinishing.Visible = Not titleMode
    lblFinishingHelp.Visible = Not titleMode
    lblStyle.Caption = IIf(titleMode, "제목", "표 유형")
    Me.Caption = DialogCaption()
    mUpdating = False
End Sub

Private Sub ApplyDefaults()
    cmdExecute.Default = True
    cmdCancel.Cancel = True
    cmdExecute.Enabled = True
End Sub

Private Sub cboDisplayMode_Change(): OptionsChanged: End Sub
Private Sub cboTableStyle_Change(): OptionsChanged: End Sub
Private Sub cboHeaderRows_Change(): OptionsChanged: End Sub
Private Sub cboHeaderSize_Change(): OptionsChanged: End Sub
Private Sub cboHeaderAlign_Change(): OptionsChanged: End Sub
Private Sub cboBodySize_Change(): OptionsChanged: End Sub
Private Sub cboBodyAlign_Change(): OptionsChanged: End Sub
Private Sub cboFinishing_Change(): OptionsChanged: End Sub

Private Sub cmdResetDefaults_Click()
    If mInit.IsInitializing Then Exit Sub
    On Error GoTo Failed
    RequireContext
    ConfigureStyleChoices
    mUpdating = True
    cboDisplayMode.ListIndex = 0
    cboHeaderRows.ListIndex = 0
    cboBodySize.Value = "본문 크기: 9"
    cboBodyAlign.ListIndex = 0
    cboFinishing.ListIndex = 0
    mUpdating = False
    InvalidatePreview
    RefreshPreview
    Exit Sub
Failed:
    mUpdating = False
    InvalidatePreview
    ShowError
End Sub

Private Sub OptionsChanged()
    If mInit Is Nothing Then Exit Sub
    If mInit.IsInitializing Then Exit Sub
    If mUpdating Then Exit Sub
    ' Native combo controls may raise Change more than once for the same choice.
    ' Explicit Preview always re-reads cells; only identical option events are skipped.
    If mPreviewReady Then
        If mPreviewOptions = OptionFingerprint() Then Exit Sub
    End If
    InvalidatePreview False
    RefreshPreview
End Sub

Private Sub txtRange_Change()
    If mInit Is Nothing Then Exit Sub
    If mInit.IsInitializing Or mUpdating Then Exit Sub
    InvalidatePreview
    lblPreviewStatus.Caption = "범위 입력을 마치면 미리보기를 갱신합니다."
End Sub

Private Sub txtRange_AfterUpdate()
    RefreshPreview
End Sub

Private Sub cmdPickRange_Click()
    If mInit.IsInitializing Then Exit Sub
    On Error GoTo Failed
    RequireContext
    If NxPickRange(Me, mRangeSession) Then
        txtRange.Value = mRangeSession.DisplayAddress
        RefreshPreview
    End If
    Exit Sub
Failed:
    InvalidatePreview
    ShowError
End Sub

Private Sub cmdPreview_Click()
    RefreshPreview
End Sub

Private Sub RefreshPreview()
    Dim source As Range
    If mContext Is Nothing Or mUpdating Then Exit Sub
    If mInit.IsInitializing Then Exit Sub
    On Error GoTo Failed
    RequireContext
    UpdateHeaderControls
    UpdateFinishingHelp
    UpdateRangeFromText
    Set source = mRangeSession.ResolveRange()
    NxDrawValidateSelection source
    If HeaderRowCount() > source.Rows.Count Then NxRaiseContractError "머리글 행 수가 선택 범위보다 많습니다."
    mPreviewAddress = source.Address(External:=True)
    mPreviewOptions = OptionFingerprint()
    RenderPreview source
    Me.Caption = DialogCaption()
    lblPreviewStatus.Caption = source.Address(False, False) & " · " & CStr(source.CountLarge) & "셀 · 원본 변경 없음"
    If HeaderRowCount() = 0 Then lblPreviewStatus.Caption = lblPreviewStatus.Caption & vbCrLf & "제목행 없음 · 첫 행부터 본문 서식 적용"
    If AutoFitColumns() Then lblPreviewStatus.Caption = lblPreviewStatus.Caption & vbCrLf & "열맞춤 너비는 표본 추정 · 실제 적용은 선택 범위 내용 기준"
    mPreviewReady = True
    cmdExecute.Enabled = True
    Exit Sub
Failed:
    InvalidatePreview
    lblPreviewStatus.Caption = NxUserErrorText(Err.Description)
    Err.Clear
End Sub

Private Sub UpdateHeaderControls()
    Dim enabled As Boolean
    enabled = (HeaderRowCount() > 0)
    cboHeaderSize.Enabled = enabled
    cboHeaderAlign.Enabled = enabled
    lblHeaderSize.Enabled = enabled
    lblHeaderAlign.Enabled = enabled
End Sub

Private Sub cmdExecute_Click()
    Dim selected As Range
    Dim result As CNxResult
    If mInit.IsInitializing Then Exit Sub
    On Error GoTo Failed
    RequireContext
    UpdateRangeFromText
    Set selected = mRangeSession.ResolveRange()
    If HeaderRowCount() > selected.Rows.Count Then NxRaiseContractError "머리글 행 수가 선택 범위보다 많습니다."
    NxDrawValidateSelection selected
    ' ResolveRange already verifies the original workbook/sheet identity.
    ' Bind the typed/picked range to the router's current-selection context.
    selected.Select
    Set result = NxDrawRun(mContext.FeatureId, selected, IncludeTotalRow(), _
        (BodyAlignmentValue() = xlGeneral), (HeaderRowCount() > 0), AutoFitColumns(), _
        DisplayModeValue(), TableStyleValue(), HeaderRowCount(), HeaderFontSizeValue(), _
        HeaderAlignmentValue(), BodyFontSizeValue(), BodyAlignmentValue())
    If result Is Nothing Then NxRaiseContractError "표 그리기 결과를 확인할 수 없습니다."
    If result.Outcome <> NxSuccess And result.Outcome <> NxCancelled Then NxRaiseContractError result.Recovery
    If result.Outcome = NxSuccess Then RequestCancel False
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdCancel_Click()
    If mInit.IsInitializing Then Exit Sub
    RequestCancel
End Sub

Private Function DisplayModeValue() As String
    If InStr(1, CStr(cboDisplayMode.Value), "색상", vbBinaryCompare) > 0 Then DisplayModeValue = NX_ROLE_STYLE_MODE_COLOR Else DisplayModeValue = NX_ROLE_STYLE_MODE_MONO
End Function

Private Function TableStyleValue() As String
    If InStr(1, CStr(cboTableStyle.Value), "소제목", vbBinaryCompare) > 0 Then
        TableStyleValue = "부제목"
    ElseIf InStr(1, CStr(cboTableStyle.Value), "제목", vbBinaryCompare) > 0 Then
        TableStyleValue = "제목"
    ElseIf InStr(1, CStr(cboTableStyle.Value), "열린", vbBinaryCompare) > 0 Then
        TableStyleValue = "열린 표"
    Else
        TableStyleValue = "닫힌 표"
    End If
End Function

Private Function HeaderRowCount() As Long
    If mContext.FeatureId = NX_FEATURE_DRAW_TITLE_TABLE Then HeaderRowCount = 1: Exit Function
    If CStr(cboHeaderRows.Value) = "없음 (0행)" Then HeaderRowCount = 0: Exit Function
    HeaderRowCount = CLng(OptionNumber(CStr(cboHeaderRows.Value)))
    If HeaderRowCount < 1 Then NxRaiseContractError "머리글 행 수를 선택하세요."
End Function

Private Function HeaderAlignmentValue() As Long
    If InStr(1, CStr(cboHeaderAlign.Value), "왼쪽", vbBinaryCompare) > 0 Then HeaderAlignmentValue = xlLeft: Exit Function
    If InStr(1, CStr(cboHeaderAlign.Value), "가운데", vbBinaryCompare) > 0 Then HeaderAlignmentValue = xlCenter: Exit Function
    If InStr(1, CStr(cboHeaderAlign.Value), "오른쪽", vbBinaryCompare) > 0 Then HeaderAlignmentValue = xlRight: Exit Function
    NxRaiseContractError "머리글 정렬을 선택하세요."
End Function

Private Function BodyAlignmentValue() As Long
    If InStr(1, CStr(cboBodyAlign.Value), "자동", vbBinaryCompare) > 0 Then BodyAlignmentValue = xlGeneral: Exit Function
    If InStr(1, CStr(cboBodyAlign.Value), "왼쪽", vbBinaryCompare) > 0 Then BodyAlignmentValue = xlLeft: Exit Function
    If InStr(1, CStr(cboBodyAlign.Value), "가운데", vbBinaryCompare) > 0 Then BodyAlignmentValue = xlCenter: Exit Function
    If InStr(1, CStr(cboBodyAlign.Value), "오른쪽", vbBinaryCompare) > 0 Then BodyAlignmentValue = xlRight: Exit Function
    NxRaiseContractError "본문 정렬을 선택하세요."
End Function

Private Function HeaderFontSizeValue() As Double
    HeaderFontSizeValue = OptionNumber(CStr(cboHeaderSize.Value))
End Function

Private Function BodyFontSizeValue() As Double
    BodyFontSizeValue = OptionNumber(CStr(cboBodySize.Value))
End Function

Private Function OptionNumber(ByVal optionText As String) As Double
    Dim separator As Long
    separator = InStrRev(optionText, ":")
    If separator = 0 Or Not IsNumeric(Trim$(Mid$(optionText, separator + 1))) Then _
        NxRaiseContractError "표 옵션 숫자 값을 확인하세요."
    OptionNumber = CDbl(Trim$(Mid$(optionText, separator + 1)))
End Function

Private Function IncludeTotalRow() As Boolean
    IncludeTotalRow = (InStr(1, CStr(cboFinishing.Value), "합계행", vbBinaryCompare) > 0)
End Function

Private Function AutoFitColumns() As Boolean
    AutoFitColumns = (InStr(1, CStr(cboFinishing.Value), "열맞춤", vbBinaryCompare) > 0)
End Function

Private Sub UpdateFinishingHelp()
    Dim detail As String
    detail = "기본: 선택 범위 서식만 적용 · 값·수식 유지"
    If IncludeTotalRow() Then detail = "합계행: 마지막 기존 행을 굵게·음영·윗선 강조 (행 추가·SUM 계산 없음)"
    If AutoFitColumns() Then detail = detail & vbCrLf & "열맞춤: 선택 범위 내용 기준 · 너비는 해당 열 전체에 적용"
    lblFinishingHelp.Caption = detail
End Sub

Private Function OptionFingerprint() As String
    OptionFingerprint = DisplayModeValue() & "|" & TableStyleValue() & "|" & _
        CStr(HeaderRowCount()) & "|" & CStr(HeaderFontSizeValue()) & "|" & CStr(HeaderAlignmentValue()) & "|" & _
        CStr(BodyFontSizeValue()) & "|" & CStr(BodyAlignmentValue()) & "|" & CStr(IncludeTotalRow()) & "|" & _
        CStr(AutoFitColumns())
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    RequestCancel
End Sub

Private Sub RequireContext()
    mInit.RequireReady
    If mContext Is Nothing Then NxRaiseContractError "Draw table context must be bound before use"
End Sub

Private Sub InvalidatePreview(Optional ByVal hideCells As Boolean = True)
    mPreviewReady = False
    mPreviewAddress = vbNullString
    mPreviewOptions = vbNullString
    cmdExecute.Enabled = True
    If hideCells Then ClearPreview
    If Not mContext Is Nothing Then Me.Caption = DialogCaption()
End Sub

Private Function DialogCaption() As String
    If mContext.FeatureId = NX_FEATURE_DRAW_TITLE_TABLE Then
        DialogCaption = "내엑셀 - 타이틀표 그리기"
    Else
        DialogCaption = "내엑셀 - 표 그리기"
    End If
End Function

Private Sub UpdateRangeFromText()
    Dim current As Range, selected As Range, address As String, prefix As String
    Dim external As String, splitAt As Long, parts As Variant, part As Variant
    Dim i As Long, column As Long, row As Double, ch As String
    Set current = mRangeSession.ResolveRange()
    address = Trim$(CStr(txtRange.Value))
    splitAt = InStrRev(address, "!")
    If splitAt > 0 Then
        prefix = Left$(address, splitAt - 1)
        external = current.Address(External:=True)
        external = Left$(external, InStrRev(external, "!") - 1)
        If prefix <> current.Worksheet.Name And _
            prefix <> "'" & Replace$(current.Worksheet.Name, "'", "''") & "'" And _
            prefix <> external Then NxRaiseContractError "범위 선택 버튼으로 대상 시트를 다시 선택하세요."
        address = Mid$(address, splitAt + 1)
    End If
    address = UCase$(Replace$(address, "$", vbNullString))
    parts = Split(address, ":")
    If UBound(parts) > 1 Then NxRaiseContractError "A1 또는 A1:B5 형식으로 입력하세요."
    For Each part In parts
        column = 0: i = 1
        Do While i <= Len(part)
            ch = Mid$(part, i, 1)
            If ch < "A" Or ch > "Z" Then Exit Do
            column = column * 26 + Asc(ch) - Asc("A") + 1
            If column > 16384 Then NxRaiseContractError "열 주소가 Excel 범위를 벗어났습니다."
            i = i + 1
        Loop
        If column = 0 Or i > Len(part) Then NxRaiseContractError "A1 또는 A1:B5 형식으로 입력하세요."
        row = 0
        Do While i <= Len(part)
            ch = Mid$(part, i, 1)
            If ch < "0" Or ch > "9" Then NxRaiseContractError "A1 또는 A1:B5 형식으로 입력하세요."
            row = row * 10 + Asc(ch) - Asc("0")
            If row > 1048576 Then NxRaiseContractError "행 주소가 Excel 범위를 벗어났습니다."
            i = i + 1
        Loop
        If row < 1 Then NxRaiseContractError "행 번호는 1 이상이어야 합니다."
    Next part
    Set selected = current.Worksheet.Range(address)
    mRangeSession.UpdateFromRange selected
End Sub

Private Sub ClearPreview()
    NxNativePreviewHide Me
    Dim index As Long
    For index = Me.Controls.Count - 1 To 0 Step -1
        If Left$(Me.Controls(index).Name, 6) = "nxPrv_" Then Me.Controls(index).Visible = False
    Next index
End Sub

Private Sub RenderPreview(ByVal source As Range)
    Dim r As Long, c As Long, rows As Long, cols As Long, sourceRow As Long
    Dim item As MSForms.Label, cell As Range, titleMode As Boolean, previewFont As StdFont
    Dim header As Boolean, total As Boolean, alignment As Long, x As Single
    Dim hasMergedCells As Boolean, merged As Range
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    Dim heights(1 To 5) As Single, xs(0 To 5) As Single, ys(0 To 5) As Single
    Dim heightSum As Single, previewScale As Single
    Dim widths(1 To 5) As Single, widthSum As Single, weight As Single
    Const previewTop As Single = 288
    Const previewLeft As Single = 46
    Const previewWidth As Single = 366
    ClearPreview
    rows = Application.Min(5, source.Rows.Count)
    cols = Application.Min(5, source.Columns.Count)
    titleMode = (mContext.FeatureId = NX_FEATURE_DRAW_TITLE_TABLE)
    hasMergedCells = PreviewHasMerge(source.Cells(1, 1).Resize(rows, cols))
    If IncludeTotalRow() And source.Rows.Count > rows Then _
        hasMergedCells = hasMergedCells Or PreviewHasMerge(source.Cells(source.Rows.Count, 1).Resize(1, cols))
    For c = 1 To cols
        widths(c) = Application.Max(1, source.Columns(c).Width)
        If AutoFitColumns() Then
            widths(c) = 1
            For r = 1 To rows
                sourceRow = PreviewSourceRow(source, r, rows)
                weight = Len(CStr(source.Cells(sourceRow, c).Text)) + 2
                weight = weight * IIf(titleMode Or sourceRow <= HeaderRowCount(), HeaderFontSizeValue(), BodyFontSizeValue())
                widths(c) = Application.Max(widths(c), weight)
            Next r
        End If
        widthSum = widthSum + widths(c)
    Next c
    previewScale = Application.Min(1, previewWidth / widthSum)
    xs(0) = previewLeft
    For c = 1 To cols
        widths(c) = previewWidth * widths(c) / widthSum
        xs(c) = xs(c - 1) + widths(c)
    Next c
    Do
        heightSum = PreviewRowHeights(source, rows, cols, widths, heights, previewScale)
        If heightSum <= 84 Or rows = 1 Then Exit Do
        rows = rows - 1
    Loop
    If heightSum > 84 Then heights(1) = 84
    ys(0) = previewTop
    For r = 1 To rows
        ys(r) = ys(r - 1) + heights(r)
    Next r
    For c = 1 To cols
        Set item = PreviewLabel("column_" & c)
        item.Left = xs(c - 1): item.Top = previewTop - 16
        item.Width = widths(c): item.Height = 16
        item.Caption = Split(source.Cells(1, c).Address(True, False), "$")(0)
        item.BackColor = RGB(240, 240, 240): item.ForeColor = RGB(70, 70, 70)
        item.TextAlign = 2: item.Font.Size = 8
    Next c
    For r = 1 To rows
        Set item = PreviewLabel("row_" & r)
        item.Left = 18: item.Top = ys(r - 1): item.Width = 27: item.Height = heights(r)
        item.Caption = CStr(source.Row + PreviewSourceRow(source, r, rows) - 1)
        item.BackColor = RGB(240, 240, 240): item.ForeColor = RGB(70, 70, 70)
        item.TextAlign = 2: item.Font.Size = 8
    Next r
    lblPreview.Caption = "미리보기 · 처음 " & rows & "행 × " & cols & "열 · 현재 열 너비 비율"
    If PreviewSourceRow(source, rows, rows) <> rows Then _
        lblPreview.Caption = "미리보기 · 처음 " & (rows - 1) & "행 + 마지막 " & source.Rows.Count & "행 · " & cols & "열"
    If AutoFitColumns() Then lblPreview.Caption = Replace$(lblPreview.Caption, "현재 열 너비 비율", "열 너비 표본 추정")
    lblPreview.ControlTipText = "병합·줄바꿈을 반영한 일부 표본입니다. 표본 밖 병합 영역은 경계에서 잘립니다. 글자에 마우스를 올리면 원문을 확인할 수 있습니다."
    For r = 1 To rows
        x = previewLeft
        sourceRow = PreviewSourceRow(source, r, rows)
        For c = 1 To cols
            Set cell = source.Cells(sourceRow, c)
            Set item = PreviewLabel("cell_" & r & "_" & c)
            header = titleMode Or sourceRow <= HeaderRowCount()
            total = Not titleMode And IncludeTotalRow() And Not header And sourceRow = source.Rows.Count
            With item
                .Left = xs(c - 1)
                .Top = ys(r - 1)
                .Width = widths(c)
                .Height = heights(r)
                If cell.MergeCells Then
                    Set merged = cell.MergeArea
                    firstRow = Application.Max(1, merged.Row - source.Row + 1)
                    lastRow = Application.Min(rows, merged.Row - source.Row + merged.Rows.Count)
                    firstCol = Application.Max(1, merged.Column - source.Column + 1)
                    lastCol = Application.Min(cols, merged.Column - source.Column + merged.Columns.Count)
                    .Visible = (r = firstRow And c = firstCol)
                    If .Visible Then
                        .Width = xs(lastCol) - xs(firstCol - 1)
                        .Height = ys(lastRow) - ys(firstRow - 1)
                        Set cell = merged.Cells(1, 1)
                    End If
                End If
                .Caption = CStr(cell.Text)
                .Tag = cell.Address(False, False)
                .WordWrap = CBool(cell.WrapText)
                .ControlTipText = CStr(cell.Text)
                If Not .Visible Then .Caption = vbNullString
                Set previewFont = New StdFont
                previewFont.Name = "맑은 고딕"
                previewFont.Size = IIf(header, HeaderFontSizeValue(), BodyFontSizeValue())
                previewFont.Bold = total Or (header And (Not titleMode Or TableStyleValue() = "제목"))
                previewFont.Italic = cell.Font.Italic
                previewFont.Underline = (cell.Font.Underline <> xlUnderlineStyleNone)
                previewFont.Strikethrough = cell.Font.Strikethrough
                Set .Font = previewFont
                .BackColor = RGB(255, 255, 255)
                .ForeColor = RGB(32, 32, 32)
                If DisplayModeValue() = NX_ROLE_STYLE_MODE_COLOR Then
                    .ForeColor = IIf(header, NxHeaderFontColor(), NxBodyFontColor())
                    If titleMode Then .BackColor = NxHeaderFillColor()
                End If
                If header And Not titleMode Then
                    .BackColor = NxDrawTableHeaderFill(DisplayModeValue())
                End If
                If total Then
                    .BackColor = RGB(242, 242, 242)
                    If DisplayModeValue() = NX_ROLE_STYLE_MODE_COLOR Then .BackColor = NxTotalFillColor()
                End If
                alignment = IIf(header, HeaderAlignmentValue(), BodyAlignmentValue())
                If alignment = xlGeneral Then
                    alignment = cell.HorizontalAlignment
                    If alignment = xlGeneral Then
                        alignment = xlLeft
                        If IsError(cell.Value2) Or VarType(cell.Value2) = vbBoolean Then
                            alignment = xlCenter
                        ElseIf VarType(cell.Value2) <> vbString And Not IsEmpty(cell.Value2) Then
                            If IsNumeric(cell.Value2) Then alignment = xlRight
                        End If
                    End If
                End If
                .TextAlign = IIf(alignment = xlLeft, 1, IIf(alignment = xlRight, 3, 2))
            End With
            x = x + widths(c)
        Next c
    Next r
    If Not titleMode Then
        PreviewLine "top", previewLeft, previewTop, previewWidth, 1
        x = previewLeft
        For c = 1 To cols - 1
            x = x + widths(c)
            If Not hasMergedCells Then
                PreviewLine "v" & c, x, previewTop, 1, ys(rows) - previewTop
            Else
                For r = 1 To rows
                    sourceRow = PreviewSourceRow(source, r, rows)
                    If Not PreviewSharesMerge(source.Cells(sourceRow, c), source.Cells(sourceRow, c + 1)) Then _
                        PreviewLine "v" & c & "_" & r, x, ys(r - 1), 1, heights(r)
                Next r
            End If
        Next c
        For r = 1 To rows - 1
            If Not hasMergedCells Then
                PreviewLine "h" & r, previewLeft, ys(r), previewWidth, IIf(r = HeaderRowCount(), 3, 1)
            Else
                For c = 1 To cols
                    If Not PreviewSharesMerge(source.Cells(PreviewSourceRow(source, r, rows), c), source.Cells(PreviewSourceRow(source, r + 1, rows), c)) Then _
                        PreviewLine "h" & r & "_" & c, xs(c - 1), ys(r), widths(c), IIf(r = HeaderRowCount(), 3, 1)
                Next c
            End If
        Next r
        If TableStyleValue() = "닫힌 표" Then
            PreviewLine "left", previewLeft, previewTop, 1, ys(rows) - previewTop
            PreviewLine "right", previewLeft + previewWidth, previewTop, 1, ys(rows) - previewTop
        End If
        If IncludeTotalRow() And source.Rows.Count > HeaderRowCount() + 1 And PreviewSourceRow(source, rows, rows) = source.Rows.Count Then _
            PreviewLine "total", previewLeft, ys(rows - 1), previewWidth, 2, True
    End If
    PreviewLine "bottom", previewLeft, ys(rows), previewWidth, IIf(titleMode, IIf(TableStyleValue() = "제목", 2, 1), 3), titleMode
    NxNativePreviewLabels Me, "nxPrv_", 18, 270, 394, 104
End Sub

Private Function PreviewHasMerge(ByVal source As Range) As Boolean
    Dim value As Variant
    value = source.MergeCells
    If IsNull(value) Then PreviewHasMerge = True Else PreviewHasMerge = CBool(value)
End Function

Private Function PreviewSharesMerge(ByVal leftCell As Range, ByVal rightCell As Range) As Boolean
    If Not leftCell.MergeCells Or Not rightCell.MergeCells Then Exit Function
    PreviewSharesMerge = (leftCell.MergeArea.Address = rightCell.MergeArea.Address)
End Function

Private Function PreviewRowHeights(ByVal source As Range, ByVal rows As Long, ByVal cols As Long, _
    ByRef widths() As Single, ByRef heights() As Single, ByVal previewScale As Single) As Single
    Dim r As Long, c As Long, k As Long, sourceRow As Long, cell As Range
    Dim fontSize As Single, available As Single, lines As Long, part As Variant, required As Single
    For r = 1 To rows
        sourceRow = PreviewSourceRow(source, r, rows)
        heights(r) = Application.Max(20, source.Rows(sourceRow).RowHeight * previewScale)
        fontSize = IIf(mContext.FeatureId = NX_FEATURE_DRAW_TITLE_TABLE Or sourceRow <= HeaderRowCount(), HeaderFontSizeValue(), BodyFontSizeValue())
        For c = 1 To cols
            Set cell = source.Cells(sourceRow, c)
            available = widths(c)
            If cell.MergeCells Then
                If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then GoTo NextCell
                For k = c + 1 To Application.Min(cols, c + cell.MergeArea.Columns.Count - 1)
                    available = available + widths(k)
                Next k
            End If
            lines = 0
            For Each part In Split(Replace$(CStr(cell.Text), vbCr, vbNullString), vbLf)
                If cell.WrapText Then
                    ' Conservative mixed Korean/Latin estimate. Native labels perform actual wrapping.
                    lines = lines + Application.Max(1, -Int(-PreviewTextWidth(CStr(part), fontSize) / Application.Max(1, available - 6)))
                Else
                    lines = lines + 1
                End If
            Next part
            required = lines * fontSize * 1.4 + 6
            If cell.MergeCells Then required = required / cell.MergeArea.Rows.Count
            heights(r) = Application.Max(heights(r), required)
NextCell:
        Next c
        PreviewRowHeights = PreviewRowHeights + heights(r)
    Next r
End Function

Private Function PreviewTextWidth(ByVal value As String, ByVal fontSize As Single) As Single
    Dim index As Long, unit As Long
    For index = 1 To Len(value)
        unit = AscW(Mid$(value, index, 1))
        PreviewTextWidth = PreviewTextWidth + fontSize * IIf(unit >= 0 And unit < 128, 0.6, 1)
    Next index
End Function

Private Function PreviewSourceRow(ByVal source As Range, ByVal previewRow As Long, ByVal previewRows As Long) As Long
    PreviewSourceRow = previewRow
    If mContext.FeatureId = NX_FEATURE_DRAW_TITLE_TABLE Then Exit Function
    ' Only the final sample row can represent the total. Avoid querying merge
    ' geometry for every ordinary sample cell and every option refresh.
    If Not IncludeTotalRow() Then Exit Function
    If previewRow <> previewRows Then Exit Function
    If source.Rows.Count <= previewRows Then Exit Function
    ' Keep merged geometry contiguous; never join two separated pieces of one merge.
    If previewRows < 2 Then Exit Function
    If PreviewHasMerge(source.Cells(1, 1).Resize(previewRows - 1, Application.Min(5, source.Columns.Count))) Then Exit Function
    If PreviewHasMerge(source.Cells(source.Rows.Count, 1).Resize(1, Application.Min(5, source.Columns.Count))) Then Exit Function
    If IncludeTotalRow() And source.Rows.Count > HeaderRowCount() Then
        If source.Rows.Count > previewRows And previewRow = previewRows Then PreviewSourceRow = source.Rows.Count
    End If
End Function

Private Function PreviewLabel(ByVal key As String) As MSForms.Label
    Dim item As MSForms.Label
    On Error Resume Next
    Set item = Me.Controls("nxPrv_" & key)
    On Error GoTo 0
    If item Is Nothing Then Set item = Me.Controls.Add("Forms.Label.1", "nxPrv_" & key, True)
    item.Visible = True
    Set PreviewLabel = item
End Function

Private Sub PreviewLine(ByVal key As String, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, Optional ByVal emphasis As Boolean = False)
    Dim line As MSForms.Label
    Set line = PreviewLabel("line_" & key)
    line.Left = x: line.Top = y: line.Width = w: line.Height = h
    line.BackColor = NxRoleStyleTokenColor(IIf(emphasis, "EMPHASIS", "LINE"), DisplayModeValue())
    line.Caption = vbNullString
    line.ZOrder 0
End Sub

Private Sub RequestCancel(Optional ByVal restoreSelection As Boolean = True)
    mAllowClose = True
    Unload Me
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 표 그리기"
End Sub
