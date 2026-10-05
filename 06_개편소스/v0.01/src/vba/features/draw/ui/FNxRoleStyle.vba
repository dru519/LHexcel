Option Explicit

Private mInit As CNxFormInitializationGuard
Private mContext As CNxFeatureDialogContext, mPreview As CNxRoleStylePreview, mPreviewReady As Boolean, mAllowClose As Boolean
Private mRangeSession As CNxRangeSelectionSession
Private mPreviewControls As Collection
Private Const NX_PREVIEW_MAX_AXIS As Long = 5
Private Const NX_PREVIEW_CELL_WIDTH As Single = 58
Private Const NX_PREVIEW_CELL_HEIGHT As Single = 22

Public Sub BindFeatureContext(ByVal context As CNxFeatureDialogContext)
    If context Is Nothing Or Not context.IsConfigured Then NxRaiseContractError "셀스타일 대화상자 컨텍스트가 필요합니다."
    If context.FeatureId <> "NX-DRAW-ROLE-STYLE" Then NxRaiseContractError "셀스타일 기능 ID가 일치하지 않습니다."
    If context.DialogVariant <> "role-style" Then NxRaiseContractError "셀스타일 대화상자 유형이 일치하지 않습니다."
    Set mContext = context: mPreviewReady = False: mAllowClose = False
    InvalidatePreview
    Set mRangeSession = NxRangeSessionFromSelection()
    txtRange.Value = mRangeSession.DisplayAddress
End Sub

Private Sub UserForm_Initialize()
    Set mInit = New CNxFormInitializationGuard
    mInit.BeginInitialization
    NxInitializeStaticComboChoices Me
    ApplyDefaults
    mInit.CompleteInitialization
End Sub

Private Sub ApplyDefaults()
    mPreviewReady = False
    mAllowClose = False
    cmdExecute.Default = True
    cmdCancel.Cancel = True
    cmdExecute.Enabled = True
    Set mPreviewControls = New Collection
End Sub

Private Sub cmdPreview_Click()
    If mInit.IsInitializing Then Exit Sub
    Dim target As Range, failure As String
    On Error GoTo Failed
    InvalidatePreview
    RequireContext
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set target = mRangeSession.ResolveRange()
    NxRoleStyleValidateTarget target, SelectedAxisCode(), txtRelativeIndexes.Value
    Set mPreview = NxRoleStyleBuildPreview(target, SelectedRoleCode(), SelectedDisplayMode(), SelectedAxisCode(), txtRelativeIndexes.Value)
    RenderVisualPreview target
    mPreviewReady = mPreview.IsSealed: cmdExecute.Enabled = mPreviewReady
    lblStatus.Caption = "원본은 변경되지 않았습니다. 범위와 샘플을 검토한 뒤 적용을 누르세요."
    Exit Sub
Failed:
    failure = Err.Description
    InvalidatePreview
    lblStatus.Caption = "오류: " & NxUserErrorText(failure)
End Sub

Private Sub cmdExecute_Click()
    If mInit.IsInitializing Then Exit Sub
    Dim result As CNxResult, target As Range, currentPreview As CNxRoleStylePreview, failure As String
    On Error GoTo Failed
    RequireContext
    NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Set target = mRangeSession.ResolveRange()
    NxRoleStyleValidateTarget target, SelectedAxisCode(), txtRelativeIndexes.Value
    Set result = NxRoleStyleRun(target, SelectedRoleCode(), SelectedDisplayMode(), SelectedAxisCode(), txtRelativeIndexes.Value)
    If result.Outcome = NxCancelled Then lblStatus.Caption = "취소되었습니다."
    mAllowClose = True: Unload Me
    Exit Sub
Failed:
    failure = Err.Description
    InvalidatePreview
    lblStatus.Caption = "오류: " & NxUserErrorText(failure)
End Sub

Private Sub cmdCancel_Click()
    If mInit.IsInitializing Then Exit Sub
    InvalidatePreview
    mAllowClose = True: Unload Me
End Sub

Private Sub cmdResetDefaults_Click()
    If mInit.IsInitializing Then Exit Sub
    cboRole.ListIndex = 0
    cboDisplayMode.ListIndex = 0
    cboAxis.ListIndex = 0
    txtRelativeIndexes.Value = vbNullString
    txtPreview.Value = "미리보기 전용"
    lblStatus.Caption = "옵션을 확인하고 적용하세요."
    InvalidatePreview
End Sub

Private Sub cboRole_Change()
    If mInit.IsInitializing Then Exit Sub
    InvalidatePreview
End Sub

Private Sub cboDisplayMode_Change()
    If mInit.IsInitializing Then Exit Sub
    InvalidatePreview
End Sub

Private Sub cboAxis_Change()
    If mInit.IsInitializing Then Exit Sub
    InvalidatePreview
End Sub

Private Sub txtRelativeIndexes_Change()
    If mInit.IsInitializing Then Exit Sub
    InvalidatePreview
End Sub

Private Sub txtRange_Change()
    If mInit Is Nothing Then Exit Sub
    If mInit.IsInitializing Then Exit Sub
    InvalidatePreview
End Sub

Private Sub cmdPickRange_Click()
    If mInit.IsInitializing Then Exit Sub
    On Error GoTo Failed
    RequireContext
    InvalidatePreview
    If NxPickRange(Me, mRangeSession) Then txtRange.Value = mRangeSession.DisplayAddress
    Exit Sub
Failed:
    lblStatus.Caption = "오류: " & NxUserErrorText(Err.Description)
End Sub

Private Sub InvalidatePreview()
    mPreviewReady = False: Set mPreview = Nothing: cmdExecute.Enabled = True
    ClearVisualPreview
    txtPreview.Value = "범위와 옵션을 확인하고 적용하세요. 미리보기는 선택 사항입니다."
    lblStatus.Caption = "옵션을 확인하고 적용하세요."
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    InvalidatePreview
    Set mRangeSession = Nothing: Set mContext = Nothing
End Sub

Private Sub ClearVisualPreview()
    NxNativePreviewHide Me
    Dim controlName As Variant
    If Not mPreviewControls Is Nothing Then
        For Each controlName In mPreviewControls
            Me.Controls.Remove CStr(controlName)
        Next controlName
    End If
    Set mPreviewControls = New Collection
    lblPreviewArea.Caption = "미리보기 전용 · 원본 셀은 변경하지 않습니다."
End Sub

Private Sub RenderVisualPreview(ByVal target As Range)
    Dim affected As Range, sample As Range, area As Range, visibleArea As Range, cell As Range
    Dim label As MSForms.Label, rowCount As Long, columnCount As Long, firstRow As Long, firstColumn As Long
    Dim r As Long, c As Long, leftEdge As Single, topEdge As Single, selected As Boolean, axisSummary As String
    Set affected = NxRoleStyleResolveAffectedRange(target, mPreview.AxisCode, mPreview.RelativeIndexes)
    firstRow = 1: firstColumn = 1
    If mPreview.AxisCode = "ROWS" Then firstRow = affected.Row - target.Row + 1
    If mPreview.AxisCode = "COLUMNS" Then firstColumn = affected.Column - target.Column + 1
    rowCount = target.Rows.Count - firstRow + 1: columnCount = target.Columns.Count - firstColumn + 1
    If rowCount > NX_PREVIEW_MAX_AXIS Then rowCount = NX_PREVIEW_MAX_AXIS
    If columnCount > NX_PREVIEW_MAX_AXIS Then columnCount = NX_PREVIEW_MAX_AXIS
    Set sample = target.Cells(firstRow, firstColumn).Resize(rowCount, columnCount)
    leftEdge = lblPreviewArea.Left + 38: topEdge = lblPreviewArea.Top + 22
    lblPreviewArea.Caption = vbNullString
    For c = 1 To columnCount
        Set label = AddPreviewLabel("nxRoleCol_" & c, leftEdge + (c - 1) * NX_PREVIEW_CELL_WIDTH, lblPreviewArea.Top + 3, NX_PREVIEW_CELL_WIDTH, 18)
        label.Caption = CStr(firstColumn + c - 1): label.TextAlign = fmTextAlignCenter
    Next c
    For r = 1 To rowCount
        Set label = AddPreviewLabel("nxRoleRow_" & r, lblPreviewArea.Left + 3, topEdge + (r - 1) * NX_PREVIEW_CELL_HEIGHT, 32, NX_PREVIEW_CELL_HEIGHT)
        label.Caption = CStr(firstRow + r - 1)
        For c = 1 To columnCount
            Set cell = sample.Cells(r, c)
            selected = Not Application.Intersect(cell, affected) Is Nothing
            Set label = AddPreviewLabel("nxRoleCell_" & r & "_" & c, leftEdge + (c - 1) * NX_PREVIEW_CELL_WIDTH, topEdge + (r - 1) * NX_PREVIEW_CELL_HEIGHT, NX_PREVIEW_CELL_WIDTH, NX_PREVIEW_CELL_HEIGHT)
            label.Caption = Left$(Replace(Replace(CStr(cell.Text), vbCr, " "), vbLf, " "), 10)
            label.ControlTipText = cell.Address(False, False) & " · " & IIf(selected, "적용 대상", "적용 제외") & " · " & CStr(cell.Text)
            label.Tag = IIf(selected, "affected", "unchanged")
            StylePreviewCell label, cell, selected
        Next c
    Next r
    For Each area In affected.Areas
        Set visibleArea = Application.Intersect(area, sample)
        If Not visibleArea Is Nothing Then
            DrawPreviewBorders leftEdge + (visibleArea.Column - sample.Column) * NX_PREVIEW_CELL_WIDTH, _
                topEdge + (visibleArea.Row - sample.Row) * NX_PREVIEW_CELL_HEIGHT, _
                visibleArea.Columns.Count * NX_PREVIEW_CELL_WIDTH, visibleArea.Rows.Count * NX_PREVIEW_CELL_HEIGHT
        End If
    Next area
    axisSummary = CStr(cboAxis.Value)
    If mPreview.AxisCode <> "ALL" Then axisSummary = axisSummary & " " & mPreview.RelativeIndexes
    txtPreview.Value = mPreview.TargetAddress & vbCrLf & CStr(cboRole.Value) & " · " & CStr(cboDisplayMode.Value) & _
        " · 적용 대상: " & axisSummary & " / " & CStr(affected.CountLarge) & "셀" & vbCrLf & _
        PreviewStyleSummary() & vbCrLf & _
        "최대 5×5 축약 샘플: 상대 " & firstRow & "~" & firstRow + rowCount - 1 & "행, " & firstColumn & "~" & firstColumn + columnCount - 1 & "열." & vbCrLf & _
        "테두리는 축약 범위 기준입니다. 병합·셀 크기·기존 제외 셀 테두리는 재현하지 않습니다."
    txtPreview.ControlTipText = CStr(txtPreview.Value)
    NxNativePreviewLabels Me, "nxRole", lblPreviewArea.Left, lblPreviewArea.Top, lblPreviewArea.Width, lblPreviewArea.Height
End Sub

Private Function PreviewStyleSummary() As String
    Dim detail As String, fontName As String, bodySize As String
    fontName = NxBodyFontName(): bodySize = CStr(NxBodyFontSize()) & "pt "
    Select Case mPreview.RoleCode
        Case NX_ROLE_STYLE_TITLE: fontName = NxHeaderFontName(): detail = "14pt 굵게 · 하단 굵은 선"
        Case NX_ROLE_STYLE_SUBTITLE: fontName = NxHeaderFontName(): detail = "11pt 굵게 · 채우기·하단 선"
        Case NX_ROLE_STYLE_TABLE_HEADER: detail = bodySize & "굵게 · 채우기·외곽선"
        Case NX_ROLE_STYLE_TABLE_BODY: detail = bodySize & "본문 · 테두리 없음"
        Case NX_ROLE_STYLE_EMPHASIS_CELL: detail = bodySize & "굵게 · 반전색·외곽선"
        Case NX_ROLE_STYLE_TOTAL_ROW: detail = bodySize & "굵게 · 채우기·아래 이중선"
    End Select
    PreviewStyleSummary = fontName & " " & detail
End Function

Private Sub StylePreviewCell(ByVal label As MSForms.Label, ByVal cell As Range, ByVal selected As Boolean)
    Dim mode As String, value As Variant, formatText As String
    label.Font.Italic = cell.Font.Italic
    label.Font.Underline = (cell.Font.Underline <> xlUnderlineStyleNone)
    label.Font.Strikethrough = cell.Font.Strikethrough
    Select Case cell.HorizontalAlignment
        Case xlCenter, xlCenterAcrossSelection: label.TextAlign = fmTextAlignCenter
        Case xlRight: label.TextAlign = fmTextAlignRight
        Case xlGeneral
            If Not IsError(cell.Value2) And Not IsEmpty(cell.Value2) Then
                If IsNumeric(cell.Value2) Then label.TextAlign = fmTextAlignRight
            End If
    End Select
    If Not selected Then
        label.Font.Name = cell.Font.Name: label.Font.Size = cell.Font.Size
        label.Font.Bold = cell.Font.Bold: label.ForeColor = cell.Font.Color
        If cell.Interior.Pattern <> xlNone Then label.BackColor = cell.Interior.Color
        Exit Sub
    End If
    mode = mPreview.DisplayMode
    label.Font.Name = NxBodyFontName(): label.Font.Size = NxBodyFontSize()
    label.Font.Bold = False: label.ForeColor = NxRoleStyleTokenColor("FONT_BODY", mode)
    Select Case mPreview.RoleCode
        Case NX_ROLE_STYLE_TITLE, NX_ROLE_STYLE_SUBTITLE
            label.TextAlign = fmTextAlignLeft
            label.Font.Name = NxHeaderFontName(): label.Font.Bold = True
            label.ForeColor = NxRoleStyleTokenColor("FONT", mode)
            If mPreview.RoleCode = NX_ROLE_STYLE_TITLE Then
                label.Font.Size = 14#
            Else
                label.Font.Size = 11#: label.BackColor = NxRoleStyleTokenColor("FILL", mode)
            End If
        Case NX_ROLE_STYLE_TABLE_HEADER
            label.Font.Bold = True: label.ForeColor = NxRoleStyleTokenColor("FONT", mode)
            label.BackColor = NxRoleStyleTokenColor("FILL", mode): label.TextAlign = fmTextAlignCenter
        Case NX_ROLE_STYLE_TABLE_BODY
            label.TextAlign = fmTextAlignLeft
            value = cell.Value2
            If Not IsError(value) And Not IsNull(value) And Not IsEmpty(value) Then
                formatText = LCase$(CStr(cell.NumberFormatLocal))
                If IsDate(value) And (InStr(formatText, "y") > 0 Or InStr(formatText, "년") > 0) And _
                    (InStr(formatText, "m") > 0 Or InStr(formatText, "월") > 0) And _
                    (InStr(formatText, "d") > 0 Or InStr(formatText, "일") > 0) Then
                    label.TextAlign = fmTextAlignCenter
                ElseIf IsNumeric(value) Then
                    label.TextAlign = fmTextAlignRight
                End If
            End If
        Case NX_ROLE_STYLE_EMPHASIS_CELL
            label.Font.Bold = True: label.ForeColor = RGB(255, 255, 255)
            label.BackColor = NxRoleStyleTokenColor("EMPHASIS", mode)
        Case NX_ROLE_STYLE_TOTAL_ROW
            label.Font.Bold = True: label.TextAlign = fmTextAlignRight
            label.BackColor = NxRoleStyleTokenColor("TOTAL", mode)
    End Select
End Sub

Private Sub DrawPreviewBorders(ByVal leftEdge As Single, ByVal topEdge As Single, ByVal spanWidth As Single, ByVal spanHeight As Single)
    Dim colorValue As Long, thickness As Single, bottomEdge As Single
    colorValue = NxRoleStyleTokenColor("EMPHASIS", mPreview.DisplayMode)
    thickness = PreviewLineThickness(NxBaseLineWeight()): bottomEdge = topEdge + spanHeight
    Select Case mPreview.RoleCode
        Case NX_ROLE_STYLE_TITLE, NX_ROLE_STYLE_SUBTITLE
            If mPreview.RoleCode = NX_ROLE_STYLE_TITLE Then thickness = PreviewLineThickness(NxEmphasisLineWeight())
            AddPreviewLine leftEdge, bottomEdge - thickness, spanWidth, thickness, colorValue
        Case NX_ROLE_STYLE_TABLE_HEADER, NX_ROLE_STYLE_EMPHASIS_CELL
            AddPreviewLine leftEdge, topEdge, spanWidth, thickness, colorValue
            AddPreviewLine leftEdge, bottomEdge - thickness, spanWidth, thickness, colorValue
            AddPreviewLine leftEdge, topEdge, thickness, spanHeight, colorValue
            AddPreviewLine leftEdge + spanWidth - thickness, topEdge, thickness, spanHeight, colorValue
        Case NX_ROLE_STYLE_TOTAL_ROW
            colorValue = NxEmphasisLineColor()
            AddPreviewLine leftEdge, topEdge, spanWidth, PreviewLineThickness(NxEmphasisLineWeight()), colorValue
            ' xlDouble is two separate strokes, matching the actual total-row bottom edge.
            AddPreviewLine leftEdge, bottomEdge - 4, spanWidth, 1, colorValue
            AddPreviewLine leftEdge, bottomEdge - 1, spanWidth, 1, colorValue
    End Select
End Sub

Private Function PreviewLineThickness(ByVal weight As XlBorderWeight) As Single
    If weight = xlMedium Or weight = xlThick Then PreviewLineThickness = 2 Else PreviewLineThickness = 1
End Function

Private Sub AddPreviewLine(ByVal leftEdge As Single, ByVal topEdge As Single, ByVal spanWidth As Single, ByVal spanHeight As Single, ByVal colorValue As Long)
    Dim label As MSForms.Label
    Set label = AddPreviewLabel("nxRoleLine_" & mPreviewControls.Count, leftEdge, topEdge, spanWidth, spanHeight)
    label.BackColor = colorValue
End Sub

Private Function AddPreviewLabel(ByVal controlName As String, ByVal leftEdge As Single, ByVal topEdge As Single, ByVal spanWidth As Single, ByVal spanHeight As Single) As MSForms.Label
    Dim label As MSForms.Label
    Set label = Me.Controls.Add("Forms.Label.1", controlName, True)
    mPreviewControls.Add controlName
    With label
        .Left = leftEdge: .Top = topEdge: .Width = spanWidth: .Height = spanHeight
        .Caption = vbNullString: .BackStyle = fmBackStyleOpaque: .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleNone: .TextAlign = fmTextAlignLeft: .WordWrap = False
        .Font.Name = NxBodyFontName(): .Font.Size = NxBodyFontSize()
        .ForeColor = NxBodyFontColor()
    End With
    Set AddPreviewLabel = label
End Function

Private Sub RequireContext()
    If mContext Is Nothing Or Not mContext.IsConfigured Then NxRaiseContractError "셀스타일 대화상자 컨텍스트가 필요합니다."
End Sub

Private Function SelectedRoleCode() As String
    SelectedRoleCode = Choose(cboRole.ListIndex + 1, "TITLE", "SUBTITLE", "TABLE_HEADER", "TABLE_BODY", "EMPHASIS_CELL", "TOTAL_ROW")
End Function

Private Function SelectedDisplayMode() As String
    SelectedDisplayMode = Choose(cboDisplayMode.ListIndex + 1, "MONO", "COLOR")
End Function

Private Function SelectedAxisCode() As String
    SelectedAxisCode = Choose(cboAxis.ListIndex + 1, "ALL", "ROWS", "COLUMNS")
End Function

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
