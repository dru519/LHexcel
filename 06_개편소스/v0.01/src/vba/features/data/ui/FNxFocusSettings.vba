Option Explicit

Private mAllowClose As Boolean
Private mLoading As Boolean
Private mPreviewReady As Boolean
Private mFirstRow As Long, mFirstColumn As Long
Private mSelectedTop As Long, mSelectedBottom As Long
Private mSelectedLeft As Long, mSelectedRight As Long
Private mSampleText(1 To 6, 1 To 5) As String
Private mSampleColor(1 To 6, 1 To 5) As Long

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
    cmdApply.Default = True
    cmdCancel.Cancel = True
    CapturePreview
    LoadModel NxFocusLoadSettings()
End Sub
Private Sub LoadModel(ByVal model As CNxFocusSettingsModel)
    If model Is Nothing Then Set model = NxFocusDefaultSettings()
    mLoading = True
    Select Case model.Shape
        Case NX_FOCUS_SHAPE_CROSS: cboShape.ListIndex = 0
        Case NX_FOCUS_SHAPE_HORIZONTAL: cboShape.ListIndex = 1
        Case NX_FOCUS_SHAPE_VERTICAL: cboShape.ListIndex = 2
    End Select
    cboStyle.ListIndex = 0
    txtColor.Value = model.ColorHex
    txtIntensity.Value = CStr(model.IntensityPercent)
    chkSelectedArea.Value = model.HighlightSelectedArea
    chkEnableOnApply.Value = model.EnableOnApply
    mLoading = False
    UpdatePreview
End Sub

Private Function ModelFromControls() As CNxFocusSettingsModel
    Dim intensity As Double
    If Not IsNumeric(txtIntensity.Value) Then NxRaiseContractError "농도는 0부터 100 사이의 정수로 입력하세요."
    intensity = CDbl(txtIntensity.Value)
    If intensity <> Fix(intensity) Then NxRaiseContractError "농도는 정수로 입력하세요."
    Set ModelFromControls = NxFocusCreateSettings(CStr(cboShape.Value), CStr(cboStyle.Value), _
        CStr(txtColor.Value), CLng(intensity), CBool(chkSelectedArea.Value), CBool(chkEnableOnApply.Value))
End Function

Private Sub UpdatePreview()
    Dim model As CNxFocusSettingsModel
    If mLoading Then Exit Sub
    On Error GoTo Invalid
    Set model = ModelFromControls()
    DrawGridPreview model
    Exit Sub
Invalid:
    Err.Clear
    lblPreview.BackColor = RGB(255, 255, 255)
    lblPreview.Caption = "입력값을 확인하세요."
    HideGridPreview
End Sub

Private Sub CapturePreview()
    Dim target As Range, sample As Range, cell As Range, r As Long, c As Long
    mFirstRow = 1: mFirstColumn = 1
    mSelectedTop = 3: mSelectedBottom = 3
    mSelectedLeft = 3: mSelectedRight = 3
    For r = 1 To 6
        For c = 1 To 5
            mSampleColor(r, c) = vbWhite
        Next c
    Next r
    On Error GoTo Done
    If TypeName(Application.Selection) <> "Range" Then GoTo Done
    Set target = Application.Selection.Areas(1)
    mFirstRow = Application.Max(1, Application.Min(target.Row - 2, target.Worksheet.Rows.Count - 5))
    mFirstColumn = Application.Max(1, Application.Min(target.Column - 2, target.Worksheet.Columns.Count - 4))
    mSelectedTop = target.Row - mFirstRow + 1
    mSelectedLeft = target.Column - mFirstColumn + 1
    mSelectedBottom = Application.Min(6, mSelectedTop + target.Rows.Count - 1)
    mSelectedRight = Application.Min(5, mSelectedLeft + target.Columns.Count - 1)
    Set sample = target.Worksheet.Cells(mFirstRow, mFirstColumn).Resize(6, 5)
    ' Bounded read once. Option changes never read or write the worksheet.
    For r = 1 To 6
        For c = 1 To 5
            Set cell = sample.Cells(r, c)
            mSampleText(r, c) = Left$(CStr(cell.Text), 18)
            If cell.Interior.ColorIndex <> xlColorIndexNone Then mSampleColor(r, c) = cell.Interior.Color
        Next c
    Next r
Done:
    Err.Clear
    mPreviewReady = True
End Sub

Private Function PreviewLabel(ByVal key As String, ByVal x As Single, ByVal y As Single, _
        ByVal width As Single, ByVal height As Single) As Object
    Dim item As Object
    On Error Resume Next
    Set item = Me.Controls(key)
    On Error GoTo 0
    If item Is Nothing Then Set item = Me.Controls.Add("Forms.Label.1", key, True)
    With item
        .Left = x: .Top = y: .Width = width: .Height = height
        .BackStyle = 1: .BackColor = vbWhite: .ForeColor = RGB(45, 45, 45)
        .Font.Name = "맑은 고딕": .Font.Size = 8: .Font.Bold = False
        .TextAlign = 2: .WordWrap = False: .Caption = vbNullString
        .Visible = True: .ZOrder 0
    End With
    Set PreviewLabel = item
End Function

Private Sub DrawGridPreview(ByVal model As CNxFocusSettingsModel)
    Const gridLeft As Single = 274, gridTop As Single = 104
    Const cellWidth As Single = 36, cellHeight As Single = 17
    Dim r As Long, c As Long, item As Object, selected As Boolean, highlighted As Boolean
    Dim color As Long, x As Single, y As Single, w As Single, h As Single
    If Not mPreviewReady Then CapturePreview
    lblPreview.Caption = vbNullString
    lblPreview.BackColor = RGB(215, 220, 223)
    For c = 1 To 5
        Set item = PreviewLabel("nxFocus_column_" & c, gridLeft + (c - 1) * cellWidth, 88, cellWidth - 1, 15)
        item.Caption = PreviewColumnName(mFirstColumn + c - 1)
        item.BackColor = RGB(242, 244, 245)
    Next c
    For r = 1 To 6
        Set item = PreviewLabel("nxFocus_row_" & r, 250, gridTop + (r - 1) * cellHeight, 23, cellHeight - 1)
        item.Caption = CStr(mFirstRow + r - 1): item.BackColor = RGB(242, 244, 245)
        For c = 1 To 5
            selected = r >= mSelectedTop And r <= mSelectedBottom And c >= mSelectedLeft And c <= mSelectedRight
            highlighted = Not selected And (((model.Shape <> NX_FOCUS_SHAPE_VERTICAL) And r >= mSelectedTop And r <= mSelectedBottom) Or _
                ((model.Shape <> NX_FOCUS_SHAPE_HORIZONTAL) And c >= mSelectedLeft And c <= mSelectedRight))
            color = mSampleColor(r, c)
            If highlighted Then color = PreviewBlend(color, NxFocusEffectiveColor(model), model.IntensityPercent)
            Set item = PreviewLabel("nxFocus_cell_" & r & "_" & c, gridLeft + (c - 1) * cellWidth, gridTop + (r - 1) * cellHeight, cellWidth - 1, cellHeight - 1)
            item.Caption = mSampleText(r, c): item.BackColor = color
        Next c
    Next r
    x = gridLeft + (mSelectedLeft - 1) * cellWidth: y = gridTop + (mSelectedTop - 1) * cellHeight
    w = (mSelectedRight - mSelectedLeft + 1) * cellWidth - 1
    h = (mSelectedBottom - mSelectedTop + 1) * cellHeight - 1
    color = RGB(33, 115, 70)
    If model.HighlightSelectedArea Then color = NxFocusConditionalColor(model)
    Set item = PreviewLabel("nxFocus_top", x, y, w, 2): item.BackColor = color
    Set item = PreviewLabel("nxFocus_bottom", x, y + h - 2, w, 2): item.BackColor = color
    Set item = PreviewLabel("nxFocus_left", x, y, 2, h): item.BackColor = color
    Set item = PreviewLabel("nxFocus_right", x + w - 2, y, 2, h): item.BackColor = color
    lblPreviewNote.Caption = "선택 주변 6×5 표본 · " & cboShape.Value & " · 농도 " & model.IntensityPercent & vbCrLf & "간략 표시이며 원본 셀은 바뀌지 않습니다."
    NxNativePreviewLabels Me, "nxFocus_", 250, 88, 204, 118
End Sub

Private Function PreviewColumnName(ByVal number As Long) As String
    Do While number > 0
        number = number - 1
        PreviewColumnName = Chr$(65 + number Mod 26) & PreviewColumnName
        number = number \ 26
    Loop
End Function

Private Function PreviewBlend(ByVal background As Long, ByVal foreground As Long, ByVal intensity As Long) As Long
    PreviewBlend = RGB(((background And 255&) * (100 - intensity) + (foreground And 255&) * intensity + 50) \ 100, _
        (((background \ 256) And 255&) * (100 - intensity) + ((foreground \ 256) And 255&) * intensity + 50) \ 100, _
        (((background \ 65536) And 255&) * (100 - intensity) + ((foreground \ 65536) And 255&) * intensity + 50) \ 100)
End Function

Private Sub HideGridPreview()
    NxNativePreviewHide Me
    Dim item As Object
    For Each item In Me.Controls
        If Left$(item.Name, 8) = "nxFocus_" Then item.Visible = False
    Next item
End Sub

Private Sub cmdApply_Click()
    Dim model As CNxFocusSettingsModel
    On Error GoTo Failed
    Set model = ModelFromControls()
    NxFocusControllerCommitSettings model
    RequestCancel
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdDefaults_Click()
    LoadModel NxFocusDefaultSettings()
End Sub

Private Sub cmdSwatch1_Click(): SetColor "#5FC8D8": End Sub
Private Sub cmdSwatch2_Click(): SetColor "#78C6A3": End Sub
Private Sub cmdSwatch3_Click(): SetColor "#F2CF63": End Sub
Private Sub cmdSwatch4_Click(): SetColor "#F4A261": End Sub
Private Sub cmdSwatch5_Click(): SetColor "#E5989B": End Sub
Private Sub cmdSwatch6_Click(): SetColor "#A69CAC": End Sub

Private Sub SetColor(ByVal value As String)
    txtColor.Value = value
    UpdatePreview
End Sub

Private Sub cboShape_Change(): UpdatePreview: End Sub
Private Sub cboStyle_Change(): UpdatePreview: End Sub
Private Sub txtColor_Change(): UpdatePreview: End Sub
Private Sub txtIntensity_Change(): UpdatePreview: End Sub
Private Sub chkSelectedArea_Click(): UpdatePreview: End Sub
Private Sub cmdCancel_Click(): RequestCancel: End Sub

Private Sub RequestCancel()
    mAllowClose = True
    Unload Me
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
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 포커스셀 설정"
End Sub
