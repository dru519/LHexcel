Option Explicit

Private mInit As CNxFormInitializationGuard
Private mContext As CNxFeatureDialogContext
Private mPendingColors As String
Private mSelectedColor As String
Private mPendingMaxItems As Long
Private mUpdatingPalette As Boolean

Private Sub UserForm_Initialize()
    Set mInit = New CNxFormInitializationGuard
    mInit.BeginInitialization
    NxInitializeStaticComboChoices Me
    ApplyDefaults
    SetPaletteChoices NxFocusLoadColors(), 10
    mInit.CompleteInitialization
End Sub
Public Sub BindFeatureContext(ByVal context As CNxFeatureDialogContext)
    If context Is Nothing Then NxRaiseContractError "Focus palette context is required"
    If Not context.IsConfigured Then NxRaiseContractError "Focus palette context is required"
    If context.FeatureId <> "NX-DATA-FOCUS-CELL" Then NxRaiseContractError "Focus palette accepts only NX-DATA-FOCUS-CELL"
    If context.DialogId <> "NX-DLG-FOCUS-PALETTE" Then NxRaiseContractError "Focus palette dialog binding is invalid"
    ' This form is reachable only from the hybrid focus route; direct/dialog contexts fail closed above.
    If Not mContext Is Nothing Then NxRaiseContractError "Focus palette context can be bound only once"
    Set mContext = context
    lblContext.Caption = context.FeatureId & " · " & context.DialogVariant
End Sub

Private Sub ApplyDefaults()
    cmdToggle.Default = True
    cmdCancel.Cancel = True
    txtCustomColors.Value = NxFocusLoadColors()
    mPendingMaxItems = 10
    mSelectedColor = vbNullString
    lblColorPreview.Caption = "선택 없음"
    If NxFocusIsEnabled Then
        cmdToggle.Caption = "끄기"
        lblState.Caption = "포커스셀이 켜져 있습니다."
    Else
        cmdToggle.Caption = "켜기"
        lblState.Caption = "포커스셀이 꺼져 있습니다."
    End If
End Sub

Private Sub cmdThemeSwatch_Click()
    If mInit.IsInitializing Then Exit Sub
    RequireContext
    SetPaletteChoices NxFocusThemeColors(), 0
End Sub

Private Sub cmdPresetSwatch_Click()
    If mInit.IsInitializing Then Exit Sub
    RequireContext
    SetPaletteChoices NxFocusPresetColors("STAEDTLER"), 0
End Sub

Private Sub cmdSignoSwatch_Click()
    If mInit.IsInitializing Then Exit Sub
    RequireContext
    SetPaletteChoices NxFocusPresetColors("SIGNO"), 0
End Sub

Private Sub cmdStandardSwatch_Click()
    If mInit.IsInitializing Then Exit Sub
    RequireContext
    SetPaletteChoices NxFocusStandardColors(), 0
End Sub

Private Sub cmdMildPreset_Click()
    If mInit.IsInitializing Then Exit Sub
    RequireContext
    SetPaletteChoices NxFocusPresetColors("MILD"), 0
End Sub

Private Sub cmdUserSwatch_Click()
    If mInit.IsInitializing Then Exit Sub
    RequireContext
    SetPaletteChoices NxFocusLoadColors(), 10
End Sub

Private Sub cmbPaletteColor_Change()
    If mInit.IsInitializing Then Exit Sub
    Dim selected As String
    RequireContext
    selected = NxFocusNormalizeStrictColors(cmbPaletteColor.Value, 1)
    If Len(selected) = 0 Then
        mSelectedColor = vbNullString
        lblColorPreview.Caption = "선택 없음"
        lblState.Caption = "유효한 색상을 선택하세요."
        Exit Sub
    End If
    mSelectedColor = selected
    lblColorPreview.Caption = selected
    lblColorPreview.BackColor = NxFocusFirstColor(selected)
    lblState.Caption = "선택 색상을 적용할 준비가 되었습니다."
End Sub

Private Sub txtCustomColors_Change()
    If mInit.IsInitializing Then Exit Sub
    If mUpdatingPalette Then Exit Sub
    SetPaletteChoices txtCustomColors.Value, 10
End Sub

Private Sub cmdToggle_Click()
    If mInit.IsInitializing Then Exit Sub
    RequireContext
    If NxFocusIsEnabled Then
        NxFocusDisable
        cmdToggle.Caption = "켜기"
        lblState.Caption = "포커셀이 꺼졌습니다."
    Else
        Dim colors As String
        colors = NxFocusNormalizeStrictColors(txtCustomColors.Value, mPendingMaxItems)
        If Len(colors) = 0 Then NxRaiseContractError "#RRGGBB 색상을 하나 이상 입력하세요."
        If Len(mSelectedColor) = 0 Then NxRaiseContractError "적용할 색상을 하나 선택하세요."
        colors = NxFocusMoveColorToFirst(colors, mSelectedColor)
        If Len(colors) = 0 Then NxRaiseContractError "선택한 색상이 현재 팔레트에 없습니다."
        NxFocusEnable NxFocusFirstColor(colors)
        NxFocusSaveColors colors
        mPendingColors = colors
        cmdToggle.Caption = "끄기"
        lblState.Caption = "포커스셀이 켜졌습니다."
    End If
End Sub

Private Sub cmdCancel_Click()
    If mInit.IsInitializing Then Exit Sub
    Unload Me
End Sub

Private Sub RequireContext()
    mInit.RequireReady
    If mContext Is Nothing Then NxRaiseContractError "Focus palette context must be bound before use"
End Sub

Private Sub SetPaletteChoices(ByVal rawColors As String, ByVal maxItems As Long)
    Dim colors As String
    Dim item As Variant

    colors = NxFocusNormalizeStrictColors(rawColors, maxItems)
    mPendingMaxItems = maxItems
    mSelectedColor = vbNullString
    mPendingColors = colors
    cmbPaletteColor.Clear
    lblColorPreview.Caption = "선택 없음"
    lblColorPreview.BackColor = &H8000000F
    If Len(colors) = 0 Then
        lblState.Caption = "유효한 색상 목록이 없습니다."
        Exit Sub
    End If

    mUpdatingPalette = True
    txtCustomColors.Value = colors
    mUpdatingPalette = False
    For Each item In Split(colors, ",")
        cmbPaletteColor.AddItem CStr(item)
    Next item
    cmbPaletteColor.ListIndex = -1
    lblState.Caption = "색상을 하나 선택한 뒤 적용하세요."
End Sub
