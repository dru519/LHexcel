Attribute VB_Name = "NxFocusSettings"
Option Explicit

Private Const NX_FOCUS_SETTINGS_FILE As String = "focus-cell-v2.cfg"
Private Const NX_FOCUS_LEGACY_FILE As String = "focus-palette.cfg"
Private Const NX_FOCUS_BACKUP_SUFFIX As String = ".bak"
Private Const NX_FOCUS_TEMP_SUFFIX As String = ".tmp"
Private Const NX_FOCUS_LOCK_SUFFIX As String = ".lck"
Private Const NX_FOCUS_DEFAULT_COLOR As String = "#5FC8D8"
Private Const NX_FOCUS_SWATCHES As String = "#5FC8D8,#78C6A3,#F2CF63,#F4A261,#E5989B,#A69CAC"
Private Const NX_FOCUS_STANDARD_COLORS As String = "#C00000,#FF0000,#FFC000,#FFFF00,#92D050,#00B050,#00B0F0,#0070C0,#002060,#7030A0"
Private Const NX_FOCUS_PRIOR_BACKUP_SUFFIX As String = ".prev"

Private mStageActive As Boolean
Private mStageCommitted As Boolean
Private mStageHadOriginal As Boolean
Private mStageHadBackup As Boolean
Private mStagePath As String
Private mStageTemporary As String
Private mStageBackup As String
Private mStagePriorBackup As String
Private mStageLockPath As String
Private mStageLockHandle As Integer

Public Function NxFocusSettingsPath() As String
    NxFocusSettingsPath = NxLHexcelProfileRoot() & "\Settings\" & NX_FOCUS_SETTINGS_FILE
End Function

Private Function NxFocusLegacySettingsPath() As String
    NxFocusLegacySettingsPath = NxLHexcelProfileRoot() & "\Settings\" & NX_FOCUS_LEGACY_FILE
End Function

Public Function NxFocusDefaultSettings() As CNxFocusSettingsModel
    Dim model As New CNxFocusSettingsModel
    model.Configure NX_FOCUS_SHAPE_CROSS, NX_FOCUS_STYLE_WIDE, NX_FOCUS_DEFAULT_COLOR, 34, False, True
    Set NxFocusDefaultSettings = model
End Function

Public Function NxFocusLoadSettings() As CNxFocusSettingsModel
    Dim model As CNxFocusSettingsModel
    Dim legacyColors As String
    On Error Resume Next
    Set model = NxFocusParseSettings(NxFocusReadText(NxFocusSettingsPath()))
    If model Is Nothing Then Set model = NxFocusParseSettings(NxFocusReadText(NxFocusSettingsPath() & NX_FOCUS_BACKUP_SUFFIX))
    On Error GoTo 0
    If Not model Is Nothing Then Set NxFocusLoadSettings = model: Exit Function

    Set model = NxFocusDefaultSettings()
    legacyColors = NxFocusNormalizeColors(NxFocusReadText(NxFocusLegacySettingsPath()))
    If Len(legacyColors) > 0 Then
        Set model = NxFocusCreateSettings(model.Shape, model.Style, CStr(Split(legacyColors, ",")(0)), _
            model.IntensityPercent, model.HighlightSelectedArea, model.EnableOnApply)
        On Error Resume Next
        NxFocusSaveSettingsModel model
        On Error GoTo 0
    End If
    Set NxFocusLoadSettings = model
End Function

Public Sub NxFocusSaveSettingsModel(ByVal model As CNxFocusSettingsModel)
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    NxFocusStageSettingsModel model
    NxFocusCommitStagedSettings
    NxFocusCompleteSettingsTransaction
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    NxFocusRollbackSettingsFile
    If Len(failureDescription) = 0 Then failureDescription = "설정 파일을 쓸 수 없습니다."
    Err.Raise failureNumber, "NxFocusSettings", "포커스셀 설정을 저장하지 못했습니다. " & failureDescription
End Sub

Public Sub NxFocusStageSettingsModel(ByVal model As CNxFocusSettingsModel)
    Dim outputHandle As Integer
    Dim payload As String
    Dim verified As CNxFocusSettingsModel
    Dim failureNumber As Long
    Dim failureDescription As String
    If model Is Nothing Then NxRaiseContractError "저장할 포커스셀 설정이 없습니다."
    If Not model.IsSealed Then NxRaiseContractError "저장할 포커스셀 설정이 없습니다."
    If mStageActive Then NxRaiseContractError "다른 포커스셀 설정 저장이 진행 중입니다."
    mStagePath = NxFocusSettingsPath()
    mStageTemporary = mStagePath & NX_FOCUS_TEMP_SUFFIX
    mStageBackup = mStagePath & NX_FOCUS_BACKUP_SUFFIX
    mStagePriorBackup = mStageBackup & NX_FOCUS_PRIOR_BACKUP_SUFFIX
    mStageLockPath = mStagePath & NX_FOCUS_LOCK_SUFFIX
    payload = NxFocusSettingsPayload(model)
    On Error GoTo Failed
    NxFocusEnsureSettingsFolder
    mStageLockHandle = FreeFile
    Open mStageLockPath For Binary Access Read Write Lock Read Write As #mStageLockHandle
    mStageActive = True
    If NxFocusFileExists(mStageTemporary) Then Kill mStageTemporary
    outputHandle = FreeFile
    Open mStageTemporary For Output Access Write Lock Read Write As #outputHandle
    Print #outputHandle, payload;
    Close #outputHandle
    outputHandle = 0
    Set verified = NxFocusParseSettings(NxFocusReadText(mStageTemporary))
    If verified Is Nothing Then NxRaiseContractError "임시 설정 파일을 확인하지 못했습니다."
    If NxFocusSettingsPayload(verified) <> payload Then NxRaiseContractError "임시 설정 파일을 확인하지 못했습니다."
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If outputHandle > 0 Then Close #outputHandle
    On Error GoTo 0
    NxFocusRollbackSettingsFile
    Err.Raise failureNumber, "NxFocusSettings", failureDescription
End Sub

Public Sub NxFocusCommitStagedSettings()
    Dim failureNumber As Long
    Dim failureDescription As String
    If Not mStageActive Then NxRaiseContractError "저장할 임시 포커스셀 설정이 없습니다."
    On Error GoTo Failed
    mStageHadOriginal = NxFocusFileExists(mStagePath)
    mStageHadBackup = NxFocusFileExists(mStageBackup)
    If NxFocusFileExists(mStagePriorBackup) Then Kill mStagePriorBackup
    If mStageHadBackup Then Name mStageBackup As mStagePriorBackup
    If mStageHadOriginal Then Name mStagePath As mStageBackup
    Name mStageTemporary As mStagePath
    mStageCommitted = True
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    NxFocusRollbackSettingsFile
    Err.Raise failureNumber, "NxFocusSettings", failureDescription
End Sub

Public Sub NxFocusCompleteSettingsTransaction()
    On Error Resume Next
    If NxFocusFileExists(mStagePriorBackup) Then Kill mStagePriorBackup
    If mStageLockHandle > 0 Then Close #mStageLockHandle
    mStageLockHandle = 0
    If NxFocusFileExists(mStageLockPath) Then Kill mStageLockPath
    NxFocusResetStageState
    On Error GoTo 0
End Sub

Public Sub NxFocusRollbackSettingsFile()
    On Error Resume Next
    If Len(mStageTemporary) > 0 And NxFocusFileExists(mStageTemporary) Then Kill mStageTemporary
    If mStageHadOriginal And Not NxFocusFileExists(mStagePath) And NxFocusFileExists(mStageBackup) Then Name mStageBackup As mStagePath
    If mStageCommitted Then
        If NxFocusFileExists(mStagePath) Then Kill mStagePath
        If mStageHadOriginal And NxFocusFileExists(mStageBackup) Then Name mStageBackup As mStagePath
    End If
    If mStageHadBackup And NxFocusFileExists(mStagePriorBackup) Then Name mStagePriorBackup As mStageBackup
    If mStageLockHandle > 0 Then Close #mStageLockHandle
    mStageLockHandle = 0
    If Len(mStageLockPath) > 0 And NxFocusFileExists(mStageLockPath) Then Kill mStageLockPath
    NxFocusResetStageState
    On Error GoTo 0
End Sub

Public Function NxFocusSettingsTransactionActive() As Boolean
    NxFocusSettingsTransactionActive = mStageActive
End Function

Private Sub NxFocusResetStageState()
    mStageActive = False
    mStageCommitted = False
    mStageHadOriginal = False
    mStageHadBackup = False
    mStagePath = vbNullString
    mStageTemporary = vbNullString
    mStageBackup = vbNullString
    mStagePriorBackup = vbNullString
    mStageLockPath = vbNullString
End Sub

Public Sub NxFocusResetSettings()
    NxFocusSaveSettingsModel NxFocusDefaultSettings()
End Sub

Public Function NxFocusCreateSettings(ByVal shape As String, ByVal style As String, _
    ByVal colorHex As String, ByVal intensityPercent As Long, _
    ByVal highlightSelectedArea As Boolean, ByVal enableOnApply As Boolean) As CNxFocusSettingsModel
    Dim model As New CNxFocusSettingsModel
    model.Configure shape, style, colorHex, intensityPercent, highlightSelectedArea, enableOnApply
    Set NxFocusCreateSettings = model
End Function

Public Function NxFocusNormalizeShape(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case NX_FOCUS_SHAPE_CROSS, "교차": NxFocusNormalizeShape = NX_FOCUS_SHAPE_CROSS
        Case NX_FOCUS_SHAPE_HORIZONTAL, "가로": NxFocusNormalizeShape = NX_FOCUS_SHAPE_HORIZONTAL
        Case NX_FOCUS_SHAPE_VERTICAL, "세로": NxFocusNormalizeShape = NX_FOCUS_SHAPE_VERTICAL
    End Select
End Function

Public Function NxFocusNormalizeStyle(ByVal value As String) As String
    Select Case LCase$(Trim$(value))
        Case NX_FOCUS_STYLE_WIDE, "넓은 띠": NxFocusNormalizeStyle = NX_FOCUS_STYLE_WIDE
    End Select
End Function

Public Function NxFocusNormalizeHexColor(ByVal value As String) As String
    value = UCase$(Trim$(value))
    If NxFocusIsHexColor(value) Then NxFocusNormalizeHexColor = value
End Function

Public Function NxFocusColorValue(ByVal colorHex As String) As Long
    colorHex = NxFocusNormalizeHexColor(colorHex)
    If Len(colorHex) = 0 Then NxRaiseContractError "색상은 #RRGGBB 형식으로 입력하세요."
    NxFocusColorValue = RGB(CLng("&H" & Mid$(colorHex, 2, 2)), _
        CLng("&H" & Mid$(colorHex, 4, 2)), CLng("&H" & Mid$(colorHex, 6, 2)))
End Function

Public Function NxFocusColorHexFromLong(ByVal colorValue As Long) As String
    NxFocusColorHexFromLong = "#" & Right$("0" & Hex$(colorValue And &HFF&), 2) & _
        Right$("0" & Hex$((colorValue \ &H100&) And &HFF&), 2) & _
        Right$("0" & Hex$((colorValue \ &H10000) And &HFF&), 2)
End Function

Public Function NxFocusEffectiveColor(ByVal model As CNxFocusSettingsModel) As Long
    If model Is Nothing Then NxRaiseContractError "포커스셀 설정이 없습니다."
    NxFocusEffectiveColor = NxFocusColorValue(model.ColorHex)
End Function

Public Function NxFocusConditionalColor(ByVal model As CNxFocusSettingsModel) As Long
    Dim colorValue As Long
    Dim intensity As Long
    Dim redValue As Long
    Dim greenValue As Long
    Dim blueValue As Long
    If model Is Nothing Then NxRaiseContractError "포커스셀 설정이 없습니다."
    colorValue = NxFocusColorValue(model.ColorHex)
    intensity = model.IntensityPercent
    redValue = 255 - ((255 - (colorValue And &HFF&)) * intensity + 50) \ 100
    greenValue = 255 - ((255 - ((colorValue \ &H100&) And &HFF&)) * intensity + 50) \ 100
    blueValue = 255 - ((255 - ((colorValue \ &H10000) And &HFF&)) * intensity + 50) \ 100
    NxFocusConditionalColor = RGB(redValue, greenValue, blueValue)
End Function

Public Function NxFocusAlphaByte(ByVal model As CNxFocusSettingsModel) As Byte
    Dim alphaValue As Long
    If model Is Nothing Then NxRaiseContractError "포커스셀 설정이 없습니다."
    alphaValue = (model.IntensityPercent * 255& + 50&) \ 100&
    If alphaValue < 0 Then alphaValue = 0
    If alphaValue > 255 Then alphaValue = 255
    NxFocusAlphaByte = CByte(alphaValue)
End Function

Private Function NxFocusSettingsPayload(ByVal model As CNxFocusSettingsModel) As String
    If model Is Nothing Then Exit Function
    NxFocusSettingsPayload = "schema_version=2" & vbLf & _
        "shape=" & model.Shape & vbLf & _
        "style=" & model.Style & vbLf & _
        "color=" & model.ColorHex & vbLf & _
        "intensity_percent=" & CStr(model.IntensityPercent) & vbLf & _
        "highlight_selected_area=" & CStr(Abs(CLng(model.HighlightSelectedArea))) & vbLf & _
        "enable_on_apply=" & CStr(Abs(CLng(model.EnableOnApply))) & vbLf
End Function

Private Function NxFocusParseSettings(ByVal payload As String) As CNxFocusSettingsModel
    Dim values As Object
    Dim line As Variant
    Dim parts As Variant
    Dim intensity As Long
    On Error GoTo Invalid
    If Len(payload) = 0 Then Exit Function
    Set values = CreateObject("Scripting.Dictionary")
    payload = Replace(payload, vbCr, vbNullString)
    For Each line In Split(payload, vbLf)
        If Len(CStr(line)) > 0 Then
            parts = Split(CStr(line), "=", 2)
            If UBound(parts) <> 1 Or values.Exists(parts(0)) Then Exit Function
            values.Add parts(0), parts(1)
        End If
    Next line
    If values.Count <> 7 Or values("schema_version") <> "2" Then Exit Function
    If Not IsNumeric(values("intensity_percent")) Then Exit Function
    intensity = CLng(values("intensity_percent"))
    If CStr(intensity) <> values("intensity_percent") Then Exit Function
    If values("highlight_selected_area") <> "0" And values("highlight_selected_area") <> "1" Then Exit Function
    If values("enable_on_apply") <> "0" And values("enable_on_apply") <> "1" Then Exit Function
    Set NxFocusParseSettings = NxFocusCreateSettings(values("shape"), values("style"), values("color"), _
        intensity, values("highlight_selected_area") = "1", values("enable_on_apply") = "1")
    Exit Function
Invalid:
    Set NxFocusParseSettings = Nothing
End Function

Private Function NxFocusReadText(ByVal path As String) As String
    Dim handle As Integer
    Dim lineText As String
    Dim result As String
    On Error GoTo Failed
    If Not NxFocusFileExists(path) Then Exit Function
    handle = FreeFile
    Open path For Input Access Read As #handle
    Do Until EOF(handle)
        Line Input #handle, lineText
        result = result & lineText & vbLf
    Loop
    NxFocusReadText = result
    Close #handle
    Exit Function
Failed:
    On Error Resume Next
    If handle > 0 Then Close #handle
    On Error GoTo 0
End Function

Private Function NxFocusIsHexColor(ByVal value As String) As Boolean
    If Len(value) <> 7 Then Exit Function
    If Left$(value, 1) <> "#" Then Exit Function
    NxFocusIsHexColor = Mid$(value, 2) Like _
        "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]"
End Function

Private Function NxFocusFileExists(ByVal path As String) As Boolean
    On Error GoTo Missing
    NxFocusFileExists = (Len(path) > 0 And Len(Dir$(path)) > 0)
Missing:
End Function

Private Sub NxFocusEnsureSettingsFolder()
    Dim root As String
    Dim settingsFolder As String
    root = NxLHexcelProfileRoot()
    settingsFolder = root & "\Settings"
    If Not NxFocusFolderExists(root) Then MkDir root
    If Not NxFocusFolderExists(settingsFolder) Then MkDir settingsFolder
End Sub

Private Function NxFocusFolderExists(ByVal path As String) As Boolean
    On Error GoTo Missing
    NxFocusFolderExists = ((GetAttr(path) And vbDirectory) = vbDirectory)
Missing:
End Function

' Compatibility for the retired palette source. Normal product paths use the v2 model above.
Public Function NxFocusLoadColors() As String
    NxFocusLoadColors = NxFocusLoadSettings().ColorHex & "," & NX_FOCUS_SWATCHES
End Function

Public Function NxFocusPresetColors(ByVal preset As String) As String
    NxFocusPresetColors = NX_FOCUS_SWATCHES
End Function

Public Function NxFocusThemeColors() As String: NxFocusThemeColors = NX_FOCUS_SWATCHES: End Function
Public Function NxFocusStandardColors() As String: NxFocusStandardColors = NX_FOCUS_STANDARD_COLORS: End Function

Public Function NxFocusNormalizeStrictColors(ByVal colors As String, ByVal maxItems As Long) As String
    Dim item As Variant
    Dim value As String
    Dim count As Long
    For Each item In Split(colors, ",")
        value = NxFocusNormalizeHexColor(CStr(item))
        If Len(value) = 0 Then Exit Function
        count = count + 1
        If maxItems > 0 And count > maxItems Then Exit Function
        If Len(NxFocusNormalizeStrictColors) > 0 Then NxFocusNormalizeStrictColors = NxFocusNormalizeStrictColors & ","
        NxFocusNormalizeStrictColors = NxFocusNormalizeStrictColors & value
    Next item
End Function

Public Function NxFocusNormalizeColors(ByVal colors As String) As String
    Dim item As Variant
    Dim value As String
    For Each item In Split(colors, ",")
        value = NxFocusNormalizeHexColor(CStr(item))
        If Len(value) > 0 Then
            If Len(NxFocusNormalizeColors) > 0 Then NxFocusNormalizeColors = NxFocusNormalizeColors & ","
            NxFocusNormalizeColors = NxFocusNormalizeColors & value
        End If
    Next item
End Function

Public Function NxFocusMoveColorToFirst(ByVal colors As String, ByVal selectedColor As String) As String
    Dim normalized As String
    normalized = NxFocusNormalizeStrictColors(colors, 0)
    selectedColor = NxFocusNormalizeHexColor(selectedColor)
    If Len(normalized) = 0 Or Len(selectedColor) = 0 Then Exit Function
    NxFocusMoveColorToFirst = selectedColor
    If InStr(1, "," & normalized & ",", "," & selectedColor & ",", vbTextCompare) = 0 Then _
        NxFocusMoveColorToFirst = selectedColor & "," & normalized
End Function

Public Sub NxFocusSaveColors(ByVal colors As String)
    Dim model As CNxFocusSettingsModel
    Dim normalized As String
    normalized = NxFocusNormalizeColors(colors)
    If Len(normalized) = 0 Then NxRaiseContractError "유효한 색상이 필요합니다."
    Set model = NxFocusLoadSettings()
    NxFocusSaveSettingsModel NxFocusCreateSettings(model.Shape, model.Style, CStr(Split(normalized, ",")(0)), _
        model.IntensityPercent, model.HighlightSelectedArea, model.EnableOnApply)
End Sub

Public Function NxFocusFirstColor(ByVal colors As String) As Long
    Dim normalized As String
    normalized = NxFocusNormalizeColors(colors)
    If Len(normalized) = 0 Then NxRaiseContractError "유효한 색상이 필요합니다."
    NxFocusFirstColor = NxFocusColorValue(CStr(Split(normalized, ",")(0)))
End Function
