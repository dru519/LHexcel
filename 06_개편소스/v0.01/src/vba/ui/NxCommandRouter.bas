Attribute VB_Name = "NxCommandRouter"
Option Explicit

Public Sub NxRouteRibbonTarget(ByVal target As CNxRibbonTarget)
    If target Is Nothing Then NxRaiseContractError "Ribbon target is required"
    If Not target.IsSealed Then NxRaiseContractError "Ribbon target must be sealed"
    If Not NxDistributionEntryAllowed(target.Kind, target.Identifier) Then NxDistributionEnsureExecutable
    Select Case target.Kind
        Case "entry": NxRouteEntry target.Identifier
        Case "category": NxRouteCategory target.Identifier
        Case "feature": NxRouteFeature target.Identifier
        Case "command": NxRouteCommand target.Identifier
        Case Else: NxRaiseContractError "Unknown ribbon target kind"
    End Select
End Sub

Public Sub NxRouteCommand(ByVal commandId As String)
    Dim definition As CNxCommandDefinition
    Dim canonical As String
    canonical = NxCanonicalRouteKey("command:" & commandId)
    If Left$(canonical, 8) = "feature:" Then
        NxRouteFeature Mid$(canonical, 9)
        Exit Sub
    End If
    commandId = Mid$(canonical, 9)
    NxRouteAvailabilityEnsureExecutable NxRouteKey("command", commandId)
    Set definition = NxCommandDefinition(commandId)
    NxExecuteRegisteredCommand definition
End Sub

Public Function NxCanonicalRouteKey(ByVal routeKey As String) As String
    Select Case routeKey
        Case "command:NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-COPYVISIBLE": NxCanonicalRouteKey = "feature:NX-DATA-COPY-VISIBLE"
        Case "command:NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-PASTEVISIBLEVALUES": NxCanonicalRouteKey = "feature:NX-DATA-PASTE-VISIBLE-VALUES"
        Case "command:NX-CMD-STYLE-RB-LHESTYLE-OPENOPTIONS": NxCanonicalRouteKey = "feature:NX-DRAW-ROLE-STYLE"
        Case Else: NxCanonicalRouteKey = routeKey
    End Select
End Function

Public Sub NxRouteFeature(ByVal featureId As String)
    Dim registry As CNxFeatureRegistry
    Dim definition As CNxFeatureDefinition
    NxRouteAvailabilityEnsureExecutable NxRouteKey("feature", featureId)
    Set registry = NxCreateProductRegistry()
    Set definition = registry.FeatureById(featureId)
    NxRouteFeatureSurface definition
End Sub

Private Sub NxRouteEntry(ByVal entryId As String)
    Select Case entryId
        Case "NX-MGMT-DISTRIBUTION-STATUS": MsgBox NxDistributionMessage(), vbInformation, "내엑셀 - 사용기한 및 보호 안내"
        Case "NX-ENTRY-START": NxProductSettingsOpenMenu
        Case "NX-ENTRY-AI": NxAiOpenInput
        Case "NX-ENTRY-TEMPLATE": NxRaiseContractError "템플릿 목록에서 실행할 기능을 선택하세요."
        Case "NX-ENTRY-ALL": NxProductSettingsOpenMenu
        Case "NX-ENTRY-FAVORITES": NxFavoritesOpenMenu
        Case "NX-ENTRY-MANAGEMENT": NxProductSettingsOpenMenu
        Case "NX-ENTRY-FOCUS-SETTINGS": NxFocusOpenSettings
        Case "NX-ENTRY-FOCUS-RESET": NxFocusResetDefaults
        Case "NX-MGMT-INSTALL": NxProductInstall
        Case "NX-MGMT-REMOVE": NxProductRemove
        Case "NX-MGMT-SAVED-FOLDER": NxProductOpenSavedFolder
        Case "NX-MGMT-INSTALL-FOLDER": NxProductOpenInstallFolder
        Case "NX-MGMT-QAT": NxProductOpenQatSettings
        Case "NX-MGMT-EXCEL-ADDINS": NxProductOpenExcelAddInsSettings
        Case "NX-MGMT-FILE-INFO": NxProductShowFileInfo
        Case "NX-MGMT-SHORTCUTS": NxProductOpenShortcutManager
        Case "NX-ENTRY-HANGUL-SETTINGS": NxHangulOpenSettings
        Case Else: NxRaiseContractError "Unknown ribbon entry"
    End Select
End Sub

Public Sub NxRouteCategory(ByVal categoryId As String, Optional ByVal featureId As String = vbNullString)
    Select Case categoryId
        Case "NX-CAT-AI", "NX-CAT-TEMPLATE", "NX-CAT-DATA", _
             "NX-CAT-DRAW", "NX-CAT-FILE", "NX-CAT-SYMBOLS", "NX-CAT-CALCULATOR", "NX-CAT-UTIL"
            NxRaiseContractError "전체기능 목록에서 실행할 기능을 선택하세요."
        Case Else: NxRaiseContractError "Unknown product category"
    End Select
End Sub
