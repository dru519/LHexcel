Attribute VB_Name = "NxShortcuts"
Option Explicit

Private Const NX_SHORTCUT_FILE As String = "shortcuts-v2.cfg"
Private Const NX_SHORTCUT_VERSION As String = "NXSHORT2"
Private Const NX_SHORTCUT_LEGACY_FILE As String = "shortcuts-v1.cfg"
Private Const NX_SHORTCUT_LEGACY_VERSION As String = "NXSHORT1"
Private Const NX_SHORTCUT_SLOT_COUNT As Long = 64

Private mRouteKeys(1 To NX_SHORTCUT_SLOT_COUNT) As String
Private mModifiers(1 To NX_SHORTCUT_SLOT_COUNT) As String
Private mKeys(1 To NX_SHORTCUT_SLOT_COUNT) As String
Private mLoaded As Boolean
Private mMigratedFromV1 As Boolean

Public Sub NxShortcutsOpenManager()
    NxShortcutsShowManager
End Sub

Public Sub NxShortcutsShowManager()
    Dim manager As New FNxShortcutManager
    manager.Show
End Sub

' Dialog drafts never change the saved arrays, profile or Excel OnKey bindings.
Public Function NxShortcutsCreateDraft() As Object
    Dim draft As Object, slot As Long
    EnsureLoaded
    Set draft = CreateObject("Scripting.Dictionary")
    draft.CompareMode = vbBinaryCompare
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        If Len(mRouteKeys(slot)) > 0 Then draft.Add mRouteKeys(slot), Array(mModifiers(slot), mKeys(slot))
    Next slot
    Set NxShortcutsCreateDraft = draft
End Function

Public Function NxShortcutsSnapshot() As String
    Dim slot As Long
    EnsureLoaded
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        NxShortcutsSnapshot = NxShortcutsSnapshot & CStr(slot) & "|" & mRouteKeys(slot) & "|" & mModifiers(slot) & "|" & mKeys(slot) & vbLf
    Next slot
End Function

Public Function NxShortcutDraftDisplay(ByVal draft As Object, ByVal routeKey As String) As String
    Dim binding As Variant
    If draft Is Nothing Then NxRaiseContractError "단축키 편집 상태가 없습니다."
    If Not draft.Exists(routeKey) Then Exit Function
    binding = draft.Item(routeKey)
    NxShortcutDraftDisplay = NxShortcutDisplayKey(CStr(binding(0)), CStr(binding(1)))
End Function

Public Function NxShortcutDraftConflictReason(ByVal draft As Object, ByVal modifierName As String, _
    ByVal keyName As String, Optional ByVal routeKey As String = vbNullString) As String
    Dim candidate As Variant, binding As Variant
    Dim existingModifier As String, existingKey As String
    On Error GoTo Invalid
    If draft Is Nothing Then NxRaiseContractError "단축키 편집 상태가 없습니다."
    NxShortcutNormalizeChord modifierName, keyName
    If NxShortcutIsReservedCombination(modifierName, keyName) Then
        NxShortcutDraftConflictReason = "Excel 기본 동작과 충돌하는 예약 단축키입니다."
        Exit Function
    End If
    For Each candidate In draft.Keys
        If StrComp(CStr(candidate), routeKey, vbBinaryCompare) <> 0 Then
            binding = draft.Item(candidate)
            existingModifier = CStr(binding(0))
            existingKey = CStr(binding(1))
            NxShortcutNormalizeChord existingModifier, existingKey
            If existingModifier = modifierName And existingKey = keyName Then
                NxShortcutDraftConflictReason = "이미 " & NxShortcutsRouteLabel(CStr(candidate)) & " 항목에서 사용 중입니다."
                Exit Function
            End If
        End If
    Next candidate
    Exit Function
Invalid:
    NxShortcutDraftConflictReason = Err.Description
    Err.Clear
End Function

Public Sub NxShortcutDraftAssign(ByVal draft As Object, ByVal routeKey As String, _
    ByVal modifierName As String, ByVal keyName As String)
    Dim reason As String
    If draft Is Nothing Then NxRaiseContractError "단축키 편집 상태가 없습니다."
    If Not NxShortcutsRouteIsEligible(routeKey) Then NxRaiseContractError "이 항목은 단축키를 지정할 수 없습니다."
    reason = NxShortcutDraftConflictReason(draft, modifierName, keyName, routeKey)
    If Len(reason) > 0 Then NxRaiseContractError reason
    If Not draft.Exists(routeKey) Then
        If draft.Count >= NX_SHORTCUT_SLOT_COUNT Then NxRaiseContractError "단축키는 최대 64개까지 지정할 수 있습니다."
    End If
    NxShortcutNormalizeChord modifierName, keyName
    draft.Item(routeKey) = Array(modifierName, keyName)
End Sub

Public Sub NxShortcutDraftRemove(ByVal draft As Object, ByVal routeKey As String)
    If draft Is Nothing Then NxRaiseContractError "단축키 편집 상태가 없습니다."
    If Not draft.Exists(routeKey) Then NxRaiseContractError "선택한 기능에 지정된 단축키가 없습니다."
    draft.Remove routeKey
End Sub

Public Sub NxShortcutsApplyDraft(ByVal draft As Object, ByVal expectedSnapshot As String)
    Dim oldRoutes(1 To NX_SHORTCUT_SLOT_COUNT) As String
    Dim oldModifiers(1 To NX_SHORTCUT_SLOT_COUNT) As String
    Dim oldKeys(1 To NX_SHORTCUT_SLOT_COUNT) As String
    Dim validated As Object, route As Variant, binding As Variant, slot As Long
    Dim failureNumber As Long, failureDescription As String, restoreError As String
    EnsureLoaded
    If StrComp(NxShortcutsSnapshot(), expectedSnapshot, vbBinaryCompare) <> 0 Then _
        NxRaiseContractError "다른 작업에서 단축키 설정이 변경되었습니다. 창을 다시 열어 확인하세요."
    If draft Is Nothing Then NxRaiseContractError "단축키 편집 상태가 없습니다."
    Set validated = CreateObject("Scripting.Dictionary")
    validated.CompareMode = vbBinaryCompare
    ' Complete validation, including cross-draft conflicts, before any mutation.
    For Each route In draft.Keys
        binding = draft.Item(route)
        NxShortcutDraftAssign validated, CStr(route), CStr(binding(0)), CStr(binding(1))
    Next route
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        oldRoutes(slot) = mRouteKeys(slot)
        oldModifiers(slot) = mModifiers(slot)
        oldKeys(slot) = mKeys(slot)
    Next slot
    On Error GoTo Rollback
    NxShortcutsClearCurrentBindings
    Erase mRouteKeys
    Erase mModifiers
    Erase mKeys
    slot = 0
    For Each route In validated.Keys
        slot = slot + 1
        binding = validated.Item(route)
        mRouteKeys(slot) = CStr(route)
        mModifiers(slot) = CStr(binding(0))
        mKeys(slot) = CStr(binding(1))
    Next route
    NxShortcutsApplyCurrentBindings
    NxShortcutsSave
    Exit Sub
Rollback:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    NxShortcutsClearCurrentBindings
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        mRouteKeys(slot) = oldRoutes(slot)
        mModifiers(slot) = oldModifiers(slot)
        mKeys(slot) = oldKeys(slot)
    Next slot
    Err.Clear
    NxShortcutsApplyCurrentBindings
    If Err.Number <> 0 Then restoreError = " 이전 실행키 복구도 실패했습니다. Excel 재시작 후 설정을 확인하세요."
    On Error GoTo 0
    Err.Raise failureNumber, "NxShortcuts.ApplyDraft", failureDescription & restoreError
End Sub

Public Function NxShortcutsEligibleRoutes() As Collection
    Dim item As Variant
    Dim result As New Collection
    For Each item In NxGeneratedNavigationItems()
        If NxShortcutsRouteIsEligible(CStr(item(0))) Then result.Add item
    Next item
    Set NxShortcutsEligibleRoutes = result
End Function

Public Function NxShortcutsRouteIsEligible(ByVal routeKey As String) As Boolean
    Dim definition As CNxCommandDefinition
    On Error GoTo Ineligible
    If Not NxGeneratedNavigationRouteExists(routeKey) Then Exit Function
    If routeKey = "feature:NX-DATA-DUPLICATE-LIST" Then Exit Function
    If NxGeneratedNavigationRouteField(routeKey, "item_type") = "command" Then
        Set definition = NxCommandDefinition(NxGeneratedNavigationRouteField(routeKey, "id"))
        If Not definition.ShortcutEligible Then Exit Function
    End If
    NxShortcutsRouteIsEligible = True
Ineligible:
End Function

Public Function NxShortcutConflictReason(ByVal modifierName As String, ByVal keyName As String, _
    Optional ByVal routeKey As String = vbNullString) As String
    Dim slot As Long
    On Error GoTo Invalid
    EnsureLoaded
    NxShortcutNormalizeChord modifierName, keyName
    If NxShortcutIsReservedCombination(modifierName, keyName) Then
        NxShortcutConflictReason = "Excel 기본 동작과 충돌하는 예약 단축키입니다."
        Exit Function
    End If
    slot = NxShortcutFindKeySlot(modifierName, keyName)
    ' VBA And/Or evaluate both operands; an unused chord has no array slot.
    If slot = 0 Then Exit Function
    If Len(routeKey) = 0 Or StrComp(mRouteKeys(slot), routeKey, vbBinaryCompare) <> 0 Then
        NxShortcutConflictReason = "이미 " & NxShortcutsRouteLabel(mRouteKeys(slot)) & " 항목에서 사용 중입니다."
    End If
    Exit Function
Invalid:
    NxShortcutConflictReason = Err.Description
    Err.Clear
End Function

Public Function NxShortcutCaptureKey(ByVal keyCode As Long, ByVal shift As Integer, _
    ByRef modifierName As String, ByRef keyName As String) As Boolean
    On Error GoTo Unsupported
    modifierName = vbNullString
    keyName = vbNullString
    If (shift And 2) <> 0 Then modifierName = "Ctrl+"
    If (shift And 4) <> 0 Then modifierName = modifierName & "Alt+"
    If (shift And 1) <> 0 Then modifierName = modifierName & "Shift+"
    If Len(modifierName) > 0 Then modifierName = Left$(modifierName, Len(modifierName) - 1)
    Select Case keyCode
        Case vbKeyA To vbKeyZ: keyName = Chr$(keyCode)
        Case vbKey0 To vbKey9: keyName = Chr$(keyCode)
        ' Keep the existing digit alias; do not silently change saved numpad bindings.
        Case vbKeyNumpad0 To vbKeyNumpad9: keyName = CStr(keyCode - vbKeyNumpad0)
        Case vbKeyF1 To 126: keyName = "F" & CStr(keyCode - vbKeyF1 + 1)
        Case 186: keyName = "Semicolon"
        Case 187: keyName = "Equals"
        Case 188: keyName = "Comma"
        Case 189: keyName = "Minus"
        Case 190: keyName = "Period"
        Case 191: keyName = "Slash"
        Case 192: keyName = "Grave"
        Case 219: keyName = "LeftBracket"
        Case 220: keyName = "Backslash"
        Case 221: keyName = "RightBracket"
        Case 222: keyName = "Apostrophe"
        Case vbKeyTab: keyName = "Tab"
        Case vbKeyReturn: keyName = "Enter"
        Case vbKeyEscape: keyName = "Escape"
        Case vbKeySpace: keyName = "Space"
        Case vbKeyBack: keyName = "Backspace"
        Case vbKeyDelete: keyName = "Delete"
        Case vbKeyInsert: keyName = "Insert"
        Case vbKeyHome: keyName = "Home"
        Case vbKeyEnd: keyName = "End"
        Case vbKeyPageUp: keyName = "PageUp"
        Case vbKeyPageDown: keyName = "PageDown"
        Case vbKeyLeft: keyName = "Left"
        Case vbKeyRight: keyName = "Right"
        Case vbKeyUp: keyName = "Up"
        Case vbKeyDown: keyName = "Down"
        Case Else: GoTo Unsupported
    End Select
    NxShortcutNormalizeChord modifierName, keyName
    NxShortcutCaptureKey = True
    Exit Function
Unsupported:
    modifierName = vbNullString
    keyName = vbNullString
    Err.Clear
End Function

Public Function NxShortcutCaptureDisplay(ByVal keyCode As Long, ByVal shift As Integer) As String
    Dim keyName As String
    Dim modifierName As String
    If NxShortcutCaptureKey(keyCode, shift, modifierName, keyName) Then _
        NxShortcutCaptureDisplay = NxShortcutDisplayKey(modifierName, keyName)
End Function

Public Function NxShortcutIsReservedCombination(ByVal modifierName As String, ByVal keyName As String) As Boolean
    Dim combination As String
    NxShortcutNormalizeChord modifierName, keyName
    combination = modifierName & "+" & keyName
    ' Fail closed outside the reviewed key families. Navigation, editing, menu and
    ' Windows system chords are never assignable, even with additional modifiers.
    Select Case keyName
        Case "Tab", "Enter", "Escape", "Space", "Backspace", "Delete", "Insert", _
             "Home", "End", "PageUp", "PageDown", "Left", "Right", "Up", "Down"
            NxShortcutIsReservedCombination = True
            Exit Function
    End Select
    ' F13-F15 are documented OnKey keys, with no native Excel action in the
    ' reviewed Windows shortcut table. F1-F12 retain their native families.
    If keyName = "F13" Or keyName = "F14" Or keyName = "F15" Then Exit Function
    If modifierName = "None" Or modifierName = "Shift" Or modifierName = "Alt" Or _
        modifierName = "Ctrl" Or modifierName = "Alt+Shift" Then
        NxShortcutIsReservedCombination = True
        Exit Function
    End If
    ' Microsoft Windows Excel shortcuts (including contextual formula/formatting
    ' actions); no warning-only override. Plain Ctrl and Alt/Ribbon are blocked
    ' as complete families because their native list depends on locale/context.
    ' https://support.microsoft.com/en-us/office/1798d9d5-842a-42b8-9c99-9b7213f0040f
    Select Case combination
        Case "Ctrl+C"
            NxShortcutIsReservedCombination = True
        Case "Ctrl+Alt+V", "Ctrl+Alt+P", "Ctrl+Alt+F5", "Ctrl+Alt+F9", _
             "Ctrl+Alt+5", "Ctrl+Alt+Equals", "Ctrl+Alt+Minus", _
             "Ctrl+Alt+Shift+Equals", "Ctrl+Alt+Shift+Minus"
            NxShortcutIsReservedCombination = True
        Case "Ctrl+Shift+A", "Ctrl+Shift+C", "Ctrl+Shift+F", _
             "Ctrl+Shift+G", "Ctrl+Shift+L", "Ctrl+Shift+M", _
             "Ctrl+Shift+O", "Ctrl+Shift+P", "Ctrl+Shift+U", "Ctrl+Shift+V", _
             "Ctrl+Shift+F1", "Ctrl+Shift+F2", "Ctrl+Shift+F3", "Ctrl+Shift+F4", _
             "Ctrl+Shift+F5", "Ctrl+Shift+F6", "Ctrl+Shift+F10", "Ctrl+Shift+F11", "Ctrl+Shift+F12", _
             "Ctrl+Shift+0", "Ctrl+Shift+1", "Ctrl+Shift+2", _
             "Ctrl+Shift+3", "Ctrl+Shift+4", "Ctrl+Shift+5", _
             "Ctrl+Shift+6", "Ctrl+Shift+7", "Ctrl+Shift+8", "Ctrl+Shift+9", _
             "Ctrl+Alt+Shift+F2", "Ctrl+Alt+Shift+F9", _
             "Ctrl+Alt+Shift+C", "Ctrl+Alt+Shift+H", "Ctrl+Alt+Shift+M", _
             "Ctrl+Alt+Shift+P", "Ctrl+Alt+Shift+T"
            NxShortcutIsReservedCombination = True
    End Select
    ' Ctrl+Shift punctuation includes time/copy/insert/outline/precedent actions;
    ' reserve the entire family so shifted symbol aliases cannot bypass it.
    If modifierName = "Ctrl+Shift" And Len(keyName) > 1 And Left$(keyName, 1) <> "F" Then _
        NxShortcutIsReservedCombination = True
End Function

Public Sub NxShortcutsApplySavedBindings()
    If Not NxDistributionCanExecute() Then Exit Sub
    EnsureLoaded
    NxShortcutsClearCurrentBindings
    NxShortcutsApplyCurrentBindings
End Sub

Public Sub NxShortcutsAssign(ByVal routeKey As String, ByVal modifierName As String, ByVal keyName As String)
    Dim existingKeySlot As Long
    Dim existingRouteSlot As Long
    Dim failureDescription As String
    Dim failureNumber As Long
    Dim failureSource As String
    Dim previousKey As String
    Dim previousModifier As String
    Dim previousRouteKey As String
    Dim reason As String
    Dim slot As Long

    EnsureLoaded
    If Not NxShortcutsRouteIsEligible(routeKey) Then NxRaiseContractError "이 항목은 단축키를 지정할 수 없습니다."
    NxShortcutNormalizeChord modifierName, keyName
    If NxShortcutIsReservedCombination(modifierName, keyName) Then _
        NxRaiseContractError "Excel 기본 동작과 충돌하는 예약 단축키입니다."
    existingRouteSlot = NxShortcutFindRouteSlot(routeKey)
    existingKeySlot = NxShortcutFindKeySlot(modifierName, keyName)
    If existingKeySlot > 0 And existingKeySlot <> existingRouteSlot Then
        reason = NxShortcutConflictReason(modifierName, keyName, routeKey)
        If Len(reason) = 0 Then reason = "이미 사용 중인 단축키입니다."
        NxRaiseContractError reason
    End If
    If existingRouteSlot > 0 Then slot = existingRouteSlot Else slot = NxShortcutFirstFreeSlot()
    If slot = 0 Then NxRaiseContractError "단축키는 최대 64개까지 지정할 수 있습니다."

    previousRouteKey = mRouteKeys(slot)
    previousModifier = mModifiers(slot)
    previousKey = mKeys(slot)
    On Error GoTo Rollback
    NxShortcutsClearCurrentBindings
    mRouteKeys(slot) = routeKey
    mModifiers(slot) = modifierName
    mKeys(slot) = keyName
    NxShortcutsApplyCurrentBindings
    NxShortcutsSave
    Exit Sub

Rollback:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    NxShortcutsClearCurrentBindings
    mRouteKeys(slot) = previousRouteKey
    mModifiers(slot) = previousModifier
    mKeys(slot) = previousKey
    NxShortcutsApplyCurrentBindings
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Sub

Public Sub NxShortcutsRemoveBindingAt(ByVal position As Long)
    NxShortcutsRemoveSlot NxShortcutsBindingSlotAt(position)
End Sub

Public Sub NxShortcutsRemoveRoute(ByVal routeKey As String)
    Dim slot As Long
    EnsureLoaded
    slot = NxShortcutFindRouteSlot(routeKey)
    If slot = 0 Then NxRaiseContractError "선택한 기능에 지정된 단축키가 없습니다."
    NxShortcutsRemoveSlot slot
End Sub

Private Sub NxShortcutsRemoveSlot(ByVal slot As Long)
    Dim failureDescription As String
    Dim failureNumber As Long
    Dim failureSource As String
    Dim previousKey As String
    Dim previousModifier As String
    Dim previousRouteKey As String
    EnsureLoaded
    If slot < 1 Or slot > NX_SHORTCUT_SLOT_COUNT Then NxRaiseContractError "해제할 단축키를 찾을 수 없습니다."
    previousRouteKey = mRouteKeys(slot)
    previousModifier = mModifiers(slot)
    previousKey = mKeys(slot)
    On Error GoTo Rollback
    NxShortcutsClearCurrentBindings
    mRouteKeys(slot) = vbNullString
    mModifiers(slot) = vbNullString
    mKeys(slot) = vbNullString
    NxShortcutsApplyCurrentBindings
    NxShortcutsSave
    Exit Sub

Rollback:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    NxShortcutsClearCurrentBindings
    mRouteKeys(slot) = previousRouteKey
    mModifiers(slot) = previousModifier
    mKeys(slot) = previousKey
    NxShortcutsApplyCurrentBindings
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Sub

Public Function NxShortcutsBindingCount() As Long
    Dim slot As Long
    EnsureLoaded
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        If Len(mRouteKeys(slot)) > 0 Then NxShortcutsBindingCount = NxShortcutsBindingCount + 1
    Next slot
End Function

Public Function NxShortcutsBindingSlotAt(ByVal position As Long) As Long
    Dim current As Long
    Dim slot As Long
    EnsureLoaded
    If position < 1 Then NxRaiseContractError "해제할 단축키를 선택하세요."
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        If Len(mRouteKeys(slot)) > 0 Then
            current = current + 1
            If current = position Then NxShortcutsBindingSlotAt = slot: Exit Function
        End If
    Next slot
    NxRaiseContractError "선택한 단축키를 찾을 수 없습니다."
End Function

Public Function NxShortcutsBindingRouteKey(ByVal position As Long) As String
    NxShortcutsBindingRouteKey = mRouteKeys(NxShortcutsBindingSlotAt(position))
End Function

' Compatibility alias retained for older form attestations.
Public Function NxShortcutsBindingCommandId(ByVal position As Long) As String
    NxShortcutsBindingCommandId = NxGeneratedNavigationRouteField(NxShortcutsBindingRouteKey(position), "id")
End Function

Public Function NxShortcutsBindingDisplayKey(ByVal position As Long) As String
    Dim slot As Long
    slot = NxShortcutsBindingSlotAt(position)
    NxShortcutsBindingDisplayKey = NxShortcutDisplayKey(mModifiers(slot), mKeys(slot))
End Function

Public Function NxShortcutsTryGetRouteBinding(ByVal routeKey As String, _
    ByRef modifierName As String, ByRef keyName As String) As Boolean
    Dim slot As Long
    modifierName = vbNullString
    keyName = vbNullString
    EnsureLoaded
    slot = NxShortcutFindRouteSlot(routeKey)
    If slot = 0 Then Exit Function
    modifierName = mModifiers(slot)
    keyName = mKeys(slot)
    NxShortcutsTryGetRouteBinding = True
End Function

Public Function NxShortcutsBindingDisplayForRoute(ByVal routeKey As String) As String
    Dim keyName As String
    Dim modifierName As String
    If NxShortcutsTryGetRouteBinding(routeKey, modifierName, keyName) Then _
        NxShortcutsBindingDisplayForRoute = NxShortcutDisplayKey(modifierName, keyName)
End Function

Public Function NxShortcutsRouteLabel(ByVal routeKey As String) As String
    If NxGeneratedNavigationRouteExists(routeKey) Then
        NxShortcutsRouteLabel = NxGeneratedNavigationRouteField(routeKey, "label_ko")
    Else
        NxShortcutsRouteLabel = "사용할 수 없는 항목 · " & routeKey
    End If
End Function

Public Sub NxShortcutsShutdown()
    If Not mLoaded Then Exit Sub
    NxShortcutsClearCurrentBindings
End Sub

Public Sub Auto_Close()
    NxDistributionStop
    NxShortcutsShutdown
    NxRouteAvailabilityStopEvents
    NxNavigatorShutdown
    NxDocumentNavigatorShutdown
    NxHostBridgeReset
End Sub

Public Sub NxShortcutDispatchSlot(ByVal slot As Long)
    Dim identifier As String
    Dim itemType As String
    Dim routeKey As String
    EnsureLoaded
    If slot < 1 Or slot > NX_SHORTCUT_SLOT_COUNT Then NxRaiseContractError "단축키 슬롯이 올바르지 않습니다."
    routeKey = mRouteKeys(slot)
    If Len(routeKey) = 0 Then NxRaiseContractError "지정되지 않은 단축키입니다."
    If Not NxShortcutsRouteIsEligible(routeKey) Then NxRaiseContractError "현재 버전에서 실행할 수 없는 단축키 항목입니다."
    itemType = NxGeneratedNavigationRouteField(routeKey, "item_type")
    identifier = NxGeneratedNavigationRouteField(routeKey, "id")
    Select Case itemType
        Case "feature": NxRouteFeature identifier
        Case "command": NxRouteCommand identifier
        Case Else: NxRaiseContractError "지원하지 않는 단축키 경로입니다."
    End Select
End Sub

Public Sub NxShortcutDispatch01(): NxShortcutDispatchSlot 1: End Sub
Public Sub NxShortcutDispatch02(): NxShortcutDispatchSlot 2: End Sub
Public Sub NxShortcutDispatch03(): NxShortcutDispatchSlot 3: End Sub
Public Sub NxShortcutDispatch04(): NxShortcutDispatchSlot 4: End Sub
Public Sub NxShortcutDispatch05(): NxShortcutDispatchSlot 5: End Sub
Public Sub NxShortcutDispatch06(): NxShortcutDispatchSlot 6: End Sub
Public Sub NxShortcutDispatch07(): NxShortcutDispatchSlot 7: End Sub
Public Sub NxShortcutDispatch08(): NxShortcutDispatchSlot 8: End Sub
Public Sub NxShortcutDispatch09(): NxShortcutDispatchSlot 9: End Sub
Public Sub NxShortcutDispatch10(): NxShortcutDispatchSlot 10: End Sub
Public Sub NxShortcutDispatch11(): NxShortcutDispatchSlot 11: End Sub
Public Sub NxShortcutDispatch12(): NxShortcutDispatchSlot 12: End Sub
Public Sub NxShortcutDispatch13(): NxShortcutDispatchSlot 13: End Sub
Public Sub NxShortcutDispatch14(): NxShortcutDispatchSlot 14: End Sub
Public Sub NxShortcutDispatch15(): NxShortcutDispatchSlot 15: End Sub
Public Sub NxShortcutDispatch16(): NxShortcutDispatchSlot 16: End Sub
Public Sub NxShortcutDispatch17(): NxShortcutDispatchSlot 17: End Sub
Public Sub NxShortcutDispatch18(): NxShortcutDispatchSlot 18: End Sub
Public Sub NxShortcutDispatch19(): NxShortcutDispatchSlot 19: End Sub
Public Sub NxShortcutDispatch20(): NxShortcutDispatchSlot 20: End Sub
Public Sub NxShortcutDispatch21(): NxShortcutDispatchSlot 21: End Sub
Public Sub NxShortcutDispatch22(): NxShortcutDispatchSlot 22: End Sub
Public Sub NxShortcutDispatch23(): NxShortcutDispatchSlot 23: End Sub
Public Sub NxShortcutDispatch24(): NxShortcutDispatchSlot 24: End Sub
Public Sub NxShortcutDispatch25(): NxShortcutDispatchSlot 25: End Sub
Public Sub NxShortcutDispatch26(): NxShortcutDispatchSlot 26: End Sub
Public Sub NxShortcutDispatch27(): NxShortcutDispatchSlot 27: End Sub
Public Sub NxShortcutDispatch28(): NxShortcutDispatchSlot 28: End Sub
Public Sub NxShortcutDispatch29(): NxShortcutDispatchSlot 29: End Sub
Public Sub NxShortcutDispatch30(): NxShortcutDispatchSlot 30: End Sub
Public Sub NxShortcutDispatch31(): NxShortcutDispatchSlot 31: End Sub
Public Sub NxShortcutDispatch32(): NxShortcutDispatchSlot 32: End Sub
Public Sub NxShortcutDispatch33(): NxShortcutDispatchSlot 33: End Sub
Public Sub NxShortcutDispatch34(): NxShortcutDispatchSlot 34: End Sub
Public Sub NxShortcutDispatch35(): NxShortcutDispatchSlot 35: End Sub
Public Sub NxShortcutDispatch36(): NxShortcutDispatchSlot 36: End Sub
Public Sub NxShortcutDispatch37(): NxShortcutDispatchSlot 37: End Sub
Public Sub NxShortcutDispatch38(): NxShortcutDispatchSlot 38: End Sub
Public Sub NxShortcutDispatch39(): NxShortcutDispatchSlot 39: End Sub
Public Sub NxShortcutDispatch40(): NxShortcutDispatchSlot 40: End Sub
Public Sub NxShortcutDispatch41(): NxShortcutDispatchSlot 41: End Sub
Public Sub NxShortcutDispatch42(): NxShortcutDispatchSlot 42: End Sub
Public Sub NxShortcutDispatch43(): NxShortcutDispatchSlot 43: End Sub
Public Sub NxShortcutDispatch44(): NxShortcutDispatchSlot 44: End Sub
Public Sub NxShortcutDispatch45(): NxShortcutDispatchSlot 45: End Sub
Public Sub NxShortcutDispatch46(): NxShortcutDispatchSlot 46: End Sub
Public Sub NxShortcutDispatch47(): NxShortcutDispatchSlot 47: End Sub
Public Sub NxShortcutDispatch48(): NxShortcutDispatchSlot 48: End Sub
Public Sub NxShortcutDispatch49(): NxShortcutDispatchSlot 49: End Sub
Public Sub NxShortcutDispatch50(): NxShortcutDispatchSlot 50: End Sub
Public Sub NxShortcutDispatch51(): NxShortcutDispatchSlot 51: End Sub
Public Sub NxShortcutDispatch52(): NxShortcutDispatchSlot 52: End Sub
Public Sub NxShortcutDispatch53(): NxShortcutDispatchSlot 53: End Sub
Public Sub NxShortcutDispatch54(): NxShortcutDispatchSlot 54: End Sub
Public Sub NxShortcutDispatch55(): NxShortcutDispatchSlot 55: End Sub
Public Sub NxShortcutDispatch56(): NxShortcutDispatchSlot 56: End Sub
Public Sub NxShortcutDispatch57(): NxShortcutDispatchSlot 57: End Sub
Public Sub NxShortcutDispatch58(): NxShortcutDispatchSlot 58: End Sub
Public Sub NxShortcutDispatch59(): NxShortcutDispatchSlot 59: End Sub
Public Sub NxShortcutDispatch60(): NxShortcutDispatchSlot 60: End Sub
Public Sub NxShortcutDispatch61(): NxShortcutDispatchSlot 61: End Sub
Public Sub NxShortcutDispatch62(): NxShortcutDispatchSlot 62: End Sub
Public Sub NxShortcutDispatch63(): NxShortcutDispatchSlot 63: End Sub
Public Sub NxShortcutDispatch64(): NxShortcutDispatchSlot 64: End Sub

Private Sub EnsureLoaded()
    Dim slot As Long
    If mLoaded Then Exit Sub
    mLoaded = True
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        mRouteKeys(slot) = vbNullString
        mModifiers(slot) = vbNullString
        mKeys(slot) = vbNullString
    Next slot
    NxShortcutsLoad
End Sub

Private Sub NxShortcutsLoad()
    If NxShortcutFileExists(NxShortcutsPath()) Then
        NxShortcutsLoadFile NxShortcutsPath(), False
    ElseIf NxShortcutFileExists(NxShortcutsLegacyPath()) Then
        NxShortcutsLoadFile NxShortcutsLegacyPath(), True
    End If
End Sub

Private Sub NxShortcutsLoadFile(ByVal path As String, ByVal legacy As Boolean)
    Dim expectedHeader As String
    Dim handle As Integer
    Dim header As String
    Dim lineText As String
    On Error GoTo Rejected
    If legacy Then expectedHeader = NX_SHORTCUT_LEGACY_VERSION Else expectedHeader = NX_SHORTCUT_VERSION
    handle = FreeFile
    Open path For Input Access Read Lock Write As #handle
    If EOF(handle) Then GoTo Rejected
    Line Input #handle, header
    If StrComp(header, expectedHeader, vbBinaryCompare) <> 0 Then GoTo Rejected
    Do While Not EOF(handle)
        Line Input #handle, lineText
        NxShortcutsTryLoadRow lineText, legacy
    Loop
    Close #handle
    mMigratedFromV1 = legacy
    Exit Sub
Rejected:
    On Error Resume Next
    If handle > 0 Then Close #handle
    On Error GoTo 0
End Sub

Private Sub NxShortcutsTryLoadRow(ByVal lineText As String, ByVal legacy As Boolean)
    Dim keyName As String
    Dim modifierName As String
    Dim parts As Variant
    Dim routeKey As String
    Dim slot As Long
    On Error GoTo Rejected
    parts = Split(lineText, "|", -1, vbBinaryCompare)
    If UBound(parts) <> 3 Then Exit Sub
    If Not IsNumeric(parts(0)) Then Exit Sub
    slot = CLng(parts(0))
    If slot < 1 Or slot > NX_SHORTCUT_SLOT_COUNT Then Exit Sub
    If Len(mRouteKeys(slot)) > 0 Then Exit Sub
    modifierName = CStr(parts(1))
    keyName = CStr(parts(2))
    NxShortcutNormalizeChord modifierName, keyName
    If NxShortcutIsReservedCombination(modifierName, keyName) Then Exit Sub
    If legacy Then routeKey = "command:" & CStr(parts(3)) Else routeKey = CStr(parts(3))
    routeKey = NxCanonicalRouteKey(routeKey)
    If Not NxShortcutsRouteIsEligible(routeKey) Then Exit Sub
    If NxShortcutFindKeySlot(modifierName, keyName) > 0 Then Exit Sub
    If NxShortcutFindRouteSlot(routeKey) > 0 Then Exit Sub
    mRouteKeys(slot) = routeKey
    mModifiers(slot) = modifierName
    mKeys(slot) = keyName
    Exit Sub
Rejected:
    Err.Clear
End Sub

Private Sub NxShortcutsSave()
    Dim backup As String
    Dim failureDescription As String
    Dim failureNumber As Long
    Dim failureSource As String
    Dim handle As Integer
    Dim movedOriginal As Boolean
    Dim path As String
    Dim slot As Long
    Dim temporary As String
    On Error GoTo Failed
    NxShortcutsEnsureSettingsFolder
    path = NxShortcutsPath()
    temporary = path & ".tmp"
    backup = path & ".bak"
    If NxShortcutFileExists(temporary) Then Kill temporary
    If NxShortcutFileExists(backup) Then Kill backup

    handle = FreeFile
    Open temporary For Output Access Write Lock Read Write As #handle
    Print #handle, NX_SHORTCUT_VERSION
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        If Len(mRouteKeys(slot)) > 0 Then
            If Not NxShortcutsRouteIsEligible(mRouteKeys(slot)) Then NxRaiseContractError "저장할 수 없는 단축키 항목입니다."
            Print #handle, CStr(slot) & "|" & mModifiers(slot) & "|" & mKeys(slot) & "|" & mRouteKeys(slot)
        End If
    Next slot
    Close #handle
    handle = 0

    If NxShortcutFileExists(path) Then
        Name path As backup
        movedOriginal = True
    End If
    Name temporary As path
    On Error Resume Next
    If NxShortcutFileExists(backup) Then Kill backup
    On Error GoTo 0
    mMigratedFromV1 = False
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    If handle > 0 Then Close #handle
    If movedOriginal And Not NxShortcutFileExists(path) Then
        If NxShortcutFileExists(backup) Then Name backup As path
    End If
    If Len(temporary) > 0 And NxShortcutFileExists(temporary) Then Kill temporary
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Sub

Private Sub NxShortcutsApplyCurrentBindings()
    Dim slot As Long
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        If Len(mRouteKeys(slot)) > 0 Then NxShortcutApplySlot slot
    Next slot
End Sub

Private Sub NxShortcutApplySlot(ByVal slot As Long)
    Dim macroName As String
    macroName = "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!" & NxShortcutMacroName(slot)
    Application.OnKey NxShortcutOnKeySpec(mModifiers(slot), mKeys(slot)), macroName
End Sub

Private Sub NxShortcutsClearCurrentBindings()
    Dim slot As Long
    On Error Resume Next
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        If Len(mRouteKeys(slot)) > 0 Then Application.OnKey NxShortcutOnKeySpec(mModifiers(slot), mKeys(slot))
    Next slot
    On Error GoTo 0
End Sub

Private Function NxShortcutFindRouteSlot(ByVal routeKey As String) As Long
    Dim slot As Long
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        If StrComp(mRouteKeys(slot), routeKey, vbBinaryCompare) = 0 Then NxShortcutFindRouteSlot = slot: Exit Function
    Next slot
End Function

Private Function NxShortcutFindKeySlot(ByVal modifierName As String, ByVal keyName As String) As Long
    Dim slot As Long
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        If StrComp(mModifiers(slot), modifierName, vbBinaryCompare) = 0 And _
            StrComp(mKeys(slot), keyName, vbBinaryCompare) = 0 Then
            NxShortcutFindKeySlot = slot
            Exit Function
        End If
    Next slot
End Function

Private Function NxShortcutFirstFreeSlot() As Long
    Dim slot As Long
    For slot = 1 To NX_SHORTCUT_SLOT_COUNT
        If Len(mRouteKeys(slot)) = 0 Then NxShortcutFirstFreeSlot = slot: Exit Function
    Next slot
End Function

Private Function NxShortcutNormalizeModifier(ByVal value As String) As String
    Dim part As Variant, ctrl As Boolean, alt As Boolean, shift As Boolean
    value = UCase$(Replace$(Trim$(value), " ", vbNullString))
    If Len(value) = 0 Or value = "NONE" Then NxShortcutNormalizeModifier = "None": Exit Function
    For Each part In Split(value, "+")
        Select Case CStr(part)
            Case "CTRL", "CONTROL"
                If ctrl Then NxRaiseContractError "같은 보조키를 두 번 지정할 수 없습니다."
                ctrl = True
            Case "ALT"
                If alt Then NxRaiseContractError "같은 보조키를 두 번 지정할 수 없습니다."
                alt = True
            Case "SHIFT"
                If shift Then NxRaiseContractError "같은 보조키를 두 번 지정할 수 없습니다."
                shift = True
            Case Else: NxRaiseContractError "보조키는 Ctrl, Alt, Shift 조합을 사용하세요."
        End Select
    Next part
    If ctrl Then NxShortcutNormalizeModifier = "Ctrl+"
    If alt Then NxShortcutNormalizeModifier = NxShortcutNormalizeModifier & "Alt+"
    If shift Then NxShortcutNormalizeModifier = NxShortcutNormalizeModifier & "Shift+"
    NxShortcutNormalizeModifier = Left$(NxShortcutNormalizeModifier, Len(NxShortcutNormalizeModifier) - 1)
End Function

Private Function NxShortcutNormalizeKey(ByVal value As String) As String
    Dim canonical As String
    value = UCase$(Trim$(value))
    If Len(value) = 3 And Left$(value, 1) = "[" And Right$(value, 1) = "]" Then value = Mid$(value, 2, 1)
    If Len(value) = 1 Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", value, vbBinaryCompare) > 0 Or _
            InStr(1, "0123456789", value, vbBinaryCompare) > 0 Then
            NxShortcutNormalizeKey = value
            Exit Function
        End If
    End If
    Select Case value
        Case "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", "F13", "F14", "F15"
            canonical = value
        Case ";", "SEMICOLON": canonical = "Semicolon"
        Case "=", "EQUALS", "EQUALSSIGN": canonical = "Equals"
        Case ",", "COMMA": canonical = "Comma"
        Case "-", "MINUS", "MINUSSIGN": canonical = "Minus"
        Case ".", "PERIOD", "DOT": canonical = "Period"
        Case "/", "SLASH": canonical = "Slash"
        Case "`", "GRAVE", "GRAVEACCENT": canonical = "Grave"
        Case "[", "LEFTBRACKET", "LEFTSQUAREBRACKET": canonical = "LeftBracket"
        Case "\", "₩", "BACKSLASH", "WONBACKSLASH": canonical = "Backslash"
        Case "]", "RIGHTBRACKET", "RIGHTSQUAREBRACKET": canonical = "RightBracket"
        Case "'", "APOSTROPHE": canonical = "Apostrophe"
        Case "TAB": canonical = "Tab"
        Case "ENTER", "RETURN": canonical = "Enter"
        Case "ESC", "ESCAPE": canonical = "Escape"
        Case "SPACE", "SPACEBAR": canonical = "Space"
        Case "BACK", "BS", "BACKSPACE": canonical = "Backspace"
        Case "DEL", "DELETE": canonical = "Delete"
        Case "INS", "INSERT": canonical = "Insert"
        Case "HOME": canonical = "Home"
        Case "END": canonical = "End"
        Case "PGUP", "PAGEUP": canonical = "PageUp"
        Case "PGDN", "PAGEDOWN": canonical = "PageDown"
        Case "LEFT", "LEFTARROW": canonical = "Left"
        Case "RIGHT", "RIGHTARROW": canonical = "Right"
        Case "UP", "UPARROW": canonical = "Up"
        Case "DOWN", "DOWNARROW": canonical = "Down"
        Case Else
            NxRaiseContractError "문자·숫자·기호 또는 F1~F15 한 키를 지정하세요. 지원하지 않는 키 조합은 등록하지 않습니다."
    End Select
    NxShortcutNormalizeKey = canonical
End Function

' All entry points normalize physical shifted punctuation before conflict checks.
' For example Ctrl+Alt++ and Ctrl+Alt+Shift+= are the same binding.
Private Sub NxShortcutNormalizeChord(ByRef modifierName As String, ByRef keyName As String)
    Dim shifted As Boolean
    modifierName = NxShortcutNormalizeModifier(modifierName)
    keyName = Trim$(keyName)
    If Len(keyName) = 3 And Left$(keyName, 1) = "[" And Right$(keyName, 1) = "]" Then keyName = Mid$(keyName, 2, 1)
    shifted = True
    Select Case UCase$(keyName)
        Case "+", "PLUS", "PLUSSIGN": keyName = "Equals"
        Case "!": keyName = "1"
        Case "@": keyName = "2"
        Case "#": keyName = "3"
        Case "$": keyName = "4"
        Case "%": keyName = "5"
        Case "^": keyName = "6"
        Case "&": keyName = "7"
        Case "*": keyName = "8"
        Case "(": keyName = "9"
        Case ")": keyName = "0"
        Case "_": keyName = "Minus"
        Case ":": keyName = "Semicolon"
        Case "<": keyName = "Comma"
        Case ">": keyName = "Period"
        Case "?": keyName = "Slash"
        Case "~": keyName = "Grave"
        Case "{": keyName = "LeftBracket"
        Case "}": keyName = "RightBracket"
        Case "|": keyName = "Backslash"
        Case """": keyName = "Apostrophe"
        Case Else: shifted = False
    End Select
    If shifted And InStr(1, modifierName, "Shift", vbBinaryCompare) = 0 Then
        If modifierName = "None" Then modifierName = "Shift" Else modifierName = modifierName & "+Shift"
    End If
    keyName = NxShortcutNormalizeKey(keyName)
End Sub

' Pure parser: accepts syntax separately from the non-overridable reservation gate.
Public Function NxShortcutParseCombination(ByVal value As String, ByRef modifierName As String, _
    ByRef keyName As String, Optional ByRef reason As String = vbNullString) As Boolean
    Dim separator As Long, modifierText As String, token As String
    On Error GoTo Invalid
    modifierName = vbNullString
    keyName = vbNullString
    reason = vbNullString
    value = Trim$(value)
    If Len(value) = 0 Or Len(value) > 80 Then NxRaiseContractError "단축키 조합을 입력하세요. 예: Ctrl+Alt+Q"
    Do
        separator = InStr(1, value, "+", vbBinaryCompare)
        If separator = 0 Then Exit Do
        token = UCase$(Trim$(Left$(value, separator - 1)))
        If token <> "CTRL" And token <> "CONTROL" And token <> "ALT" And token <> "SHIFT" Then Exit Do
        If Len(modifierText) > 0 Then modifierText = modifierText & "+"
        modifierText = modifierText & token
        value = Trim$(Mid$(value, separator + 1))
    Loop
    modifierName = modifierText
    keyName = value
    NxShortcutNormalizeChord modifierName, keyName
    NxShortcutParseCombination = True
    Exit Function
Invalid:
    reason = Err.Description
    modifierName = vbNullString
    keyName = vbNullString
    Err.Clear
End Function

Public Function NxShortcutOnKeySpec(ByVal modifierName As String, ByVal keyName As String) As String
    Dim keySpec As String
    NxShortcutNormalizeChord modifierName, keyName
    If NxShortcutIsReservedCombination(modifierName, keyName) Then _
        NxRaiseContractError "Excel·시스템 기본 동작을 보호하는 예약 단축키입니다."
    Select Case keyName
        Case "Semicolon": keySpec = ";"
        Case "Equals": keySpec = "{=}"
        Case "Comma": keySpec = ","
        Case "Minus": keySpec = "{-}"
        Case "Period": keySpec = "."
        Case "Slash": keySpec = "/"
        Case "Grave": keySpec = "{`}"
        Case "LeftBracket": keySpec = "{[}"
        Case "Backslash": keySpec = "\"
        Case "RightBracket": keySpec = "{]}"
        Case "Apostrophe": keySpec = "{'}"
        Case Else
            If Len(keyName) > 1 And Left$(keyName, 1) = "F" Then keySpec = "{" & keyName & "}" Else keySpec = LCase$(keyName)
    End Select
    Select Case modifierName
        Case "None": NxShortcutOnKeySpec = keySpec
        Case "Shift": NxShortcutOnKeySpec = "+" & keySpec
        Case "Alt": NxShortcutOnKeySpec = "%" & keySpec
        Case "Alt+Shift": NxShortcutOnKeySpec = "%+" & keySpec
        Case "Ctrl": NxShortcutOnKeySpec = "^" & keySpec
        Case "Ctrl+Shift": NxShortcutOnKeySpec = "^+" & keySpec
        Case "Ctrl+Alt": NxShortcutOnKeySpec = "^%" & keySpec
        Case "Ctrl+Alt+Shift": NxShortcutOnKeySpec = "^%+" & keySpec
    End Select
End Function

Private Function NxShortcutDisplayKey(ByVal modifierName As String, ByVal keyName As String) As String
    If modifierName = "None" Then NxShortcutDisplayKey = keyName Else NxShortcutDisplayKey = modifierName & "+" & keyName
End Function

Private Function NxShortcutMacroName(ByVal slot As Long) As String
    NxShortcutMacroName = "NxShortcutDispatch" & Format$(slot, "00")
End Function

Private Function NxShortcutsPath() As String
    NxShortcutsPath = NxLHexcelProfileRoot() & "\Settings\" & NX_SHORTCUT_FILE
End Function

Private Function NxShortcutsLegacyPath() As String
    NxShortcutsLegacyPath = NxLHexcelProfileRoot() & "\Settings\" & NX_SHORTCUT_LEGACY_FILE
End Function

Private Sub NxShortcutsEnsureSettingsFolder()
    Dim root As String
    Dim settingsFolder As String
    root = NxLHexcelProfileRoot()
    settingsFolder = root & "\Settings"
    If Not NxShortcutFolderExists(root) Then MkDir root
    If Not NxShortcutFolderExists(settingsFolder) Then MkDir settingsFolder
End Sub

Private Function NxShortcutFolderExists(ByVal path As String) As Boolean
    On Error GoTo Missing
    NxShortcutFolderExists = ((GetAttr(path) And vbDirectory) = vbDirectory)
Missing:
End Function

Private Function NxShortcutFileExists(ByVal path As String) As Boolean
    On Error GoTo Missing
    NxShortcutFileExists = (Len(Dir$(path, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
Missing:
End Function
