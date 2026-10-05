Option Explicit

Private mCapturedKey As String
Private mCapturedModifier As String
Private mItems As Collection
Private mReady As Boolean
Private mReloading As Boolean
Private mSelectedRouteKey As String
Private mDraft As Object
Private mInitialBindings As String
Private mDirty As Boolean
Private mDiscardConfirmed As Boolean
Private mCaptureReason As String

Private Sub UserForm_Initialize()
    Set mItems = NxGeneratedNavigationItems()
    Set mDraft = NxShortcutsCreateDraft()
    mInitialBindings = NxShortcutsSnapshot()
    LoadCategories
    cboState.Clear
    cboState.AddItem "전체"
    cboState.AddItem "지정됨"
    cboState.AddItem "미지정"
    cboState.ListIndex = 0
    txtDescription.ScrollBars = fmScrollBarsVertical
    lblPreview.WordWrap = True
    lblStatus.WordWrap = True
    mReady = True
    ReloadRoutes
End Sub

Private Sub UserForm_Activate()
    NxFormWheelAttach Me
End Sub

Public Sub HandleWheelDelta(ByVal wheelDelta As Long)
    NxFormWheelScrollActiveList Me, "lstCommands", wheelDelta
End Sub

Private Sub lstCommands_Change()
    If Not mReady Or mReloading Then Exit Sub
    If lstCommands.ListIndex >= 0 Then _
        mSelectedRouteKey = CStr(lstCommands.List(lstCommands.ListIndex, 0))
    LoadSelectedBinding
    UpdatePreview
End Sub

Private Sub txtShortcut_Enter()
    txtShortcut.SelStart = 0
    txtShortcut.SelLength = Len(CStr(txtShortcut.Value))
End Sub

Private Sub txtShortcut_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim keyName As String
    Dim modifierName As String

    If KeyCode = vbKeyEscape Then
        cmdClose_Click
        Exit Sub
    End If

    ' Leave navigation/default-button keys and modifier state to MSForms.
    ' Swallowing them traps focus and clears a captured chord before assignment.
    Select Case CLng(KeyCode)
        Case vbKeyTab, vbKeyReturn, vbKeyControl, vbKeyShift, vbKeyMenu
            Exit Sub
    End Select

    ' Keep standard text editing available in the now-editable input. These
    ' native combinations can still be typed explicitly to see their rejection.
    If Shift = 2 Then
        Select Case CLng(KeyCode)
            Case vbKeyA, vbKeyC, vbKeyV, vbKeyX, vbKeyZ, vbKeyBack, vbKeyDelete
                Exit Sub
        End Select
    End If
    If (Shift And 6) = 2 Then
        Select Case CLng(KeyCode)
            Case vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd
                Exit Sub
        End Select
    End If

    ' Ordinary text (including Shift punctuation) edits the combination directly.
    ' Ctrl/Alt chords and function keys still support immediate key capture.
    If (Shift And 6) = 0 And Not (CLng(KeyCode) >= vbKeyF1 And CLng(KeyCode) <= 126) Then Exit Sub
    If NxShortcutCaptureKey(CLng(KeyCode), Shift, modifierName, keyName) Then
        mCapturedModifier = modifierName
        mCapturedKey = keyName
        txtShortcut.Value = NxShortcutCaptureDisplay(CLng(KeyCode), Shift)
        lblStatus.Caption = "입력한 조합을 확인했습니다. Excel·시스템 예약키는 등록할 수 없습니다."
    Else
        ClearCapture
        lblStatus.Caption = "조합을 직접 입력하거나 키를 함께 누르세요. 예: Ctrl+Alt+Q, Shift+F13"
    End If
    UpdatePreview
    KeyCode = 0
End Sub

Private Sub txtShortcut_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = vbKeyTab Or KeyAscii = vbKeyReturn Then Exit Sub
    ' Direct text input is intentional; captured KeyDown events already consume their key.
End Sub

Private Sub txtShortcut_Change()
    If Not mReady Or mReloading Then Exit Sub
    mCaptureReason = vbNullString
    If Len(Trim$(CStr(txtShortcut.Value))) = 0 Then
        mCapturedModifier = vbNullString
        mCapturedKey = vbNullString
    Else
        NxShortcutParseCombination CStr(txtShortcut.Value), mCapturedModifier, mCapturedKey, mCaptureReason
    End If
    UpdatePreview
End Sub

Private Sub txtSearch_Change()
    If Not mReady Or mReloading Then Exit Sub
    ReloadRoutes
End Sub

Private Sub cboCategory_Change()
    If Not mReady Or mReloading Then Exit Sub
    ReloadRoutes
End Sub

Private Sub cboState_Change()
    If Not mReady Or mReloading Then Exit Sub
    ReloadRoutes
End Sub

Private Sub cmdAssign_Click()
    Dim routeKey As String
    On Error GoTo Failed
    If lstCommands.ListIndex < 0 Then NxRaiseContractError "단축키를 지정할 기능을 선택하세요."
    If Len(mCapturedModifier) = 0 Or Len(mCapturedKey) = 0 Then _
        NxRaiseContractError "단축키 조합을 직접 입력하거나 보조키와 키를 함께 누르세요."

    routeKey = CStr(lstCommands.List(lstCommands.ListIndex, 0))
    NxShortcutDraftAssign mDraft, routeKey, mCapturedModifier, mCapturedKey
    mDirty = True
    ReloadRoutesAndReselect routeKey
    LoadSelectedBinding
    UpdatePreview
    lblStatus.Caption = "미리보기 반영 · 적용을 누르면 저장됩니다."
    Exit Sub
Failed:
    lblStatus.Caption = NxUserErrorText(Err.Description)
End Sub

Private Sub cmdRemove_Click()
    Dim routeKey As String
    On Error GoTo Failed
    If lstCommands.ListIndex < 0 Then NxRaiseContractError "단축키 지정을 해제할 기능을 선택하세요."
    routeKey = CStr(lstCommands.List(lstCommands.ListIndex, 0))
    If MsgBox(NxShortcutsRouteLabel(routeKey) & "의 단축키 지정을 해제하시겠습니까?", _
        vbYesNo + vbQuestion + vbDefaultButton2, "내엑셀 - 선택 단축키 해제") <> vbYes Then Exit Sub
    NxShortcutDraftRemove mDraft, routeKey
    mDirty = True
    ReloadRoutesAndReselect routeKey
    ClearCapture
    UpdatePreview
    lblStatus.Caption = "해제 대기 · 적용을 누르면 저장됩니다."
    Exit Sub
Failed:
    lblStatus.Caption = NxUserErrorText(Err.Description)
End Sub

Private Sub cmdClose_Click()
    If Not ConfirmDiscard() Then Exit Sub
    mDiscardConfirmed = True
    Unload Me
End Sub

Private Sub cmdResetAll_Click()
    If mDraft.Count = 0 Then Exit Sub
    If MsgBox("전체 " & CStr(mDraft.Count) & "개 단축키 지정을 해제하시겠습니까?" & vbCrLf & _
        "적용 전에는 저장되지 않으며 취소하면 기존 설정을 유지합니다.", _
        vbYesNo + vbExclamation + vbDefaultButton2, "내엑셀 - 전체 단축키 해제") <> vbYes Then Exit Sub
    mDraft.RemoveAll
    mDirty = True
    ReloadRoutes
    lblStatus.Caption = "전체 해제 대기 · 적용을 누르면 저장됩니다."
End Sub

Private Sub cmdApply_Click()
    On Error GoTo Failed
    ' Apply the current captured chord too; preview is optional, not a hidden prerequisite.
    If cmdAssign.Enabled Then
        NxShortcutDraftAssign mDraft, mSelectedRouteKey, mCapturedModifier, mCapturedKey
        mDirty = True
    End If
    If Not mDirty Then Exit Sub
    NxShortcutsApplyDraft mDraft, mInitialBindings
    mInitialBindings = NxShortcutsSnapshot()
    mDirty = False
    ReloadRoutes
    lblStatus.Caption = "단축키 설정을 저장하고 적용했습니다."
    cmdApply.Enabled = False
    Exit Sub
Failed:
    lblStatus.Caption = NxUserErrorText(Err.Description)
End Sub

Private Function ConfirmDiscard() As Boolean
    ConfirmDiscard = True
    If Not mDirty Then Exit Function
    ConfirmDiscard = (MsgBox("적용하지 않은 단축키 변경을 버리고 닫으시겠습니까?", _
        vbYesNo + vbQuestion + vbDefaultButton2, "내엑셀 - 변경 취소") = vbYes)
End Function

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then cmdClose_Click
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not mDiscardConfirmed Then
        If Not ConfirmDiscard() Then Cancel = True: Exit Sub
    End If
    NxFormWheelDetach Me
End Sub

Private Sub UserForm_Terminate()
    NxFormWheelDetach Me
End Sub

Private Sub LoadCategories()
    Dim item As Variant
    Dim categoryId As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbBinaryCompare
    cboCategory.Clear
    cboCategory.ColumnCount = 2
    cboCategory.ColumnWidths = "108 pt;0 pt"
    cboCategory.AddItem "전체"
    cboCategory.List(0, 1) = vbNullString
    For Each item In mItems
        categoryId = CStr(item(5))
        If Not seen.Exists(categoryId) Then
            cboCategory.AddItem CStr(item(6))
            cboCategory.List(cboCategory.ListCount - 1, 1) = categoryId
            seen.Add categoryId, True
        End If
    Next item
    cboCategory.ListIndex = 0
End Sub

Private Sub ReloadRoutes()
    Dim item As Variant
    Dim query As String
    Dim routeKey As String
    Dim categoryId As String
    Dim bindingText As String

    If Not mReady Or mReloading Then Exit Sub
    query = Trim$(CStr(txtSearch.Value))
    If cboCategory.ListIndex >= 0 Then categoryId = CStr(cboCategory.List(cboCategory.ListIndex, 1))
    mReloading = True
    lstCommands.Clear
    For Each item In mItems
        routeKey = CStr(item(0))
        bindingText = NxShortcutDraftDisplay(mDraft, routeKey)
        If NxShortcutsRouteIsEligible(routeKey) And ItemMatchesCategory(item, categoryId) And ItemMatchesQuery(item, query) And ItemMatchesState(bindingText) Then
            lstCommands.AddItem routeKey
            lstCommands.List(lstCommands.ListCount - 1, 1) = CStr(item(3))
            lstCommands.List(lstCommands.ListCount - 1, 2) = bindingText
            lstCommands.List(lstCommands.ListCount - 1, 3) = DraftBindingStatus(routeKey, bindingText)
            If StrComp(routeKey, mSelectedRouteKey, vbBinaryCompare) = 0 Then _
                lstCommands.ListIndex = lstCommands.ListCount - 1
        End If
    Next item

    If lstCommands.ListCount = 0 Then
        lblStatus.Caption = "검색 결과가 없습니다."
    Else
        lblStatus.Caption = CStr(lstCommands.ListCount) & "개 기능에서 선택할 수 있습니다."
    End If
    mReloading = False
    LoadSelectedBinding
    UpdatePreview
End Sub

Private Function DraftBindingStatus(ByVal routeKey As String, ByVal bindingText As String) As String
    If StrComp(bindingText, NxShortcutsBindingDisplayForRoute(routeKey), vbBinaryCompare) <> 0 Then
        If Len(bindingText) = 0 Then DraftBindingStatus = "해제 대기" Else DraftBindingStatus = "지정 대기"
    ElseIf Len(bindingText) = 0 Then
        DraftBindingStatus = "미지정"
    Else
        DraftBindingStatus = "지정됨"
    End If
End Function

Private Function ItemMatchesState(ByVal bindingText As String) As Boolean
    Select Case cboState.ListIndex
        Case 1: ItemMatchesState = (Len(bindingText) > 0)
        Case 2: ItemMatchesState = (Len(bindingText) = 0)
        Case Else: ItemMatchesState = True
    End Select
End Function

Private Sub ReloadRoutesAndReselect(ByVal routeKey As String)
    mSelectedRouteKey = routeKey
    ReloadRoutes
End Sub

Private Function ItemMatchesCategory(ByVal item As Variant, ByVal categoryId As String) As Boolean
    ItemMatchesCategory = Len(categoryId) = 0 Or StrComp(CStr(item(5)), categoryId, vbBinaryCompare) = 0
End Function

Private Function ItemMatchesQuery(ByVal item As Variant, ByVal query As String) As Boolean
    If Len(query) = 0 Then
        ItemMatchesQuery = True
        Exit Function
    End If
    ItemMatchesQuery = _
        InStr(1, CStr(item(2)), query, vbTextCompare) > 0 Or _
        InStr(1, CStr(item(3)), query, vbTextCompare) > 0 Or _
        InStr(1, CStr(item(4)), query, vbTextCompare) > 0 Or _
        InStr(1, CStr(item(6)), query, vbTextCompare) > 0 Or _
        InStr(1, NxShortcutDraftDisplay(mDraft, CStr(item(0))), query, vbTextCompare) > 0
End Function

Private Sub LoadSelectedBinding()
    Dim keyName As String
    Dim modifierName As String
    Dim routeKey As String
    Dim binding As Variant

    ClearCapture
    If lstCommands.ListIndex < 0 Then Exit Sub
    routeKey = CStr(lstCommands.List(lstCommands.ListIndex, 0))
    If mDraft.Exists(routeKey) Then
        binding = mDraft.Item(routeKey)
        modifierName = CStr(binding(0))
        keyName = CStr(binding(1))
        mCapturedModifier = modifierName
        mCapturedKey = keyName
        txtShortcut.Value = NxShortcutDraftDisplay(mDraft, routeKey)
    End If
End Sub

Private Sub ClearCapture()
    mCapturedModifier = vbNullString
    mCapturedKey = vbNullString
    mCaptureReason = vbNullString
    txtShortcut.Value = vbNullString
End Sub

Private Sub UpdatePreview()
    Dim reason As String
    Dim routeKey As String

    cmdAssign.Enabled = False
    cmdRemove.Enabled = False
    cmdApply.Enabled = mDirty
    cmdResetAll.Enabled = (mDraft.Count > 0)
    txtDescription.Value = vbNullString
    If lstCommands.ListIndex < 0 Then
        lblPreview.Caption = "기능을 고른 뒤 조합을 입력하거나 키를 누르세요."
        Exit Sub
    End If

    routeKey = CStr(lstCommands.List(lstCommands.ListIndex, 0))
    txtDescription.Value = NxShortcutsRouteLabel(routeKey) & vbCrLf & vbCrLf & _
        NxGeneratedNavigationRouteField(routeKey, "description_ko")
    cmdRemove.Enabled = mDraft.Exists(routeKey)
    If Len(mCapturedModifier) = 0 Or Len(mCapturedKey) = 0 Then
        If Len(mCaptureReason) > 0 Then lblPreview.Caption = mCaptureReason Else lblPreview.Caption = "단축키 입력 대기"
        Exit Sub
    End If

    reason = NxShortcutDraftConflictReason(mDraft, mCapturedModifier, mCapturedKey, routeKey)
    If Len(reason) > 0 Then
        lblPreview.Caption = reason
    Else
        If mCapturedModifier = "None" Then
            lblPreview.Caption = "사용 가능 · " & mCapturedKey
        Else
            lblPreview.Caption = "사용 가능 · " & mCapturedModifier & "+" & mCapturedKey
        End If
        cmdAssign.Enabled = True
        cmdApply.Enabled = True
    End If
End Sub
