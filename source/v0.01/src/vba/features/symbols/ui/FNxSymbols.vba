Option Explicit

Private Const NX_SYMBOL_PAGE_SIZE As Long = 50
Private Const NX_SYMBOL_LAST_GLYPH As String = "cmdGlyph50"
Private Const NX_SYMBOL_ACTIVE_COLOR As Long = 12611584
Private Const NX_SYMBOL_INACTIVE_COLOR As Long = 15790320
Private Const NX_SYMBOL_GLYPH_COLUMNS As Long = 10
Private Const NX_SYMBOL_GLYPH_ROWS As Long = 5
Private Const NX_SYMBOL_GLYPH_WIDTH As Long = 23
Private Const NX_SYMBOL_GLYPH_HEIGHT As Long = 20
Private Const NX_SYMBOL_GAP_X As Long = 1
Private Const NX_SYMBOL_GAP_Y As Long = 1
Private Const NX_SYMBOL_FONT_SIZE As Long = 9
Private Const NX_SYMBOL_FALLBACK_FONT_SIZE As Long = 9
Private Const NX_SYMBOL_WHEEL_ROWS As Long = 3
Private Const NX_SYMBOL_BUFFER_TIP As String = "기호를 누르면 이 입력칸에 바로 추가됩니다."

Private WithEvents mScrollBar As MSForms.ScrollBar
Private mRecent As CNxSymbolRecent
Private mResults As Collection
Private mSelected As CNxSymbolRecord
Private mHandlers As Collection
Private mTabId As String
Private mCategoryId As String
Private mTopRow As Long
Private mTargetIdentity As String
Private mInitializing As Boolean
Private mBindingScroll As Boolean

Private Sub UserForm_Initialize()
    mInitializing = True
    Set mHandlers = New Collection
    CreateGlyphGrid
    CreateScrollBar
    mTabId = "practical"
    RefreshTarget
    RefreshCategories
    SetTabState
    RefreshResults
    mInitializing = False
End Sub

Private Sub UserForm_Activate()
    On Error Resume Next
    NxFormWheelAttach Me
    On Error GoTo 0
End Sub

Public Sub HandleWheelDelta(ByVal wheelDelta As Long)
    ScrollWheelDelta wheelDelta
End Sub

Public Sub BindSession(ByVal recent As CNxSymbolRecent)
    If recent Is Nothing Then NxRaiseContractError "기호표 최근 기록이 필요합니다."
    Set mRecent = recent
    RefreshResults
End Sub

Public Sub RefreshTarget()
    mTargetIdentity = NxSymbolActiveSelectionIdentity()
End Sub

Public Function SelectTab(ByVal tabId As String) As Boolean
    If tabId <> "practical" And tabId <> "unicode" And tabId <> "recent" Then Exit Function
    SetTab tabId
    SelectTab = True
End Function

Public Property Get RuntimeGlyphCount() As Long
    RuntimeGlyphCount = mHandlers.Count
End Property

Public Property Get RuntimeBufferValue() As String
    RuntimeBufferValue = CStr(txtBuffer.Value)
End Property

Public Property Get RuntimeScrollMaximum() As Long
    If Not mScrollBar Is Nothing Then RuntimeScrollMaximum = CLng(mScrollBar.Max)
End Property

Public Property Get RuntimeScrollValue() As Long
    RuntimeScrollValue = mTopRow
End Property

Public Function RuntimeActivateGlyph(ByVal oneBasedIndex As Long) As Boolean
    If oneBasedIndex < 1 Or oneBasedIndex > NX_SYMBOL_PAGE_SIZE Then Exit Function
    If Not Me.Controls(GlyphControlName(oneBasedIndex)).Enabled Then Exit Function
    HandleGlyphActivate oneBasedIndex
    RuntimeActivateGlyph = True
End Function

Public Function RuntimeSetScrollValue(ByVal requestedValue As Long) As Boolean
    If mScrollBar Is Nothing Then Exit Function
    If requestedValue < mScrollBar.Min Then requestedValue = mScrollBar.Min
    If requestedValue > mScrollBar.Max Then requestedValue = mScrollBar.Max
    mScrollBar.Value = requestedValue
    RuntimeSetScrollValue = (mTopRow = requestedValue)
End Function

Public Function GlyphControlName(ByVal oneBasedIndex As Long) As String
    If oneBasedIndex < 1 Or oneBasedIndex > NX_SYMBOL_PAGE_SIZE Then NxRaiseContractError "기호 격자 위치가 잘못되었습니다."
    GlyphControlName = "cmdGlyph" & Right$("00" & CStr(oneBasedIndex), 2)
End Function

Public Sub HandleGlyphClick(ByVal oneBasedIndex As Long)
    HandleGlyphActivate oneBasedIndex
End Sub

Public Sub HandleGlyphActivate(ByVal oneBasedIndex As Long)
    SelectGlyph oneBasedIndex
    AddSelectedToBuffer
End Sub

Public Sub HandleGlyphKey(ByVal oneBasedIndex As Long, ByRef KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim nextIndex As Long
    If KeyCode = 13 Or KeyCode = 32 Then
        HandleGlyphActivate oneBasedIndex
        KeyCode = 0
        Exit Sub
    End If
    nextIndex = oneBasedIndex
    Select Case KeyCode
        Case 37: nextIndex = oneBasedIndex - 1
        Case 38: nextIndex = oneBasedIndex - NX_SYMBOL_GLYPH_COLUMNS
        Case 39: nextIndex = oneBasedIndex + 1
        Case 40: nextIndex = oneBasedIndex + NX_SYMBOL_GLYPH_COLUMNS
        Case Else: Exit Sub
    End Select
    If nextIndex < 1 Then
        ScrollRows -1
        nextIndex = nextIndex + NX_SYMBOL_GLYPH_COLUMNS
    ElseIf nextIndex > NX_SYMBOL_PAGE_SIZE Then
        ScrollRows 1
        nextIndex = nextIndex - NX_SYMBOL_GLYPH_COLUMNS
    End If
    If nextIndex < 1 Or nextIndex > NX_SYMBOL_PAGE_SIZE Then Exit Sub
    If Not Me.Controls(GlyphControlName(nextIndex)).Enabled Then Exit Sub
    Me.Controls(GlyphControlName(nextIndex)).SetFocus
    SelectGlyph nextIndex
    KeyCode = 0
End Sub

Public Sub ScrollRows(ByVal rowDelta As Long)
    SetTopRow mTopRow + rowDelta
End Sub

Public Sub ScrollWheelDelta(ByVal wheelDelta As Long)
    If wheelDelta > 0 Then
        ScrollRows -NX_SYMBOL_WHEEL_ROWS
    ElseIf wheelDelta < 0 Then
        ScrollRows NX_SYMBOL_WHEEL_ROWS
    End If
End Sub

Private Sub CreateGlyphGrid()
    Dim index As Long, rowIndex As Long, columnIndex As Long
    Dim button As MSForms.CommandButton, handler As CNxSymbolButtonHandler
    For index = 1 To NX_SYMBOL_PAGE_SIZE
        rowIndex = (index - 1) \ NX_SYMBOL_GLYPH_COLUMNS
        columnIndex = (index - 1) Mod NX_SYMBOL_GLYPH_COLUMNS
        Set button = Me.Controls.Add("Forms.CommandButton.1", GlyphControlName(index), True)
        With button
            .Left = lblGridAnchor.Left + columnIndex * (NX_SYMBOL_GLYPH_WIDTH + NX_SYMBOL_GAP_X)
            .Top = lblGridAnchor.Top + rowIndex * (NX_SYMBOL_GLYPH_HEIGHT + NX_SYMBOL_GAP_Y)
            .Width = NX_SYMBOL_GLYPH_WIDTH
            .Height = NX_SYMBOL_GLYPH_HEIGHT
            .TabIndex = 20 + index
            .TabStop = True
            .Enabled = False
            .Font.Name = "맑은 고딕"
            .Font.Size = NX_SYMBOL_FONT_SIZE
        End With
        Set handler = New CNxSymbolButtonHandler
        handler.Bind Me, button, index
        mHandlers.Add handler
    Next index
    If mHandlers.Count <> NX_SYMBOL_PAGE_SIZE Or Me.Controls(NX_SYMBOL_LAST_GLYPH).Name <> NX_SYMBOL_LAST_GLYPH Then NxRaiseContractError "기호 격자를 만들지 못했습니다."
End Sub

Private Sub CreateScrollBar()
    Set mScrollBar = Me.Controls.Add("Forms.ScrollBar.1", "scrSymbols", True)
    With mScrollBar
        .Left = lblScrollAnchor.Left
        .Top = lblScrollAnchor.Top
        .Width = lblScrollAnchor.Width
        .Height = lblScrollAnchor.Height
        .Orientation = fmOrientationVertical
        .TabIndex = 8
        .TabStop = True
        .Min = 0
        .Max = 0
        .SmallChange = 1
        .LargeChange = NX_SYMBOL_GLYPH_ROWS
        .Value = 0
        .Enabled = False
        .ControlTipText = "끌거나 휠을 굴려 기호 목록을 이동합니다."
    End With
End Sub

Private Sub RefreshCategories()
    Dim rows As Collection, row As Variant, parts As Variant
    lstCategories.Clear
    lstCategories.ColumnCount = 2
    lstCategories.ColumnWidths = "72 pt;0 pt"
    If mTabId = "practical" Then
        Set rows = NxSymbolPracticalCategories()
    ElseIf mTabId = "unicode" Then
        Set rows = NxSymbolUnicodeBlocks()
    Else
        Set rows = New Collection
        rows.Add "recent|최근 사용"
    End If
    For Each row In rows
        parts = Split(CStr(row), "|")
        lstCategories.AddItem CStr(parts(1))
        lstCategories.List(lstCategories.ListCount - 1, 1) = CStr(parts(0))
    Next row
    If lstCategories.ListCount > 0 Then
        lstCategories.ListIndex = 0
        mCategoryId = CStr(lstCategories.List(0, 1))
    Else
        mCategoryId = vbNullString
    End If
End Sub

Private Sub RefreshResults()
    Dim recentSnapshot As Collection
    If mRecent Is Nothing Then Set recentSnapshot = New Collection Else Set recentSnapshot = mRecent.Snapshot()
    Set mResults = NxSymbolSearch(mTabId, mCategoryId, vbNullString, recentSnapshot)
    If mTopRow > MaxTopRow() Then mTopRow = MaxTopRow()
    RefreshScrollBar
    BindGrid
End Sub

Private Sub RefreshScrollBar()
    Dim maximum As Long
    maximum = MaxTopRow()
    If mTopRow < 0 Then mTopRow = 0
    If mTopRow > maximum Then mTopRow = maximum
    mBindingScroll = True
    With mScrollBar
        .Min = 0
        .Max = maximum
        .SmallChange = 1
        .LargeChange = NX_SYMBOL_GLYPH_ROWS
        .Value = mTopRow
        .Enabled = maximum > 0
    End With
    mBindingScroll = False
End Sub

Private Function MaxTopRow() As Long
    Dim count As Long, rowCount As Long
    If Not mResults Is Nothing Then count = mResults.Count
    rowCount = (count + NX_SYMBOL_GLYPH_COLUMNS - 1) \ NX_SYMBOL_GLYPH_COLUMNS
    MaxTopRow = rowCount - NX_SYMBOL_GLYPH_ROWS
    If MaxTopRow < 0 Then MaxTopRow = 0
End Function

Private Sub SetTopRow(ByVal requestedRow As Long)
    Dim maximum As Long
    maximum = MaxTopRow()
    If requestedRow < 0 Then requestedRow = 0
    If requestedRow > maximum Then requestedRow = maximum
    If requestedRow = mTopRow Then Exit Sub
    mTopRow = requestedRow
    If Not mScrollBar Is Nothing Then
        mBindingScroll = True
        mScrollBar.Value = mTopRow
        mBindingScroll = False
    End If
    BindGrid
End Sub

Private Sub BindGrid()
    Dim slot As Long, resultIndex As Long, button As MSForms.CommandButton
    Dim record As CNxSymbolRecord
    For slot = 1 To NX_SYMBOL_PAGE_SIZE
        Set button = Me.Controls(GlyphControlName(slot))
        resultIndex = (mTopRow * NX_SYMBOL_GLYPH_COLUMNS) + slot
        If Not mResults Is Nothing And resultIndex <= mResults.Count Then
            Set record = mResults(resultIndex)
            button.Caption = record.Character
            button.Tag = CStr(record.CodePoint)
            button.Font.Name = "맑은 고딕"
            button.Font.Size = IIf(Len(record.Character) > 1, NX_SYMBOL_FALLBACK_FONT_SIZE, NX_SYMBOL_FONT_SIZE)
            button.Enabled = True
            button.ControlTipText = record.NameKo
        Else
            button.Caption = vbNullString
            button.Tag = vbNullString
            button.Enabled = False
            button.ControlTipText = vbNullString
        End If
    Next slot
    If mResults Is Nothing Or mResults.Count = 0 Then Set mSelected = Nothing
End Sub

Private Sub SelectGlyph(ByVal oneBasedIndex As Long)
    Dim resultIndex As Long
    resultIndex = (mTopRow * NX_SYMBOL_GLYPH_COLUMNS) + oneBasedIndex
    If mResults Is Nothing Or resultIndex < 1 Or resultIndex > mResults.Count Then Exit Sub
    Set mSelected = mResults(resultIndex)
End Sub

Private Sub AddSelectedToBuffer()
    If mSelected Is Nothing Then Exit Sub
    txtBuffer.Value = CStr(txtBuffer.Value) & mSelected.Character
    If Not mRecent Is Nothing Then mRecent.Touch mSelected
    If mTabId = "recent" Then RefreshResults
End Sub

Private Sub SetTab(ByVal tabId As String)
    If tabId <> "practical" And tabId <> "unicode" And tabId <> "recent" Then Exit Sub
    mTabId = tabId
    mTopRow = 0
    RefreshCategories
    SetTabState
    RefreshResults
End Sub

Private Sub SetTabState()
    cmdTabPractical.BackColor = IIf(mTabId = "practical", NX_SYMBOL_ACTIVE_COLOR, NX_SYMBOL_INACTIVE_COLOR)
    cmdTabUnicode.BackColor = IIf(mTabId = "unicode", NX_SYMBOL_ACTIVE_COLOR, NX_SYMBOL_INACTIVE_COLOR)
    cmdTabRecent.BackColor = IIf(mTabId = "recent", NX_SYMBOL_ACTIVE_COLOR, NX_SYMBOL_INACTIVE_COLOR)
    cmdTabPractical.ControlTipText = IIf(mTabId = "practical", "선택됨: 실무 기호", "실무 기호 탭")
    cmdTabUnicode.ControlTipText = IIf(mTabId = "unicode", "선택됨: 유니코드", "유니코드 탭")
    cmdTabRecent.ControlTipText = IIf(mTabId = "recent", "선택됨: 최근 사용", "최근 사용 탭")
End Sub

Private Sub mScrollBar_Change()
    If Not mBindingScroll Then SetTopRow CLng(mScrollBar.Value)
End Sub

Private Sub mScrollBar_Scroll()
    If Not mBindingScroll Then SetTopRow CLng(mScrollBar.Value)
End Sub

Private Sub cmdTabPractical_Click(): SetTab "practical": End Sub
Private Sub cmdTabUnicode_Click(): SetTab "unicode": End Sub
Private Sub cmdTabRecent_Click(): SetTab "recent": End Sub

Private Sub lstCategories_Click()
    If mInitializing Or lstCategories.ListIndex < 0 Then Exit Sub
    mCategoryId = CStr(lstCategories.List(lstCategories.ListIndex, 1))
    mTopRow = 0
    RefreshResults
End Sub

Private Sub cmdCopy_Click()
    If Not NxDistributionCanExecute() Then Exit Sub
    If Len(CStr(txtBuffer.Value)) = 0 Then Exit Sub
    NxClipboardWriteUnicode CStr(txtBuffer.Value)
End Sub

Private Sub cmdInsertCell_Click()
    Dim prepared As CNxSymbolCellWrite
    On Error GoTo Failed
    NxDistributionEnsureExecutable
    If NxSymbolActiveSelectionIdentity() <> mTargetIdentity Then NxRaiseContractError "입력할 셀 범위가 바뀌었습니다. 범위를 다시 선택해 주세요."
    Set prepared = NxSymbolPrepareActiveWrite(CStr(txtBuffer.Value), mTargetIdentity)
    prepared.Commit
    txtBuffer.ControlTipText = NX_SYMBOL_BUFFER_TIP
    RefreshTarget
    Exit Sub
Failed:
    txtBuffer.ControlTipText = NxUserErrorText(Err.Description)
    Beep
    Err.Clear
End Sub

Private Sub cmdClose_Click(): Unload Me: End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    NxFormWheelDetach Me
End Sub

Private Sub UserForm_Terminate()
    NxFormWheelDetach Me
    NxSymbolsFormClosed Me
End Sub
