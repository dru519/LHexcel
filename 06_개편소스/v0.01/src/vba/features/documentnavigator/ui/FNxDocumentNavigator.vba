Option Explicit

Private mItems As Collection
Private mRendering As Boolean
Private mRefreshing As Boolean
Private mClosing As Boolean
Private mActiveList As String
Private mWheelList As String
Private mSheetBookToken As Long

Private Sub UserForm_Initialize()
    Set mItems = New Collection
    mActiveList = "books"
    mWheelList = "lstBooks"
End Sub

Private Sub UserForm_Activate()
    If mClosing Then Exit Sub
    RefreshNavigator
    NxFormWheelAttach Me
End Sub

Public Sub FocusNavigator()
    NxFormWheelAttach Me
    txtBooks.SetFocus
End Sub

Public Sub RefreshNavigator()
    If mRefreshing Or mClosing Then Exit Sub
    On Error GoTo Failed
    mRefreshing = True
    Set mItems = NxDocumentNavigatorSnapshot(mItems)
    RenderBooks
    mRefreshing = False
    Exit Sub
Failed:
    mRendering = True
    lstBooks.Clear
    lstSheets.Clear
    Set mItems = New Collection
    cmdMove.Enabled = False
    lblStatus.Caption = "목록을 읽지 못했습니다: " & NxUserErrorText(Err.Description)
    mRendering = False
    mRefreshing = False
End Sub

Private Function SelectedToken(ByVal list As MSForms.ListBox) As Long
    If list.ListIndex >= 0 Then SelectedToken = CLng(list.List(list.ListIndex, 0))
End Function

Private Function ItemByToken(ByVal token As Long) As CNxDocumentNavigatorItem
    Dim item As CNxDocumentNavigatorItem
    If token = 0 Then Exit Function
    For Each item In mItems
        If item.TokenId = token Then Set ItemByToken = item: Exit Function
    Next item
End Function

Private Function SelectedItem() As CNxDocumentNavigatorItem
    If mActiveList = "sheets" Then
        Set SelectedItem = ItemByToken(SelectedToken(lstSheets))
    Else
        Set SelectedItem = ItemByToken(SelectedToken(lstBooks))
    End If
End Function

Private Sub RestoreTop(ByVal list As MSForms.ListBox, ByVal priorTop As Long)
    If list.ListCount = 0 Then Exit Sub
    If priorTop < 0 Then priorTop = 0
    If priorTop >= list.ListCount Then priorTop = list.ListCount - 1
    list.TopIndex = priorTop
End Sub

Private Sub RenderBooks()
    Dim item As CNxDocumentNavigatorItem, token As Long, priorTop As Long
    Dim query As String
    token = SelectedToken(lstBooks)
    priorTop = lstBooks.TopIndex
    query = Trim$(CStr(txtBooks.Value))
    mRendering = True
    lstBooks.Clear
    For Each item In mItems
        If item.SheetRef Is Nothing Then
            If Len(query) = 0 Or InStr(1, item.DisplayName, query, vbTextCompare) > 0 Then
                lstBooks.AddItem CStr(item.TokenId)
                lstBooks.List(lstBooks.ListCount - 1, 1) = item.DisplayName
                If item.TokenId = token Then lstBooks.ListIndex = lstBooks.ListCount - 1
            End If
        End If
    Next item
    RestoreTop lstBooks, priorTop
    lblBooks.Caption = "열린 문서 · " & CStr(lstBooks.ListCount) & "개"
    mRendering = False
    RenderSheets
End Sub

Private Sub RenderSheets()
    Dim bookItem As CNxDocumentNavigatorItem, item As CNxDocumentNavigatorItem
    Dim token As Long, priorTop As Long, query As String
    Set bookItem = ItemByToken(SelectedToken(lstBooks))
    token = SelectedToken(lstSheets)
    priorTop = lstSheets.TopIndex
    query = Trim$(CStr(txtSheets.Value))
    If bookItem Is Nothing Then
        token = 0
        mSheetBookToken = 0
    ElseIf bookItem.TokenId <> mSheetBookToken Then
        token = 0
        priorTop = 0
        mSheetBookToken = bookItem.TokenId
    End If
    mRendering = True
    lstSheets.Clear
    If Not bookItem Is Nothing Then
        For Each item In mItems
            If Not item.ParentBook Is Nothing Then
                If item.ParentBook Is bookItem Then
                    If Len(query) = 0 Or InStr(1, item.DisplayName, query, vbTextCompare) > 0 Then
                        lstSheets.AddItem CStr(item.TokenId)
                        lstSheets.List(lstSheets.ListCount - 1, 1) = item.DisplayName
                        lstSheets.List(lstSheets.ListCount - 1, 2) = item.StateText
                        If item.TokenId = token Then lstSheets.ListIndex = lstSheets.ListCount - 1
                    End If
                End If
            End If
        Next item
    End If
    RestoreTop lstSheets, priorTop
    lblSheets.Caption = "시트 · " & CStr(lstSheets.ListCount) & "개"
    mRendering = False
    RenderSelection
End Sub

Private Sub RenderSelection()
    Dim item As CNxDocumentNavigatorItem
    Set item = SelectedItem()
    cmdMove.Enabled = False
    If item Is Nothing Then
        If mItems.Count = 0 Then
            lblStatus.Caption = "열린 문서가 없습니다. 문서를 연 뒤 새로고침하세요."
        Else
            lblStatus.Caption = "목록에서 문서 또는 시트를 선택하세요."
        End If
    ElseIf Not item.Available Then
        lblStatus.Caption = "숨긴 시트는 이동할 수 없습니다."
    Else
        cmdMove.Enabled = True
        lblStatus.Caption = "Enter · 더블클릭 · 이동 버튼으로 이동합니다."
    End If
End Sub

Private Sub txtBooks_Change()
    If mRendering Then Exit Sub
    mActiveList = "books"
    RenderBooks
End Sub

Private Sub txtSheets_Change()
    If mRendering Then Exit Sub
    mActiveList = "sheets"
    RenderSheets
End Sub

Private Sub lstBooks_Click()
    If mRendering Then Exit Sub
    mActiveList = "books"
    mWheelList = "lstBooks"
    RenderSheets
End Sub

Private Sub lstBooks_Change()
    lstBooks_Click
End Sub

Private Sub lstSheets_Click()
    If mRendering Then Exit Sub
    mActiveList = "sheets"
    mWheelList = "lstSheets"
    RenderSelection
End Sub

Private Sub lstSheets_Change()
    lstSheets_Click
End Sub

Private Sub lstBooks_Enter()
    mActiveList = "books"
    mWheelList = "lstBooks"
    RenderSelection
End Sub

Private Sub lstSheets_Enter()
    mActiveList = "sheets"
    mWheelList = "lstSheets"
    RenderSelection
End Sub

Private Sub MoveSelected()
    Dim item As CNxDocumentNavigatorItem, reason As String, moved As Boolean
    Set item = SelectedItem()
    If item Is Nothing Then Exit Sub
    moved = NxDocumentNavigatorMove(item, reason)
    If mClosing Then Exit Sub
    RefreshNavigator
    If Not moved Then lblStatus.Caption = reason
    If moved Then NxFormWheelActivateHost Me
End Sub

Private Sub cmdMove_Click()
    MoveSelected
End Sub

Private Sub cmdRefresh_Click()
    RefreshNavigator
End Sub

Private Sub lstBooks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    mActiveList = "books"
    MoveSelected
End Sub

Private Sub lstSheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    mActiveList = "sheets"
    MoveSelected
End Sub

Private Sub txtBooks_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    FocusResults keyCode, lstBooks
End Sub

Private Sub txtSheets_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    FocusResults keyCode, lstSheets
End Sub

Private Sub FocusResults(ByVal keyCode As MSForms.ReturnInteger, ByVal list As MSForms.ListBox)
    If keyCode = vbKeyReturn Then
        keyCode = 0
        list.SetFocus
    ElseIf keyCode = vbKeyEscape Then
        keyCode = 0
        CloseNavigator
    End If
End Sub

Private Sub lstBooks_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    mActiveList = "books"
    HandleResultKey keyCode
End Sub

Private Sub lstSheets_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    mActiveList = "sheets"
    HandleResultKey keyCode
End Sub

Private Sub HandleResultKey(ByVal keyCode As MSForms.ReturnInteger)
    If keyCode = vbKeyReturn Then
        keyCode = 0
        MoveSelected
    ElseIf keyCode = vbKeyEscape Then
        keyCode = 0
        CloseNavigator
    End If
End Sub

Private Sub lstBooks_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mWheelList = "lstBooks"
End Sub

Private Sub lstSheets_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    mWheelList = "lstSheets"
End Sub

Public Sub HandleWheelDelta(ByVal wheelDelta As Long)
    NxFormWheelScrollList Me.Controls(mWheelList), wheelDelta
End Sub

Private Sub cmdClose_Click()
    CloseNavigator
End Sub

Private Sub CloseNavigator()
    mClosing = True
    NxFormWheelDetach Me
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    mClosing = True
    NxFormWheelDetach Me
    Set mItems = Nothing
    NxDocumentNavigatorRelease Me
End Sub

Private Sub UserForm_Terminate()
    NxFormWheelDetach Me
    Set mItems = Nothing
    NxDocumentNavigatorRelease Me
End Sub
