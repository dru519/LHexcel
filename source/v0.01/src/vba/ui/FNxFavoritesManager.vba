Option Explicit

Private mItems As Collection

Private Sub UserForm_Initialize()
    Set mItems = NxGeneratedNavigationItems()
    LoadCategories
    ReloadSelected
    ReloadAvailable
End Sub

Private Sub UserForm_Activate()
    NxFormWheelAttach Me
End Sub

Public Sub HandleWheelDelta(ByVal wheelDelta As Long)
    NxFormWheelScrollActiveList Me, "lstAvailable", wheelDelta
End Sub

Private Sub cboCategory_Change(): ReloadAvailable: End Sub
Private Sub txtSearch_Change(): ReloadAvailable: End Sub

Private Sub lstAvailable_Click()
    cmdAdd.Enabled = (lstAvailable.ListIndex >= 0)
End Sub

Private Sub lstAvailable_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdAdd_Click
End Sub

Private Sub lstSelected_Click()
    Dim routeKey As String
    Dim selected As Boolean
    selected = (lstSelected.ListIndex >= 0)
    cmdRemove.Enabled = selected
    cmdMoveUp.Enabled = selected And lstSelected.ListIndex > 0
    cmdMoveDown.Enabled = selected And lstSelected.ListIndex < lstSelected.ListCount - 1
    If selected Then
        routeKey = CStr(lstSelected.List(lstSelected.ListIndex, 0))
        If Not NxFavoritesRouteIsKnown(routeKey) Then _
            lblStatus.Caption = "사용할 수 없는 항목입니다. 제거하거나 순서를 바꿀 수 있습니다."
    End If
End Sub

Private Sub lstSelected_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdRemove_Click
End Sub

Private Sub cmdAdd_Click()
    If lstAvailable.ListIndex < 0 Then Exit Sub
    AddSelectedRow CStr(lstAvailable.List(lstAvailable.ListIndex, 0))
    ReloadAvailable
    lblStatus.Caption = "즐겨찾기에 추가할 항목을 넣었습니다. 적용을 눌러 저장하세요."
End Sub

Private Sub cmdRemove_Click()
    Dim index As Long
    index = lstSelected.ListIndex
    If index < 0 Then Exit Sub
    lstSelected.RemoveItem index
    If lstSelected.ListCount > 0 Then
        If index >= lstSelected.ListCount Then index = lstSelected.ListCount - 1
        lstSelected.ListIndex = index
    End If
    ReloadAvailable
    lstSelected_Click
    lblStatus.Caption = "즐겨찾기에서 항목을 뺐습니다. 적용을 눌러 저장하세요."
End Sub

Private Sub cmdMoveUp_Click(): MoveSelected -1: End Sub
Private Sub cmdMoveDown_Click(): MoveSelected 1: End Sub

Private Sub cmdReset_Click()
    lstSelected.Clear
    ReloadAvailable
    lstSelected_Click
    lblStatus.Caption = "목록을 비웠습니다. 적용을 누르면 저장됩니다."
End Sub

Private Sub cmdApply_Click()
    Dim routes As New Collection
    Dim index As Long
    On Error GoTo Failed
    For index = 0 To lstSelected.ListCount - 1
        routes.Add CStr(lstSelected.List(index, 0))
    Next index
    NxFavoritesReplace routes
    NxRibbonInvalidateFavorites
    lblStatus.Caption = "즐겨찾기와 표시 순서를 저장했습니다."
    Exit Sub
Failed:
    lblStatus.Caption = NxUserErrorText(Err.Description)
End Sub

Private Sub cmdClose_Click(): Unload Me: End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    NxFormWheelDetach Me
End Sub

Private Sub UserForm_Terminate()
    NxFormWheelDetach Me
End Sub

Private Sub LoadCategories()
    Dim item As Variant
    Dim categoryId As String
    Dim categoryLabel As String
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbBinaryCompare
    cboCategory.Clear
    cboCategory.ColumnCount = 2
    cboCategory.ColumnWidths = "120 pt;0 pt"
    cboCategory.AddItem "전체"
    cboCategory.List(0, 1) = vbNullString
    For Each item In mItems
        categoryId = CStr(item(5))
        If Not seen.Exists(categoryId) Then
            categoryLabel = CStr(item(6))
            cboCategory.AddItem categoryLabel
            cboCategory.List(cboCategory.ListCount - 1, 1) = categoryId
            seen.Add categoryId, True
        End If
    Next item
    cboCategory.ListIndex = 0
End Sub

Private Sub ReloadSelected()
    Dim routeKey As Variant
    lstSelected.Clear
    For Each routeKey In NxFavoritesRouteKeys()
        AddSelectedRow CStr(routeKey)
    Next routeKey
    lstSelected_Click
End Sub

Private Sub ReloadAvailable()
    Dim item As Variant
    Dim query As String
    Dim routeKey As String
    If mItems Is Nothing Then Exit Sub
    query = Trim$(CStr(txtSearch.Value))
    lstAvailable.Clear
    For Each item In mItems
        routeKey = CStr(item(0))
        If NxFavoritesIsEligible(routeKey) And Not SelectedContains(routeKey) Then
            If CategoryMatches(CStr(item(5))) And ItemMatchesQuery(item, query) Then
                lstAvailable.AddItem routeKey
                lstAvailable.List(lstAvailable.ListCount - 1, 1) = CStr(item(3))
            End If
        End If
    Next item
    cmdAdd.Enabled = False
    lblStatus.Caption = "추가 가능 " & CStr(lstAvailable.ListCount) & "개 / 즐겨찾기 " & CStr(lstSelected.ListCount) & "개"
End Sub

Private Function CategoryMatches(ByVal categoryId As String) As Boolean
    If cboCategory.ListIndex <= 0 Then
        CategoryMatches = True
    Else
        CategoryMatches = (StrComp(CStr(cboCategory.List(cboCategory.ListIndex, 1)), categoryId, vbBinaryCompare) = 0)
    End If
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
        InStr(1, CStr(item(6)), query, vbTextCompare) > 0
End Function

Private Function SelectedContains(ByVal routeKey As String) As Boolean
    Dim index As Long
    For index = 0 To lstSelected.ListCount - 1
        If StrComp(CStr(lstSelected.List(index, 0)), routeKey, vbBinaryCompare) = 0 Then
            SelectedContains = True
            Exit Function
        End If
    Next index
End Function

Private Sub AddSelectedRow(ByVal routeKey As String)
    If SelectedContains(routeKey) Then Exit Sub
    lstSelected.AddItem routeKey
    lstSelected.List(lstSelected.ListCount - 1, 1) = NxFavoritesRouteLabel(routeKey)
End Sub

Private Sub MoveSelected(ByVal direction As Long)
    Dim sourceIndex As Long
    Dim targetIndex As Long
    Dim routeKey As String
    Dim label As String
    sourceIndex = lstSelected.ListIndex
    targetIndex = sourceIndex + direction
    If sourceIndex < 0 Or targetIndex < 0 Or targetIndex >= lstSelected.ListCount Then Exit Sub
    routeKey = CStr(lstSelected.List(sourceIndex, 0))
    label = CStr(lstSelected.List(sourceIndex, 1))
    lstSelected.RemoveItem sourceIndex
    lstSelected.AddItem routeKey, targetIndex
    lstSelected.List(targetIndex, 1) = label
    lstSelected.ListIndex = targetIndex
    lstSelected_Click
End Sub
