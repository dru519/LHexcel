Option Explicit

Private mItems As Collection
Private mVisibleItems As Collection
Private mClosing As Boolean

Private Sub UserForm_Initialize()
    Set mItems = NxGeneratedNavigationItems()
    Set mVisibleItems = New Collection
    LoadCategories
    ReloadResults
End Sub

Private Sub UserForm_Activate()
    NxFormWheelAttach Me
End Sub

Public Sub HandleWheelDelta(ByVal wheelDelta As Long)
    NxFormWheelScrollActiveList Me, "lstResults", wheelDelta
End Sub

Public Sub ActivateNavigator()
    On Error Resume Next
    txtSearch.SetFocus
    On Error GoTo 0
End Sub

Public Sub RefreshAvailability()
    RenderSelection
End Sub

Private Sub LoadCategories()
    Dim item As Variant
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbBinaryCompare
    cboCategory.Clear
    cboCategory.AddItem "전체"
    For Each item In mItems
        If Not seen.Exists(CStr(item(5))) Then
            seen.Add CStr(item(5)), True
            cboCategory.AddItem CStr(item(6))
            cboCategory.List(cboCategory.ListCount - 1, 1) = CStr(item(5))
        End If
    Next item
    cboCategory.ListIndex = 0
End Sub

Private Sub txtSearch_Change()
    ReloadResults
End Sub

Private Sub cboCategory_Change()
    ReloadResults
End Sub

Private Sub ReloadResults()
    Dim categoryId As String
    Dim item As Variant
    Dim query As String
    Dim searchValue As String
    Dim selectedRoute As String, priorTop As Long
    If lstResults.ListIndex >= 0 Then selectedRoute = CStr(lstResults.List(lstResults.ListIndex, 0))
    priorTop = lstResults.TopIndex
    query = Trim$(CStr(txtSearch.Value))
    categoryId = SelectedCategoryId()
    lstResults.Clear
    Set mVisibleItems = New Collection
    For Each item In mItems
        searchValue = CStr(item(14)) & " | " & CStr(item(6))
        If (Len(query) = 0 Or InStr(1, searchValue, query, vbTextCompare) > 0) And _
            (Len(categoryId) = 0 Or CStr(item(5)) = categoryId) Then
            mVisibleItems.Add item
            lstResults.AddItem CStr(item(0))
            lstResults.List(lstResults.ListCount - 1, 1) = CStr(item(3))
            lstResults.List(lstResults.ListCount - 1, 2) = CStr(item(6))
            If CStr(item(0)) = selectedRoute Then lstResults.ListIndex = lstResults.ListCount - 1
        End If
    Next item
    cmdRun.Enabled = False
    lblDescription.Caption = "기능을 선택하면 설명이 표시됩니다."
    lblRequirements.Caption = "요구 입력: -"
    lblAvailability.Caption = CStr(lstResults.ListCount) & "개 항목"
    If lstResults.ListCount > 0 Then
        If priorTop > lstResults.ListCount - 1 Then priorTop = lstResults.ListCount - 1
        If priorTop >= 0 Then lstResults.TopIndex = priorTop
    End If
    RenderSelection
End Sub

Private Function SelectedCategoryId() As String
    If cboCategory.ListIndex <= 0 Then Exit Function
    SelectedCategoryId = CStr(cboCategory.List(cboCategory.ListIndex, 1))
End Function

Private Sub lstResults_Click()
    RenderSelection
End Sub

Private Sub lstResults_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    RunSelected
End Sub

Private Sub RenderSelection()
    Dim item As Variant
    Dim routeKey As String
    Dim state As String
    If lstResults.ListIndex < 0 Then Exit Sub
    item = mVisibleItems(lstResults.ListIndex + 1)
    routeKey = CStr(item(0))
    state = NxRouteAvailabilityState(routeKey)
    lblDescription.Caption = CStr(item(4))
    lblRequirements.Caption = NxNavigationDisplayText(CStr(item(8))) & " · " & NxNavigationDisplayText(CStr(item(13)))
    lblAvailability.Caption = AvailabilityCaption(state, NxRouteAvailabilityReason(routeKey))
    cmdRun.Enabled = (state <> NX_ROUTE_UNAVAILABLE)
End Sub

Private Function AvailabilityCaption(ByVal state As String, ByVal reason As String) As String
    Select Case state
        Case NX_ROUTE_AVAILABLE: AvailabilityCaption = "실행 가능"
        Case NX_ROUTE_NEEDS_INPUT: AvailabilityCaption = "입력 필요"
        Case Else: AvailabilityCaption = "실행 불가"
    End Select
    If Len(reason) > 0 Then AvailabilityCaption = AvailabilityCaption & " · " & reason
End Function

Private Sub cmdRun_Click()
    RunSelected
End Sub

Private Sub RunSelected()
    Dim item As Variant
    Dim routeKey As String
    On Error GoTo Failed
    If lstResults.ListIndex < 0 Then Exit Sub
    item = mVisibleItems(lstResults.ListIndex + 1)
    routeKey = CStr(item(0))
    NxRouteAvailabilityEnsureExecutable routeKey
    If CStr(item(1)) = "feature" Then
        NxRouteFeature CStr(item(2))
    Else
        NxRouteCommand CStr(item(2))
    End If
    lblAvailability.Caption = "실행 요청을 전달했습니다."
    Exit Sub
Failed:
    lblAvailability.Caption = NxUserErrorText(Err.Description)
    Err.Clear
End Sub

Private Sub cmdClose_Click()
    HideNavigator
End Sub

Private Sub txtSearch_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    HandleKey keyCode
End Sub

Private Sub cboCategory_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    HandleKey keyCode
End Sub

Private Sub lstResults_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    HandleKey keyCode
End Sub

Private Sub UserForm_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    HandleKey keyCode
End Sub

Private Sub HandleKey(ByRef keyCode As MSForms.ReturnInteger)
    Select Case keyCode
        Case vbKeyReturn
            RunSelected
            keyCode = 0
        Case vbKeyEscape
            HideNavigator
            keyCode = 0
    End Select
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    NxFormWheelDetach Me
    If CloseMode = vbFormControlMenu And Not mClosing Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub UserForm_Terminate()
    NxFormWheelDetach Me
    mClosing = True
    NxNavigatorRelease Me
    Set mVisibleItems = Nothing
    Set mItems = Nothing
End Sub

Private Sub HideNavigator()
    NxFormWheelDetach Me
    Me.Hide
End Sub
