Attribute VB_Name = "T_R61ShortcutCatalog"
Option Explicit

' Import into a disposable Product test copy, then call this function from the
' Windows BAT-owned Excel harness. binding_display needs a route pre-bound by
' that harness in its isolated profile. This fixture never writes bindings.
Public Function NxR61ShortcutCatalogRunCase(ByVal kind As String, Optional ByVal assignedRouteKey As String = vbNullString) As String
    Dim view As FNxShortcutManager
    Dim item As Variant, items As Collection
    Dim categoryIndex As Long, rowIndex As Long, expectedCount As Long
    Dim routeKey As String, categoryId As String, description As String, binding As String
    Dim categories As Object
    On Error GoTo Failed
    Set items = NxGeneratedNavigationItems()
    Set view = New FNxShortcutManager
    Load view
    Require view.lstCommands.ListCount = NxShortcutsEligibleRoutes().Count, "initial eligible count"
    Require view.lstCommands.ListIndex = -1, "no implicit initial assignment target"
    Select Case kind
        Case "categories"
            Set categories = CreateObject("Scripting.Dictionary")
            categories.CompareMode = vbBinaryCompare
            For Each item In items
                categoryId = CStr(item(5))
                If Not categories.Exists(categoryId) Then categories.Add categoryId, CStr(item(6))
            Next item
            Require view.cboCategory.ListCount = categories.Count + 1, "catalog category count"
            Require CStr(view.cboCategory.List(0, 1)) = vbNullString, "all category hidden key"
            For categoryIndex = 1 To view.cboCategory.ListCount - 1
                categoryId = CStr(view.cboCategory.List(categoryIndex, 1))
                Require categories.Exists(categoryId), "unknown category"
                Require CStr(view.cboCategory.List(categoryIndex, 0)) = CStr(categories(categoryId)), "category label"
            Next categoryIndex
        Case "category_search"
            For categoryIndex = 1 To view.cboCategory.ListCount - 1
                view.txtSearch.Value = vbNullString
                view.cboCategory.ListIndex = categoryIndex
                categoryId = CStr(view.cboCategory.List(categoryIndex, 1))
                expectedCount = 0
                For Each item In items
                    If CStr(item(5)) = categoryId And NxShortcutsRouteIsEligible(CStr(item(0))) Then expectedCount = expectedCount + 1
                Next item
                Require view.lstCommands.ListCount = expectedCount, "category eligible count"
                For rowIndex = 0 To view.lstCommands.ListCount - 1
                    routeKey = CStr(view.lstCommands.List(rowIndex, 0))
                    Require NxGeneratedNavigationRouteField(routeKey, "category_id") = categoryId, "category leaked route"
                Next rowIndex
                If expectedCount > 0 Then
                    routeKey = CStr(view.lstCommands.List(0, 0))
                    view.txtSearch.Value = NxGeneratedNavigationRouteField(routeKey, "label_ko")
                    Require FindRoute(view, routeKey) >= 0, "combined category and label search"
                    For rowIndex = 0 To view.lstCommands.ListCount - 1
                        Require NxGeneratedNavigationRouteField(CStr(view.lstCommands.List(rowIndex, 0)), "category_id") = categoryId, "search escaped category"
                    Next rowIndex
                End If
            Next categoryIndex
        Case "empty_results"
            view.lstCommands.ListIndex = 0
            view.txtSearch.Value = "__r61_no_catalog_match__"
            Require view.lstCommands.ListCount = 0 And view.lstCommands.ListIndex = -1, "empty query result"
            Require Not view.cmdAssign.Enabled And Not view.cmdRemove.Enabled, "empty result actions"
            Require CStr(view.txtDescription.Value) = vbNullString And CStr(view.txtShortcut.Value) = vbNullString, "empty result details"
            view.txtSearch.Value = vbNullString
            view.cboCategory.AddItem "r61 empty category"
            view.cboCategory.List(view.cboCategory.ListCount - 1, 1) = "__r61_empty_category__"
            view.cboCategory.ListIndex = view.cboCategory.ListCount - 1
            Require view.lstCommands.ListCount = 0, "empty category result"
            Require Not view.cmdAssign.Enabled And Not view.cmdRemove.Enabled, "empty category actions"
        Case "reselection"
            view.lstCommands.ListIndex = view.lstCommands.ListCount - 1
            routeKey = CStr(view.lstCommands.List(view.lstCommands.ListIndex, 0))
            view.txtSearch.Value = NxGeneratedNavigationRouteField(routeKey, "label_ko")
            Require SelectedRoute(view) = routeKey, "selected route after row reordering"
            view.txtSearch.Value = "__r61_no_catalog_match__"
            Require view.lstCommands.ListIndex = -1, "hidden selection is not substituted"
            view.txtSearch.Value = vbNullString
            Require SelectedRoute(view) = routeKey, "restore hidden selection by route key"
            categoryId = NxGeneratedNavigationRouteField(routeKey, "category_id")
            For categoryIndex = 1 To view.cboCategory.ListCount - 1
                If CStr(view.cboCategory.List(categoryIndex, 1)) <> categoryId Then
                    view.cboCategory.ListIndex = categoryIndex
                    Require view.lstCommands.ListIndex = -1, "different category has no substitute target"
                    Exit For
                End If
            Next categoryIndex
            view.cboCategory.ListIndex = 0
            Require SelectedRoute(view) = routeKey, "restore category-hidden selection"
        Case "description"
            For rowIndex = 0 To view.lstCommands.ListCount - 1
                view.lstCommands.ListIndex = rowIndex
                routeKey = CStr(view.lstCommands.List(rowIndex, 0))
                description = NxGeneratedNavigationRouteField(routeKey, "label_ko") & vbCrLf & vbCrLf & _
                    NxGeneratedNavigationRouteField(routeKey, "description_ko")
                Require CStr(view.txtDescription.Value) = description, "selected command catalog description"
            Next rowIndex
            Require view.txtDescription.Locked And view.txtDescription.MultiLine, "read-only description"
            Require view.txtDescription.ScrollBars = fmScrollBarsVertical, "description overflow"
        Case "binding_display"
            Require Len(assignedRouteKey) > 0, "harness must supply a route pre-bound in its isolated profile"
            binding = NxShortcutsBindingDisplayForRoute(assignedRouteKey)
            Require Len(binding) > 0, "fixture route must already have a binding"
            rowIndex = FindRoute(view, assignedRouteKey)
            Require rowIndex >= 0, "assigned route is eligible"
            view.lstCommands.ListIndex = rowIndex
            Require CStr(view.lstCommands.List(rowIndex, 2)) = binding, "binding column"
            Require CStr(view.txtShortcut.Value) = binding, "selected binding capture"
            view.txtSearch.Value = "__r61_no_catalog_match__"
            view.txtSearch.Value = vbNullString
            Require SelectedRoute(view) = assignedRouteKey, "bound route reselected"
            Require CStr(view.txtShortcut.Value) = binding And view.cmdRemove.Enabled, "reselected binding and remove target"
            Require NxShortcutsBindingDisplayForRoute(assignedRouteKey) = binding, "filter preserved persisted binding"
        Case "lifecycle"
            Require Not NxFormWheelIsAttached(), "fixture requires no pre-existing wheel owner"
            Require view.cmdAssign.Default And view.cmdClose.Cancel, "Enter and Esc buttons"
            Require view.cboCategory.TabIndex < view.txtSearch.TabIndex And view.txtSearch.TabIndex < view.lstCommands.TabIndex, "filter Tab order"
            Require view.lstCommands.TabIndex < view.txtDescription.TabIndex And view.txtDescription.TabIndex < view.txtShortcut.TabIndex, "detail Tab order"
            view.Show vbModeless
            DoEvents
            Require NxFormWheelIsAttached(), "wheel attached on activate"
            Unload view
            Set view = Nothing
            Require Not NxFormWheelIsAttached(), "wheel detached on close"
        Case Else
            Err.Raise vbObjectError + 761, "T_R61ShortcutCatalog", "Unknown case: " & kind
    End Select
    NxR61ShortcutCatalogRunCase = "PASS|" & kind
    GoTo CleanUp
Failed:
    NxR61ShortcutCatalogRunCase = "FAIL|" & kind & "|" & Err.Description
CleanUp:
    On Error Resume Next
    If Not view Is Nothing Then Unload view
    Set view = Nothing
End Function

Private Function FindRoute(ByVal view As FNxShortcutManager, ByVal routeKey As String) As Long
    Dim index As Long
    FindRoute = -1
    For index = 0 To view.lstCommands.ListCount - 1
        If StrComp(CStr(view.lstCommands.List(index, 0)), routeKey, vbBinaryCompare) = 0 Then
            FindRoute = index
            Exit Function
        End If
    Next index
End Function

Private Function SelectedRoute(ByVal view As FNxShortcutManager) As String
    If view.lstCommands.ListIndex >= 0 Then SelectedRoute = CStr(view.lstCommands.List(view.lstCommands.ListIndex, 0))
End Function

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 762, "T_R61ShortcutCatalog", message
End Sub
