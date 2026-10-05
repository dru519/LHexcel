Attribute VB_Name = "NxRibbonCallbacks"
Option Explicit

Private mRibbon As Object

Public Sub NxRibbonOnLoad(ByVal ribbon As Object)
    Set mRibbon = ribbon
    On Error Resume Next
    NxRouteAvailabilityStartEvents
    NxDistributionStart
    NxShortcutsApplySavedBindings
    On Error GoTo 0
End Sub

Public Sub NxRibbonExecute(ByVal control As Object)
    Dim detail As String
    On Error GoTo Failed
    If control Is Nothing Then NxRaiseContractError "Ribbon control is required"
    NxRibbonExecuteTag CStr(control.Tag)
    NxRibbonInvalidateAvailability
    Exit Sub
Failed:
    detail = Err.Description
    Err.Clear
    NxRibbonInvalidateAvailability
    If Len(detail) = 0 Then detail = "작업을 실행하지 못했습니다."
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 작업 오류"
End Sub

Public Sub NxRibbonExecuteTag(ByVal tag As String)
    Dim target As CNxRibbonTarget
    Set target = NxParseRibbonTag(tag)
    NxRouteRibbonTarget target
End Sub

Public Sub NxRibbonOpenStart(ByVal control As Object)
    NxRibbonExecuteTag "nx1|entry|NX-ENTRY-MANAGEMENT"
End Sub

Public Sub NxRibbonOpenAi(ByVal control As Object)
    NxRibbonExecuteTag "nx1|entry|NX-ENTRY-AI"
End Sub
Public Sub NxRibbonOpenAll(ByVal control As Object)
    NxRibbonExecuteTag "nx1|entry|NX-ENTRY-ALL"
End Sub

Public Sub NxRibbonGetAllFunctionsContent(ByVal control As Object, ByRef returnedVal)
    returnedVal = NxRibbonScopedMenuXml(control, NxRibbonAllFunctionsMenuXml())
End Sub

Public Sub NxRibbonGetFavoritesContent(ByVal control As Object, ByRef returnedVal)
    returnedVal = NxRibbonScopedMenuXml(control, NxRibbonFavoritesMenuXml())
End Sub

Public Sub NxRibbonGetFocusContent(ByVal control As Object, ByRef returnedVal)
    returnedVal = NxRibbonScopedMenuXml(control, NxRibbonFocusMenuXml())
End Sub

Public Sub NxRibbonGetManagementContent(ByVal control As Object, ByRef returnedVal)
    returnedVal = NxRibbonScopedMenuXml(control, NxRibbonManagementMenuXml())
End Sub

Public Sub NxRibbonGetEnabled(ByVal control As Object, ByRef returnedVal)
    On Error GoTo Unavailable
    If control Is Nothing Then GoTo Unavailable
    If Not NxDistributionCanExecute() Then GoTo Unavailable
    ' Excel caches this callback. An idle read-only utility must not be disabled
    ' permanently because Ready was transiently False during Ribbon initialization.
    If CStr(control.Tag) = "nx1|feature|NX-UTIL-CALCULATOR" Then
        returnedVal = True
        Exit Sub
    End If
    returnedVal = (NxRouteAvailabilityState(NxRouteKeyFromRibbonTag(CStr(control.Tag))) <> NX_ROUTE_UNAVAILABLE)
    Exit Sub
Unavailable:
    Err.Clear
    returnedVal = False
End Sub

Public Sub NxRibbonGetDistributionEnabled(ByVal control As Object, ByRef returnedVal)
    returnedVal = NxDistributionCanExecute()
End Sub

Public Sub NxRibbonInvalidateAvailability()
    If mRibbon Is Nothing Then Exit Sub
    On Error Resume Next
    mRibbon.Invalidate
    On Error GoTo 0
End Sub

Public Sub NxRibbonFavoriteAdd(ByVal control As Object)
    NxDistributionEnsureExecutable
    If control Is Nothing Then NxRaiseContractError "Favorite control is required"
    NxFavoritesAdd NxFavoriteTagRouteKey(CStr(control.Tag), "add")
    NxRibbonInvalidateFavorites
End Sub

Public Sub NxRibbonFavoriteRemove(ByVal control As Object)
    NxDistributionEnsureExecutable
    If control Is Nothing Then NxRaiseContractError "Favorite control is required"
    NxFavoritesRemove NxFavoriteTagRouteKey(CStr(control.Tag), "remove")
    NxRibbonInvalidateFavorites
End Sub

Public Sub NxRibbonInvalidateFavorites()
    If mRibbon Is Nothing Then Exit Sub
    On Error Resume Next
    mRibbon.InvalidateControl "NX-HOME-FAVORITES"
    mRibbon.InvalidateControl "NX-PROD-FAVORITES"
    On Error GoTo 0
End Sub

Public Sub NxRibbonToggleFocus(ByVal control As Object, ByVal pressed As Boolean)
    NxDistributionEnsureExecutable
    NxFocusSetEnabled pressed
End Sub

Public Sub NxRibbonGetFocusPressed(ByVal control As Object, ByRef returnedVal)
    returnedVal = NxFocusIsEnabled()
End Sub

Public Sub NxRibbonInvalidateFocus()
    If mRibbon Is Nothing Then Exit Sub
    On Error Resume Next
    mRibbon.InvalidateControl "NX-HOME-FOCUS"
    mRibbon.InvalidateControl "NX-HOME-FOCUS-TOGGLE"
    mRibbon.InvalidateControl "NX-PROD-FOCUS"
    mRibbon.InvalidateControl "NX-PROD-FOCUS-TOGGLE"
    mRibbon.InvalidateControl "NX-UTIL-FOCUS-TOGGLE"
    On Error GoTo 0
End Sub

' Short aliases are retained for RibbonX declarations authored by older shells.
Public Sub NxRibbonGetAllContent(ByVal control As Object, ByRef returnedVal)
    NxRibbonGetAllFunctionsContent control, returnedVal
End Sub

Public Sub NxRibbonGetFavorites(ByVal control As Object, ByRef returnedVal)
    NxRibbonGetFavoritesContent control, returnedVal
End Sub
