Attribute VB_Name = "NxFavorites"
Option Explicit

Private Const NX_FAVORITES_FILE As String = "favorites-v2.cfg"
Private Const NX_FAVORITES_VERSION As String = "NXFAV2"
Private Const NX_FAVORITES_LEGACY_FILE As String = "favorites-v1.cfg"
Private Const NX_FAVORITES_LEGACY_VERSION As String = "NXFAV1"
' Stored route examples: feature:NX-... and command:NX-CMD-...

Private mFavorites As Collection
Private mMigratedFromV1 As Boolean

Private Sub EnsureFavorites()
    If mFavorites Is Nothing Then
        Set mFavorites = New Collection
        NxFavoritesLoad
    End If
End Sub

Public Sub NxFavoritesOpenMenu()
    NxFavoritesShowManager
End Sub

Public Sub NxFavoritesShowManager()
    Dim manager As New FNxFavoritesManager
    manager.Show vbModal
End Sub

Public Function NxFavoritesCount() As Long
    EnsureFavorites
    NxFavoritesCount = mFavorites.Count
End Function

Public Function NxFavoritesRouteKeys() As Collection
    Dim result As New Collection
    Dim value As Variant
    EnsureFavorites
    For Each value In mFavorites
        result.Add CStr(value)
    Next value
    Set NxFavoritesRouteKeys = result
End Function

' Compatibility alias for older internal callers. Values are route keys in v2.
Public Function NxFavoritesIds() As Collection
    Set NxFavoritesIds = NxFavoritesRouteKeys()
End Function

Public Function NxFavoriteRouteKind(ByVal routeKey As String) As String
    Dim separator As Long
    separator = InStr(1, routeKey, ":", vbBinaryCompare)
    If separator <= 1 Or InStr(separator + 1, routeKey, ":", vbBinaryCompare) > 0 Then _
        NxRaiseContractError "즐겨찾기 경로 형식이 올바르지 않습니다."
    NxFavoriteRouteKind = Left$(routeKey, separator - 1)
    If NxFavoriteRouteKind <> "feature" And NxFavoriteRouteKind <> "command" Then _
        NxRaiseContractError "즐겨찾기 경로 종류가 올바르지 않습니다."
End Function

Public Function NxFavoriteRouteIdentifier(ByVal routeKey As String) As String
    Dim separator As Long
    Call NxFavoriteRouteKind(routeKey)
    separator = InStr(1, routeKey, ":", vbBinaryCompare)
    NxFavoriteRouteIdentifier = Mid$(routeKey, separator + 1)
    If Not NxRibbonIdentifierIsSafe(NxFavoriteRouteIdentifier) Then _
        NxRaiseContractError "즐겨찾기 항목 ID가 올바르지 않습니다."
End Function

Public Function NxFavoritesRouteIsKnown(ByVal routeKey As String) As Boolean
    On Error GoTo Unknown
    Call NxFavoriteRouteIdentifier(routeKey)
    NxFavoritesRouteIsKnown = NxGeneratedNavigationRouteExists(routeKey)
Unknown:
End Function

Public Function NxFavoritesRouteLabel(ByVal routeKey As String) As String
    If NxFavoritesRouteIsKnown(routeKey) Then
        NxFavoritesRouteLabel = NxGeneratedNavigationRouteField(routeKey, "label_ko")
    Else
        NxFavoritesRouteLabel = "사용할 수 없는 항목 · " & routeKey
    End If
End Function

Public Function NxFavoritesIsEligible(ByVal routeKey As String) As Boolean
    If Not NxFavoritesRouteIsKnown(routeKey) Then Exit Function
    Select Case routeKey
        Case "feature:NX-DATA-DUPLICATE-LIST"
            NxFavoritesIsEligible = False
        Case Else
            NxFavoritesIsEligible = True
    End Select
End Function

Public Function NxFavoritesContains(ByVal routeKey As String) As Boolean
    EnsureFavorites
    NxFavoritesContains = NxFavoritesContainsLoaded(routeKey)
End Function

Public Sub NxFavoritesAdd(ByVal routeKey As String)
    Dim replacement As Collection
    EnsureFavorites
    Call NxFavoriteRouteIdentifier(routeKey)
    If Not NxFavoritesRouteIsKnown(routeKey) Then NxRaiseContractError "현재 버전에 없는 항목은 새로 등록할 수 없습니다."
    If Not NxFavoritesIsEligible(routeKey) Then NxRaiseContractError "이 항목은 즐겨찾기에 등록할 수 없습니다."
    If NxFavoritesContainsLoaded(routeKey) Then Exit Sub
    Set replacement = NxFavoritesRouteKeys()
    replacement.Add routeKey
    NxFavoritesReplace replacement
End Sub

Public Sub NxFavoritesRemove(ByVal routeKey As String)
    Dim replacement As New Collection
    Dim value As Variant
    EnsureFavorites
    Call NxFavoriteRouteIdentifier(routeKey)
    If Not NxFavoritesContainsLoaded(routeKey) Then Exit Sub
    For Each value In mFavorites
        If StrComp(CStr(value), routeKey, vbBinaryCompare) <> 0 Then replacement.Add CStr(value)
    Next value
    NxFavoritesReplace replacement
End Sub

Public Function NxFavoriteTagRouteKey(ByVal rawTag As String, ByVal expectedAction As String) As String
    Dim parts As Variant
    Dim routeKey As String
    parts = Split(rawTag, "|", -1, vbBinaryCompare)
    If UBound(parts) <> 2 Then NxRaiseContractError "즐겨찾기 태그 형식이 올바르지 않습니다."
    If CStr(parts(0)) <> "fav2" Or CStr(parts(1)) <> expectedAction Then _
        NxRaiseContractError "즐겨찾기 작업 태그가 거부되었습니다."
    routeKey = CStr(parts(2))
    Call NxFavoriteRouteIdentifier(routeKey)
    If expectedAction = "add" Then
        If Not NxFavoritesRouteIsKnown(routeKey) Or Not NxFavoritesIsEligible(routeKey) Then _
            NxRaiseContractError "즐겨찾기에 등록할 수 없는 항목입니다."
    End If
    NxFavoriteTagRouteKey = routeKey
End Function

Public Sub NxFavoritesReplace(ByVal routeKeys As Collection)
    Dim previous As Collection
    Dim replacement As New Collection
    Dim routeKey As Variant
    Dim seen As Object
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureDescription As String

    If routeKeys Is Nothing Then NxRaiseContractError "즐겨찾기 목록이 필요합니다."
    EnsureFavorites
    Set previous = NxFavoritesRouteKeys()
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbBinaryCompare

    For Each routeKey In routeKeys
        Call NxFavoriteRouteIdentifier(CStr(routeKey))
        If NxFavoritesRouteIsKnown(CStr(routeKey)) Then
            If Not NxFavoritesIsEligible(CStr(routeKey)) Then _
                NxRaiseContractError "즐겨찾기에 등록할 수 없는 항목이 포함되어 있습니다."
        ElseIf Not NxCollectionContainsRoute(previous, CStr(routeKey)) Then
            NxRaiseContractError "현재 버전에 없는 항목은 새로 등록할 수 없습니다."
        End If
        If seen.Exists(CStr(routeKey)) Then NxRaiseContractError "즐겨찾기에 중복 항목이 포함되어 있습니다."
        seen.Add CStr(routeKey), True
        replacement.Add CStr(routeKey)
    Next routeKey

    On Error GoTo Failed
    Set mFavorites = replacement
    NxFavoritesSave
    Exit Sub

Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    Set mFavorites = previous
    Err.Raise failureNumber, failureSource, failureDescription
End Sub

Private Sub NxFavoritesLoad()
    Dim path As String
    path = NxFavoritesPath()
    If NxFavoritesFileExists(path) Then
        NxFavoritesLoadV2 path
    ElseIf NxFavoritesFileExists(NxFavoritesLegacyPath()) Then
        NxFavoritesLoadLegacyV1 NxFavoritesLegacyPath()
    End If
End Sub

Private Sub NxFavoritesLoadV2(ByVal path As String)
    Dim handle As Integer
    Dim header As String
    Dim routeKey As String
    On Error GoTo Rejected
    handle = FreeFile
    Open path For Input Access Read Lock Write As #handle
    If EOF(handle) Then GoTo Rejected
    Line Input #handle, header
    If StrComp(header, NX_FAVORITES_VERSION, vbBinaryCompare) <> 0 Then GoTo Rejected
    Do While Not EOF(handle)
        Line Input #handle, routeKey
        routeKey = Trim$(routeKey)
        If Len(routeKey) > 0 Then NxFavoritesTryAddLoaded routeKey
    Loop
    Close #handle
    Exit Sub

Rejected:
    On Error Resume Next
    If handle > 0 Then Close #handle
    Set mFavorites = New Collection
    On Error GoTo 0
End Sub

Private Sub NxFavoritesLoadLegacyV1(ByVal path As String)
    Dim featureId As String
    Dim handle As Integer
    Dim header As String
    On Error GoTo Rejected
    handle = FreeFile
    Open path For Input Access Read Lock Write As #handle
    If EOF(handle) Then GoTo Rejected
    Line Input #handle, header
    If StrComp(header, NX_FAVORITES_LEGACY_VERSION, vbBinaryCompare) <> 0 Then GoTo Rejected
    Do While Not EOF(handle)
        Line Input #handle, featureId
        featureId = Trim$(featureId)
        If Len(featureId) > 0 And NxRibbonIdentifierIsSafe(featureId) Then _
            NxFavoritesTryAddLoaded "feature:" & featureId
    Loop
    Close #handle
    mMigratedFromV1 = True
    Exit Sub

Rejected:
    On Error Resume Next
    If handle > 0 Then Close #handle
    Set mFavorites = New Collection
    mMigratedFromV1 = False
    On Error GoTo 0
End Sub

Private Sub NxFavoritesTryAddLoaded(ByVal routeKey As String)
    On Error GoTo Rejected
    If routeKey = "command:NX-CMD-SHEET-RB-SHEET-SAVE-TO-FILE" Then Exit Sub
    routeKey = NxCanonicalRouteKey(routeKey)
    Call NxFavoriteRouteIdentifier(routeKey)
    If NxFavoritesRouteIsKnown(routeKey) Then
        If Not NxFavoritesIsEligible(routeKey) Then Exit Sub
    End If
    If Not NxFavoritesContainsLoaded(routeKey) Then mFavorites.Add routeKey
Rejected:
End Sub

Private Function NxFavoritesContainsLoaded(ByVal routeKey As String) As Boolean
    Dim value As Variant
    For Each value In mFavorites
        If StrComp(CStr(value), routeKey, vbBinaryCompare) = 0 Then
            NxFavoritesContainsLoaded = True
            Exit Function
        End If
    Next value
End Function

Private Function NxCollectionContainsRoute(ByVal values As Collection, ByVal routeKey As String) As Boolean
    Dim value As Variant
    For Each value In values
        If StrComp(CStr(value), routeKey, vbBinaryCompare) = 0 Then
            NxCollectionContainsRoute = True
            Exit Function
        End If
    Next value
End Function

Private Sub NxFavoritesSave()
    Dim backup As String
    Dim errorDescription As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim handle As Integer
    Dim movedOriginal As Boolean
    Dim path As String
    Dim temporary As String
    Dim value As Variant
    On Error GoTo Failed

    NxFavoritesEnsureSettingsFolder
    path = NxFavoritesPath()
    temporary = path & ".tmp"
    backup = path & ".bak"
    If NxFavoritesFileExists(temporary) Then Kill temporary
    If NxFavoritesFileExists(backup) Then Kill backup

    handle = FreeFile
    Open temporary For Output Access Write Lock Read Write As #handle
    Print #handle, NX_FAVORITES_VERSION
    For Each value In mFavorites
        Call NxFavoriteRouteIdentifier(CStr(value))
        Print #handle, CStr(value)
    Next value
    Close #handle
    handle = 0

    If NxFavoritesFileExists(path) Then
        Name path As backup
        movedOriginal = True
    End If
    Name temporary As path
    On Error Resume Next
    If NxFavoritesFileExists(backup) Then Kill backup
    On Error GoTo 0
    mMigratedFromV1 = False
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    On Error Resume Next
    If handle > 0 Then Close #handle
    If movedOriginal And Not NxFavoritesFileExists(path) Then
        If NxFavoritesFileExists(backup) Then Name backup As path
    End If
    If Len(temporary) > 0 And NxFavoritesFileExists(temporary) Then Kill temporary
    On Error GoTo 0
    Err.Raise errorNumber, errorSource, errorDescription
End Sub

Private Function NxFavoritesPath() As String
    NxFavoritesPath = NxLHexcelProfileRoot() & "\Settings\" & NX_FAVORITES_FILE
End Function

Private Function NxFavoritesLegacyPath() As String
    NxFavoritesLegacyPath = NxLHexcelProfileRoot() & "\Settings\" & NX_FAVORITES_LEGACY_FILE
End Function

Private Sub NxFavoritesEnsureSettingsFolder()
    Dim root As String
    Dim settingsFolder As String
    root = NxLHexcelProfileRoot()
    settingsFolder = root & "\Settings"
    If Not NxFavoritesFolderExists(root) Then MkDir root
    If Not NxFavoritesFolderExists(settingsFolder) Then MkDir settingsFolder
End Sub

Private Function NxFavoritesFolderExists(ByVal path As String) As Boolean
    On Error GoTo Missing
    NxFavoritesFolderExists = ((GetAttr(path) And vbDirectory) = vbDirectory)
Missing:
End Function

Private Function NxFavoritesFileExists(ByVal path As String) As Boolean
    On Error GoTo Missing
    NxFavoritesFileExists = (Len(Dir$(path, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
Missing:
End Function
