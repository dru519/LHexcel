Attribute VB_Name = "NxRibbonMenuCatalog"
Option Explicit

Private Const NX_RIBBON_NAMESPACE As String = "http://schemas.microsoft.com/office/2009/07/customui"

Public Function NxRibbonAllFunctionsMenuXml() As String
    NxRibbonAllFunctionsMenuXml = NxGeneratedNavigationMenuXml()
End Function

Public Function NxRibbonScopedMenuXml(ByVal control As Object, ByVal xml As String) As String
    Dim idPrefix As String
    If control Is Nothing Then NxRaiseContractError "Ribbon control is required"
    idPrefix = "dyn_" & NxRibbonSafeId(CStr(control.Id)) & "_"
    xml = Replace$(xml, " id=""", " id=""" & idPrefix)
    xml = Replace$(xml, " enabled=""true""", " getEnabled=""NxRibbonGetEnabled""")
    xml = Replace$(xml, " enabled=""false""", " getEnabled=""NxRibbonGetEnabled""")
    NxRibbonScopedMenuXml = NxRibbonBrandMenuXml(xml)
End Function

Private Function NxRibbonBrandMenuXml(ByVal xml As String) As String
    Dim parts As Variant, i As Long, tagStart As Long, tagEnd As Long
    Dim imageStart As Long, imageEnd As Long, imageId As String, part As String
    parts = Split(xml, ">")
    For i = LBound(parts) To UBound(parts)
        part = CStr(parts(i))
        tagStart = InStr(1, part, " tag=""", vbBinaryCompare)
        imageStart = InStr(1, part, " imageMso=""", vbBinaryCompare)
        If tagStart > 0 And imageStart > 0 Then
            tagStart = tagStart + Len(" tag=""")
            tagEnd = InStr(tagStart, part, """")
            If tagEnd > tagStart Then
                imageId = NxGeneratedBrandImage(Mid$(part, tagStart, tagEnd - tagStart))
                If Len(imageId) > 0 Then
                    imageEnd = InStr(imageStart + Len(" imageMso="""), part, """")
                    If imageEnd > imageStart Then parts(i) = Left$(part, imageStart - 1) & " image=""" & imageId & """" & Mid$(part, imageEnd + 1)
                End If
            End If
        End If
    Next i
    NxRibbonBrandMenuXml = Join(parts, ">")
End Function

Public Function NxRibbonFocusMenuXml() As String
    NxRibbonFocusMenuXml = NxRibbonMenuRoot() & _
        "<button id=""focus_toggle"" label=""포커스셀 켜기/끄기"" imageMso=""TableStyleBandedColumns"" " & _
        "tag=""nx1|feature|NX-DATA-FOCUS-CELL"" onAction=""NxRibbonExecute""/>" & _
        "<button id=""focus_settings"" label=""포커스셀 설정"" imageMso=""ControlProperties"" " & _
        "tag=""nx1|entry|NX-ENTRY-FOCUS-SETTINGS"" onAction=""NxRibbonExecute""/>" & _
        "<button id=""focus_reset"" label=""기본값 복원"" imageMso=""RefreshAll"" " & _
        "tag=""nx1|entry|NX-ENTRY-FOCUS-RESET"" onAction=""NxRibbonExecute""/>" & _
        "</menu>"
End Function

Public Function NxRibbonFavoritesMenuXml() As String
    Dim routeKey As Variant
    Dim routes As Collection
    Dim xml As String
    Set routes = NxFavoritesRouteKeys()
    xml = NxRibbonMenuRoot() & NxFavoritesEditorButtonXml() & _
        "<menuSeparator id=""fav_editor_separator""/>"

    If routes.Count = 0 Then
        xml = xml & "<button id=""fav_empty"" label=""등록된 즐겨찾기가 없습니다."" enabled=""false""/>"
    Else
        For Each routeKey In routes
            xml = xml & NxFavoriteRouteButtonXml(CStr(routeKey))
        Next routeKey
    End If

    xml = xml & NxFavoritesAddMenuXml()
    If routes.Count > 0 Then xml = xml & NxFavoritesRemoveMenuXml(routes)
    NxRibbonFavoritesMenuXml = xml & "</menu>"
End Function

Private Function NxFavoritesEditorButtonXml() As String
    NxFavoritesEditorButtonXml = "<button id=""fav_editor"" label=""즐겨찾기 편집"" " & _
        "imageMso=""AddToFavorites"" tag=""nx1|entry|NX-ENTRY-FAVORITES"" " & _
        "onAction=""NxRibbonExecute""/>"
End Function

Private Function NxFavoriteRouteButtonXml(ByVal routeKey As String) As String
    Dim identifier As String
    Dim routeKind As String
    If Not NxFavoritesRouteIsKnown(routeKey) Then
        NxFavoriteRouteButtonXml = "<button id=""fav_missing_" & NxRibbonSafeId(routeKey) & _
            """ label=""" & NxRibbonXmlEscape(NxFavoritesRouteLabel(routeKey)) & _
            """ imageMso=""Info"" enabled=""false"" supertip=""즐겨찾기 편집에서 이 항목을 제거할 수 있습니다.""/>"
        Exit Function
    End If

    routeKind = NxFavoriteRouteKind(routeKey)
    identifier = NxFavoriteRouteIdentifier(routeKey)
    NxFavoriteRouteButtonXml = "<button id=""fav_" & NxRibbonSafeId(routeKey) & _
        """ label=""" & NxRibbonXmlEscape(NxGeneratedNavigationRouteField(routeKey, "label_ko")) & _
        """ imageMso=""" & NxRibbonXmlEscape(NxGeneratedNavigationRouteField(routeKey, "image_mso")) & _
        """ tag=""nx1|" & routeKind & "|" & identifier & _
        """ enabled=""" & NxRouteAvailabilityRibbonValue(routeKey) & _
        """ supertip=""" & NxRibbonXmlEscape(NxRouteAvailabilityMenuTip(routeKey)) & _
        """ onAction=""NxRibbonExecute""/>"
End Function

Private Function NxFavoritesAddMenuXml() As String
    Dim categoryId As Variant
    Dim categories As Variant
    Dim innerXml As String
    Dim xml As String
    categories = Array("NX-GRP-SAVE", "NX-GRP-PRINT", "NX-GRP-COPY", "NX-GRP-INSERT", _
        "NX-GRP-FILE", "NX-GRP-TEMPLATE", "NX-GRP-DATA", "NX-GRP-CELL-FIT", _
        "NX-GRP-FORMULA", "NX-GRP-SHEET", "NX-GRP-VIEW", "NX-GRP-STYLE", _
        "NX-GRP-UTIL", "NX-GRP-INFO")
    xml = "<menu id=""fav_add"" label=""즐겨찾기 추가"" imageMso=""AddToFavorites"">"
    For Each categoryId In categories
        innerXml = NxFavoritesAddCategoryItemsXml(CStr(categoryId))
        If Len(innerXml) > 0 Then
            xml = xml & "<menu id=""fav_add_" & NxRibbonSafeId(CStr(categoryId)) & _
                """ label=""" & NxRibbonXmlEscape(NxRibbonCategoryLabel(CStr(categoryId))) & _
                """ imageMso=""" & NxRibbonCategoryImage(CStr(categoryId)) & """>" & innerXml & "</menu>"
        End If
    Next categoryId
    NxFavoritesAddMenuXml = xml & "</menu>"
End Function

Private Function NxFavoritesAddCategoryItemsXml(ByVal categoryId As String) As String
    Dim item As Variant
    Dim routeKey As String
    Dim xml As String
    For Each item In NxGeneratedNavigationItems()
        routeKey = CStr(item(0))
        If CStr(item(5)) = categoryId Then
            If NxFavoritesIsEligible(routeKey) And Not NxFavoritesContains(routeKey) Then _
                xml = xml & NxFavoriteActionButtonXml(routeKey, "add")
        End If
    Next item
    NxFavoritesAddCategoryItemsXml = xml
End Function

Private Function NxFavoritesRemoveMenuXml(ByVal routes As Collection) As String
    Dim routeKey As Variant
    Dim xml As String
    xml = "<menu id=""fav_remove"" label=""즐겨찾기 제거"" imageMso=""Delete"">"
    For Each routeKey In routes
        xml = xml & NxFavoriteActionButtonXml(CStr(routeKey), "remove")
    Next routeKey
    NxFavoritesRemoveMenuXml = xml & "</menu>"
End Function

Private Function NxFavoriteActionButtonXml(ByVal routeKey As String, ByVal actionName As String) As String
    Dim callbackName As String
    Dim imageMso As String
    If actionName = "add" Then callbackName = "NxRibbonFavoriteAdd" Else callbackName = "NxRibbonFavoriteRemove"
    If NxFavoritesRouteIsKnown(routeKey) Then
        imageMso = NxGeneratedNavigationRouteField(routeKey, "image_mso")
    Else
        imageMso = "Info"
    End If
    NxFavoriteActionButtonXml = "<button id=""fav_" & actionName & "_" & NxRibbonSafeId(routeKey) & _
        """ label=""" & NxRibbonXmlEscape(NxFavoritesRouteLabel(routeKey)) & _
        """ imageMso=""" & NxRibbonXmlEscape(imageMso) & _
        """ tag=""fav2|" & actionName & "|" & routeKey & """ onAction=""" & callbackName & """/>"
End Function

Public Function NxRibbonManagementMenuXml() As String
    NxRibbonManagementMenuXml = NxRibbonMenuRoot()
    If Not NxDistributionCanExecute() Then
        NxRibbonManagementMenuXml = NxRibbonManagementMenuXml & _
            "<button id=""distribution_status"" label=""사용기한 및 보호 안내"" imageMso=""Info"" tag=""nx1|entry|NX-MGMT-DISTRIBUTION-STATUS"" onAction=""NxRibbonExecute""/>"
        If NxDistributionIsStandaloneXlam() Then
            NxRibbonManagementMenuXml = NxRibbonManagementMenuXml & NxManagementButton("NX-MGMT-REMOVE", "내엑셀 제거", "Delete")
        End If
        NxRibbonManagementMenuXml = NxRibbonManagementMenuXml & _
            NxManagementButton("NX-MGMT-FILE-INFO", "내엑셀 정보", "Info") & "</menu>"
        Exit Function
    End If
    If NxDistributionIsStandaloneXlam() Then
        NxRibbonManagementMenuXml = NxRibbonManagementMenuXml & _
            "<menuSeparator id=""mgmt_install_separator"" title=""내엑셀 설치/삭제""/>" & _
            NxManagementButton("NX-MGMT-INSTALL", "내엑셀 설치", "AddInManager") & _
            NxManagementButton("NX-MGMT-REMOVE", "내엑셀 제거", "Delete")
    End If
    NxRibbonManagementMenuXml = NxRibbonManagementMenuXml & _
        "<menuSeparator id=""mgmt_environment_separator"" title=""내엑셀 환경설정""/>" & _
        NxManagementButton("NX-MGMT-INSTALL-FOLDER", "내엑셀 설치 폴더 열기", "FileOpen") & _
        NxManagementButton("NX-MGMT-QAT", "빠른 실행 도구 모음 설정", "QuickAccessToolbarMoreCommands") & _
        NxManagementButton("NX-MGMT-EXCEL-ADDINS", "Excel 추가 기능 설정", "AddInManager") & _
        NxManagementButton("NX-MGMT-SHORTCUTS", "단축키 관리", "ControlProperties") & _
        NxManagementButton("NX-MGMT-FILE-INFO", "내엑셀 정보", "Info") & "</menu>"
End Function

Private Function NxManagementButton(ByVal entryId As String, ByVal label As String, ByVal imageMso As String) As String
    NxManagementButton = "<button id=""" & entryId & """ label=""" & NxRibbonXmlEscape(label) & _
        """ imageMso=""" & imageMso & """ tag=""nx1|entry|" & entryId & _
        """ onAction=""NxRibbonExecute""/>"
End Function

Private Function NxRibbonMenuRoot() As String
    NxRibbonMenuRoot = "<menu xmlns=""" & NX_RIBBON_NAMESPACE & """>"
End Function

Private Function NxRibbonCategoryLabel(ByVal categoryId As String) As String
    Select Case categoryId
        Case "NX-GRP-SAVE": NxRibbonCategoryLabel = "저장"
        Case "NX-GRP-PRINT": NxRibbonCategoryLabel = "인쇄"
        Case "NX-GRP-COPY": NxRibbonCategoryLabel = "복붙"
        Case "NX-GRP-INSERT": NxRibbonCategoryLabel = "삽입"
        Case "NX-GRP-FILE": NxRibbonCategoryLabel = "파일관리"
        Case "NX-GRP-TEMPLATE": NxRibbonCategoryLabel = "템플릿"
        Case "NX-GRP-DATA": NxRibbonCategoryLabel = "데이터"
        Case "NX-GRP-CELL-FIT": NxRibbonCategoryLabel = "셀맞춤"
        Case "NX-GRP-FORMULA": NxRibbonCategoryLabel = "수식·참조"
        Case "NX-GRP-SHEET": NxRibbonCategoryLabel = "시트"
        Case "NX-GRP-VIEW": NxRibbonCategoryLabel = "보기창"
        Case "NX-GRP-STYLE": NxRibbonCategoryLabel = "스타일"
        Case "NX-GRP-UTIL": NxRibbonCategoryLabel = "추가기능"
        Case "NX-GRP-INFO": NxRibbonCategoryLabel = "정보진단"
        Case Else: NxRaiseContractError "Unknown Ribbon category label"
    End Select
End Function

Private Function NxRibbonCategoryImage(ByVal categoryId As String) As String
    Select Case categoryId
        Case "NX-GRP-SAVE": NxRibbonCategoryImage = "FileSave"
        Case "NX-GRP-PRINT": NxRibbonCategoryImage = "PrintPreviewAndPrint"
        Case "NX-GRP-COPY": NxRibbonCategoryImage = "Copy"
        Case "NX-GRP-INSERT": NxRibbonCategoryImage = "CellsInsertDialog"
        Case "NX-GRP-FILE": NxRibbonCategoryImage = "FileOpen"
        Case "NX-GRP-TEMPLATE": NxRibbonCategoryImage = "TableInsert"
        Case "NX-GRP-DATA": NxRibbonCategoryImage = "DatabaseInsert"
        Case "NX-GRP-CELL-FIT": NxRibbonCategoryImage = "TableAutoFitContents"
        Case "NX-GRP-FORMULA": NxRibbonCategoryImage = "FunctionWizard"
        Case "NX-GRP-SHEET": NxRibbonCategoryImage = "CellsInsertDialog"
        Case "NX-GRP-VIEW": NxRibbonCategoryImage = "ViewNormalViewExcel"
        Case "NX-GRP-STYLE": NxRibbonCategoryImage = "TableStylesGallery"
        Case "NX-GRP-UTIL": NxRibbonCategoryImage = "AddInManager"
        Case "NX-GRP-INFO": NxRibbonCategoryImage = "FileProperties"
        Case Else: NxRibbonCategoryImage = "AddToFavorites"
    End Select
End Function

Private Function NxRibbonSafeId(ByVal value As String) As String
    value = Replace$(value, ":", "_")
    NxRibbonSafeId = Replace$(value, "-", "_")
End Function

Public Function NxRibbonXmlEscape(ByVal value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    value = Replace(value, """", "&quot;")
    NxRibbonXmlEscape = Replace(value, "'", "&apos;")
End Function
