Attribute VB_Name = "NxCommandHandlers"
Option Explicit

Private Function NxSelectedRange() As Excel.Range
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "A cell selection is required"
    Set NxSelectedRange = Application.Selection
End Function

Private Function NxActiveUserWorkbook() As Excel.Workbook
    If Application.ActiveWorkbook Is Nothing Then NxRaiseContractError "An active workbook is required"
    If Application.ActiveWorkbook.ProtectStructure Then NxRaiseContractError "The workbook structure is protected"
    Set NxActiveUserWorkbook = Application.ActiveWorkbook
End Function

Private Sub NxReject(ByVal key As String)
    NxRaiseContractError "Command argument rejected: " & key
End Sub

Public Sub NxCmdClipboard(ByVal key As String)
    Dim r As Excel.Range
    Set r = NxSelectedRange()
    Select Case key
        Case "RB_CLIPBOARD_COPY_BYFORMULA": NxClipboardWriteUnicode NxFormulaClipboardText(r)
        Case "RB_CLIPBOARD_COPY_BYREFERENCE": NxClipboardWriteUnicode r.Cells(1, 1).Address(External:=True)
        Case "RB_CLIPBOARD_COPY_BYTEXT": NxClipboardWriteUnicode NxTextClipboardText(r)
        Case "RB_LHECLIPBOARD_COPYVISIBLE": SetVisibleCopyResult NxVisibleCellsCopy(r)
        Case "RB_LHECLIPBOARD_PASTEVISIBLEVALUES": SetVisiblePasteResult NxVisibleCellsPasteValues(r)
        Case "RB_LHECLIPBOARD_PASTEVISIBLEFORMULAS": NxPasteVisibleExcelClipboard r, xlPasteFormulas
        Case "RB_LHECLIPBOARD_PASTEVISIBLEFORMATS": NxPasteVisibleExcelClipboard r, xlPasteFormats
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdPrint(ByVal key As String)
    Dim r As Excel.Range
    Set r = NxSelectedRange()
    Select Case key
        Case "RB_PRINT_SETUP_QUICK": NxConfigurePrint r, xlPortrait, False, True
        Case "RB_PRINT_SETUP_QUICK_PORTRAIT": NxConfigurePrint r, xlPortrait, False, False
        Case "RB_PRINT_SETUP_QUICK_PORTRAIT_1BY1": NxConfigurePrint r, xlPortrait, True, False
        Case "RB_PRINT_SETUP_QUICK_LANDSCAPE": NxConfigurePrint r, xlLandscape, False, False
        Case "RB_PRINT_SETUP_QUICK_LANDSCAPE_1BY1": NxConfigurePrint r, xlLandscape, True, False
        Case "RB_PRINT_SETUP_REPEAT"
            If r.Areas.Count <> 1 Then NxRaiseContractError "Repeat print rows require one contiguous selection"
            r.Worksheet.PageSetup.PrintTitleRows = r.EntireRow.Address
        Case "RB_LHEPRINT_ADD1MMSPACER": r.Worksheet.PageSetup.LeftMargin = r.Worksheet.PageSetup.LeftMargin + Application.InchesToPoints(0.04)
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdStyle(ByVal key As String)
    Dim r As Excel.Range
    Set r = NxSelectedRange()
    Select Case key
        Case "NX_STYLE_FONT_SIZE_9": NxQuickFormatApply r, "font_size", 9
        Case "NX_STYLE_FONT_SIZE_10": NxQuickFormatApply r, "font_size", 10
        Case "NX_STYLE_FONT_SIZE_12": NxQuickFormatApply r, "font_size", 12
        Case "NX_STYLE_FONT_SIZE_15": NxQuickFormatApply r, "font_size", 15
        Case "NX_STYLE_FONT_COLOR_BLACK": NxQuickFormatApply r, "font_color", RGB(0, 0, 0)
        Case "NX_STYLE_FONT_COLOR_RED": NxQuickFormatApply r, "font_color", RGB(255, 0, 0)
        Case "NX_STYLE_FONT_COLOR_BLUE": NxQuickFormatApply r, "font_color", RGB(0, 0, 255)
        Case "NX_STYLE_FONT_COLOR_GREEN": NxQuickFormatApply r, "font_color", RGB(0, 128, 0)
        Case "NX_STYLE_FILL_GRAY": NxQuickFormatApply r, "fill_color", RGB(219, 219, 219)
        Case "NX_STYLE_FILL_LIGHT_RED": NxQuickFormatApply r, "fill_color", RGB(255, 202, 208)
        Case "NX_STYLE_FILL_LIGHT_YELLOW": NxQuickFormatApply r, "fill_color", RGB(255, 236, 161)
        Case "NX_STYLE_FILL_LIGHT_GREEN": NxQuickFormatApply r, "fill_color", RGB(201, 240, 208)
        Case "RB_LHESTYLE_RESETWORKBOOKSTYLES": NxResetCustomStyles NxActiveUserWorkbook()
        Case "RB_LHESTYLE_MONO_TITLE": NxRoleStyleApply r, NX_ROLE_STYLE_TITLE, NX_ROLE_STYLE_MODE_MONO
        Case "RB_LHESTYLE_COLOR_TITLE": NxRoleStyleApply r, NX_ROLE_STYLE_TITLE, NX_ROLE_STYLE_MODE_COLOR
        Case "RB_LHESTYLE_MONO_SUBTITLE": NxRoleStyleApply r, NX_ROLE_STYLE_SUBTITLE, NX_ROLE_STYLE_MODE_MONO
        Case "RB_LHESTYLE_COLOR_SUBTITLE": NxRoleStyleApply r, NX_ROLE_STYLE_SUBTITLE, NX_ROLE_STYLE_MODE_COLOR
        Case "RB_LHESTYLE_MONO_TABLEHEADER": NxRoleStyleApply r, NX_ROLE_STYLE_TABLE_HEADER, NX_ROLE_STYLE_MODE_MONO
        Case "RB_LHESTYLE_COLOR_TABLEHEADER": NxRoleStyleApply r, NX_ROLE_STYLE_TABLE_HEADER, NX_ROLE_STYLE_MODE_COLOR
        Case "RB_LHESTYLE_MONO_TABLEBODY": NxRoleStyleApply r, NX_ROLE_STYLE_TABLE_BODY, NX_ROLE_STYLE_MODE_MONO
        Case "RB_LHESTYLE_COLOR_TABLEBODY": NxRoleStyleApply r, NX_ROLE_STYLE_TABLE_BODY, NX_ROLE_STYLE_MODE_COLOR
        Case "RB_LHESTYLE_MONO_EMPHASISCELL": NxRoleStyleApply r, NX_ROLE_STYLE_EMPHASIS_CELL, NX_ROLE_STYLE_MODE_MONO
        Case "RB_LHESTYLE_COLOR_EMPHASISCELL": NxRoleStyleApply r, NX_ROLE_STYLE_EMPHASIS_CELL, NX_ROLE_STYLE_MODE_COLOR
        Case "RB_LHESTYLE_MONO_TOTALROW": NxRoleStyleApply r, NX_ROLE_STYLE_TOTAL_ROW, NX_ROLE_STYLE_MODE_MONO
        Case "RB_LHESTYLE_COLOR_TOTALROW": NxRoleStyleApply r, NX_ROLE_STYLE_TOTAL_ROW, NX_ROLE_STYLE_MODE_COLOR
        Case "RB_LHESTYLE_OPENOPTIONS": NxRouteFeature NX_FEATURE_DRAW_ROLE_STYLE
        Case "RB_EDIT_MEMO_ADD_LHEXCELFORMULA": NxAddFormulaComments r
        Case Else: NxReject key
    End Select
End Sub

Private Sub NxQuickFormatApply(ByVal selected As Excel.Range, ByVal propertyKey As String, ByVal value As Long)
    Dim area As Excel.Range
    Dim failure As Long, detail As String
    If selected.Worksheet.ProtectContents Then NxRaiseContractError "Quick formatting is unavailable on a protected worksheet"
    If selected.Worksheet.Parent.ReadOnly Then NxRaiseContractError "Quick formatting is unavailable in a read-only workbook"
    NxQuickFormatUndoCapture selected, propertyKey
    On Error GoTo Failed
    ' Touch only the requested property; values, formulas and other formats stay intact.
    For Each area In selected.Areas
        Select Case propertyKey
            Case "font_size": area.Font.Size = value
            Case "font_color": area.Font.Color = value
            Case "fill_color": area.Interior.Color = value
            Case Else: NxReject propertyKey
        End Select
    Next area
    NxQuickFormatUndoCommit
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description
    On Error Resume Next
    NxQuickFormatUndoRestore
    If Err.Number <> 0 Then detail = detail & " / 서식 복원: " & Err.Description
    On Error GoTo 0
    Err.Raise failure, "NxQuickFormatApply", detail
End Sub

Public Sub NxCmdNumberFormat(ByVal key As String)
    Dim r As Excel.Range
    Set r = NxSelectedRange()
    Select Case key
        Case "RB_EDIT_NUMBERFORMAT_NUMBER": r.NumberFormat = "#,##0"
        Case "RB_EDIT_NUMBERFORMAT_PERCENTAGE": NxNumberFormatOpen r, True
        Case "RB_EDIT_NUMBERFORMAT_DECIMAL": NxNumberFormatOpen r, False
        Case "NX_NUMBER_EMPHASIS": NxNumberFormatOpen r, False, True
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdSelection(ByVal key As String)
    Dim r As Excel.Range
    Set r = NxSelectedRange()
    Select Case key
        Case "NX_DATA_DATE_CONVERT": NxOpenDateConversion
        Case "RB_SELECTION_CELL_USEDRANGE": r.Worksheet.UsedRange.Select
        Case "RB_SELECTION_CELL_MOVE_LEFT": NxMoveSelection r, 0, -1, xlToLeft
        Case "RB_SELECTION_CELL_MOVE_RIGHT": NxMoveSelection r, 0, 1, xlToRight
        Case "RB_SELECTION_CELL_MOVE_DOWN": NxMoveSelection r, 1, 0, xlDown
        Case "RB_SELECTION_CELL_MOVE_UP": NxMoveSelection r, -1, 0, xlUp
        Case "RB_SELECTION_CELL_EXPAND_COL": r.EntireColumn.Select
        Case "RB_SELECTION_CELL_EXPAND_ROW": r.EntireRow.Select
        Case "RB_APP_OPTION_EDITDIRECTION_CHANGE": NxCycleEnterDirection
        Case "RB_SELECTION_CELL_TEXT_SPECIFIC": NxSelectCellsContainingText r.Worksheet
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdRowsColumns(ByVal key As String)
    Dim r As Excel.Range
    Set r = NxSelectedRange()
    Select Case key
        Case "RB_EDIT_CELL_RESIZE": NxResizeSelectionByDialog r
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdAlignment(ByVal key As String)
    Select Case key
        Case "RB_EDIT_ALIGN_CENTER_OVERCELLS": NxModelMergeOpen
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdInsertDelete(ByVal key As String)
    Dim r As Excel.Range
    Set r = NxSelectedRange()
    Select Case key
        Case "RB_EDIT_CELL_MAKEGROUP_HIDDEN_ROWCOLUMN": If r.Columns.Count >= r.Rows.Count Then r.EntireColumn.Group Else r.EntireRow.Group
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdFilterSort(ByVal key As String)
    Dim r As Excel.Range
    Set r = NxSelectedRange()
    Select Case key
        Case "RB_DATA_SORT_ASCENDING": NxSortCurrentRange r, xlAscending
        Case "RB_DATA_SORT_DESCENDING": NxSortCurrentRange r, xlDescending
        Case "RB_FILTER_AUTOFILTER_CANCEL": NxClearActiveColumnFilter r
        Case "RB_FILTER_AUTOFILTER_SHOWALL": NxShowAllFilteredData r.Worksheet
        Case "RB_FILTER_AUTOFILTER_FILTERING": NxFilterByActiveCell r, False
        Case "RB_FILTER_AUTOFILTER_FILTERING_OPTIONAL": NxFilterByActiveCell r, True
        Case "RB_DATA_SHOW_SOURCE": NxShowPivotSourceOrFilter r
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdFormula(ByVal key As String)
    Dim r As Excel.Range
    Set r = NxSelectedRange()
    Select Case key
        Case "NX_FUNCTION_ROUND": NxFunctionWrapOpen r, False
        Case "NX_FUNCTION_IFERROR": NxFunctionWrapOpen r, True
        Case "RB_FORMULA_PRECEDENTS_LIST": NxCreatePrecedentsIndex r
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdSheet(ByVal key As String)
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    Set wb = NxActiveUserWorkbook()
    Set ws = ActiveSheet
    If Not ws.Parent Is wb Then NxRaiseContractError "An active worksheet is required"
    Select Case key
        Case "RB_SHEET_SELECT_END": wb.Worksheets(wb.Worksheets.Count).Select
        Case "RB_SHEET_SELECT_HOME": wb.Worksheets(1).Select
        Case "RB_SHEET_SAVE_TO_FILE": NxRaiseContractError "삭제된 기능입니다."
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdView(ByVal key As String)
    Dim w As Excel.Window
    Set w = ActiveWindow
    If w Is Nothing Then NxRaiseContractError "An active window is required"
    Select Case key
        Case "NX_VIEW_PRESET_1": NxViewPresetApply w, 1
        Case "NX_VIEW_PRESET_2": NxViewPresetApply w, 2
        Case "NX_VIEW_PRESET_3": NxViewPresetApply w, 3
        Case "NX_VIEW_PRESET_4": NxViewPresetApply w, 4
        Case "NX_VIEW_PRESET_5": NxViewPresetApply w, 5
        Case "NX_VIEW_PRESET_SETTINGS": NxViewPresetSettingsOpen
        Case "RB_WINDOWS_VIEW_FULLSCREEN": Application.DisplayFullScreen = Not Application.DisplayFullScreen
        Case "RB_WINDOWS_VIEW_SCROLL_DOWN": w.SmallScroll Down:=10
        Case "RB_WINDOWS_VIEW_SCROLL_LEFT": w.SmallScroll ToLeft:=5
        Case "RB_WINDOWS_VIEW_SCROLL_RIGHT": w.SmallScroll ToRight:=5
        Case "RB_WINDOWS_VIEW_SCROLL_UP": w.SmallScroll Up:=10
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdInfo(ByVal key As String)
    Dim wb As Excel.Workbook
    Set wb = NxActiveUserWorkbook()
    Select Case key
        Case "RB_WORKBOOK_INFOMATION": NxProductShowAbout
        Case "RB_SHOW_WORKSHEET_LIST": NxCreateSheetIndex wb
        Case "RB_APP_NAMEUNHIDE": NxUnhideWorkbookNames wb
        Case "RB_INFO_MEMO_LIST": NxCreateMemoIndex wb
        Case "RB_INFO_SHEET_LIST": NxCreateSheetIndex wb
        Case Else: NxReject key
    End Select
End Sub

Public Sub NxCmdPivot(ByVal key As String)
    Dim r As Excel.Range
    Set r = NxSelectedRange()
    Select Case key
        Case "RB_PIVOTLISTOFWB": NxCreatePivotIndex NxActiveUserWorkbook()
        Case "RB_PIVOT_BUILDUP": NxBuildPivotFromSelection r
        Case Else: NxReject key
    End Select
End Sub

Private Function NxVisibleRange(ByVal source As Excel.Range) As Excel.Range
    On Error GoTo NoVisibleCells
    Set NxVisibleRange = source.SpecialCells(xlCellTypeVisible)
    Exit Function
NoVisibleCells:
    NxRaiseContractError "The selected range has no visible cells"
End Function

Private Sub SetVisibleCopyResult(ByVal result As CNxResult)
    If result Is Nothing Or result.Outcome <> NxSuccess Then NxRaiseContractError "Visible-cell copy did not complete"
End Sub

Private Sub SetVisiblePasteResult(ByVal result As CNxResult)
    If result Is Nothing Or result.Outcome <> NxSuccess Then NxRaiseContractError "Visible-cell paste did not complete"
End Sub

Private Sub NxPasteVisibleExcelClipboard(ByVal target As Excel.Range, ByVal pasteType As XlPasteType)
    Dim visibleTarget As Excel.Range
    Dim errorNumber As Long
    Dim errorDescription As String
    On Error GoTo Failed
    Set visibleTarget = NxVisibleRange(target)
    visibleTarget.PasteSpecial Paste:=pasteType
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    If errorNumber = 0 Then errorNumber = NX_CONTRACT_ERROR
    Err.Raise errorNumber, "NxPasteVisibleExcelClipboard", _
        "Excel 형식 클립보드를 보이는 셀에 붙여넣지 못했습니다: " & errorDescription
End Sub

Private Function NxFormulaClipboardText(ByVal target As Excel.Range) As String
    Dim formulaText As String
    formulaText = CStr(target.Cells(1, 1).Formula)
    If Left$(formulaText, 1) <> "=" Then NxRaiseContractError "The first selected cell does not contain a formula"
    NxFormulaClipboardText = formulaText
End Function

Private Function NxTextClipboardText(ByVal target As Excel.Range) As String
    Dim cell As Excel.Range
    Dim valueText As String
    For Each cell In target.Cells
        If Len(valueText) > 0 Then valueText = valueText & vbTab
        valueText = valueText & CStr(cell.Text)
    Next cell
    If Len(valueText) = 0 Then NxRaiseContractError "The selection has no text to copy"
    NxTextClipboardText = valueText
End Function

Private Function NxScalarColumnWidth(ByVal target As Excel.Range) As Double
    If IsNull(target.ColumnWidth) Or IsEmpty(target.ColumnWidth) Then
        NxScalarColumnWidth = CDbl(target.Cells(1, 1).ColumnWidth)
    Else
        NxScalarColumnWidth = CDbl(target.ColumnWidth)
    End If
End Function

Private Function NxScalarRowHeight(ByVal target As Excel.Range) As Double
    If IsNull(target.RowHeight) Or IsEmpty(target.RowHeight) Then
        NxScalarRowHeight = CDbl(target.Cells(1, 1).RowHeight)
    Else
        NxScalarRowHeight = CDbl(target.RowHeight)
    End If
End Function

Private Function NxScalarFontSize(ByVal target As Excel.Range) As Double
    If IsNull(target.Font.Size) Or IsEmpty(target.Font.Size) Then
        NxScalarFontSize = CDbl(target.Cells(1, 1).Font.Size)
    Else
        NxScalarFontSize = CDbl(target.Font.Size)
    End If
End Function

Private Sub NxConfigurePrint(ByVal target As Excel.Range, ByVal orientationValue As XlPageOrientation, ByVal fitOnePage As Boolean, ByVal preview As Boolean)
    With target.Worksheet.PageSetup
        .PrintArea = target.Address
        .Orientation = orientationValue
        .Zoom = False
        .FitToPagesWide = 1
        If fitOnePage Then .FitToPagesTall = 1 Else .FitToPagesTall = False
    End With
    If preview Then target.Worksheet.PrintPreview
End Sub

Private Sub NxResetCustomStyles(ByVal wb As Excel.Workbook)
    Dim styleItem As Excel.Style
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim previousAlerts As Boolean
    previousAlerts = Application.DisplayAlerts
    On Error GoTo CleanUp
    Application.DisplayAlerts = False
    For Each styleItem In wb.Styles
        If Not styleItem.BuiltIn Then styleItem.Delete
    Next styleItem
CleanUp:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Application.DisplayAlerts = previousAlerts
    If errorNumber <> 0 Then Err.Raise errorNumber, "NxResetCustomStyles", errorDescription
End Sub

Private Sub NxAddFormulaComments(ByVal target As Excel.Range)
    NxFormulaToolsOpen target, False
End Sub

Public Sub NxFormulaCommentsApply(ByVal target As Excel.Range, ByVal showNotes As Boolean)
    NxFormulaNotesApply target, showNotes
End Sub

Private Sub NxCycleEnterDirection()
    Select Case Application.MoveAfterReturnDirection
        Case xlDown: Application.MoveAfterReturnDirection = xlToRight
        Case xlToRight: Application.MoveAfterReturnDirection = xlUp
        Case xlUp: Application.MoveAfterReturnDirection = xlToLeft
        Case Else: Application.MoveAfterReturnDirection = xlDown
    End Select
End Sub

Private Sub NxMoveSelection(ByVal target As Excel.Range, ByVal rowDelta As Long, ByVal columnDelta As Long, ByVal endDirection As XlDirection)
    If target.Cells.CountLarge = 1 Then
        target.Cells(1, 1).End(endDirection).Select
    ElseIf NxCanOffsetRange(target, rowDelta, columnDelta) Then
        target.Offset(rowDelta, columnDelta).Select
    End If
End Sub

Private Function NxCanOffsetRange(ByVal target As Excel.Range, ByVal rowDelta As Long, ByVal columnDelta As Long) As Boolean
    Dim area As Excel.Range
    For Each area In target.Areas
        If area.Row + rowDelta < 1 Then Exit Function
        If area.Column + columnDelta < 1 Then Exit Function
        If area.Row + area.Rows.Count - 1 + rowDelta > area.Worksheet.Rows.Count Then Exit Function
        If area.Column + area.Columns.Count - 1 + columnDelta > area.Worksheet.Columns.Count Then Exit Function
    Next area
    NxCanOffsetRange = True
End Function

Private Sub NxSelectCellsContainingText(ByVal ws As Excel.Worksheet)
    Dim keyword As Variant
    Dim searchRange As Excel.Range
    Dim found As Excel.Range
    Dim firstAddress As String
    Dim result As Excel.Range
    keyword = Application.InputBox("Text to find", "LHexcel", Type:=2)
    If VarType(keyword) = vbBoolean Or Len(CStr(keyword)) = 0 Then Exit Sub
    Set searchRange = ws.UsedRange
    Set found = searchRange.Find(What:=CStr(keyword), LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
    If found Is Nothing Then Exit Sub
    firstAddress = found.Address
    Do
        If result Is Nothing Then Set result = found Else Set result = Union(result, found)
        Set found = searchRange.FindNext(found)
    Loop While Not found Is Nothing And found.Address <> firstAddress
    result.Select
End Sub

Private Sub NxResizeSelectionByDialog(ByVal target As Excel.Range)
    NxResizeOpen target
End Sub

Private Function NxAutoFilterRange(ByVal target As Excel.Range) As Excel.Range
    Dim filterRange As Excel.Range
    If target.Worksheet.AutoFilterMode Then Set filterRange = target.Worksheet.AutoFilter.Range
    If filterRange Is Nothing Then
        If target.Cells.CountLarge = 1 Then Set filterRange = target.CurrentRegion Else Set filterRange = target
    End If
    If filterRange.Rows.Count < 2 Then NxRaiseContractError "A header row and data row are required for filtering"
    Set NxAutoFilterRange = filterRange
End Function

Private Function NxActiveFilterField(ByVal filterRange As Excel.Range, ByVal selectedTarget As Excel.Range) As Long
    If selectedTarget Is Nothing Then NxRaiseContractError "A filter target is required"
    If Application.Intersect(filterRange, selectedTarget) Is Nothing Then NxRaiseContractError "The filter target is outside the filter range"
    NxActiveFilterField = selectedTarget.Cells(1, 1).Column - filterRange.Column + 1
    If NxActiveFilterField < 1 Or NxActiveFilterField > filterRange.Columns.Count Then NxActiveFilterField = 1
End Function

Private Sub NxSortCurrentRange(ByVal target As Excel.Range, ByVal sortOrder As XlSortOrder)
    Dim source As Excel.Range
    Set source = NxAutoFilterRange(target)
    source.Sort Key1:=source.Cells(1, NxActiveFilterField(source, target)), Order1:=sortOrder, Header:=xlYes
End Sub

Private Sub NxClearActiveColumnFilter(ByVal target As Excel.Range)
    Dim source As Excel.Range
    Set source = NxAutoFilterRange(target)
    source.AutoFilter Field:=NxActiveFilterField(source, target)
End Sub

Private Sub NxShowAllFilteredData(ByVal ws As Excel.Worksheet)
    If Not ws.FilterMode Then Exit Sub
    On Error GoTo Failed
    ws.ShowAllData
    Exit Sub
Failed:
    NxRaiseContractError "Excel could not clear the active filter"
End Sub

Private Sub NxFilterByActiveCell(ByVal target As Excel.Range, ByVal askValue As Boolean)
    Dim source As Excel.Range
    Dim criteriaText As Variant
    Dim fieldIndex As Long
    Set source = NxAutoFilterRange(target)
    fieldIndex = NxActiveFilterField(source, target)
    If askValue Then
        If IsError(target.Cells(1, 1).Value2) Then NxRaiseContractError "An error cell cannot be used as a filter criterion"
        criteriaText = Application.InputBox("Filter value", "LHexcel", CStr(target.Cells(1, 1).Value2), Type:=2)
        If VarType(criteriaText) = vbBoolean Then Exit Sub
    Else
        If IsError(target.Cells(1, 1).Value2) Then NxRaiseContractError "An error cell cannot be used as a filter criterion"
        criteriaText = target.Cells(1, 1).Value2
    End If
    If Len(CStr(criteriaText)) = 0 Then source.AutoFilter Field:=fieldIndex, Criteria1:="=" Else source.AutoFilter Field:=fieldIndex, Criteria1:=CStr(criteriaText)
End Sub

Private Sub NxShowPivotSourceOrFilter(ByVal target As Excel.Range)
    Dim pivotCell As Excel.PivotCell
    On Error GoTo Failed
    Set pivotCell = target.Cells(1, 1).PivotCell
    If pivotCell Is Nothing Then NxRaiseContractError "The selected cell is not a pivot cell"
    target.Cells(1, 1).ShowDetail = True
    Exit Sub
Failed:
    NxRaiseContractError "Pivot drill-down is available only for a detail-enabled pivot cell"
End Sub







Private Sub NxCreatePrecedentsIndex(ByVal target As Excel.Range)
    NxFormulaToolsOpen target, True
End Sub

Private Function NxDefaultOutputFolder(ByVal wb As Excel.Workbook) As String
    If Len(wb.Path) > 0 Then
        NxDefaultOutputFolder = wb.Path
    Else
        NxDefaultOutputFolder = CurDir$
    End If
End Function

Private Function NxSafeFileBase(ByVal fileName As String) As String
    Dim blocked As Variant
    Dim item As Variant
    Dim dotPosition As Long
    dotPosition = InStrRev(fileName, ".")
    If dotPosition > 1 Then NxSafeFileBase = Left$(fileName, dotPosition - 1) Else NxSafeFileBase = fileName
    blocked = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each item In blocked
        NxSafeFileBase = Replace(NxSafeFileBase, CStr(item), "_")
    Next item
End Function


Private Sub NxCreateSheetIndex(ByVal wb As Excel.Workbook)
    Dim output As Excel.Worksheet
    Dim ws As Excel.Worksheet
    Dim rowIndex As Long
    Set output = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    output.Name = NxUniqueSheetName(wb, "내엑셀_시트목록")
    output.Range("A1:C1").Value = Array("번호", "시트명", "표시 상태")
    rowIndex = 2
    For Each ws In wb.Worksheets
        If Not ws Is output Then
            output.Cells(rowIndex, 1).Value = rowIndex - 1
            output.Cells(rowIndex, 2).Value = ws.Name
            output.Cells(rowIndex, 3).Value = NxSheetVisibilityLabel(ws.Visible)
            rowIndex = rowIndex + 1
        End If
    Next ws
    output.Columns.AutoFit
End Sub

Private Function NxSheetVisibilityLabel(ByVal visibility As XlSheetVisibility) As String
    Select Case visibility
        Case xlSheetVisible: NxSheetVisibilityLabel = "표시"
        Case xlSheetHidden: NxSheetVisibilityLabel = "숨김"
        Case xlSheetVeryHidden: NxSheetVisibilityLabel = "완전 숨김"
        Case Else: NxRaiseContractError "알 수 없는 시트 표시 상태입니다."
    End Select
End Function

Private Sub NxUnhideWorkbookNames(ByVal wb As Excel.Workbook)
    Dim nameItem As Excel.Name
    For Each nameItem In wb.Names
        nameItem.Visible = True
    Next nameItem
End Sub

Private Sub NxCreateMemoIndex(ByVal wb As Excel.Workbook)
    Dim output As Excel.Worksheet
    Dim ws As Excel.Worksheet
    Dim cell As Excel.Range
    Dim noteText As String
    Dim rowIndex As Long
    Set output = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    output.Name = NxUniqueSheetName(wb, "LHExcel_MemoIndex")
    output.Range("A1:E1").Value = Array("No", "Sheet", "Cell", "Value", "Comment")
    rowIndex = 2
    For Each ws In wb.Worksheets
        For Each cell In ws.UsedRange.Cells
            noteText = vbNullString
            On Error Resume Next
            If Not cell.Comment Is Nothing Then noteText = cell.Comment.Text
            On Error GoTo 0
            If Len(noteText) > 0 Then
                output.Cells(rowIndex, 1).Value = rowIndex - 1
                output.Cells(rowIndex, 2).Value = ws.Name
                output.Cells(rowIndex, 3).Value = cell.Address(False, False)
                output.Cells(rowIndex, 4).Value = cell.Value2
                output.Cells(rowIndex, 5).Value = noteText
                rowIndex = rowIndex + 1
            End If
        Next cell
    Next ws
    output.Columns.AutoFit
End Sub

Private Sub NxCreatePivotIndex(ByVal wb As Excel.Workbook)
    Dim output As Excel.Worksheet
    Dim ws As Excel.Worksheet
    Dim pivotTable As Excel.PivotTable
    Dim rowIndex As Long
    Set output = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    output.Name = NxUniqueSheetName(wb, "LHExcel_PivotIndex")
    output.Range("A1:D1").Value = Array("No", "Sheet", "Pivot", "Source")
    rowIndex = 2
    For Each ws In wb.Worksheets
        For Each pivotTable In ws.PivotTables
            output.Cells(rowIndex, 1).Value = rowIndex - 1
            output.Cells(rowIndex, 2).Value = ws.Name
            output.Cells(rowIndex, 3).Value = pivotTable.Name
            output.Cells(rowIndex, 4).Value = pivotTable.SourceData
            rowIndex = rowIndex + 1
        Next pivotTable
    Next ws
    output.Columns.AutoFit
End Sub

Private Sub NxBuildPivotFromSelection(ByVal selectionRange As Excel.Range)
    Dim wb As Excel.Workbook
    Dim source As Excel.Range
    Dim output As Excel.Worksheet
    Dim cache As Excel.PivotCache
    Dim pivotTable As Excel.PivotTable
    Dim dataFieldIndex As Long
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim previousAlerts As Boolean
    On Error GoTo Failed
    Set wb = NxActiveUserWorkbook()
    If selectionRange.Cells.CountLarge = 1 Then Set source = selectionRange.CurrentRegion Else Set source = selectionRange
    If source.Rows.Count < 2 Or source.Columns.Count < 2 Then NxRaiseContractError "Select a header row and at least one data row to build a pivot table"
    NxValidatePivotHeaders source
    Set output = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    output.Name = NxUniqueSheetName(wb, "LHExcel_Pivot")
    Set cache = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=source.Address(True, True, xlR1C1, True))
    Set pivotTable = cache.CreatePivotTable(TableDestination:=output.Range("A3"), TableName:=NxUniquePivotName(wb, "LH_Pivot"))
    pivotTable.PivotFields(1).Orientation = xlRowField
    dataFieldIndex = NxFirstNumericColumn(source)
    If dataFieldIndex = 0 Then dataFieldIndex = 1
    pivotTable.AddDataField pivotTable.PivotFields(dataFieldIndex), "Total", IIf(dataFieldIndex = 1, xlCount, xlSum)
    On Error Resume Next
    pivotTable.TableStyle2 = "PivotStyleMedium9"
    On Error GoTo Failed
    output.Range("A1").Value = "LHexcel Pivot"
    output.Range("A1").Font.Bold = True
    output.Columns.AutoFit
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    If Not output Is Nothing Then
        previousAlerts = Application.DisplayAlerts
        On Error Resume Next
        Application.DisplayAlerts = False
        output.Delete
        Application.DisplayAlerts = previousAlerts
        On Error GoTo 0
    End If
    If errorNumber = 0 Then errorNumber = NX_CONTRACT_ERROR
    Err.Raise errorNumber, "NxBuildPivotFromSelection", errorDescription
End Sub

Private Sub NxValidatePivotHeaders(ByVal source As Excel.Range)
    Dim seen As Object
    Dim columnIndex As Long
    Dim headerText As String
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For columnIndex = 1 To source.Columns.Count
        headerText = Trim$(CStr(source.Cells(1, columnIndex).Value2))
        If Len(headerText) = 0 Then NxRaiseContractError "Pivot source headers cannot be blank"
        If seen.Exists(headerText) Then NxRaiseContractError "Pivot source headers must be unique"
        seen.Add headerText, True
    Next columnIndex
End Sub

Private Function NxFirstNumericColumn(ByVal source As Excel.Range) As Long
    Dim columnIndex As Long
    Dim rowIndex As Long
    For columnIndex = 1 To source.Columns.Count
        For rowIndex = 2 To source.Rows.Count
            If IsNumeric(source.Cells(rowIndex, columnIndex).Value2) And Len(CStr(source.Cells(rowIndex, columnIndex).Value2)) > 0 Then
                NxFirstNumericColumn = columnIndex
                Exit Function
            End If
        Next rowIndex
    Next columnIndex
End Function

Private Function NxUniquePivotName(ByVal wb As Excel.Workbook, ByVal baseName As String) As String
    Dim ws As Excel.Worksheet
    Dim pivotTable As Excel.PivotTable
    Dim candidate As String
    Dim indexValue As Long
    Dim exists As Boolean
    Do
        indexValue = indexValue + 1
        candidate = baseName & "_" & Format$(indexValue, "000")
        exists = False
        For Each ws In wb.Worksheets
            For Each pivotTable In ws.PivotTables
                If StrComp(pivotTable.Name, candidate, vbTextCompare) = 0 Then exists = True
            Next pivotTable
        Next ws
    Loop While exists
    NxUniquePivotName = candidate
End Function

Private Function NxUniqueSheetName(ByVal wb As Excel.Workbook, ByVal baseName As String) As String
    Dim candidate As String
    Dim indexValue As Long
    candidate = baseName
    Do While NxSheetExists(wb, candidate)
        indexValue = indexValue + 1
        candidate = baseName & "_" & CStr(indexValue)
    Loop
    NxUniqueSheetName = candidate
End Function

Private Function NxSheetExists(ByVal wb As Excel.Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Excel.Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            NxSheetExists = True
            Exit Function
        End If
    Next ws
End Function
