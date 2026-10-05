Attribute VB_Name = "T_R60DocumentNavigator"
Option Explicit

Private mVisualFirst As Workbook
Private mVisualSecond As Workbook

' Imported only into a disposable, never-saved copy of the product XLAM.
' Control-event wrappers exercise native VBA handlers, not physical UI input.
Public Function NxR60DocumentNavigatorCase(ByVal kind As String, ByVal outputRoot As String) As String
    Dim firstBook As Workbook, secondBook As Workbook, hiddenBook As Workbook
    Dim originalWindow As Window, extraWindow As Window
    Dim items As Collection, refreshed As Collection
    Dim item As CNxDocumentNavigatorItem, bookItem As CNxDocumentNavigatorItem
    Dim oldBookItem As CNxDocumentNavigatorItem, oldSheetItem As CNxDocumentNavigatorItem
    Dim panel As FNxDocumentNavigator, sheet As Object
    Dim eventProbe As CNxR60NavigatorEvents
    Dim before As String, dataBefore As String, reason As String, path As String
    Dim bookToken As Long, sheetToken As Long, index As Long, moved As Boolean
    Dim initialEvents As Boolean, initialAlerts As Boolean
    initialEvents = Application.EnableEvents
    initialAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    TraceStep outputRoot, kind, "start"
    Require Not NxFormWheelIsAttached(), "wheel must start detached"

    If kind = "no_documents" Then
        Require ThisWorkbook.IsAddin, "probe host must be an addin"
        Require UserDocumentCount() = 0, "real zero-document fixture: no user document may be open"
        Require Application.ActiveWorkbook Is Nothing, "zero documents must have no active workbook"
        Set items = NxDocumentNavigatorSnapshot(Nothing)
        Require items.Count = 0, "empty document snapshot"
        NxDocumentNavigatorOpen
        Set panel = NxR60ProbeNavigator()
        Require panel.Visible, "empty-state form is visible"
        Require panel.lstBooks.ListCount = 0 And panel.lstSheets.ListCount = 0, "empty lists"
        Require Not panel.cmdMove.Enabled, "empty move disabled"
        panel.RefreshNavigator
        Require UserDocumentCount() = 0, "empty refresh must not create workbook"
        Require Application.ActiveWorkbook Is Nothing, "empty form must not create active document"
        panel.NxR60TestClose
        Set panel = Nothing
        Require Not NxFormWheelIsAttached(), "empty close detaches wheel"
        GoTo Passed
    End If

    Set firstBook = CreateFixture(outputRoot & Application.PathSeparator & kind & "_First.xlsx", True)
    Set secondBook = CreateFixture(outputRoot & Application.PathSeparator & kind & "_Second.xlsx", False)
    firstBook.Activate
    firstBook.Worksheets("Start").Activate
    firstBook.Worksheets("Start").Range("D5").Select
    Set items = NxDocumentNavigatorSnapshot(Nothing)
    Set bookItem = FindItem(items, firstBook, Nothing)
    Require Not bookItem Is Nothing, "first workbook represented"
    Set originalWindow = Application.ActiveWindow
    Set extraWindow = firstBook.Windows(1)
    TraceStep outputRoot, kind, "window_is=" & CStr(originalWindow Is extraWindow) & ";active_hwnd=" & CStr(originalWindow.Hwnd) & ";enumerated_hwnd=" & CStr(extraWindow.Hwnd)
    Set extraWindow = Nothing

    Select Case kind
        Case "snapshot"
            Set hiddenBook = CreateFixture(outputRoot & Application.PathSeparator & "snapshot_HiddenBook.xlsx", False)
            hiddenBook.Windows(1).Visible = False
            firstBook.Activate
            ' Both clean and dirty documents must retain their original Saved flags.
            secondBook.Worksheets("Start").Range("B3").Value2 = "unsaved fixture edit"
            before = CaptureState(firstBook, secondBook)
            For index = 1 To 20
                Set refreshed = NxDocumentNavigatorSnapshot(items)
                Set items = refreshed
            Next index
            Require CountBooks(items) = 2, "only two visible user workbooks listed"
            Require CountBookItems(items, firstBook) = 6, "all first workbook sheets including chart represented"
            Require FindItem(items, ThisWorkbook, Nothing) Is Nothing, "addin excluded"
            Require FindItem(items, hiddenBook, Nothing) Is Nothing, "window-hidden workbook excluded"
            Require before = CaptureState(firstBook, secondBook), "repeated snapshots changed workbook state"
            Require hiddenBook.Windows(1).Visible = False, "snapshot must not reveal hidden workbook"
            RequireUniqueTokens items

        Case "form_filter_click"
            before = CaptureState(firstBook, secondBook)
            NxDocumentNavigatorOpen
            Set panel = NxR60ProbeNavigator()
            Require panel.Visible, "navigator shown modeless"
            Require panel.Caption = "내엑셀 - 문서·시트 탐색", "navigator caption"
            Require panel.lstBooks.ListCount = 2, "initial open populates documents before any manual refresh"
            panel.txtBooks.Value = "no_match_r60_zzzz"
            Require panel.lstBooks.ListCount = 0 And panel.lstSheets.ListCount = 0, "book no-match clears both lists"
            Require Not panel.cmdMove.Enabled, "no-match move disabled"
            panel.txtBooks.Value = secondBook.Name
            Require panel.lstBooks.ListCount = 1, "book filter selects one matching row"
            panel.NxR60TestBookClick 0
            panel.txtSheets.Value = "no_match_r60_zzzz"
            Require panel.lstSheets.ListCount = 0, "sheet no-match"
            Require Not panel.cmdMove.Enabled, "sheet no-match move disabled"
            panel.txtSheets.Value = "Target"
            Require panel.lstSheets.ListCount = 1, "sheet filter selects matching row"
            panel.NxR60TestSheetClick 0
            panel.RefreshNavigator
            Require before = CaptureState(firstBook, secondBook), "open/search/list selection/refresh changed document state"
            ' Re-select after refresh; this is an explicit move-button handler call.
            panel.txtBooks.Value = secondBook.Name
            panel.NxR60TestBookClick FindRow(panel.lstBooks, secondBook.Name)
            panel.txtSheets.Value = "Target"
            panel.NxR60TestSheetClick FindRow(panel.lstSheets, "Target")
            Require panel.cmdMove.Enabled, "explicit selected target is movable"
            dataBefore = BookDataState(firstBook) & BookDataState(secondBook)
            TraceStep outputRoot, kind, "selected=" & panel.NxR60TestSelectionInfo()
            panel.NxR60TestMove
            Require Application.ActiveWorkbook Is secondBook, "Move button chooses requested book"
            Require Application.ActiveSheet Is secondBook.Worksheets("Target"), "Move button chooses requested sheet; actual=" & Application.ActiveSheet.Name & "; status=" & panel.lblStatus.Caption
            Require dataBefore = BookDataState(firstBook) & BookDataState(secondBook), "Move button changed data or formats"

        Case "hidden_protected_chart"
            firstBook.Protect Password:="r60-fixture", Structure:=True
            Set items = NxDocumentNavigatorSnapshot(items)
            dataBefore = BookDataState(firstBook)
            For Each sheet In firstBook.Sheets
                Set item = FindItem(items, firstBook, sheet)
                Require Not item Is Nothing, "every sheet including hidden/chart has an item"
                If sheet.Visible <> xlSheetVisible Then
                    Require Not item.Available And Not NxDocumentNavigatorCanMove(item), "hidden sheet unavailable"
                    Require InStr(item.StateText, "숨김") > 0, "hidden state is disclosed"
                    before = CaptureState(firstBook, secondBook)
                    moved = NxDocumentNavigatorMove(item, reason)
                    Require Not moved And Len(reason) > 0, "hidden movement must give failure reason"
                    Require before = CaptureState(firstBook, secondBook), "hidden rejection changed state"
                End If
            Next sheet
            Set item = FindItem(items, firstBook, firstBook.Worksheets("Target"))
            Require item.Available And NxDocumentNavigatorCanMove(item), "protected visible sheet is movable"
            moved = NxDocumentNavigatorMove(item, reason)
            Require moved And Len(reason) = 0, "protected sheet move succeeds"
            Require Application.ActiveSheet Is firstBook.Worksheets("Target"), "protected sheet identity"
            Require firstBook.Worksheets("Target").ProtectContents, "sheet protection retained"
            Require firstBook.ProtectStructure, "workbook structure protection retained"
            Set item = FindItem(items, firstBook, firstBook.Charts("ChartOnly"))
            Require item.Available And InStr(item.StateText, "차트") > 0, "chart represented as chart"
            moved = NxDocumentNavigatorMove(item, reason)
            Require moved And Application.ActiveSheet Is firstBook.Charts("ChartOnly"), "chart navigation identity"
            Require dataBefore = BookDataState(firstBook), "navigation changed protected data or formats"

        Case "rename_closed_identity"
            Set oldBookItem = bookItem
            Set oldSheetItem = FindItem(items, firstBook, firstBook.Worksheets("Start"))
            bookToken = oldBookItem.TokenId
            sheetToken = oldSheetItem.TokenId
            firstBook.Worksheets("Start").Name = "RenamedStart"
            path = outputRoot & Application.PathSeparator & "rename_RenamedBook.xlsx"
            firstBook.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
            Set refreshed = NxDocumentNavigatorSnapshot(items)
            Set item = FindItem(refreshed, firstBook, Nothing)
            Require item Is oldBookItem, "workbook object identity reused after rename"
            Require item.TokenId = bookToken And item.DisplayName = firstBook.Name, "renamed workbook token and caption"
            Set item = FindItem(refreshed, firstBook, firstBook.Worksheets("RenamedStart"))
            Require item Is oldSheetItem, "sheet object identity reused after rename"
            Require item.TokenId = sheetToken And item.DisplayName = "RenamedStart", "renamed sheet token and caption"
            secondBook.Activate
            firstBook.Close SaveChanges:=False
            Set firstBook = Nothing
            Require Not NxDocumentNavigatorCanMove(oldBookItem), "closed workbook reference rejected"
            Require Not NxDocumentNavigatorCanMove(oldSheetItem), "closed sheet reference rejected"
            Set firstBook = Application.Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=False)
            secondBook.Activate
            before = CaptureState(firstBook, secondBook)
            moved = NxDocumentNavigatorMove(oldBookItem, reason)
            Require Not moved And Len(reason) > 0, "same path reopen must not retarget stale workbook"
            moved = NxDocumentNavigatorMove(oldSheetItem, reason)
            Require Not moved And Len(reason) > 0, "same name reopen must not retarget stale sheet"
            Require before = CaptureState(firstBook, secondBook), "stale reference rejection changed state"
            Set items = NxDocumentNavigatorSnapshot(refreshed)
            Set item = FindItem(items, firstBook, Nothing)
            Require item.TokenId <> bookToken, "reopened workbook receives distinct session token"
            Set item = FindItem(items, firstBook, firstBook.Worksheets("RenamedStart"))
            Require item.TokenId <> sheetToken, "reopened sheet receives distinct session token"
            Require NxDocumentNavigatorMove(item, reason), "fresh reopened sheet can move"
            Require Application.ActiveSheet Is firstBook.Worksheets("RenamedStart"), "fresh reopened identity"

        Case "multiwindow"
            Set originalWindow = firstBook.Windows(1)
            originalWindow.Activate
            originalWindow.Zoom = 75
            Set extraWindow = firstBook.NewWindow
            extraWindow.Activate
            firstBook.Worksheets("Target").Activate
            extraWindow.Zoom = 125
            firstBook.Worksheets("Target").Range("G11").Select
            Set items = NxDocumentNavigatorSnapshot(items)
            Set bookItem = FindItem(items, firstBook, Nothing)
            Require bookItem.RecentWindow.Hwnd = extraWindow.Hwnd, "snapshot remembers active native window within workbook"
            secondBook.Activate
            before = CaptureState(firstBook, secondBook)
            Set refreshed = NxDocumentNavigatorSnapshot(items)
            Require before = CaptureState(firstBook, secondBook), "multiwindow snapshot changed state"
            Set item = FindItem(refreshed, firstBook, firstBook.Worksheets("Target"))
            Require NxDocumentNavigatorMove(item, reason), "multiwindow move succeeds"
            Require Application.ActiveWorkbook Is firstBook, "recent window workbook identity"
            Require Application.ActiveWindow.Hwnd = extraWindow.Hwnd, "recent native window preferred"
            Require Application.ActiveSheet Is firstBook.Worksheets("Target"), "recent window exact target"
            Require Application.ActiveCell.Address = "$G$11" And extraWindow.Zoom = 125, "recent window view retained"
            secondBook.Activate
            extraWindow.Close
            Set extraWindow = Nothing
            Require NxDocumentNavigatorMove(item, reason), "closed recent window falls back to live window"
            Require Application.ActiveWorkbook Is firstBook, "fallback workbook identity"
            Require Application.ActiveWindow.Hwnd = originalWindow.Hwnd, "fallback chooses live original native window"
            Require Application.ActiveSheet Is firstBook.Worksheets("Target"), "fallback exact target"

        Case "event_redirect", "event_reenter", "event_close"
            Require Application.EnableEvents, "real Application events must be enabled"
            secondBook.Activate
            Set item = FindItem(items, firstBook, firstBook.Worksheets("Target"))
            If kind = "event_close" Then
                NxDocumentNavigatorOpen
                Set panel = NxR60ProbeNavigator()
                TraceStep outputRoot, kind, panel.NxR60TestLifecycleInfo()
                panel.NxR60TestBookClick FindRow(panel.lstBooks, firstBook.Name)
                panel.NxR60TestSheetClick FindRow(panel.lstSheets, "Target")
            End If
            Set eventProbe = New CNxR60NavigatorEvents
            eventProbe.Start kind, firstBook, secondBook.Windows(1), item
            dataBefore = BookDataState(firstBook) & BookDataState(secondBook)
            If kind = "event_close" Then
                panel.NxR60TestMove
            Else
                moved = NxDocumentNavigatorMove(item, reason)
            End If
            Require eventProbe.Hits > 0, "native WindowActivate handler must execute"
            Require Len(eventProbe.Failure) = 0, "event handler error: " & eventProbe.Failure
            Select Case kind
                Case "event_redirect"
                    Require Not moved And Len(reason) > 0, "redirect must reject movement with reason"
                    Require Application.ActiveWorkbook Is secondBook, "redirected window must not be stolen back"
                    Require firstBook.Windows(1).ActiveSheet Is firstBook.Worksheets("Start"), "redirect must stop before target sheet activation"
                Case "event_reenter"
                    Require moved, "outer move completes after blocked reentry"
                    Require Not eventProbe.ReentrantSucceeded And Len(eventProbe.ReentrantReason) > 0, "nested move must be blocked explicitly"
                    Require Application.ActiveWorkbook Is firstBook, "outer book identity retained"
                    Require Application.ActiveSheet Is firstBook.Worksheets("Target"), "outer sheet identity retained"
                Case "event_close"
                    Set panel = Nothing
                    Require NavigatorFormCount() = 0, "event shutdown leaves no navigator form"
                    Require Not NxFormWheelIsAttached(), "event shutdown detaches wheel"
                    Require Application.ActiveWorkbook Is firstBook, "event shutdown does not prevent requested workbook move"
                    Require Application.ActiveSheet Is firstBook.Worksheets("Target"), "event shutdown does not prevent requested sheet move"
            End Select
            eventProbe.StopListening
            Set eventProbe = Nothing
            Require dataBefore = BookDataState(firstBook) & BookDataState(secondBook), "event-edge move changed data or formats"
            ' A completed or failed move must release its reentrancy guard for later use.
            Set item = FindItem(items, secondBook, secondBook.Worksheets("Target"))
            Require NxDocumentNavigatorMove(item, reason), "move guard released after event edge"
            Require Application.ActiveSheet Is secondBook.Worksheets("Target"), "subsequent move exact target"

        Case "wheel_owner"
            before = CaptureState(firstBook, secondBook)
            NxDocumentNavigatorOpen
            Set panel = NxR60ProbeNavigator()
            NxSymbolsOpen
            Dim otherPanel As Object, loadedPanel As Object
            For Each loadedPanel In VBA.UserForms
                If TypeName(loadedPanel) = "FNxSymbols" Then Set otherPanel = loadedPanel: Exit For
            Next loadedPanel
            Set loadedPanel = Nothing
            Require Not otherPanel Is Nothing, "symbols coexistence form exists"
            Require NxFormWheelAttach(otherPanel), "symbols takes shared wheel ownership"
            NxDocumentNavigatorShutdown
            Require NxFormWheelIsAttached(), "inactive navigator close preserves symbols hook"
            NxDocumentNavigatorOpen
            Set panel = NxR60ProbeNavigator()
            Require NxFormWheelAttach(panel), "navigator takes shared wheel ownership"
            NxSymbolsClose
            Set otherPanel = Nothing
            Require NxFormWheelIsAttached(), "inactive symbols close preserves navigator hook"
            NxDocumentNavigatorShutdown
            Require Not NxFormWheelIsAttached(), "both forms closed leaves no wheel hook"
            Require before = CaptureState(firstBook, secondBook), "form coexistence changed document state"

        Case "lifecycle"
            before = CaptureState(firstBook, secondBook)
            For index = 1 To 6
                NxDocumentNavigatorOpen
                Set panel = NxR60ProbeNavigator()
                TraceStep outputRoot, kind, panel.NxR60TestLifecycleInfo()
                Require panel.Visible, "repeat open visible"
                Require NxFormWheelIsAttached(), "realized form wheel hook attached"
                panel.RefreshNavigator
                If index Mod 2 = 0 Then
                    NxDocumentNavigatorShutdown
                Else
                    panel.NxR60TestClose
                End If
                Set panel = Nothing
                Require Not NxFormWheelIsAttached(), "repeat close leaves no wheel hook"
                Require NavigatorFormCount() = 0, "repeat close leaves no loaded navigator"
            Next index
            NxDocumentNavigatorShutdown
            Require Not NxFormWheelIsAttached(), "shutdown is idempotent"
            Require before = CaptureState(firstBook, secondBook), "repeat open/close changed document state"

        Case Else
            Err.Raise vbObjectError + 6001, "r60 navigator probe", "Unknown case: " & kind
    End Select
Passed:
    TraceStep outputRoot, kind, "assertions_pass"
    NxR60DocumentNavigatorCase = "PASS|" & kind & "|native VBA on disposable fixtures; physical input not asserted"
    GoTo CleanUp
Failed:
    NxR60DocumentNavigatorCase = "FAIL|" & kind & "|" & CStr(Err.Number) & "|" & Err.Description
CleanUp:
    On Error Resume Next
    If Not eventProbe Is Nothing Then eventProbe.StopListening
    Set eventProbe = Nothing
    Err.Clear
    NxDocumentNavigatorShutdown
    TraceStep outputRoot, kind, "shutdown_error=" & CStr(Err.Number) & ";" & Err.Description
    NxSymbolsClose
    Set otherPanel = Nothing
    Set panel = Nothing
    TraceStep outputRoot, kind, "loaded_forms_after_shutdown=" & CStr(NavigatorFormCount())
    Set item = Nothing: Set bookItem = Nothing
    Set oldBookItem = Nothing: Set oldSheetItem = Nothing
    Set items = Nothing: Set refreshed = Nothing
    Set sheet = Nothing: Set extraWindow = Nothing: Set originalWindow = Nothing
    If Not hiddenBook Is Nothing Then hiddenBook.Close SaveChanges:=False
    If Not secondBook Is Nothing Then secondBook.Close SaveChanges:=False
    If Not firstBook Is Nothing Then firstBook.Close SaveChanges:=False
    Application.EnableEvents = initialEvents
    Application.DisplayAlerts = initialAlerts
End Function

Public Sub NxR60DocumentNavigatorDispose()
    On Error Resume Next
    NxDocumentNavigatorShutdown
    NxFocusDisable
    If Not mVisualSecond Is Nothing Then mVisualSecond.Close SaveChanges:=False
    If Not mVisualFirst Is Nothing Then mVisualFirst.Close SaveChanges:=False
    Set mVisualSecond = Nothing
    Set mVisualFirst = Nothing
End Sub

Public Function NxR60DocumentNavigatorVisual(ByVal outputRoot As String, _
    Optional ByVal visualMode As String = "off") As String
    Const longSheetName As String = "문서시트탐색기긴이름검색확인용가나다라마바사아자차카타파하끝값"
    Dim panel As FNxDocumentNavigator, sheet As Worksheet, index As Long
    Dim userRuleCount As Long, focusResult As String, backend As String
    On Error GoTo Failed
    visualMode = LCase$(Trim$(visualMode))
    Require visualMode = "off" Or visualMode = "cf", "visual mode must be off or cf"
    Require Len(longSheetName) = 31, "long search fixture must use 31 characters"
    Set mVisualFirst = CreateFixture(outputRoot & Application.PathSeparator & "visual_First.xlsx", True)
    For index = 1 To 20
        Set sheet = mVisualFirst.Worksheets.Add(After:=mVisualFirst.Sheets(mVisualFirst.Sheets.Count))
        If index = 20 Then
            sheet.Name = longSheetName
        Else
            sheet.Name = "Preview_" & Format$(index, "00")
        End If
    Next index
    mVisualFirst.Worksheets("Start").Activate
    mVisualFirst.Worksheets("Start").Range("D5").Select
    mVisualFirst.Save
    Set mVisualSecond = CreateFixture(outputRoot & Application.PathSeparator & "visual_Second.xlsx", False)
    Set sheet = mVisualSecond.Worksheets("Start")
    userRuleCount = sheet.Range("C3").FormatConditions.Count
    Require userRuleCount > 0, "visual fixture must retain its user conditional formatting"
    ' Prepare only: all edits, Copy/Paste and Undo after READY require physical UI input.
    ' Retain the verified fixture note and C3 rule; never delete user formatting rules.
    sheet.Range("B2").Value2 = 202
    sheet.Range("C2").Formula = "=B2+1"
    sheet.Range("D2").Value2 = 204
    sheet.Range("B3").Value2 = 302
    sheet.Range("C3").Formula = "=12+30"
    sheet.Range("D3").Formula = "=B3+C3"
    sheet.Range("B4").Value2 = 402
    sheet.Range("C4").Value2 = "r60 text"
    sheet.Range("D4").Formula = "=$B$2+B4"
    sheet.Range("B2:D4").NumberFormat = "0.00"
    sheet.Range("B2:D4").Font.Bold = True
    sheet.Range("B2:D2").Interior.Color = RGB(210, 230, 250)
    sheet.Range("B3:D3").Interior.Color = RGB(220, 240, 220)
    sheet.Range("B4:D4").Interior.Color = RGB(250, 230, 210)
    sheet.Range("C4").NumberFormat = "@"
    sheet.Range("D2").Font.Italic = True
    sheet.Range("K12:M14").ClearContents
    sheet.Range("H2").Value2 = "before"
    Require sheet.Range("C3").FormatConditions.Count = userRuleCount, "copy setup changed user conditional formatting"
    sheet.Activate
    sheet.Range("D5").Select
    mVisualSecond.Save
    If visualMode = "cf" Then
        focusResult = NxFocusTryCommitSettingsValues("criss-cross", "wide-stripe", "#5FC8D8", 40, False, True)
        Require focusResult = "PASS", "visual focus setup failed: " & focusResult
    Else
        NxFocusDisable
    End If
    backend = NxFocusControllerBackend()
    Require backend = visualMode, "visual focus backend mismatch: " & backend
    TraceStep outputRoot, "visual_hold", "prepared_focus=" & backend & ";source=B2:D4;destination=K12:M14;undo=H2:before"
    NxDocumentNavigatorOpen
    Set panel = NxR60ProbeNavigator()
    panel.NxR60TestBookClick FindRow(panel.lstBooks, mVisualFirst.Name)
    panel.lstSheets.SetFocus
    Require panel.lstBooks.ListCount = 2 And panel.lstSheets.ListCount >= 20, "visual fixtures ready"
    Require Application.ActiveWorkbook Is mVisualSecond, "visual list selection must not navigate"
    NxR60DocumentNavigatorVisual = "READY|two fixture books; first book sheet list; active workbook visual_Second.xlsx Start D5" & _
        ";mode=" & visualMode & ";backend=" & backend & ";source=visual_Second.xlsx!Start!B2:D4" & _
        ";destination=visual_Second.xlsx!Start!K12:M14(blank);undo=visual_Second.xlsx!Start!H2(before)" & _
        ";long_sheet=visual_First.xlsx!" & longSheetName & ";physical_input_assertions=NOT_RUN"
    Exit Function
Failed:
    NxR60DocumentNavigatorVisual = "FAIL|visual_hold|" & Err.Description
    NxR60DocumentNavigatorDispose
End Function

Private Function CreateFixture(ByVal path As String, ByVal rich As Boolean) As Workbook
    Dim book As Workbook
    Dim failureNumber As Long, failureDescription As String, phase As String, fixtureRoot As String, templateName As String
    On Error GoTo Failed
    fixtureRoot = Left$(path, InStrRev(path, Application.PathSeparator) - 1)
    If rich Then templateName = "rich.xlsx" Else templateName = "simple.xlsx"
    phase = "open_verified_fixture"
    TraceStep fixtureRoot, "fixture", phase
    ' Reuse the native candidate07 fixtures, including their notes and formatting.
    ' Creating a new note crashed this Excel build in ntdll before any navigator call.
    Set book = Application.Workbooks.Open(fixtureRoot & Application.PathSeparator & "fixtures" & Application.PathSeparator & templateName, UpdateLinks:=0, ReadOnly:=False)
    Require Not book.HasVBProject, "fixture must be macro-free"
    Require book.Worksheets("Start").Range("B2").Comment.Text = "r60 keep comment", "fixture note retained"
    Require book.Worksheets("Start").Range("C3").Formula = "=12+30", "fixture formula retained"
    phase = "save_fixture"
    TraceStep fixtureRoot, "fixture", phase
    book.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
    TraceStep fixtureRoot, "fixture", "saved"
    Set CreateFixture = book
    Exit Function
Failed:
    failureNumber = Err.Number: failureDescription = phase & ": " & Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise failureNumber, "r60 fixture setup", failureDescription
End Function

Private Function FindItem(ByVal items As Collection, ByVal book As Workbook, ByVal sheet As Object) As CNxDocumentNavigatorItem
    Dim candidate As CNxDocumentNavigatorItem
    For Each candidate In items
        If candidate.BookRef Is book Then
            If candidate.SheetRef Is sheet Then Set FindItem = candidate: Exit Function
        End If
    Next candidate
End Function

Private Function CountBooks(ByVal items As Collection) As Long
    Dim item As CNxDocumentNavigatorItem
    For Each item In items
        If item.SheetRef Is Nothing Then CountBooks = CountBooks + 1
    Next item
End Function

Private Function UserDocumentCount() As Long
    Dim book As Workbook
    For Each book In Application.Workbooks
        If Not book.IsAddin Then UserDocumentCount = UserDocumentCount + 1
    Next book
End Function

Private Function CountBookItems(ByVal items As Collection, ByVal book As Workbook) As Long
    Dim item As CNxDocumentNavigatorItem
    For Each item In items
        If item.BookRef Is book Then CountBookItems = CountBookItems + 1
    Next item
End Function

Private Sub RequireUniqueTokens(ByVal items As Collection)
    Dim first As Long, second As Long, a As CNxDocumentNavigatorItem, b As CNxDocumentNavigatorItem
    For first = 1 To items.Count
        Set a = items(first)
        Require a.TokenId > 0, "positive session token"
        For second = first + 1 To items.Count
            Set b = items(second)
            Require a.TokenId <> b.TokenId, "unique session token"
        Next second
        If Not a.SheetRef Is Nothing Then
            Require Not a.ParentBook Is Nothing, "sheet has parent book item"
            Require a.ParentBook.BookRef Is a.BookRef, "parent workbook object matches"
        End If
    Next first
End Sub

Private Function FindRow(ByVal list As Object, ByVal label As String) As Long
    Dim row As Long, column As Long
    For row = 0 To list.ListCount - 1
        For column = 0 To list.ColumnCount - 1
            If InStr(1, CStr(list.List(row, column)), label, vbTextCompare) > 0 Then FindRow = row: Exit Function
        Next column
    Next row
    Err.Raise vbObjectError + 6002, "r60 navigator probe", "List row not found: " & label
End Function

Private Function NavigatorFormCount() As Long
    Dim form As Object
    For Each form In VBA.UserForms
        If TypeName(form) = "FNxDocumentNavigator" Then NavigatorFormCount = NavigatorFormCount + 1
    Next form
End Function

Private Function CaptureState(ByVal first As Workbook, ByVal second As Workbook) As String
    Dim state As String, book As Workbook, sheet As Object, window As Window, selected As Object
    Set book = Application.ActiveWorkbook
    Set sheet = Application.ActiveSheet
    Set window = Application.ActiveWindow
    If Not book Is Nothing Then state = state & "|activebook=" & book.FullName
    If Not sheet Is Nothing Then state = state & "|activesheet=" & sheet.Name & ":" & TypeName(sheet)
    If Not window Is Nothing Then state = state & "|activewindow=" & window.Caption
    Set selected = Application.Selection
    state = state & "|selectiontype=" & TypeName(selected)
    If TypeName(selected) = "Range" Then state = state & "|selection=" & selected.Address(External:=True) & "|activecell=" & Application.ActiveCell.Address
    state = state & "|events=" & CStr(Application.EnableEvents) & "|alerts=" & CStr(Application.DisplayAlerts)
    state = state & "|calculation=" & CStr(Application.Calculation) & "|screen=" & CStr(Application.ScreenUpdating)
    state = state & "|cutcopymode=" & CStr(Application.CutCopyMode)
    CaptureState = state & BookDataState(first) & WindowState(first) & BookDataState(second) & WindowState(second)
End Function

Private Function BookDataState(ByVal book As Workbook) As String
    Dim state As String, sheet As Object, cell As Range, definedName As Excel.Name
    state = "|book=" & book.FullName & "|saved=" & CStr(book.Saved) & "|readonly=" & CStr(book.ReadOnly)
    state = state & "|sheets=" & CStr(book.Sheets.Count) & "|structure=" & CStr(book.ProtectStructure)
    For Each definedName In book.Names
        state = state & "|name=" & definedName.Name & ":" & definedName.RefersTo
    Next definedName
    For Each sheet In book.Sheets
        state = state & "|sheet=" & sheet.Name & ":" & TypeName(sheet) & ":" & CStr(sheet.Visible)
        If TypeOf sheet Is Worksheet Then
            state = state & "|protected=" & CStr(sheet.ProtectContents) & "|shapes=" & CStr(sheet.Shapes.Count)
            state = state & "|used=" & sheet.UsedRange.Address & "|cf=" & CStr(sheet.Cells.FormatConditions.Count)
            For Each cell In sheet.Range("B2,C3,E8,B3")
                state = state & "|cell=" & cell.Address & ":" & CStr(cell.Formula) & ":" & CStr(cell.NumberFormat)
                state = state & ":" & CStr(cell.Interior.Color) & ":" & CStr(cell.Font.Bold)
                state = state & ":" & CStr(cell.ColumnWidth) & ":" & CStr(cell.RowHeight)
                If Not cell.Comment Is Nothing Then state = state & ":comment=" & cell.Comment.Text
            Next cell
        End If
    Next sheet
    BookDataState = state
End Function

Private Function WindowState(ByVal book As Workbook) As String
    Dim window As Window, sheet As Object, state As String
    For Each window In book.Windows
        state = state & "|window=" & window.Caption & ":" & CStr(window.Visible) & ":" & CStr(window.Zoom)
        state = state & ":" & CStr(window.ScrollRow) & ":" & CStr(window.ScrollColumn)
        state = state & ":" & CStr(window.SplitRow) & ":" & CStr(window.SplitColumn) & ":" & CStr(window.FreezePanes)
        state = state & "|window_active_sheet=" & window.ActiveSheet.Name
        For Each sheet In window.SelectedSheets
            state = state & "|selected_sheet=" & sheet.Name
        Next sheet
    Next window
    WindowState = state
End Function

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 6003, "r60 navigator probe", message
End Sub

Private Sub TraceStep(ByVal outputRoot As String, ByVal kind As String, ByVal stage As String)
    Dim handle As Integer
    handle = FreeFile
    Open outputRoot & Application.PathSeparator & "steps.log" For Append As #handle
    Print #handle, kind & "|" & stage
    Close #handle
End Sub
