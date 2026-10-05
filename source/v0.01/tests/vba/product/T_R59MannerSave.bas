Attribute VB_Name = "T_R59MannerSave"
Option Explicit

Private Function NxDrawPopupProbe() As String
    Dim fixture As Workbook, panel As FNxDrawTable, context As CNxFeatureDialogContext
    Dim control As Object, count As Long, detail As String
    On Error GoTo Failed
    Require ThisWorkbook.CodeName = "NxProductWorkbook", "workbook module name: " & ThisWorkbook.CodeName
    Require ThisWorkbook.Worksheets(1).CodeName = "NxProductSheet001", "sheet module name: " & ThisWorkbook.Worksheets(1).CodeName
    Set fixture = Workbooks.Add(xlWBATWorksheet)
    With fixture.Worksheets(1)
        .Range("A1:F6").Value2 = "sample"
        .Range("A2").Formula = "=2+3"
        .Range("A1:F6").Select
    End With
    fixture.Saved = True
    Set context = New CNxFeatureDialogContext
    context.Configure NX_FEATURE_DRAW_BUSINESS_TABLE, "draw", "table", False
    Set panel = New FNxDrawTable
    panel.BindFeatureContext context
    Require fixture.Saved, "workbook changed during BindFeatureContext"
    panel.Show vbModeless
    DoEvents
    Require fixture.Saved, "workbook changed during Show/DoEvents"
    Require panel.cmdExecute.Enabled, "preview did not enable Apply"
    For Each control In panel.Controls
        If Left$(control.Name, 11) = "nxPrv_cell_" Then count = count + 1
    Next control
    Require count = 25, "preview must cap at 5x5"
    Require panel.cmdCancel.Top + panel.cmdCancel.Height <= panel.InsideHeight, "footer clipped"
    panel.cboTableStyle.ListIndex = 1
    Require fixture.Saved, "workbook changed during style selection"
    Require panel.cmdExecute.Enabled, "style change did not refresh preview"
    panel.txtRange.Value = "not a range!"
    Require fixture.Saved, "workbook changed during invalid range resolution"
    Require Not panel.cmdExecute.Enabled, "invalid range leaves Apply enabled"
    count = 0
    For Each control In panel.Controls
        If Left$(control.Name, 6) = "nxPrv_" Then count = count + 1
    Next control
    Require count = 0, "invalid range leaves stale preview"
    panel.txtRange.Value = fixture.Worksheets(1).Range("A1:B2").Address(External:=True)
    Require fixture.Saved, "workbook changed during manual range resolution"
    Require panel.cmdExecute.Enabled, "manual range correction failed"
    Require fixture.Saved, "preview dirtied workbook"
    Require fixture.Worksheets(1).Range("A2").Formula = "=2+3", "preview changed formula"
    Unload panel
    Set panel = Nothing
    Set context = New CNxFeatureDialogContext
    context.Configure NX_FEATURE_DRAW_TITLE_TABLE, "draw", "title", False
    Set panel = New FNxDrawTable
    panel.BindFeatureContext context
    panel.Show vbModeless
    DoEvents
    Require Not panel.cboBodySize.Visible, "title body options visible"
    Require Not panel.cboHeaderRows.Visible, "title header rows visible"
    Require panel.cboHeaderSize.Value = "머리글 크기: 13", "title font default"
    Require panel.cboHeaderAlign.ListIndex = 0, "title alignment default"
    Require panel.cmdExecute.Enabled, "title preview disabled"
    Require fixture.Saved, "title preview dirtied workbook"
    Unload panel
    fixture.Close False
    NxDrawPopupProbe = "PASS|draw_popup|names;preview;invalid_range;manual_range;title_defaults;saved"
    Exit Function
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    If Not fixture Is Nothing Then fixture.Close False
    NxDrawPopupProbe = "FAIL|draw_popup|" & detail
End Function

Public Function NxR59RunCase(ByVal kind As String, ByVal outputRoot As String) As String
    If kind = "draw_popup" Then NxR59RunCase = NxDrawPopupProbe(): Exit Function
    Dim book As Workbook, sheet As Worksheet, first As Worksheet, second As Worksheet
    Dim hidden As Worksheet, veryHidden As Worksheet, chart As Chart, extra As Window
    Dim result As CNxResult, cancelSave As CNxR59CancelSave, code As Object
    Dim path As String, extension As String, format As Long, previousEvents As Boolean
    Dim detail As String, focusWasEnabled As Boolean, view As Window, backend As String
    Dim initialAddress As String, initialSheet As String, initialZoom As Long
    Dim fixtureIndex As Long
    previousEvents = Application.EnableEvents
    On Error GoTo Failed
    If kind = "no_book" Or kind = "addin" Then
        If kind = "addin" Then Set book = ThisWorkbook
        Set result = NxRunMannerSave(book)
        Require result.Outcome = NxEnvironmentError, "missing/addin guard result"
        Set book = Nothing
        NxR59RunCase = "PASS|" & kind
        Exit Function
    End If
    If kind = "taxonomy" Then
        Require NxGeneratedNavigationItems().Count = 83, "public count"
        Require NxGeneratedNavigationRouteExists("feature:NX-UTIL-DOCUMENT-NAVIGATOR"), "document navigator route"
        Require Not NxGeneratedNavigationRouteExists("command:NX-CMD-SHEET-RB-SHEET-SAVE-TO-FILE"), "deleted route"
        Require NxGeneratedNavigationRouteField("command:NX-CMD-CLIPBOARD-RB-LHECLIPBOARD-COPYVISIBLE", "id") = NX_FEATURE_DATA_COPY_VISIBLE, "copy alias"
        Require InStr(NxRibbonAllFunctionsMenuXml(), "선택 시트들 파일로 저장") = 0, "retired menu"
        Require NxGeneratedNavigationRouteField("feature:NX-FILE-MANNER-SAVE", "launch_surface") = "direct", "direct save"
        NxR59RunCase = "PASS|taxonomy"
        Exit Function
    End If
    extension = ".xlsx": format = xlOpenXMLWorkbook
    If kind = "xlsm" Then extension = ".xlsm": format = xlOpenXMLWorkbookMacroEnabled
    If kind = "xlsb" Then extension = ".xlsb": format = xlExcel12
    path = outputRoot & Application.PathSeparator & kind & extension
    TraceStep outputRoot, kind, "fixture_create"
    ' Clone the existing note fixture: AddComment itself can crash this Excel
    ' build before the product is invoked. The original fixture stays closed.
    Set book = Workbooks.Add(outputRoot & Application.PathSeparator & "startup-fixture.xlsx")
    For fixtureIndex = book.Worksheets.Count To 2 Step -1
        book.Worksheets(fixtureIndex).Delete
    Next fixtureIndex
    TraceStep outputRoot, kind, "fixture_workbook_added"
    Set first = book.Worksheets(1)
    first.Name = "First"
    first.Range("B2").Value2 = "original value"
    first.Range("C3").Formula = "=12+30"
    TraceStep outputRoot, kind, "fixture_before_comment"
    Require Not first.Range("B2").Comment Is Nothing, "note fixture present"
    Require first.Range("B2").Comment.Text = "r60 keep comment", "note fixture text"
    TraceStep outputRoot, kind, "fixture_comment_added"
    first.Range("B2").Interior.Color = RGB(30, 120, 200)
    book.Names.Add Name:="KeepName", RefersTo:="=First!$B$2"
    book.BuiltinDocumentProperties("Author") = "r59 fixture author"
    Set second = book.Worksheets.Add(After:=first)
    TraceStep outputRoot, kind, "fixture_second_added"
    second.Name = "Second"
    Set hidden = book.Worksheets.Add(After:=second)
    hidden.Name = "Hidden"
    hidden.Visible = xlSheetHidden
    Set veryHidden = book.Worksheets.Add(After:=hidden)
    veryHidden.Name = "VeryHidden"
    veryHidden.Visible = xlSheetVeryHidden
    TraceStep outputRoot, kind, "fixture_sheets_ready"
    If kind = "xlsm" Then
        Set code = book.VBProject.VBComponents.Add(1)
        code.Name = "KeepMacro"
        code.CodeModule.AddFromString "Public Function KeepMacroValue() As Long" & vbCrLf & "KeepMacroValue = 42" & vbCrLf & "End Function"
        Set code = Nothing
    End If
    first.Select
    If kind = "frozen_protected" Or kind = "frozen" Or kind = "protected" Or kind = "cancel_frozen" Or kind = "split" Then
        first.Range("B2").Select
        If kind <> "protected" Then
            ActiveWindow.SplitRow = 1
            ActiveWindow.SplitColumn = 1
            If kind <> "split" Then ActiveWindow.FreezePanes = True
        End If
        If kind = "protected" Or kind = "frozen_protected" Then first.Protect "r59-test"
    End If
    If kind = "chart" Then
        Set chart = book.Charts.Add(After:=veryHidden)
        chart.Name = "ChartOnly"
        first.Select
    End If
    If kind = "multiwindow" Then
        Set view = ActiveWindow
        Set extra = book.NewWindow
        extra.Zoom = 125
        view.Activate
    End If
    Set view = ActiveWindow
    first.Select: view.Zoom = 75: first.Range("D9").Select
    second.Select: view.Zoom = 125: second.Range("F12").Select
    TraceStep outputRoot, kind, "fixture_save"
    book.SaveAs path, format
    If kind <> "already_clean" Then second.Range("B2").Value2 = "pending edit"
    If kind = "grouped" Or kind = "cancel_grouped" Then book.Sheets(Array("First", "Second")).Select
    If kind = "cancel" Or kind = "cancel_grouped" Or kind = "cancel_frozen" Then
        Set cancelSave = New CNxR59CancelSave
        cancelSave.Start book
    ElseIf kind = "events_disabled" Then
        Application.EnableEvents = False
    ElseIf kind = "read_only" Then
        TraceStep outputRoot, kind, "readonly_close"
        book.Close False
        ' Release fixture object references before the harness reopens it.
        Set book = Nothing
        NxR59RunCase = "READY|" & path
        GoTo CleanUp
    ElseIf kind = "focus_cf" Then
        NxFocusEnable RGB(255, 210, 40)
        focusWasEnabled = True
        backend = NxFocusControllerBackend()
        Require Left$(backend, 2) = "cf", "conditional backend expected: " & backend
    End If

    TraceStep outputRoot, kind, "manner_start"
    initialAddress = ActiveCell.Address
    initialSheet = ActiveSheet.Name
    initialZoom = ActiveWindow.Zoom
    TraceStep outputRoot, kind, "before_selection:" & initialAddress
    TraceStep outputRoot, kind, "before_group:" & CStr(ActiveWindow.SelectedSheets.Count) & "|" & initialSheet
    If kind = "cancel_grouped" Then Require ActiveWindow.SelectedSheets.Count = 2, "fixture must be grouped"
    Set result = NxRunMannerSave(book)
    TraceStep outputRoot, kind, "manner_done:" & CStr(result.Outcome)
    If kind = "cancel" Or kind = "cancel_grouped" Or kind = "cancel_frozen" Then
        Require result.Outcome = NxCancelled, "save cancel outcome: " & CStr(result.Outcome)
        Require Not book.Saved, "unsaved edits retained"
        Require first.Range("B3").Value2 = "event edit preserved", "BeforeSave edits retained"
        Require view.ActiveSheet.Name = initialSheet And view.Zoom = initialZoom, "cancel restores active view"
        Require ActiveCell.Address = initialAddress, "cancel restores selection: " & initialAddress & " -> " & ActiveCell.Address
        If kind = "cancel_grouped" Then Require view.SelectedSheets.Count = 2, "cancel restores group"
        If kind = "cancel_frozen" Then
            first.Select
            Require ActiveWindow.FreezePanes And ActiveWindow.SplitRow = 1 And ActiveWindow.SplitColumn = 1, "cancel preserves frozen layout"
            Require ActiveCell.Address = "$D$9" And ActiveWindow.Zoom = 75, "cancel restores frozen view"
        End If
    ElseIf kind = "read_only" Or kind = "events_disabled" Then
        Require result.Outcome = NxEnvironmentError, "guard rejects invalid save"
        Require view.ActiveSheet.Name = "Second" And view.Zoom = 125, "guard preserves view"
    Else
        Require result.Outcome = NxSuccess, "save failed: " & result.Recovery
        Require book.FileFormat = format, "source format"
        Require view.ActiveSheet.Name = "First" And ActiveCell.Address = "$A$1", "first sheet A1"
        Require view.Zoom = 100, "first zoom"
        If kind = "frozen_protected" Then Require first.ProtectContents And ActiveWindow.FreezePanes And ActiveWindow.SplitRow = 1 And ActiveWindow.SplitColumn = 1, "protection and freeze retained"
        Require first.Range("B2").Value2 = "original value", "value retained"
        Require first.Range("C3").Formula = "=12+30", "formula retained"
        Require first.Range("B2").Comment.Text = "r60 keep comment", "comment retained"
        Require first.Range("B2").Interior.Color = RGB(30, 120, 200), "format retained"
        Require hidden.Visible = xlSheetHidden And veryHidden.Visible = xlSheetVeryHidden, "visibility retained"
        second.Select
        Require ActiveCell.Address = "$A$1" And ActiveWindow.Zoom = 100, "second sheet saved view"
        If kind <> "already_clean" Then Require second.Range("B2").Value2 = "pending edit", "pending values saved"
        If kind = "xlsm" Then Require book.HasVBProject, "VBA retained"
        If kind = "multiwindow" Then Require extra.Zoom = 125, "other window unchanged"
        If kind = "focus_cf" Then Require NxFocusIsEnabled(), "focus restored"
    End If
    NxR59RunCase = "PASS|" & kind & "|" & path
    GoTo CleanUp
Failed:
    NxR59RunCase = "FAIL|" & kind & "|" & Err.Description
CleanUp:
    On Error Resume Next
    If Not cancelSave Is Nothing Then cancelSave.StopListening
    If focusWasEnabled Then NxFocusDisable
    Application.EnableEvents = previousEvents
    If Not book Is Nothing Then book.Close False
End Function

Public Function NxR59CheckReadOnly() As String
    Dim book As Workbook, result As CNxResult
    On Error GoTo Failed
    Set book = ActiveWorkbook
    Require book.ReadOnly, "fixture must actually be read-only"
    book.Worksheets("Second").Select
    ActiveWindow.Zoom = 125
    Range("F12").Select
    Set result = NxRunMannerSave(book)
    Require result.Outcome = NxEnvironmentError, "read-only save rejected"
    Require ActiveSheet.Name = "Second" And ActiveWindow.Zoom = 125, "read-only view retained"
    Require ActiveCell.Address = "$F$12", "read-only selection retained"
    NxR59CheckReadOnly = "PASS|read_only|" & book.FullName
    Exit Function
Failed:
    NxR59CheckReadOnly = "FAIL|read_only|" & Err.Description
End Function

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 5901, "r59 probe", message
End Sub

Private Sub TraceStep(ByVal outputRoot As String, ByVal kind As String, ByVal stage As String)
    Dim handle As Integer
    handle = FreeFile
    Open outputRoot & Application.PathSeparator & "steps.log" For Append As #handle
    Print #handle, kind & "|" & stage
    Close #handle
End Sub
