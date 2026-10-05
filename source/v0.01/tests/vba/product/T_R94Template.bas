Attribute VB_Name = "T_R94Template"
Option Explicit
Private mPanel As FNxTemplateManager

Public Function Names() As String
    Names = "register_sheet|register_workbook|register_file|metadata_compatibility|load_copy|restore|restore_failure|restore_collision|tamper_rejected|catalog_500|ui_layout|register_direct_route|large_range|format_fidelity|cancel_workflow|load_benchmark|ui_preferences|registration_filter|registration_cf_filter|print_settings"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, output As Workbook, request As CNxTemplateRequest, result As CNxResult
    Dim records As Collection, keys As Collection, record As CNxTemplateRecord, changed As CNxTemplateRecord
    Dim fs As Object, root As String, folder As String, quarantine As String, id As String, path As String
    Dim i As Long, handle As Integer, text As String, started As Double, elapsed As Double, rejected As Long
    Dim manager As FNxTemplateManager, register As FNxTemplateRegister, control As Object
    Dim beforeCopies As Object, candidateBook As Workbook
    Dim progress As C_R93CompareProgress
    Dim priorWidths As String, restoreWidths As Boolean
    Dim benchmarkValues As Variant, benchmarkRow As Long, benchmarkColumn As Long, expected As Double
    Dim oldAlerts As Boolean, oldEvents As Boolean, oldScreen As Boolean
    On Error GoTo Failed
    oldAlerts = Application.DisplayAlerts: oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    Application.DisplayAlerts = False: Application.EnableEvents = False: Application.ScreenUpdating = False
    root = Environ$("TEMP") & "\nx94-" & Left$(Replace(NxCreateRunUuid(), "-", ""), 12)
    NxTemplateEnsureStore root
    handle = FreeFile: Open Environ$("LHEXCEL_PROFILE_ROOT") & "\fixture-roots.txt" For Append As #handle
    Print #handle, name & vbTab & root: Close #handle: handle = 0
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set book = Workbooks.Add(xlWBATWorksheet)
    If name = "print_settings" Then
        RunPrintSettingsCase book, root
        RunCase = "PASS|" & name
        GoTo Cleanup
    End If
    If name = "load_benchmark" Then
        path = Environ$("LHEXCEL_TEMPLATE_BENCH_ROOT")
        If Len(path) = 0 Then Err.Raise 5, , "benchmark fixture required"
        fs.CopyFolder path & "\Packages\*", root & "\Packages\", False
        Set records = NxTemplateListMetadata(root)
        If records.Count <> 1 Then Err.Raise 5, , "benchmark fixture count"
        Set record = records(1): id = record.TemplateId
        If CDbl(record.RowCount) * CDbl(record.ColumnCount) <> 90000# Then Err.Raise 5, , "benchmark fixture size"
        started = Timer
        Set result = NxTemplateLoadForTest(id, book, True, root): RequireSuccess result
        elapsed = Timer - started
        If elapsed < 0 Then elapsed = elapsed + 86400#
        If book.Worksheets(2).Range("AX1800").Value2 <> 12 Or book.Worksheets(2).Range("B3").Value2 <> 24 Then Err.Raise 5, , "benchmark data mismatch"
        benchmarkValues = book.Worksheets(2).Range("A1:AX1800").Value2
        For benchmarkRow = 1 To 1800
            For benchmarkColumn = 1 To 50
                expected = 12
                If benchmarkRow = 3 And benchmarkColumn = 2 Then expected = 24
                If benchmarkValues(benchmarkRow, benchmarkColumn) <> expected Then Err.Raise 5, , "benchmark full-range mismatch"
            Next benchmarkColumn
        Next benchmarkRow
        If book.Worksheets(2).Range("B3").FormulaR1C1 <> "=RC[-1]*2" Then Err.Raise 5, , "benchmark formula mismatch"
        handle = FreeFile: Open Environ$("LHEXCEL_PROFILE_ROOT") & "\load-benchmark.txt" For Output As #handle
        Print #handle, "load_90000_seconds=" & CStr(elapsed)
        Print #handle, "package_sha256=" & record.PackageSha256
        Close #handle: handle = 0
        RunCase = "PASS|" & name
        GoTo Cleanup
    End If
    If name = "large_range" Then
        book.Worksheets(1).Range("A1:AX1800").Value2 = 12
        If Environ$("LHEXCEL_TEMPLATE_CF_BENCH") = "1" Then
            With book.Worksheets(1).Range("A1:AX1800")
                .Interior.Color = RGB(210, 225, 240)
                .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                .FormatConditions(1).Interior.Color = vbRed
            End With
        End If
    End If
    book.Worksheets(1).Range("A1:B3").Value2 = 12
    book.Worksheets(1).Range("B3").Formula = "=A3*2"
    If name = "registration_filter" Or name = "registration_cf_filter" Then
        With book.Worksheets(1)
            .Hyperlinks.Add Anchor:=.Range("A2"), Address:="https://example.invalid/", TextToDisplay:="제외할 내용"
            .Range("A1").AddComment "제외할 메모"
            .Shapes.AddShape msoShapeRectangle, 10, 10, 20, 20
            .Range("A1:B3").Font.Size = 11
            .Range("A1:B3").Interior.Color = RGB(210, 225, 240)
            If name = "registration_cf_filter" Then
                .Range("A1:B3").FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
                .Range("A1:B3").FormatConditions(1).Interior.Color = vbRed
            End If
        End With
    End If
    If name = "format_fidelity" Then
        With book.Worksheets(1)
            .Range("A1:B1").Merge
            .Range("A1").Value2 = "제목"
            .Range("A2").NumberFormat = "@"
            .Range("A2").Value2 = "=literal"
            .Range("A1:B3").Font.Name = "맑은 고딕"
            .Range("A1:B3").Font.Size = 9
            .Range("A1:B1").Font.Bold = True
            .Range("A1:B1").Interior.Color = RGB(210, 225, 240)
            .Range("A1:B1").Borders(xlEdgeBottom).LineStyle = xlDouble
            .Range("B2").NumberFormat = "0.00%"
            .Columns("A").ColumnWidth = 22
            .Rows(1).RowHeight = 31
        End With
    End If
    book.Saved = True
    started = Timer
    Set request = NxTemplateManagerRequest(book, book.Worksheets(1), "sheet_used_range", "업무 양식", "검색용 설명", True, True, "보고")
    If name = "registration_filter" Or name = "registration_cf_filter" Then _
        Set request = NxTemplateManagerRequest(book, book.Worksheets(1), "sheet_used_range", "업무 양식", "", False, True, "보고")
    If name = "register_workbook" Then
        book.Worksheets.Add After:=book.Worksheets(1)
        book.Worksheets(2).Range("A1").Value2 = 33
        book.Saved = True
        Set request = NxTemplateManagerRequest(book, book.Worksheets(1), "workbook", "두 시트", "", True, True, "보고")
    End If
    If name = "register_file" Then
        path = root & "\source.xlsx": book.SaveAs path, xlOpenXMLWorkbook
        book.Close False: Set book = Workbooks.Add(xlWBATWorksheet)
        Set result = NxTemplateRegisterFileForTest(path, book, "파일 양식", True, True, root, NxTemplateFileSha256(path))
    Else
        Set result = NxTemplateRegisterForTest(request, root)
    End If
    RequireSuccess result
    Set records = NxTemplateListMetadata(root, False)
    If records.Count <> 1 Then Err.Raise 5, , "registration count"
    Set record = records(1): id = record.TemplateId
    folder = NxTemplateTemplateFolderPath(id, root)
    Select Case name
        Case "registration_filter", "registration_cf_filter"
            Set output = Workbooks.Open(folder & "\template.xlsx", UpdateLinks:=0, ReadOnly:=True)
            With output.Worksheets(1)
                If Not IsEmpty(.Range("A2").Value2) Or Not IsEmpty(.Range("A3").Value2) Then Err.Raise 5, , "excluded contents copied"
                If .Range("B3").FormulaR1C1 <> "=RC[-1]*2" Then Err.Raise 5, , "formula lost"
                If .Hyperlinks.Count <> 0 Or .Shapes.Count <> 0 Or .Comments.Count <> 0 Then Err.Raise 5, , "excluded objects copied"
                If .UsedRange.FormatConditions.Count <> 0 Then Err.Raise 5, , "conditional formats copied"
                If .Range("A1").Font.Size <> 11 Then Err.Raise 5, , "base format lost"
                If .Range("A1").Interior.Color <> RGB(210, 225, 240) Then Err.Raise 5, , "conditional fill baked into base"
            End With
            If Not book.Saved Or book.Worksheets(1).Hyperlinks.Count <> 1 Then Err.Raise 5, , "source mutated"
            If name = "registration_cf_filter" Then
                If book.Worksheets(1).Range("A1:B3").FormatConditions.Count <> 1 Then Err.Raise 5, , "source conditional format removed"
            End If
        Case "ui_preferences"
            priorWidths = GetSetting("LHExcel", "TemplateLibrary", "ColumnWidths", "")
            restoreWidths = True
            SaveSetting "LHExcel", "TemplateLibrary", "ColumnWidths", "220 pt;70 pt;90 pt;114 pt;0 pt"
            Set manager = New FNxTemplateManager: Load manager
            If Val(manager.Controls("lstTemplates").ColumnWidths) <> 220 Then Err.Raise 5, , "width restore"
            manager.ResizeHeader 1, -10
            Unload manager: Set manager = Nothing
            Set manager = New FNxTemplateManager: Load manager
            If Val(manager.Controls("lstTemplates").ColumnWidths) <> 210 Then Err.Raise 5, , "width persistence"
            Unload manager: Set manager = Nothing
            SaveSetting "LHExcel", "TemplateLibrary", "ColumnWidths", "invalid"
            Set manager = New FNxTemplateManager: Load manager
            If Val(manager.Controls("lstTemplates").ColumnWidths) <> 200 Then Err.Raise 5, , "invalid width fallback"
        Case "cancel_workflow"
            If InStr(NxTemplateProgressText("검사", 1, 4), "25%") = 0 Then Err.Raise 5, , "phase progress percentage"
            NxTemplateRequestCancel
            If InStr(NxTemplateProgressText("검사", 1, 4), "취소") = 0 Then Err.Raise 5, , "cancel progress text"
            NxTemplateEndProgress
            Set progress = New C_R93CompareProgress
            progress.StopStage = "내용 불러오기"
            NxTemplateBeginProgress progress
            Set result = NxTemplateLoadForTest(id, book, True, root)
            NxTemplateEndProgress
            If Not progress.Fired Or result.Outcome <> NxCancelled Then Err.Raise 5, , "load cancellation missing"
            If book.Worksheets.Count <> 1 Then Err.Raise 5, , "cancel left a worksheet"
            If book.Worksheets(1).Range("B3").Value2 <> 24 Then Err.Raise 5, , "cancel changed original"
            progress.Fired = False
            progress.StopStage = "양식 작성"
            Set request = NxTemplateManagerRequest(book, book.Worksheets(1), "sheet_used_range", "취소 검사", "", True, True, "보고")
            NxTemplateBeginProgress progress
            Set result = NxTemplateRegisterForTest(request, root)
            NxTemplateEndProgress
            If Not progress.Fired Or result.Outcome <> NxCancelled Then Err.Raise 5, , "registration cancellation missing"
            Set records = NxTemplateListMetadata(root)
            If records.Count <> 1 Then Err.Raise 5, , "cancel changed registered templates"
            If fs.GetFolder(root & "\Staging").SubFolders.Count <> 0 Then Err.Raise 5, , "cancel left staging"
        Case "format_fidelity"
            Set result = NxTemplateLoadForTest(id, book, True, root): RequireSuccess result
            With book.Worksheets(2)
                If .Range("A1").MergeArea.Address <> "$A$1:$B$1" Then Err.Raise 5, , "merge lost"
                If .Range("A2").HasFormula Or .Range("A2").Value2 <> "=literal" Then Err.Raise 5, , "literal changed"
                If .Range("B3").FormulaR1C1 <> "=RC[-1]*2" Then Err.Raise 5, , "formula changed"
                If .Range("B2").NumberFormat <> "0.00%" Then Err.Raise 5, , "number format lost"
                If Not .Range("A1").Font.Bold Or .Range("B2").Font.Size <> 9 Then Err.Raise 5, , "font lost"
                If .Range("A1").Interior.Color <> RGB(210, 225, 240) Then Err.Raise 5, , "fill lost"
                If .Range("A1:B1").Borders(xlEdgeBottom).LineStyle <> xlDouble Then Err.Raise 5, , "border lost"
                If Abs(.Columns("A").ColumnWidth - 22) > 0.2 Or .Rows(1).RowHeight <> 31 Then Err.Raise 5, , "dimensions lost"
            End With
        Case "large_range"
            elapsed = Timer - started
            handle = FreeFile: Open Environ$("LHEXCEL_PROFILE_ROOT") & "\large-range.txt" For Output As #handle
            Print #handle, "register_90000_seconds=" & CStr(elapsed): Close #handle: handle = 0
            If Not book.Saved Then Err.Raise 5, , "registration changed source Saved"
            started = Timer
            Set result = NxTemplateLoadForTest(id, book, True, root): RequireSuccess result
            elapsed = Timer - started
            If book.Worksheets(2).Range("AX1800").Value2 <> 12 Or book.Worksheets(2).Range("B3").Value2 <> 24 Then Err.Raise 5, , "large content"
            If Environ$("LHEXCEL_TEMPLATE_CF_BENCH") = "1" Then
                If book.Worksheets(2).UsedRange.FormatConditions.Count <> 0 Then Err.Raise 5, , "large conditional format retained"
                If book.Worksheets(2).Range("AX1800").Interior.Color <> RGB(210, 225, 240) Then Err.Raise 5, , "large base fill changed"
                If book.Worksheets(1).UsedRange.FormatConditions.Count <> 1 Then Err.Raise 5, , "large source rule changed"
            End If
            benchmarkValues = book.Worksheets(2).Range("A1:AX1800").Value2
            For benchmarkRow = 1 To 1800
                For benchmarkColumn = 1 To 50
                    expected = 12
                    If benchmarkRow = 3 And benchmarkColumn = 2 Then expected = 24
                    If benchmarkValues(benchmarkRow, benchmarkColumn) <> expected Then Err.Raise 5, , "registration full-range mismatch"
                Next benchmarkColumn
            Next benchmarkRow
            handle = FreeFile: Open Environ$("LHEXCEL_PROFILE_ROOT") & "\large-range.txt" For Append As #handle
            Print #handle, "load_90000_seconds=" & CStr(elapsed): Close #handle: handle = 0
        Case "register_sheet", "register_workbook", "register_file"
            Set result = NxTemplateLoadForTest(id, book, True, root): RequireSuccess result
            If name = "register_workbook" And book.Worksheets.Count <> 4 Then Err.Raise 5, , "workbook sheets"
            If name <> "register_file" And record.Category <> "보고" Then Err.Raise 5, , "category lost"
        Case "metadata_compatibility"
            path = folder & "\metadata.ini"
            handle = FreeFile: Open path For Binary Access Read As #handle
            text = Space$(LOF(handle)): Get #handle, , text: Close #handle: handle = 0
            text = Replace(text, "schema_version=2", "schema_version=1")
            text = Left$(text, InStr(text, "category=") - 1)
            handle = FreeFile: Open path For Output As #handle: Print #handle, Left$(text, Len(text) - 2): Close #handle: handle = 0
            Set record = NxTemplateReadMetadata(path)
            If record.Category <> "" Or record.LastUsedUtc <> "" Then Err.Raise 5, , "v1 defaults"
            Set changed = NxTemplateCopyDetails(record, "새 이름", "새 설명", "분류", "", record.ModifiedUtc)
            On Error Resume Next
            NxTemplateReplaceMetadataForTest folder, changed, True
            rejected = Err.Number: Err.Clear
            On Error GoTo Failed
            If rejected = 0 Then Err.Raise 5, , "missing replacement failure"
            Set record = NxTemplateReadMetadata(path)
            If record.DisplayName = "새 이름" Then Err.Raise 5, , "rollback lost"
            NxTemplateReplaceMetadata folder, changed
            Set record = NxTemplateReadMetadata(path)
            If record.Category <> "분류" Or fs.GetFolder(root & "\Backups").Files.Count < 1 Then Err.Raise 5, , "migration/backup"
        Case "load_copy"
            Set result = NxTemplateOpenCopyForTest(id, True, root): RequireSuccess result
            Set output = ActiveWorkbook
            If output Is book Then Err.Raise 5, , "copy not activated"
            If Len(output.Path) > 0 Or output.Worksheets(1).Range("B3").Value2 <> 24 Then Err.Raise 5, , "copy content"
            Set record = NxTemplateReadMetadata(folder & "\metadata.ini")
            If Len(record.LastUsedUtc) <> 20 Then Err.Raise 5, , "recent use missing"
        Case "restore", "restore_failure", "restore_collision"
            RequireSuccess NxTemplateDeleteForTest(id, root)
            Set records = NxTemplateLibraryRecords("", "", 0, True, keys, root)
            If records.Count <> 1 Then Err.Raise 5, , "deleted listing"
            quarantine = keys(1)
            If name = "restore_collision" Then RequireSuccess NxTemplateRegisterForTest(request, root)
            Set result = NxTemplateRestore(id, quarantine, root, name = "restore_failure")
            If name = "restore" Then
                RequireSuccess result
                If Not fs.FolderExists(folder) Then Err.Raise 5, , "restore absent"
            Else
                If result.Outcome = NxSuccess Then Err.Raise 5, , "invalid restore allowed"
                If fs.FolderExists(folder) Or Not fs.FolderExists(NxTemplateQuarantinePath(quarantine, root) & "\" & id) Then Err.Raise 5, , "restore rollback"
            End If
        Case "tamper_rejected"
            handle = FreeFile: Open folder & "\template.xlsx" For Append As #handle: Print #handle, "tamper": Close #handle: handle = 0
            Set records = NxTemplateListMetadata(root, False)
            If records.Count <> 1 Then Err.Raise 5, , "catalog lost entry"
            Set result = NxTemplateLoadForTest(id, book, True, root)
            If result.Outcome = NxSuccess Or book.Worksheets.Count <> 1 Then Err.Raise 5, , "tamper accepted"
        Case "catalog_500"
            For i = 2 To 500
                path = NxTemplateTemplateFolderPath("template-bench-" & Format$(i, "0000"), root)
                fs.CreateFolder path
                Set changed = New CNxTemplateRecord
                changed.Configure "template-bench-" & Format$(i, "0000"), "양식 " & Format$(i, "0000"), "검색", record.PackageSha256, record.LogicalSha256, _
                    record.SourceKind, 3, 2, "", "healthy", record.CreatedUtc, record.ModifiedUtc, "보고"
                changed.Seal
                NxTemplateWriteMetadata path & "\metadata.ini", changed
                fs.CopyFile folder & "\template.xlsx", path & "\template.xlsx", False
            Next i
            started = Timer
            Set records = NxTemplateLibraryRecords("", "", 0, False, keys, root)
            elapsed = Timer - started
            If records.Count <> 500 Or Workbooks.Count > 3 Then Err.Raise 5, , "catalog count/opened workbooks"
            handle = FreeFile: Open Environ$("LHEXCEL_PROFILE_ROOT") & "\timing.txt" For Output As #handle: Print #handle, "catalog_500_seconds=" & CStr(elapsed): Close #handle: handle = 0
            Set records = NxTemplateLibraryRecords("0499", "보고", 0, False, keys, root)
            If records.Count <> 1 Then Err.Raise 5, , "filter"
            Set records = NxTemplateLibraryRecords("", "", 0, False, keys, root)
            started = Timer
            For i = 1 To 10
                Set result = Nothing
                Set changed = Nothing
                Dim filtered As Collection, filteredKeys As Collection
                Set filtered = NxTemplateFilterRecords(records, keys, "0499", "보고", 0, filteredKeys)
                If filtered.Count <> 1 Then Err.Raise 5, , "cached filter"
            Next i
            elapsed = (Timer - started) / 10
            handle = FreeFile: Open Environ$("LHEXCEL_PROFILE_ROOT") & "\timing.txt" For Append As #handle
            Print #handle, "cached_search_seconds=" & CStr(elapsed): Close #handle: handle = 0
            id = records(2).TemplateId
            RequireSuccess NxTemplateMove(id, -1, root)
            Set records = NxTemplateLibraryRecords("", "", 0, False, keys, root)
            If records(1).TemplateId <> id Then Err.Raise 5, , "move up persistence"
            RequireSuccess NxTemplateMove(id, 1, root)
            Set records = NxTemplateLibraryRecords("", "", 0, False, keys, root)
            If records(2).TemplateId <> id Then Err.Raise 5, , "move down persistence"
        Case "ui_layout"
            Set manager = New FNxTemplateManager: Load manager
            Set register = New FNxTemplateRegister: register.BindSource book, book.Worksheets(1), "sheet_used_range"
            For Each control In manager.Controls
                If control.Left < 0 Or control.Top < 0 Or control.Left + control.Width > manager.InsideWidth + 2 Or control.Top + control.Height > manager.InsideHeight + 2 Then Err.Raise 5, , "manager clipping " & control.Name
            Next control
            If manager.Controls("cboOutput").ListCount <> 2 Or Not manager.Controls("cmdCancel").Cancel Then Err.Raise 5, , "manager controls"
            If manager.Controls("cboOutput").ListIndex <> 0 Or manager.Controls("cboOutput").Value <> "새 통합문서" Then Err.Raise 5, , "new workbook default"
            manager.ResizeHeader 1, 20
            If Val(manager.Controls("lstTemplates").ColumnWidths) <> 220 Then Err.Raise 5, , "column resize"
            manager.ResizeHeader 1, 500
            If Val(manager.Controls("lstTemplates").ColumnWidths) <> 220 Then Err.Raise 5, , "column minimum"
            manager.ResizeHeader 1, -20
            If manager.Controls("cmdRegisterFile").Left <= manager.Controls("lstTemplates").Left + manager.Controls("lstTemplates").Width Then Err.Raise 5, , "right actions"
            For i = 1 To 50
                manager.Controls("lstTemplates").AddItem "휠 확인 " & CStr(i)
            Next i
            manager.Show vbModeless
            manager.Controls("lstTemplates").SetFocus
            manager.Controls("lstTemplates").TopIndex = 0
            If Not NxFormWheelAttach(manager) Then Err.Raise 5, , "wheel hook"
            manager.HandleWheelDelta -120
            If manager.Controls("lstTemplates").TopIndex = 0 Then Err.Raise 5, , "wheel down"
            manager.HandleWheelDelta 120
            If manager.Controls("lstTemplates").TopIndex <> 0 Then Err.Raise 5, , "wheel up"
        Case "register_direct_route"
            Set register = New FNxTemplateRegister
            register.BindSource book, book.Worksheets(1), "sheet_used_range"
            register.Controls("txtName").Value = "직접 등록"
            register.Controls("txtCategory").Value = "보고"
            register.Controls("cboContents").ListIndex = 1
            register.R94ApplyDirect
            RequireSuccess register.OperationResult
            Set records = NxTemplateRecordsForForm()
            If records.Count <> 1 Then Err.Raise 5, , "direct register"
            Set record = records(1)
            RequireSuccess NxTemplateRunManagerAction("use", book, record.TemplateId, True)
            If book.Worksheets.Count <> 2 Then Err.Raise 5, , "routed load"
            Set manager = New FNxTemplateManager
            manager.BindContext book, "NX-TPL-LIST"
            manager.Controls("txtSearch").Value = "없는항목"
            If manager.Controls("lstTemplates").ListCount <> 0 Then Err.Raise 5, , "UI filter"
            manager.Controls("txtSearch").Value = "직접"
            If manager.Controls("lstTemplates").ListCount <> 1 Then Err.Raise 5, , "UI filter restore"
            Unload manager: Set manager = Nothing
            RequireSuccess NxTemplateRunManagerAction("rename", book, record.TemplateId, True, "변경 양식", "설명", "분류")
            RequireSuccess NxTemplateRunManagerAction("move", book, record.TemplateId, True, moveDirection:=-1)
            Set beforeCopies = CreateObject("Scripting.Dictionary")
            For Each candidateBook In Application.Workbooks: beforeCopies(candidateBook.Name) = True: Next candidateBook
            RequireSuccess NxTemplateRunManagerAction("copy", book, record.TemplateId, True)
            For Each candidateBook In Application.Workbooks
                If Not beforeCopies.Exists(candidateBook.Name) Then Set output = candidateBook
            Next candidateBook
            If output Is Nothing Then Err.Raise 5, , "routed copy missing"
            If output Is book Then Err.Raise 5, , "routed copy target"
            If Len(output.Path) <> 0 Then Err.Raise 5, , "routed copy must remain unsaved"
            output.Close False: Set output = Nothing
            RequireSuccess NxTemplateRunManagerAction("delete", book, record.TemplateId, True)
            Set records = NxTemplateRecordsForForm()
            If records.Count <> 0 Then Err.Raise 5, , "routed delete"
        Case Else: Err.Raise 5, , "unknown case"
    End Select
    RunCase = "PASS|" & name
    GoTo Cleanup
Failed:
    RunCase = "FAIL|" & name & "|" & CStr(Err.Number) & "|" & Err.Description
Cleanup:
    On Error Resume Next
    NxTemplateEndProgress
    If handle <> 0 Then Close #handle
    If Not manager Is Nothing Then Unload manager
    If restoreWidths Then SaveSetting "LHExcel", "TemplateLibrary", "ColumnWidths", priorWidths
    If Not register Is Nothing Then Unload register
    If Not output Is Nothing Then output.Close False
    If Not book Is Nothing Then book.Close False
    Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
End Function

Private Sub RunPrintSettingsCase(ByVal book As Workbook, ByVal root As String)
    Dim request As CNxTemplateRequest, result As CNxResult, records As Collection, id As String
    Dim first As Worksheet, second As Worksheet, digest As String, package As Workbook
    Set first = book.Worksheets(1)
    first.Range("C4:F10").Value2 = 12
    first.Range("C4:D4").Merge
    first.Range("F10").Formula = "=C10*2"
    With first.PageSetup
        .PaperSize = xlPaperA4: .Orientation = xlLandscape
        .LeftMargin = 21: .RightMargin = 22: .TopMargin = 23: .BottomMargin = 24
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .PrintArea = "$C$4:$F$10": .PrintTitleRows = "$4:$4": .PrintTitleColumns = "$C:$C"
    End With
    Set second = book.Worksheets.Add(After:=first)
    second.Range("A1:B4").Value2 = 5
    With second.PageSetup
        .PaperSize = xlPaperA4: .Orientation = xlPortrait: .Zoom = 85
        .LeftMargin = 30: .PrintArea = "$A$1:$B$4"
    End With
    digest = NxTemplatePrintFingerprint(first)
    book.Saved = True
    Set request = NxTemplateManagerRequest(book, first, "workbook", "인쇄 양식", "", True, True)
    RequireSuccess NxTemplateRegisterForTest(request, root)
    Set records = NxTemplateListMetadata(root)
    id = records(1).TemplateId
    RequireSuccess NxTemplateLoadForTest(id, book, True, root)
    With book.Worksheets(3).PageSetup
        If .Orientation <> xlLandscape Or .PaperSize <> xlPaperA4 Then Err.Raise 5, , "paper/orientation"
        If Abs(.LeftMargin - 21) > 0.1 Or Abs(.BottomMargin - 24) > 0.1 Then Err.Raise 5, , "margins"
        If .Zoom <> False Or .FitToPagesWide <> 1 Or .FitToPagesTall <> False Then Err.Raise 5, , "fit"
        If .PrintArea <> "$A$1:$D$7" Or .PrintTitleRows <> "$1:$1" Or .PrintTitleColumns <> "$A:$A" Then Err.Raise 5, , "print range mapping"
    End With
    If book.Worksheets(4).PageSetup.Zoom <> 85 Or book.Worksheets(4).PageSetup.Orientation <> xlPortrait Then Err.Raise 5, , "second sheet"
    If book.Worksheets(3).Range("D7").Value2 <> 24 Or book.Worksheets(3).Range("A1").MergeArea.Columns.Count <> 2 Then Err.Raise 5, , "content"
    If NxTemplatePrintFingerprint(first) <> digest Then Err.Raise 5, , "source page setup changed"
    NxTemplateSetFailurePoint "after-second-load-sheet"
    Set result = NxTemplateLoadForTest(id, book, True, root)
    NxTemplateClearFailurePoint
    If result.Outcome = NxSuccess Or book.Worksheets.Count <> 4 Then Err.Raise 5, , "rollback"
    Set package = Workbooks.Open(NxTemplatePackagePath(id, root), UpdateLinks:=0, ReadOnly:=True)
    If Not NxTemplatePrintNamesOnly(package) Then Err.Raise 5, , "print names rejected"
    package.Names.Add Name:="unapproved", RefersTo:="=1"
    If NxTemplatePrintNamesOnly(package) Then Err.Raise 5, , "custom name accepted"
    package.Close False
End Sub

Private Sub RequireSuccess(ByVal result As CNxResult)
    If result Is Nothing Then Err.Raise 5, , "missing result"
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
End Sub

Public Sub OpenPanel()
    Set mPanel = New FNxTemplateManager
    mPanel.BindContext ActiveWorkbook, "NX-TPL-LIST"
    mPanel.Show vbModeless
End Sub

Public Function Caption() As String
    Caption = mPanel.Caption
End Function

Public Sub ClosePanel()
    Unload mPanel
    Set mPanel = Nothing
End Sub
