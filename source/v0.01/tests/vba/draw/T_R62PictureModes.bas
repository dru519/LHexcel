Attribute VB_Name = "T_R62PictureModes"
Option Explicit

Private mFitFaultArmed As Boolean, mFitFaultReached As Boolean, mFitFaultMutated As Boolean
Private mFitFaultBefore As String
Private Const FIT_FAULT_NUMBER As Long = vbObjectError + 6262
Private Const FIT_FAULT_SOURCE As String = "T_R62PictureModes.FitFault"
Private Const FIT_FAULT_DESCRIPTION As String = "r62 fit rollback after width mutation"

' Import into a disposable Product.xlam copy; execute only through the prepared Windows BAT.
' Native control events/engines are checked here. Physical ribbon, file picker and keyboard are NOT_RUN.
Public Function TestNames() As Collection
    Dim names As New Collection, kind As Variant
    For Each kind In Array("range_default", "shape_default", "multi_shape_default", "mode_cancel", "fit_preview", "fit_drift", "insert_preview", "engine_rollback", "picker_state")
        names.Add CStr(kind)
    Next kind
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Dim result As String
    result = NxR62PictureModesCase(name)
    Require Left$(result, 5) = "PASS|", result
End Sub

Public Sub RunAll()
    Dim kind As Variant
    For Each kind In TestNames()
        RunCase CStr(kind)
    Next kind
End Sub

Public Function NxR62PictureModesCase(ByVal kind As String) As String
    Dim book As Workbook, otherBook As Workbook, target As Range, picture As Shape, added As Shape
    Dim view As FNxPictureInsert, request As CNxDrawingRequest, context As CNxExecutionContext
    Dim insertRequest As CNxPictureInsertRequest, preview As CNxPictureInsertPreview
    Dim journal As CNxPictureInsertJournal, result As CNxResult, files As Collection
    Dim folder As String, goodPath As String, badPath As String, before As String, savedState As Variant
    Dim rejected As Boolean, countBefore As Long, mode As Variant, oldRowHeight As Double
    Dim naturalWidth As Double, naturalHeight As Double, scaleFactor As Double
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    Dim caseDetail As String
    Dim selectedPictures As Object, entryTarget As String
    On Error GoTo Failed
    folder = Environ$("TEMP") & "\NxR62Picture-" & Replace$(NxCreateRunUuid(), "-", vbNullString)
    MkDir folder
    goodPath = folder & "\01-picture.bmp": badPath = folder & "\02-broken.bmp"
    WriteBitmap goodPath, True: WriteBitmap badPath, False
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("B2:C4")
    target.Value2 = "preserve": target.Cells(1, 1).Formula = "=2+3"
    target.RowHeight = 30
    Set picture = book.Worksheets(1).Shapes.AddPicture(goodPath, msoFalse, msoTrue, 180, 120, 80, 40)
    picture.Name = "ExistingPicture"
    picture.Placement = xlFreeFloating
    book.Saved = True
    countBefore = book.Worksheets(1).Shapes.Count
    target.Select

    Select Case kind
        Case "picker_state"
            Set view = CreateView(False)
            caseDetail = PickerFormMetrics(view)
            TestPickerState book, target, picture, caseDetail
        Case "range_default"
            before = SourceState(book)
            Set view = CreateView(False)
            Require view.cboMode.ListIndex = 0, "Range entry must default to insert"
            Require view.cboSizeMode.Value = "FIT", "insert default size mode"
            Require view.chkMoveAndSize.Value, "fit default placement"
            Require Not view.cmdExecute.Enabled, "entry must not be executable"
            view.cmdPreview.Value = True
            Require Not view.cmdExecute.Enabled, "missing files must not seal insert"
            Require before = SourceState(book), "insert entry/failed preview changed source"
            view.cmdCancel.Value = True: Set view = Nothing
            Require Application.Selection.Address(External:=True) = target.Address(External:=True), "cancel lost original range selection"
            Set view = CreateView(True)
            Require view.cboMode.ListIndex = 1, "internal FIT on a range must default to fit"
            Require Not view.cmdExecute.Enabled, "FIT entry without a selected picture must not be executable"
        Case "shape_default"
            picture.Select
            before = SourceState(book)
            Set view = CreateView(False)
            Require view.cboMode.ListIndex = 1, "Shape entry through INSERT must default to fit"
            Require InStr(view.txtFiles.Value, picture.Name) > 0, "entry lost selected picture"
            Require before = SourceState(book), "shape entry changed picture or Saved"
            view.cmdCancel.Value = True: Set view = Nothing
            picture.Select
            Set view = CreateView(True)
            Require view.cboMode.ListIndex = 1, "internal FIT must use the same popup"
        Case "multi_shape_default"
            Set added = book.Worksheets(1).Shapes.AddPicture(goodPath, msoFalse, msoTrue, 280, 120, 80, 40)
            added.Name = "SecondExistingPicture"
            For Each savedState In Array(True, False)
                target.Cells(1, 1).Select
                entryTarget = ActiveCell.MergeArea.Address
                book.Worksheets(1).Shapes.Range(Array(picture.Name, added.Name)).Select
                book.Saved = CBool(savedState)
                before = SourceState(book)
                Set view = CreateView(False)
                Require view.cboMode.ListIndex = 0, "Multiple picture entry must allow new insertion"
                Require view.cmdBrowse.Visible, "Multiple picture entry hides file picker"
                Require Not view.cmdExecute.Enabled, "Multiple picture entry reused a preview"
                Require InStr(view.txtRange.Value, entryTarget) > 0, "Insertion target is not original active cell"
                Require before = SourceState(book), "Multiple picture entry mutated source or Saved"
                view.cboMode.ListIndex = 1
                view.cmdUseSelection.Value = True
                Require Left$(CStr(view.txtPreview.Value), 3) = "오류:", "Multiple pictures were accepted by FIT selection"
                Require InStr(view.txtPreview.Value, "하나") > 0, "Missing single-picture selection explanation"
                view.cmdPreview.Value = True
                Require Not view.cmdExecute.Enabled, "Multiple pictures bypassed single-picture fit guard"
                Require InStr(view.txtPreview.Value, "[선택 그림 가져오기]") > 0, "Invalid FIT selection was retained for preview"
                Require before = SourceState(book), "Rejected multi-picture FIT preview mutated source or Saved"
                view.cboMode.ListIndex = 0
                view.cmdCancel.Value = True: Set view = Nothing
                Set selectedPictures = Application.Selection.ShapeRange
                Require selectedPictures.Count = 2, "Cancel lost multiple picture selection"
                Require selectedPictures.Item(1).ID = picture.ID And selectedPictures.Item(2).ID = added.ID, "Cancel selected different pictures"
                Require before = SourceState(book), "Cancel mutated source or Saved"
            Next savedState
        Case "mode_cancel"
            For Each savedState In Array(True, False)
                book.Saved = CBool(savedState): picture.Select
                before = SourceState(book)
                Set view = CreateView(False)
                view.txtRange.Value = target.Address
                view.txtMargin.Value = "2": view.cmdPreview.Value = True
                Require view.cmdExecute.Enabled, "fit preview not ready"
                view.cboMode.ListIndex = 0
                Require Not view.cmdExecute.Enabled, "mode switch retained a seal"
                view.cboMode.ListIndex = 1
                Require Not view.cmdExecute.Enabled, "returning mode reused a seal"
                view.cmdPreview.Value = True
                view.chkMoveAndSize.Value = False
                Require Not view.cmdExecute.Enabled, "placement change retained a seal"
                view.cmdPreview.Value = True
                view.txtMargin.Value = "3"
                Require Not view.cmdExecute.Enabled, "margin change retained a seal"
                view.cmdPreview.Value = True
                view.txtRange.Value = "B2:C3"
                Require Not view.cmdExecute.Enabled, "range change retained a seal"
                view.cmdCancel.Value = True: Set view = Nothing
                Require before = SourceState(book), "mode/option change or cancel changed source/Saved"
                Require TypeName(Application.Selection) <> "Range", "cancel lost original picture selection"
                Require Application.Selection.Name = picture.Name, "cancel selected another picture"
            Next savedState
        Case "fit_preview"
            Set request = New CNxDrawingRequest
            request.ConfigurePicture picture, target, 2, True
            Set context = NxContextFactory.CaptureCurrent()
            request.RequireContext context
            picture.Select: Set view = CreateView(False)
            view.txtRange.Value = target.Address: view.txtMargin.Value = "2"
            before = SourceState(book)
            view.cmdPreview.Value = True
            Require view.cmdExecute.Enabled, "fit preview not sealed: " & view.txtPreview.Value
            Require before = SourceState(book), "fit preview mutated source"
            view.cmdCancel.Value = True: Set view = Nothing
            target.Select
            NxDrawFitPicture picture, target, 2, True
            Require picture.Placement = xlMoveAndSize, "fit did not apply moveAndSize"
            Require picture.LockAspectRatio = msoTrue, "fit did not lock aspect"
            Require picture.Width <= target.Width - 4.0001 + 0.001, "fit width outside margin"
            Require picture.Height <= target.Height - 4.0001 + 0.001, "fit height outside margin"
            Require Abs((picture.Left + picture.Width / 2) - (target.Left + target.Width / 2)) < 0.001, "fit not horizontally centered"
            Require Abs((picture.Top + picture.Height / 2) - (target.Top + target.Height / 2)) < 0.001, "fit not vertically centered"
            picture.Placement = xlFreeFloating
            NxDrawFitPicture picture, target, 2, False
            Require picture.Placement = xlFreeFloating, "unchecked moveAndSize changed original placement"
        Case "fit_drift"
            picture.Select: Set view = CreateView(False)
            view.txtRange.Value = target.Address: view.txtMargin.Value = "2"
            view.cmdPreview.Value = True
            Require view.cmdExecute.Enabled, "initial fit preview missing"
            picture.Width = picture.Width + 7
            before = SourceState(book)
            view.cmdExecute.Value = True
            Require Not view.cmdExecute.Enabled, "changed picture reused sealed fit request"
            Require before = SourceState(book), "stale fit request changed picture"
            view.cmdPreview.Value = True
            Require view.cmdExecute.Enabled, "changed picture could not be previewed again"
            book.Worksheets(1).Range("E1").Select
            before = SourceState(book)
            view.cmdExecute.Value = True
            Require Not view.cmdExecute.Enabled, "changed selection reused sealed fit request"
            Require before = SourceState(book), "selection drift changed source"
            view.cmdPreview.Value = True
            Require view.cmdExecute.Enabled, "fit preview did not recover"
            Set otherBook = Workbooks.Add(xlWBATWorksheet)
            before = SourceState(book)
            view.cmdExecute.Value = True
            Require Not view.cmdExecute.Enabled, "different workbook reused fit request"
            Require before = SourceState(book), "workbook drift changed original"
        Case "insert_preview"
            Set files = New Collection: files.Add goodPath
            Set added = book.Worksheets(1).Shapes.AddPicture(goodPath, msoFalse, msoTrue, 0, 0, -1, -1)
            naturalWidth = added.Width: naturalHeight = added.Height
            added.Delete: Set added = Nothing
            For Each mode In Array("FIT", "ORIGINAL")
                target.Select
                Set insertRequest = NxPictureInsertCreateRequest(target, files, CStr(mode))
                before = SourceState(book)
                Set preview = NxPictureInsertBuildPreview(insertRequest)
                Require preview.IsSealed And preview.InsertCount = 1, "insert preview count/seal"
                Require preview.Revalidate(insertRequest), "fresh insert preview invalid"
                Require before = SourceState(book), "insert preview changed source/Saved"
                oldRowHeight = target.Rows(1).RowHeight
                target.Rows(1).RowHeight = oldRowHeight + 1
                rejected = False
                On Error Resume Next
                rejected = Not preview.Revalidate(insertRequest)
                If Err.Number <> 0 Then rejected = True
                Err.Clear
                On Error GoTo Failed
                Require rejected, "insert target drift was accepted"
                target.Rows(1).RowHeight = oldRowHeight
                Set insertRequest = NxPictureInsertCreateRequest(target, files, CStr(mode))
                Set preview = NxPictureInsertBuildPreview(insertRequest)
                Set journal = New CNxPictureInsertJournal
                before = PictureState(picture)
                Set result = NxPictureInsertExecute(insertRequest, preview, journal)
                Require result.Outcome = NxSuccess, "insert engine result"
                Require book.Worksheets(1).Shapes.Count = countBefore + 1, "insert engine shape count"
                Set added = book.Worksheets(1).Shapes(CStr(preview.ShapeNames(1)))
                If CStr(mode) = "FIT" Then
                    Require added.Placement = xlMoveAndSize, "FIT insert placement"
                    scaleFactor = WorksheetFunction.Min(target.Cells(1, 1).Width / naturalWidth, target.Cells(1, 1).Height / naturalHeight)
                    Require Abs(added.Width - naturalWidth * scaleFactor) < 0.001, "FIT insert width"
                    Require Abs(added.Height - naturalHeight * scaleFactor) < 0.001, "FIT insert height"
                    Require Abs((added.Left + added.Width / 2) - (target.Left + target.Cells(1, 1).Width / 2)) < 0.001, "FIT insert horizontal center"
                    Require Abs((added.Top + added.Height / 2) - (target.Top + target.Cells(1, 1).Height / 2)) < 0.001, "FIT insert vertical center"
                Else
                    Require added.Placement = xlMove, "ORIGINAL insert placement"
                    Require Abs(added.Width - naturalWidth) < 0.001, "ORIGINAL changed natural width"
                    Require Abs(added.Height - naturalHeight) < 0.001, "ORIGINAL changed natural height"
                    Require Abs(added.Left - target.Left) < 0.001 And Abs(added.Top - target.Top) < 0.001, "ORIGINAL changed cell anchor"
                End If
                journal.Rollback
                Require book.Worksheets(1).Shapes.Count = countBefore, "insert rollback left created shape"
                Require before = PictureState(picture), "insert/rollback changed existing picture"
            Next mode
        Case "engine_rollback"
            TestSupportedPictureHeaders badPath
            RequireZeroBitmap badPath
            Set files = New Collection: files.Add goodPath: files.Add badPath
            Set target = book.Worksheets(1).Range("B2:B3"): target.Select
            Set insertRequest = NxPictureInsertCreateRequest(target, files, "FIT")
            Set files = insertRequest.Files
            Require StrComp(CStr(files(1)), goodPath, vbTextCompare) = 0, "first insert source is not the valid fixture"
            Require StrComp(CStr(files(2)), badPath, vbTextCompare) = 0, "second insert source is not the malformed fixture"
            Set preview = NxPictureInsertBuildPreview(insertRequest)
            Require preview.InsertCount = 2, "rollback fixture must attempt two pictures"
            Set journal = New CNxPictureInsertJournal
            before = PictureState(picture): rejected = False
            On Error Resume Next
            Set result = NxPictureInsertExecute(insertRequest, preview, journal)
            failureNumber = Err.Number: failureSource = Err.Source: failureDescription = Err.Description
            rejected = (Err.Number <> 0)
            Err.Clear
            On Error GoTo Failed
            Require rejected, "malformed second image did not fail" & PictureInsertDiagnostic(result, journal, book)
            Require failureNumber = NX_CONTRACT_ERROR, "malformed header lost original contract error"
            Require failureSource = "LHexcel.Core", "malformed header lost original error source"
            Require failureDescription = "지원하지 않거나 손상된 그림 파일 헤더입니다.", "malformed header lost original error description"
            Require journal.CreatedNames.Count = 1, "rollback fixture did not insert first image before failure"
            Require book.Worksheets(1).Shapes.Count = countBefore, "partial insert rollback left a new shape"
            Require before = PictureState(picture), "partial rollback changed existing picture"
            ' The disposable native runner injects a callback after fit mutates Width.
            ' Arm only now: insertion also calls fit and must reach its malformed second file.
            mFitFaultBefore = before: mFitFaultReached = False: mFitFaultMutated = False
            mFitFaultArmed = True
            On Error Resume Next
            NxDrawFitPicture picture, target, 2, True
            failureNumber = Err.Number: failureSource = Err.Source: failureDescription = Err.Description
            Err.Clear
            On Error GoTo Failed
            mFitFaultArmed = False
            Require mFitFaultReached And mFitFaultMutated, "fit rollback fault point was not reached after mutation"
            Require failureNumber = FIT_FAULT_NUMBER, "fit rollback lost original error number"
            Require failureSource = FIT_FAULT_SOURCE, "fit rollback lost original error source"
            Require failureDescription = FIT_FAULT_DESCRIPTION, "fit rollback lost original error description"
            Require before = PictureState(picture), "fit rollback did not restore original picture properties"
            Require book.Worksheets(1).Shapes.Count = countBefore, "fit rollback changed shape count"
            ' Invalid margin is a separate preflight-preservation check, not rollback evidence.
            rejected = False
            On Error Resume Next
            NxDrawFitPicture picture, target, target.Width, True
            rejected = (Err.Number <> 0)
            Err.Clear
            On Error GoTo Failed
            Require rejected, "invalid fit margin accepted"
            Require before = PictureState(picture), "failed fit changed original picture"
        Case Else: NxRaiseContractError "Unknown r62 picture mode case"
    End Select
    NxR62PictureModesCase = "PASS|" & kind & "|native controls/engines; physical file picker NOT_RUN" & caseDetail
    GoTo CleanUp
Failed:
    NxR62PictureModesCase = "FAIL|" & kind & "|" & Err.Number & "|" & Err.Description & caseDetail
CleanUp:
    On Error Resume Next
    mFitFaultArmed = False
    If Not view Is Nothing Then Unload view
    Set view = Nothing
    If Not otherBook Is Nothing Then otherBook.Close SaveChanges:=False
    If Not book Is Nothing Then book.Close SaveChanges:=False
    ' Only the two files and unique folder created by this fixture are removed.
    If Len(goodPath) > 0 Then Kill goodPath
    If Len(badPath) > 0 Then Kill badPath
    If Len(folder) > 0 Then RmDir folder
    On Error GoTo 0
End Function

Private Function PickerFormMetrics(ByVal view As FNxPictureInsert) As String
    Dim name As Variant, control As Object, detail As String
    detail = "|form_loaded_not_shown=True|modal_file_picker=NOT_RUN|form_zoom=" & CStr(view.Zoom)
    For Each name In Array("lblTitle", "cboMode", "cmdBrowse")
        Set control = view.Controls(CStr(name))
        detail = detail & "|" & CStr(name) & "_font=" & control.Font.Name & _
            "|" & CStr(name) & "_font_size=" & CStr(control.Font.Size) & "|" & CStr(name) & "_height=" & CStr(control.Height)
    Next name
    PickerFormMetrics = detail
End Function

Private Sub TestPickerState(ByVal book As Workbook, ByVal target As Range, ByVal picture As Shape, ByRef detail As String)
    Dim sheet As Worksheet, otherSheet As Worksheet, captured As Object, current As Object, multiple As Range
    Dim secondPicture As Shape
    Dim applicationState As Variant, saved As Variant, prefix As String, before As String, failures As String
    Dim matched As Boolean, restored As Boolean, sameObject As Boolean, sameAddress As Boolean
    Dim failureNumber As Long, failureDescription As String
    applicationState = Array(Application.ScreenUpdating, Application.EnableEvents, Application.DisplayAlerts, Application.Calculation)
    On Error GoTo Failed
    Set sheet = target.Worksheet
    Set secondPicture = picture.Duplicate
    secondPicture.Name = "SecondPickerFixture"
    Set otherSheet = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
    otherSheet.Range("B2").Value2 = "other-sheet-preserve"
    For Each saved In Array(True, False)
        sheet.Activate: target.Select: book.Saved = CBool(saved)
        prefix = "range_saved_" & CStr(saved)
        Set captured = Application.Selection: Set current = Application.Selection
        sameObject = (current Is captured)
        sameAddress = (current.Address(External:=True) = captured.Address(External:=True))
        detail = detail & "|" & prefix & "_identity=" & CStr(sameObject) & "|" & prefix & "_same_address=" & CStr(sameAddress)
        before = SourceState(book)
        matched = PickerStateMatches(sheet, captured, applicationState)
        PickerCheck matched, prefix & "_matches", detail, failures
        restored = PickerStateRestore(sheet, captured, applicationState)
        PickerCheck restored, prefix & "_restores", detail, failures
        PickerCheck SourceState(book) = before, prefix & "_source_saved_shapes_preserved", detail, failures
        PickerCheck PickerRangeEquals(captured), prefix & "_selection_preserved", detail, failures
        PickerCheck target.Cells(1, 1).Value2 = 5 And sheet.Range("C4").Value2 = "preserve", prefix & "_values_preserved", detail, failures

        sheet.Range("E1").Select
        PickerCheck Not PickerStateMatches(sheet, captured, applicationState), prefix & "_reject_other_range", detail, failures
        book.Saved = CBool(saved): before = SourceState(book)
        restored = PickerStateRestore(sheet, captured, applicationState)
        PickerCheck restored And PickerRangeEquals(captured), prefix & "_restore_other_range", detail, failures
        PickerCheck SourceState(book) = before, prefix & "_other_range_source_preserved", detail, failures

        otherSheet.Activate: otherSheet.Range(target.Address).Select
        PickerCheck Not PickerStateMatches(sheet, captured, applicationState), prefix & "_reject_other_sheet", detail, failures
        book.Saved = CBool(saved): before = SourceState(book)
        restored = PickerStateRestore(sheet, captured, applicationState)
        PickerCheck restored And PickerRangeEquals(captured), prefix & "_restore_other_sheet", detail, failures
        PickerCheck SourceState(book) = before And otherSheet.Range("B2").Value2 = "other-sheet-preserve", prefix & "_other_sheet_source_preserved", detail, failures

        sheet.Activate
        Set multiple = Application.Union(sheet.Range("B2"), sheet.Range("D4"))
        multiple.Select: Set captured = Application.Selection: book.Saved = CBool(saved)
        before = SourceState(book)
        PickerCheck captured.Areas.Count = 2, "multi_area_fixture_" & CStr(saved), detail, failures
        PickerCheck PickerStateMatches(sheet, captured, applicationState), "multi_area_matches_" & CStr(saved), detail, failures
        restored = PickerStateRestore(sheet, captured, applicationState)
        PickerCheck restored And PickerRangeEquals(captured), "multi_area_restores_" & CStr(saved), detail, failures
        PickerCheck SourceState(book) = before, "multi_area_source_preserved_" & CStr(saved), detail, failures

        picture.Select: Set captured = Application.Selection: Set current = Application.Selection
        sameObject = (current Is captured)
        detail = detail & "|shape_type=" & TypeName(captured) & "|shape_identity=" & CStr(sameObject)
        book.Saved = CBool(saved): before = SourceState(book)
        PickerCheck PickerStateMatches(sheet, captured, applicationState), "shape_matches_" & CStr(saved), detail, failures
        restored = PickerStateRestore(sheet, captured, applicationState)
        PickerCheck restored, "shape_restores_" & CStr(saved), detail, failures
        PickerCheck PickerShapeEquals(picture), "shape_selection_preserved_" & CStr(saved), detail, failures
        PickerCheck SourceState(book) = before, "shape_source_preserved_" & CStr(saved), detail, failures
        picture.Select
        secondPicture.Select Replace:=False
        Set captured = Application.Selection
        detail = detail & "|multi_shape_type=" & TypeName(captured)
        book.Saved = CBool(saved): before = SourceState(book)
        PickerCheck PickerStateMatches(sheet, captured, applicationState), "multi_shape_matches_" & CStr(saved), detail, failures
        target.Select
        restored = PickerStateRestore(sheet, captured, applicationState)
        PickerCheck restored, "multi_shape_restores_" & CStr(saved), detail, failures
        Set current = Application.Selection.ShapeRange
        PickerCheck current.Count = 2, "multi_shape_count_" & CStr(saved), detail, failures
        PickerCheck current.Item(1).ID = picture.ID And current.Item(2).ID = secondPicture.ID, "multi_shape_ids_" & CStr(saved), detail, failures
        PickerCheck SourceState(book) = before, "multi_shape_source_preserved_" & CStr(saved), detail, failures
    Next saved
    Require Len(failures) = 0, "picker state assertions failed: " & failures
    GoTo CleanUp
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description
CleanUp:
    On Error Resume Next
    If Application.ScreenUpdating <> CBool(applicationState(0)) Then Application.ScreenUpdating = CBool(applicationState(0))
    If Application.EnableEvents <> CBool(applicationState(1)) Then Application.EnableEvents = CBool(applicationState(1))
    If Application.DisplayAlerts <> CBool(applicationState(2)) Then Application.DisplayAlerts = CBool(applicationState(2))
    If Application.Calculation <> CLng(applicationState(3)) Then Application.Calculation = CLng(applicationState(3))
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_R62PictureModes.PickerState", failureDescription
End Sub

Private Function PickerStateMatches(ByVal sheet As Worksheet, ByVal selected As Object, ByVal state As Variant) As Boolean
    PickerStateMatches = NxPictureInsertController.PictureInsertPickerStateMatches(sheet.Parent, sheet, selected, CBool(state(0)), CBool(state(1)), CBool(state(2)), CLng(state(3)))
End Function

Private Function PickerStateRestore(ByVal sheet As Worksheet, ByVal selected As Object, ByVal state As Variant) As Boolean
    PickerStateRestore = NxPictureInsertController.RestorePictureInsertPickerState(sheet.Parent, sheet, selected, CBool(state(0)), CBool(state(1)), CBool(state(2)), CLng(state(3)))
End Function

Private Function PickerRangeEquals(ByVal expected As Object) As Boolean
    Dim current As Object
    On Error GoTo Different
    Set current = Application.Selection
    If TypeName(current) <> "Range" Or TypeName(expected) <> "Range" Then Exit Function
    If Not current.Worksheet Is expected.Worksheet Then Exit Function
    PickerRangeEquals = (current.Address(External:=True) = expected.Address(External:=True))
Different:
End Function

Private Function PickerShapeEquals(ByVal expected As Shape) As Boolean
    Dim current As Object
    On Error GoTo Different
    Set current = Application.Selection
    If TypeName(current) = "Range" Then Exit Function
    If Not Application.ActiveSheet Is expected.Parent Then Exit Function
    If TypeName(current) = "ShapeRange" Then
        If current.Count <> 1 Then Exit Function
        PickerShapeEquals = (current.Item(1).Name = expected.Name)
    Else
        PickerShapeEquals = (current.Name = expected.Name)
    End If
Different:
End Function

Private Sub PickerCheck(ByVal passed As Boolean, ByVal name As String, ByRef detail As String, ByRef failures As String)
    detail = detail & "|" & name & "=" & CStr(passed)
    If Not passed Then failures = failures & name & ";"
End Sub

Public Sub NxR62PictureFitFaultPoint(ByVal pictureShape As Shape)
    If Not mFitFaultArmed Then Exit Sub
    mFitFaultArmed = False
    mFitFaultReached = True
    mFitFaultMutated = (PictureState(pictureShape) <> mFitFaultBefore)
    Err.Raise FIT_FAULT_NUMBER, FIT_FAULT_SOURCE, FIT_FAULT_DESCRIPTION
End Sub

Private Function CreateView(ByVal internalFit As Boolean) As FNxPictureInsert
    Dim view As New FNxPictureInsert, context As New CNxFeatureDialogContext
    If internalFit Then
        context.Configure "NX-DRAW-FIT-PICTURE", "NX-DLG-DRAW-INSERT-PICTURE", "picture-fit", False
    Else
        context.Configure "NX-DRAW-INSERT-PICTURE", "NX-DLG-DRAW-INSERT-PICTURE", "picture-insert", False
    End If
    view.BindFeatureContext context
    Set CreateView = view
End Function

Private Function PictureState(ByVal picture As Shape) As String
    PictureState = picture.Name & "|" & CStr(ObjPtr(picture)) & "|" & picture.Left & "|" & picture.Top & "|" & picture.Width & "|" & picture.Height & "|" & picture.Placement & "|" & picture.LockAspectRatio
End Function

Private Function SourceState(ByVal book As Workbook) As String
    Dim state As String, cell As Range, picture As Shape
    state = CStr(book.Saved) & "|" & book.Worksheets(1).Shapes.Count
    For Each cell In book.Worksheets(1).Range("B2:C4")
        state = state & "|" & cell.Address & "|" & CStr(cell.Formula) & "|" & cell.NumberFormat & "|" & cell.RowHeight & "|" & cell.ColumnWidth
    Next cell
    For Each picture In book.Worksheets(1).Shapes
        state = state & "|" & PictureState(picture)
    Next picture
    SourceState = state
End Function

Private Sub WriteBitmap(ByVal path As String, ByVal valid As Boolean)
    Dim bytes(0 To 57) As Byte, handle As Integer
    If valid Then
        bytes(0) = 66: bytes(1) = 77: bytes(2) = 58: bytes(10) = 54
        bytes(14) = 40: bytes(18) = 1: bytes(22) = 1: bytes(26) = 1: bytes(28) = 24
        bytes(34) = 4: bytes(54) = 64: bytes(55) = 128: bytes(56) = 192
    End If
    handle = FreeFile
    Open path For Binary Access Write As #handle
    Put #handle, , bytes
    Close #handle
End Sub

Private Sub RequireZeroBitmap(ByVal path As String)
    Dim bytes(0 To 57) As Byte, handle As Integer, index As Long
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    handle = FreeFile
    Open path For Binary Access Read Lock Write As #handle
    Require LOF(handle) = 58, "malformed bitmap fixture length changed"
    Get #handle, , bytes
    Close #handle
    handle = 0
    For index = 0 To 57
        Require bytes(index) = 0, "malformed bitmap fixture contains nonzero bytes"
    Next index
    Exit Sub
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description
    On Error Resume Next
    If handle > 0 Then Close #handle
    On Error GoTo 0
    Err.Raise failureNumber, "T_R62PictureModes.RequireZeroBitmap", failureDescription
End Sub

Private Sub TestSupportedPictureHeaders(ByVal path As String)
    Dim headers As Variant, header As Variant
    ' Header-only fixtures exercise the guard, not Excel decoding of these formats.
    headers = Array("89504E470D0A1A0A", "FFD8FFE0", "FFD8FFE1", "FFD8FFDB", "474946383761", "474946383961", _
        "424D3A0000000000000036000000", "49492A0008000000", "4D4D002A00000008", _
        "49492B00080000001000000000000000", "4D4D002B000800000000000000000010")
    For Each header In headers
        WriteHeaderFixture path, CStr(header)
        NxPictureInsertRequireSupportedHeader path
        RequireHeaderFileUnlocked path
        WriteHeaderFixture path, Left$(CStr(header), Len(CStr(header)) - 2)
        RequireHeaderRejected path
        RequireHeaderFileUnlocked path
    Next header
    For Each header In Array(vbNullString, "0000000000000000", "6E6F7420616E20696D616765", "474946303061", "49492B00040000000000000000000000")
        WriteHeaderFixture path, CStr(header)
        RequireHeaderRejected path
        RequireHeaderFileUnlocked path
    Next header
    WriteBitmap path, False
End Sub

Private Sub WriteHeaderFixture(ByVal path As String, ByVal hexText As String)
    Dim bytes() As Byte, index As Long, handle As Integer
    Dim opened As Boolean, failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    handle = FreeFile
    Open path For Output Access Write Lock Read Write As #handle
    opened = True
    Close #handle
    opened = False
    If Len(hexText) = 0 Then Exit Sub
    ReDim bytes(0 To Len(hexText) \ 2 - 1)
    For index = 0 To UBound(bytes)
        bytes(index) = CByte("&H" & Mid$(hexText, index * 2 + 1, 2))
    Next index
    Open path For Binary Access Write Lock Read Write As #handle
    opened = True
    Put #handle, , bytes
    Close #handle
    Exit Sub
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description
    On Error Resume Next
    If opened Then Close #handle
    On Error GoTo 0
    Err.Raise failureNumber, "T_R62PictureModes.WriteHeader", failureDescription
End Sub

Private Sub RequireHeaderRejected(ByVal path As String)
    Dim failureNumber As Long, failureSource As String
    On Error Resume Next
    NxPictureInsertRequireSupportedHeader path
    failureNumber = Err.Number: failureSource = Err.Source
    Err.Clear
    On Error GoTo 0
    Require failureNumber = NX_CONTRACT_ERROR, "unknown/truncated header was not rejected"
    Require failureSource = "LHexcel.Core", "header guard lost its contract error source"
End Sub

Private Sub RequireHeaderFileUnlocked(ByVal path As String)
    Dim handle As Integer
    handle = FreeFile
    Open path For Binary Access Read Write Lock Read Write As #handle
    Close #handle
End Sub

Private Function PictureInsertDiagnostic(ByVal result As CNxResult, ByVal journal As CNxPictureInsertJournal, ByVal book As Workbook) As String
    PictureInsertDiagnostic = "|zero_bitmap_bytes=58|created=" & CStr(journal.CreatedNames.Count) & _
        "|shapes=" & CStr(book.Worksheets(1).Shapes.Count)
    If result Is Nothing Then
        PictureInsertDiagnostic = PictureInsertDiagnostic & "|result=nothing"
    Else
        PictureInsertDiagnostic = PictureInsertDiagnostic & "|outcome=" & CStr(result.Outcome) & _
            "|stage=" & result.Stage & "|message=" & result.MessageKey
    End If
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then NxRaiseContractError detail
End Sub
