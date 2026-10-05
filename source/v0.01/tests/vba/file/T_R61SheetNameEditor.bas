Attribute VB_Name = "T_R61SheetNameEditor"
Option Explicit

' Run only from the Windows BAT-owned disposable Product test copy.
' rollback cases require the existing NX_R58_TEST_BUILD rename failure seam.
Public Function NxR61SheetNameEditorRunCase(ByVal kind As String) As String
    Dim book As Workbook, otherBook As Workbook, firstSheet As Worksheet, secondSheet As Worksheet
    Dim chartSheet As Chart, replacement As Worksheet, view As FNxBatchRename
    Dim preview As Collection, draft As Collection, row As Variant, invalidName As Variant
    Dim hasDuplicate As Boolean, hasCollision As Boolean, failureDetail As String, changed As Long
    Dim savedAlerts As Boolean, stage As Variant
    Dim index As Long, priorTop As Long
    savedAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set firstSheet = book.Worksheets(1): firstSheet.Name = "Alpha"
    Set secondSheet = book.Worksheets.Add(After:=firstSheet): secondSheet.Name = "Beta"
    firstSheet.Range("A1").Value2 = "first-original"
    secondSheet.Range("A1").Value2 = "second-original"
    book.Activate
    Select Case kind
        Case "validation"
            book.Saved = True
            Set preview = BasePreview(book)
            Require CommitRejected(book, preview), "unchanged plan must not commit"
            For Each invalidName In Array("", "   ", String$(32, "x"), "bad:name", "bad\name", "bad/name", _
                "bad?name", "bad*name", "bad[name", "bad]name", "'start", "end'", "History", "bad" & vbTab & "name")
                Set draft = NxSheetBatchRenameEditPreview(book, preview, 1, CStr(invalidName), hasDuplicate, hasCollision)
                row = draft(1)
                Require Left$(CStr(row(2)), 3) = "오류:", "invalid sheet name accepted"
                Require CommitRejected(book, draft), "invalid plan committed"
            Next invalidName
            Set draft = NxSheetBatchRenameEditPreview(book, preview, 1, String$(31, "x"), hasDuplicate, hasCollision)
            NxSheetBatchRenameRequireReady book, draft
            Set draft = NxSheetBatchRenameEditPreview(book, preview, 1, "bETA", hasDuplicate, hasCollision)
            Require hasDuplicate, "case-insensitive duplicate not marked"
            Require CommitRejected(book, draft), "duplicate name committed"
            Set draft = NxSheetBatchRenameEditPreview(book, preview, 1, "Gamma", hasDuplicate, hasCollision)
            Require Not hasDuplicate And Not hasCollision, "stale validation flags remained"
            row = preview(1): Require CStr(row(1)) = "Alpha", "edit mutated the previous confirmed collection"
            Require firstSheet.Name = "Alpha" And secondSheet.Name = "Beta" And book.Saved, "draft changed workbook or Saved"
            Set chartSheet = book.Charts.Add(After:=secondSheet): chartSheet.Name = "ChartOnly"
            Set preview = BasePreview(book)
            Set draft = NxSheetBatchRenameEditPreview(book, preview, 1, "chartonly", hasDuplicate, hasCollision)
            Require hasCollision, "chart sheet collision not marked"
            Require CommitRejected(book, draft), "chart sheet collision committed"
        Case "cycles"
            TestSavedWorkbookApprovedResult
            book.Activate
            firstSheet.Activate
            firstSheet.Range("A1").Select
            Set preview = CyclePreview(book)
            changed = NxSheetBatchRenameCommit(book, preview)
            Require changed = 2, "cycle change count"
            Require firstSheet.Name = "Beta" And secondSheet.Name = "Alpha", "cycle names did not follow original objects"
            Set preview = BasePreview(book)
            Set draft = NxSheetBatchRenameEditPreview(book, preview, 1, "bETA", hasDuplicate, hasCollision)
            Require NxSheetBatchRenameCommit(book, draft) = 1 And firstSheet.Name = "bETA", "case-only rename failed"
            Require firstSheet.Range("A1").Value2 = "first-original" And secondSheet.Range("A1").Value2 = "second-original", "cycle changed cell values"
        Case "workbook_drift"
            Set preview = CyclePreview(book)
            Set otherBook = Workbooks.Add(xlWBATWorksheet)
            otherBook.Worksheets(1).Name = "Alpha"
            otherBook.Worksheets.Add(After:=otherBook.Worksheets(1)).Name = "Beta"
            Require CommitRejected(otherBook, preview), "same-named sheets in another workbook accepted"
            Require book.Worksheets(1).Name = "Alpha" And otherBook.Worksheets(1).Name = "Alpha", "workbook switch changed source"
        Case "sheet_drift"
            Set preview = CyclePreview(book)
            firstSheet.Name = "ChangedOutside"
            Require CommitRejected(book, preview), "external rename accepted"
            firstSheet.Name = "Alpha"
            Set preview = CyclePreview(book)
            secondSheet.Visible = xlSheetHidden
            Require CommitRejected(book, preview), "visibility drift accepted"
            secondSheet.Visible = xlSheetVisible
            Set preview = CyclePreview(book)
            Set chartSheet = book.Charts.Add(After:=secondSheet): chartSheet.Name = "ExtraChart"
            Require CommitRejected(book, preview), "chart-sheet addition accepted"
            Application.DisplayAlerts = False: chartSheet.Delete: Set chartSheet = Nothing
            Application.DisplayAlerts = savedAlerts
            Set preview = CyclePreview(book)
            Set replacement = book.Worksheets.Add(After:=secondSheet)
            Application.DisplayAlerts = False: firstSheet.Delete
            Application.DisplayAlerts = savedAlerts
            replacement.Name = "Alpha": replacement.Move Before:=secondSheet
            Require CommitRejected(book, preview), "same-name replacement sheet accepted"
            Require replacement.Name = "Alpha" And secondSheet.Name = "Beta", "replacement drift changed names"
        Case "rollback"
            For Each stage In Array("sheet-original-to-temporary", "sheet-temporary-to-target")
                Set preview = CyclePreview(book)
                NxBatchRenameTestInjectMoveFailure CStr(stage), 2
                On Error Resume Next
                changed = NxSheetBatchRenameCommit(book, preview)
                failureDetail = Err.Description
                Err.Clear
                On Error GoTo Failed
                Require InStr(1, failureDetail, "Injected rename move failure", vbBinaryCompare) > 0, "rollback fixture did not reach injected move"
                Require firstSheet.Name = "Alpha" And secondSheet.Name = "Beta", "rollback did not restore original object names"
                Require firstSheet.Range("A1").Value2 = "first-original" And secondSheet.Range("A1").Value2 = "second-original", "rollback changed cell values"
                NxBatchRenameTestClearMoveFailure
            Next stage
        Case "editor_cancel"
            book.Saved = True
            TestOrderedRuleDraftPreservesSource book
            Set view = New FNxBatchRename
            view.BindFeature NX_FEATURE_FILE_SHEET_BATCH_RENAME
            Require view.lstSheets.ListCount = 2 And Not view.cmdExecute.Enabled, "empty rule must offer editing without execution"
            Require view.cboRuleKind.ListCount = 7, "sheet rule editor must exclude file-only extension"
            Require view.lstRules.ListCount = 0, "new editor retained rules from a prior session"
            view.lstSheets.ListIndex = 0
            view.txtPlannedName.Value = "x_Alpha"
            view.cmdPreview.Value = True
            Require view.lstSheets.ListCount = 2 And view.cmdExecute.Enabled, "manual preview did not seal"
            view.lstSheets.ListIndex = 1
            view.txtPlannedName.Value = "ManualBeta"
            Require Not view.cmdExecute.Enabled, "editing did not invalidate the seal"
            Require CStr(view.lstSheets.List(1, 1)) = "ManualBeta", "selected row draft not updated"
            Require CStr(view.lstSheets.List(0, 1)) = "x_Alpha", "editing changed another row"
            view.cmdPreview.Value = True
            Require view.cmdExecute.Enabled And CStr(view.lstSheets.List(1, 1)) = "ManualBeta", "preview lost direct edit or did not reseal"
            view.cmdItemUp.Value = True
            Require Not view.cmdExecute.Enabled And view.lstSheets.ListCount = 2, "item order change left a stale seal"
            Require CStr(view.lstSheets.List(0, 0)) = "Beta" And CStr(view.lstSheets.List(0, 1)) = "ManualBeta", "item reorder lost object-bound manual name"
            Require CStr(view.lstSheets.List(1, 0)) = "Alpha" And CStr(view.lstSheets.List(1, 1)) = "x_Alpha", "item reorder changed another manual name"
            view.cmdPreview.Value = True
            Require view.cmdExecute.Enabled, "reordered preview did not reseal"
            view.cmdResetName.Value = True
            Require Not view.cmdExecute.Enabled And CStr(view.lstSheets.List(0, 1)) = "Beta", "manual reset left a stale seal or wrong draft"
            Require CStr(view.lstSheets.List(1, 1)) = "x_Alpha", "manual reset changed another row"
            view.cmdCancel.Value = True
            Set view = Nothing
            Require LoadedRenameFormCount() = 0, "cancel left a loaded rename form"
            Require firstSheet.Name = "Alpha" And secondSheet.Name = "Beta" And book.Saved, "cancel mutated source"
        Case "editor_lifecycle"
            Require Not NxFormWheelIsAttached(), "fixture requires no existing wheel owner"
            Set view = New FNxBatchRename
            view.BindFeature NX_FEATURE_FILE_SHEET_BATCH_RENAME
            view.Show vbModeless
            DoEvents
            Require NxFormWheelIsAttached(), "wheel not attached"
            Require view.cmdPreview.Default And view.cmdCancel.Cancel, "Enter/Esc contract"
            Require view.lstSheets.TabIndex < view.txtPlannedName.TabIndex And view.txtPlannedName.TabIndex < view.cmdPreview.TabIndex, "edit/preview Tab order"
            ' Exercise the exact late-bound callback name used by NxFormWheelProc.
            ' These rows are form-only fixtures, not added workbook sheets.
            For index = 1 To 30
                view.lstSheets.AddItem "Wheel row " & CStr(index)
                view.lstSheets.List(index - 1, 1) = "Wheel target " & CStr(index)
                view.lstSheets.List(index - 1, 2) = "변경 없음"
            Next index
            view.lstSheets.ListIndex = 0
            priorTop = view.lstSheets.TopIndex
            CallByName view, "HandleWheelDelta", VbMethod, CLng(-120)
            Require view.lstSheets.TopIndex > priorTop, "wheel callback did not scroll the list"
            Unload view: Set view = Nothing
            Require Not NxFormWheelIsAttached(), "wheel leaked after close"
            Require LoadedRenameFormCount() = 0, "Unload left a loaded rename form"
            Set view = New FNxBatchRename
            view.BindFeature NX_FEATURE_FILE_BATCH_RENAME
            view.Show vbModeless
            DoEvents
            Require view.cmdAddFiles.Visible And view.cmdAddFiles.Enabled And view.cmdAddFolder.Enabled, "file ingestion mode disabled"
            Require view.chkRecursive.Visible And Not CBool(view.chkRecursive.Value), "recursive folder ingestion must be opt-in"
            Require view.cboRuleKind.ListCount = 8, "file mode lost the extension rule"
            Require view.txtPlannedName.Visible And Not view.txtPlannedName.Enabled, "empty file mode must show an inactive per-file editor"
            Require view.lstSheets.ListCount = 0 And Not view.cmdExecute.Enabled, "empty file mode retained a sealed plan"
            Unload view: Set view = Nothing
            Require Not NxFormWheelIsAttached(), "file-mode close leaked wheel hook"
            Require LoadedRenameFormCount() = 0, "file-mode close left a loaded rename form"
        Case Else
            NxRaiseContractError "Unknown r61 sheet editor case: " & kind
    End Select
    NxR61SheetNameEditorRunCase = "PASS|" & kind
    GoTo Cleanup
Failed:
    NxR61SheetNameEditorRunCase = "FAIL|" & kind & "|" & CStr(Err.Number) & "|" & Err.Description
Cleanup:
    On Error Resume Next
    NxBatchRenameTestClearMoveFailure
    If Not view Is Nothing Then Unload view
    If Not otherBook Is Nothing Then otherBook.Close SaveChanges:=False
    If Not book Is Nothing Then book.Close SaveChanges:=False
    Application.DisplayAlerts = savedAlerts
    On Error GoTo 0
End Function

Private Sub TestOrderedRuleDraftPreservesSource(ByVal book As Workbook)
    Dim order As New Collection, rules As New Collection, overrides As Object, rule As CNxRenameRule
    Dim draft As Collection, row As Variant, boundSheet As Worksheet, hasDuplicate As Boolean, hasCollision As Boolean
    order.Add "Alpha": order.Add "Beta"
    Set overrides = CreateObject("Scripting.Dictionary"): overrides.CompareMode = vbTextCompare
    overrides.Add "Beta", "ManualBeta"
    Set rule = New CNxRenameRule: rule.Kind = "PREFIX": rule.Text = "x_": rules.Add rule
    Set rule = New CNxRenameRule: rule.Kind = "REPLACE": rule.Text = "x_": rule.Replacement = "y_": rules.Add rule
    Set draft = NxRenameSheetRulesPreview(book, order, rules, overrides, hasDuplicate, hasCollision)
    NxSheetBatchRenameRequireReady book, draft
    row = draft(1): Require CStr(row(1)) = "y_Alpha", "ordered rules were not applied in sequence"
    row = draft(2): Require CStr(row(1)) = "ManualBeta", "rule preview lost the manual override"
    NxRenameMoveItem rules, 2, -1
    NxRenameMoveItem order, 2, -1
    Set draft = NxRenameSheetRulesPreview(book, order, rules, overrides, hasDuplicate, hasCollision)
    NxSheetBatchRenameRequireReady book, draft
    row = draft(1): Set boundSheet = row(3)
    Require boundSheet Is book.Worksheets("Beta"), "preview order changed the source object binding"
    Require CStr(row(1)) = "ManualBeta", "rule or item reorder lost the manual override"
    row = draft(2): Require CStr(row(1)) = "x_Alpha", "rule reorder did not rebuild the original name"
    Require book.Worksheets(1).Name = "Alpha" And book.Worksheets(2).Name = "Beta", "draft reordered actual worksheets"
    Require book.Worksheets(1).Range("A1").Value2 = "first-original" And book.Worksheets(2).Range("A1").Value2 = "second-original", "rule draft changed values"
    Require book.Saved, "rule draft changed the workbook Saved state"
End Sub

Private Function LoadedRenameFormCount() As Long
    Dim view As Object
    For Each view In VBA.UserForms
        If TypeName(view) = "FNxBatchRename" Then LoadedRenameFormCount = LoadedRenameFormCount + 1
    Next view
End Function

Private Sub TestSavedWorkbookApprovedResult()
    Dim book As Workbook, preview As Collection, options As Object, path As String
    Dim commandObject As New CNxFileFeatureCommand, command As INxFeatureCommand
    Dim router As New CNxExecutionRouter, ticket As CNxExecutionTicket, approval As CNxApproval, result As CNxResult
    Dim hasDuplicate As Boolean, hasCollision As Boolean, detail As String, ownedPath As Boolean
    On Error GoTo Failed
    path = Environ$("TEMP") & "\LHexcel-r61-sheet-" & NxCreateRunUuid() & ".xlsx"
    If Len(Dir$(path)) > 0 Then NxRaiseContractError "Saved-sheet fixture path already exists"
    Set book = Workbooks.Add(xlWBATWorksheet)
    book.Worksheets(1).Name = "SavedSource"
    book.Worksheets(1).Range("A1").Value2 = "preserve-original"
    book.SaveAs path, xlOpenXMLWorkbook
    ownedPath = True
    book.Worksheets(1).Range("A1").Select
    Set preview = BasePreview(book)
    Set preview = NxSheetBatchRenameEditPreview(book, preview, 1, "RenamedSource", hasDuplicate, hasCollision)
    Set options = CreateObject("Scripting.Dictionary")
    Set options.Item("preview") = preview
    commandObject.Configure NX_FEATURE_FILE_SHEET_BATCH_RENAME, book.FullName, workflowOptions:=options
    Set command = commandObject
    Set ticket = router.Prepare(NxFileFeatureDefinition(NX_FEATURE_FILE_SHEET_BATCH_RENAME), command)
    Select Case ticket.Decision.ResolvedGrade
        Case NxExecutionFast: Set result = router.ExecuteFast(ticket)
        Case NxExecutionGuarded: Set result = router.ExecuteGuarded(ticket, True)
        Case NxExecutionPlanned
            Set approval = router.TakePlannedApproval(ticket, True)
            Set result = router.RunApproved(ticket, approval)
        Case Else: NxRaiseContractError "Saved-sheet fixture received invalid execution grade"
    End Select
    Require Not result Is Nothing, "saved-sheet runner returned no result"
    Require result.Outcome = NxSuccess, "saved-sheet runner failed: " & result.Recovery
    Require result.Target = book.FullName, "saved-sheet result lost approved full path"
    Require book.Worksheets(1).Name = "RenamedSource", "saved-sheet runner did not rename"
    Require book.Worksheets(1).Range("A1").Value2 = "preserve-original", "saved-sheet runner changed content"
Cleanup:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    If ownedPath Then
        If Len(Dir$(path)) > 0 Then Kill path
    End If
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume Cleanup
End Sub

Private Function BasePreview(ByVal book As Workbook) As Collection
    Dim hasDuplicate As Boolean, hasCollision As Boolean
    Set BasePreview = NxSheetBatchRenameBuildPreview(book, vbNullString, vbNullString, vbNullString, vbNullString, hasDuplicate, hasCollision)
End Function

Private Function CyclePreview(ByVal book As Workbook) As Collection
    Dim preview As Collection, hasDuplicate As Boolean, hasCollision As Boolean
    Set preview = BasePreview(book)
    Set preview = NxSheetBatchRenameEditPreview(book, preview, 1, "Beta", hasDuplicate, hasCollision)
    Set CyclePreview = NxSheetBatchRenameEditPreview(book, preview, 2, "Alpha", hasDuplicate, hasCollision)
End Function

Private Function CommitRejected(ByVal book As Workbook, ByVal preview As Collection) As Boolean
    On Error GoTo Rejected
    Call NxSheetBatchRenameCommit(book, preview)
    Exit Function
Rejected:
    CommitRejected = True
    Err.Clear
End Function

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then NxRaiseContractError "R61 sheet editor: " & message
End Sub
