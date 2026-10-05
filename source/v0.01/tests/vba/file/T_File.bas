Attribute VB_Name = "T_File"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestInputFilesRemainReadOnlyDuringPlan"
    names.Add "TestFormulaResolverAcceptsOnlyDirectImportedReference"
    names.Add "TestFullFailureLeavesNoResult"
    names.Add "TestPartialRequiresTwoApprovalsAndSuffix"
    names.Add "TestMannerSaveRejectsAddin"
    names.Add "TestRangePngIsValidAndLeavesNoTempShape"
    names.Add "TestForbiddenBackupAndWorkstateFeaturesAreAbsent"
    names.Add "TestClosedFileFeatureDispatchSummaries"
    names.Add "TestRangePngCreatesValidOutput"
    names.Add "TestChartPngCreatesValidOutput"
    names.Add "TestPartialCreatesNoticeAndFailureList"
    names.Add "TestFolderHierarchyFromRange"
    names.Add "TestRowsConsolidationPreservesOrderAndLiteralText"
    names.Add "TestChartJpgHasJpegSignature"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestInputFilesRemainReadOnlyDuringPlan", "TestFormulaResolverAcceptsOnlyDirectImportedReference", _
             "TestFullFailureLeavesNoResult", "TestPartialRequiresTwoApprovalsAndSuffix", _
            "TestMannerSaveRejectsAddin", "TestRangePngIsValidAndLeavesNoTempShape", _
            "TestForbiddenBackupAndWorkstateFeaturesAreAbsent", "TestClosedFileFeatureDispatchSummaries", _
            "TestRangePngCreatesValidOutput", "TestChartPngCreatesValidOutput", _
            "TestPartialCreatesNoticeAndFailureList"
            Select Case name
                Case "TestInputFilesRemainReadOnlyDuringPlan": TestInputFilesRemainReadOnlyDuringPlan
                Case "TestFormulaResolverAcceptsOnlyDirectImportedReference": TestFormulaResolverAcceptsOnlyDirectImportedReference
                Case "TestFullFailureLeavesNoResult": TestFullFailureLeavesNoResult
                Case "TestPartialRequiresTwoApprovalsAndSuffix": TestPartialRequiresTwoApprovalsAndSuffix
                Case "TestMannerSaveRejectsAddin": TestMannerSaveRejectsAddin
                Case "TestRangePngIsValidAndLeavesNoTempShape": TestRangePngIsValidAndLeavesNoTempShape
                Case "TestForbiddenBackupAndWorkstateFeaturesAreAbsent": TestForbiddenBackupAndWorkstateFeaturesAreAbsent
                Case "TestClosedFileFeatureDispatchSummaries": TestClosedFileFeatureDispatchSummaries
                Case "TestRangePngCreatesValidOutput": TestRangePngCreatesValidOutput
                Case "TestChartPngCreatesValidOutput": TestChartPngCreatesValidOutput
                Case "TestPartialCreatesNoticeAndFailureList": TestPartialCreatesNoticeAndFailureList
            End Select
        Case "TestFolderHierarchyFromRange": TestFolderHierarchyFromRange
        Case "TestRowsConsolidationPreservesOrderAndLiteralText": TestRowsConsolidationPreservesOrderAndLiteralText
        Case "TestChartJpgHasJpegSignature": TestChartJpgHasJpegSignature
        Case Else: NxRaiseContractError "Unknown File test"
    End Select
End Sub

Private Sub TestClosedFileFeatureDispatchSummaries()
    Dim commandObject As New CNxFileFeatureCommand
    Dim command As INxFeatureCommand
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim planOutputPath As String
    NxTestHarness.AssertTrue InStr(1, NxFileActionSummary("NX-FILE-CONSOLIDATE"), "통합", vbTextCompare) > 0, "Consolidate dispatch summary missing"
    NxTestHarness.AssertTrue InStr(1, NxFileActionSummary("NX-FILE-MANNER-SAVE"), "A1", vbTextCompare) > 0, "Manner save dispatch summary missing"
    NxTestHarness.AssertTrue InStr(1, NxFileActionSummary("NX-FILE-RANGE-PNG"), "범위", vbTextCompare) > 0, "Range PNG dispatch summary missing"
    NxTestHarness.AssertTrue InStr(1, NxFileActionSummary("NX-FILE-CHART-PNG"), "차트", vbTextCompare) > 0, "Chart PNG dispatch summary missing"
    planOutputPath = Environ$("TEMP") & "\LHexcel-file-plan-contract-" & NxCreateRunUuid() & ".png"
    commandObject.Configure "NX-FILE-RANGE-PNG", planOutputPath
    Set command = commandObject
    Set ticket = router.Prepare(NxFileFeatureDefinition("NX-FILE-RANGE-PNG"), command)
    NxTestHarness.AssertTrue ticket.Decision.ResolvedGrade = NxExecutionPlanned, "Range PNG plan did not retain planned execution"
End Sub

Private Sub TestInputFilesRemainReadOnlyDuringPlan()
    Dim paths(0 To 0) As String, plan As Collection, inputBook As Workbook
    Dim inputPath As String, detail As String
    On Error GoTo Failed
    inputPath = Environ$("TEMP") & "\LHexcel-file-plan-input-" & NxCreateRunUuid() & ".xlsx"
    Set inputBook = Workbooks.Add(xlWBATWorksheet)
    inputBook.Worksheets(1).Cells(1, 1).Value = "read-only planning source"
    inputBook.SaveAs inputPath, xlOpenXMLWorkbook
    inputBook.Close SaveChanges:=False
    Set inputBook = Nothing
    paths(0) = inputPath
    Set plan = NxConsolidationPlan(paths)
    NxTestHarness.AssertTrue plan.Count = 1, "Read-only planning did not retain the input"
    NxTestHarness.AssertTrue Len(Dir$(inputPath)) > 0, "Read-only planning changed the input file"
    DeleteTestFileEventually inputPath
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not inputBook Is Nothing Then inputBook.Close SaveChanges:=False
    If Len(inputPath) > 0 And Len(Dir$(inputPath)) > 0 Then Kill inputPath
    On Error GoTo 0
    NxRaiseContractError detail
End Sub

Private Sub TestRangePngCreatesValidOutput()
    Dim book As Workbook, source As Range, result As CNxResult, outputPath As String, detail As String
    On Error GoTo Failed
    outputPath = Environ$("TEMP") & "\LHexcel-file-range-positive-" & NxCreateRunUuid() & ".png"
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set source = book.Worksheets(1).Range("A1:B2")
    source.Cells(1, 1).Value = "A": source.Cells(1, 2).Value = "B"
    source.Cells(2, 1).Value = 1: source.Cells(2, 2).Value = 2
    book.Activate
    source.Select
    Set result = NxExportRangePng(source, outputPath)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, _
        "Range PNG positive export did not succeed; outcome=" & CStr(result.Outcome) & _
        "; stage=" & result.Stage & "; recovery=" & result.Recovery
    NxTestHarness.AssertTrue Len(Dir$(outputPath)) > 0, "Range PNG positive output was not materialized"
    NxTestHarness.AssertTrue book.Worksheets(1).ChartObjects.Count = 0, "Range PNG left a temporary chart on the source"
    Kill outputPath
    book.Close SaveChanges:=False
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    If Len(Dir$(outputPath)) > 0 Then Kill outputPath
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    NxRaiseContractError detail
End Sub

Private Sub TestPartialCreatesNoticeAndFailureList()
    Dim paths(0 To 1) As String, result As CNxResult, outputPath As String, partialPath As String
    Dim inputBook As Workbook, inputPath As String
    inputPath = Environ$("TEMP") & "\LHexcel-file-partial-input-" & NxCreateRunUuid() & ".xlsx"
    Set inputBook = Workbooks.Add(xlWBATWorksheet)
    inputBook.Worksheets(1).Cells(1, 1).Value = "partial source"
    inputBook.SaveAs inputPath, xlOpenXMLWorkbook
    inputBook.Close SaveChanges:=False
    paths(0) = inputPath
    paths(1) = Environ$("TEMP") & "\LHexcel-file-missing-" & NxCreateRunUuid() & ".xlsx"
    outputPath = Environ$("TEMP") & "\LHexcel-file-partial-positive-" & NxCreateRunUuid() & ".xlsx"
    partialPath = Left$(outputPath, Len(outputPath) - 5) & "_부분결과.xlsx"
    Set result = NxRunConsolidation(paths, outputPath, True, True)
    NxTestHarness.AssertTrue result.Outcome = NxPartialFailure, "Approved partial consolidation did not return partial failure"
    NxTestHarness.AssertTrue Len(Dir$(partialPath)) > 0, "Approved partial consolidation did not materialize output"
    NxTestHarness.AssertTrue InStr(1, result.Recovery, "Input file", vbTextCompare) > 0, "Partial result omitted failure details"
    DeleteTestFileEventually partialPath
    DeleteTestFileEventually inputPath
End Sub

Private Sub TestChartPngCreatesValidOutput()
    Dim book As Workbook, chart As ChartObject, result As CNxResult, outputPath As String, detail As String
    On Error GoTo Failed
    outputPath = Environ$("TEMP") & "\LHexcel-file-chart-positive-" & NxCreateRunUuid() & ".png"
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set chart = book.Worksheets(1).ChartObjects.Add(10, 10, 240, 160)
    book.Activate
    chart.Select
    Set result = NxExportChartPng(chart, outputPath)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, _
        "Chart PNG positive export did not succeed; outcome=" & CStr(result.Outcome) & _
        "; stage=" & result.Stage & "; recovery=" & result.Recovery
    NxTestHarness.AssertTrue Len(Dir$(outputPath)) > 0, "Chart PNG positive output was not materialized"
    NxTestHarness.AssertTrue book.Worksheets(1).ChartObjects.Count = 1, "Chart PNG removed the source chart"
    Kill outputPath
    book.Close SaveChanges:=False
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    If Len(Dir$(outputPath)) > 0 Then Kill outputPath
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    NxRaiseContractError detail
End Sub

Private Sub TestFormulaResolverAcceptsOnlyDirectImportedReference()
    Dim sheetMap As Object
    Set sheetMap = CreateObject("Scripting.Dictionary")
    sheetMap.Add "A.xlsx|Sheet1", True
    NxTestHarness.AssertTrue NxFormulaResolution("='[A.xlsx]Sheet1'!$A$1", sheetMap) = "relink", "Direct reference was not relinked"
    NxTestHarness.AssertTrue NxFormulaResolution("=SUM('[A.xlsx]Sheet1'!A1:A3)", sheetMap) = "decision", "Compound external formula bypassed decision"
    NxTestHarness.AssertTrue NxFormulaResolution("=MyName", sheetMap) = "approved", "Local formula disposition changed"
End Sub

Private Sub TestFullFailureLeavesNoResult()
    Dim paths(0 To 0) As String, result As CNxResult, outputPath As String
    paths(0) = Environ$("TEMP") & "\LHexcel-file-missing.xlsx"
    outputPath = Environ$("TEMP") & "\LHexcel-file-full-failure-" & NxCreateRunUuid() & ".xlsx"
    Set result = NxRunConsolidation(paths, outputPath, False, False)
    NxTestHarness.AssertTrue result.Outcome = NxEnvironmentError, "Full failure was not an environment error"
    NxTestHarness.AssertTrue Len(Dir$(outputPath)) = 0, "Full failure left a result file"
End Sub

Private Sub TestPartialRequiresTwoApprovalsAndSuffix()
    Dim paths(0 To 1) As String, result As CNxResult, outputPath As String
    Dim inputBook As Workbook, inputPath As String, detail As String, stage As String
    On Error GoTo Failed
    stage = "create input workbook"
    inputPath = Environ$("TEMP") & "\LHexcel-file-approval-input-" & NxCreateRunUuid() & ".xlsx"
    Set inputBook = Workbooks.Add(xlWBATWorksheet)
    inputBook.Worksheets(1).Cells(1, 1).Value = "approval source"
    stage = "save input workbook"
    inputBook.SaveAs inputPath, xlOpenXMLWorkbook
    stage = "close input workbook"
    inputBook.Close SaveChanges:=False
    Set inputBook = Nothing
    paths(0) = inputPath
    paths(1) = Environ$("TEMP") & "\LHexcel-file-missing-" & NxCreateRunUuid() & ".xlsx"
    outputPath = Environ$("TEMP") & "\LHexcel-file-partial-" & NxCreateRunUuid() & ".xlsx"
    stage = "run consolidation"
    Set result = NxRunConsolidation(paths, outputPath, True, False)
    stage = "verify cancelled result"
    NxTestHarness.AssertTrue result.Outcome = NxCancelled, _
        "Partial failure without second approval was not cancelled; outcome=" & CStr(result.Outcome) & _
        "; stage=" & result.Stage & "; recovery=" & result.Recovery
    NxTestHarness.AssertTrue InStr(1, result.Recovery, "second approval", vbTextCompare) > 0, "Partial cancellation omitted approval reason"
    NxTestHarness.AssertTrue Len(Dir$(outputPath & "_부분결과")) = 0, "Nonexistent partial result was reported"
    stage = "delete input workbook"
    DeleteTestFileEventually inputPath
    Exit Sub
Failed:
    detail = "TestPartialRequiresTwoApprovalsAndSuffix stage=" & stage & " :: " & Err.Description
    On Error Resume Next
    If Not inputBook Is Nothing Then inputBook.Close SaveChanges:=False
    If Len(inputPath) > 0 And Len(Dir$(inputPath)) > 0 Then Kill inputPath
    If Len(outputPath) > 0 And Len(Dir$(outputPath)) > 0 Then Kill outputPath
    On Error GoTo 0
    NxRaiseContractError detail
End Sub

Private Sub TestMannerSaveRejectsAddin()
    Dim result As CNxResult
    Set result = NxRunMannerSave(ThisWorkbook)
    NxTestHarness.AssertTrue result.Outcome = NxEnvironmentError, "Manner save accepted the add-in itself"
End Sub

Private Sub TestRangePngIsValidAndLeavesNoTempShape()
    Dim result As CNxResult, outputPath As String
    outputPath = Environ$("TEMP") & "\LHexcel-file-range-" & NxCreateRunUuid() & ".png"
    Set result = NxExportRangePng(Nothing, outputPath)
    NxTestHarness.AssertTrue result.Outcome <> NxSuccess, "Range PNG accepted a missing contiguous range"
    NxTestHarness.AssertTrue Len(Dir$(outputPath)) = 0, "Range PNG failure left an output"
End Sub

Private Sub TestForbiddenBackupAndWorkstateFeaturesAreAbsent()
    Dim failed As Boolean
    NxTestHarness.AssertTrue NxFileActionSummary("NX-FILE-CONSOLIDATE") <> "", "Consolidate action missing"
    NxTestHarness.AssertTrue NxFileActionSummary("NX-FILE-MANNER-SAVE") <> "", "Manner action missing"
    NxTestHarness.AssertTrue NxFileActionSummary("NX-FILE-RANGE-PNG") <> "", "Range PNG action missing"
    NxTestHarness.AssertTrue NxFileActionSummary("NX-FILE-CHART-PNG") <> "", "Chart PNG action missing"
    On Error Resume Next
    Call NxFileActionSummary("NX-FILE-BACKUP")
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Retired backup action was exposed"
End Sub

Private Sub TestFolderHierarchyFromRange()
    Dim book As Workbook, source As Range, items As Collection, item As CNxFolderCreateItem
    Dim detail As String, rejected As Boolean
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet): Set source = book.Worksheets(1).Range("A1:C3")
    source.Cells(1, 1).Value2 = "사업": source.Cells(1, 2).Value2 = "동": source.Cells(1, 3).Value2 = "세대"
    source.Cells(2, 1).Value2 = "사업": source.Cells(2, 2).Value2 = "동": source.Cells(2, 3).Value2 = "사무"
    source.Rows(3).Value2 = source.Rows(1).Value2
    Set items = NxFolderCreateItemsFromRange(source, "앞_", "_뒤")
    NxTestHarness.AssertTrue items.Count = 2, "Identical hierarchy rows were not deduplicated"
    Set item = items(1)
    NxTestHarness.AssertTrue item.RelativePath = "앞_사업_뒤\앞_동_뒤\앞_세대_뒤", "Hierarchy or per-level prefix/suffix was not preserved"
    NxTestHarness.AssertTrue NxFolderNameIsValid("한글폴더"), "Korean folder name was rejected as a control character"
    NxTestHarness.AssertTrue Not NxFolderNameIsValid("CON.txt") And Not NxFolderNameIsValid(".."), "Reserved or traversal folder name was accepted"
    source.Cells(2, 2).ClearContents
    On Error Resume Next
    Set items = NxFolderCreateItemsFromRange(source)
    rejected = Err.Number <> 0: Err.Clear
    On Error GoTo Failed
    NxTestHarness.AssertTrue rejected, "Missing parent between nonempty hierarchy columns was accepted"
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Private Sub TestRowsConsolidationPreservesOrderAndLiteralText()
    Dim firstBook As Workbook, secondBook As Workbook, resultBook As Workbook, sheet As Worksheet
    Dim paths(0 To 1) As String, outputPath As String, result As CNxResult, index As Long, detail As String
    Dim previousStatus As Variant, phase As String, defaultStatus As Variant, statusProbe As String
    On Error GoTo Failed
    previousStatus = Application.StatusBar
    phase = "status_restore"
    ' Use an owned stable message rather than Excel's transient compile/progress text.
    Application.StatusBar = "r61 owned status before file operation"
    NxFileOperationBegin 1, "r61 long progress"
    NxFileOperationReport 1, String$(600, "x")
    NxTestHarness.AssertTrue Len(CStr(Application.StatusBar)) <= 250, "Long paths overflowed the progress status"
    NxFileOperationFinish
    NxTestHarness.AssertTrue CStr(Application.StatusBar) = "r61 owned status before file operation", "File operation did not restore the owned status bar: " & CStr(Application.StatusBar)
    Application.StatusBar = False
    DoEvents
    defaultStatus = Application.StatusBar
    NxFileOperationBegin 1, "r61 default status"
    NxFileOperationFinish
    DoEvents
    statusProbe = "default=" & CStr(VarType(defaultStatus)) & ":" & CStr(defaultStatus) & "; after=" & CStr(VarType(Application.StatusBar)) & ":" & CStr(Application.StatusBar)
    NxTestHarness.AssertTrue VarType(Application.StatusBar) = VarType(defaultStatus) And CStr(Application.StatusBar) = CStr(defaultStatus), "File operation did not restore Excel default status: " & statusProbe
    Application.StatusBar = previousStatus
    phase = "output_path_guard"
    RecordRowsPhase phase
    TestRowsLongPathRejectedBeforeMutation
    RecordRowsPhase "output_path_guard_pass"
    paths(0) = Environ$("TEMP") & "\LHexcel-rows-first-" & NxCreateRunUuid() & ".xlsx"
    paths(1) = Environ$("TEMP") & "\LHexcel-rows-second-" & NxCreateRunUuid() & ".xlsx"
    outputPath = Environ$("TEMP") & "\LHexcel-rows-output-" & NxCreateRunUuid() & ".xlsx"
    Set firstBook = Workbooks.Add(xlWBATWorksheet)
    Set sheet = firstBook.Worksheets.Add(After:=firstBook.Worksheets(1))
    Set secondBook = Workbooks.Add(xlWBATWorksheet)
    Set sheet = secondBook.Worksheets.Add(After:=secondBook.Worksheets(1))
    For index = 1 To 4
        If index <= 2 Then Set sheet = firstBook.Worksheets(index) Else Set sheet = secondBook.Worksheets(index - 2)
        sheet.Range("A1:B2").NumberFormat = "@"
        sheet.Range("A1").Value2 = "번호": sheet.Range("B1").Value2 = "값"
        sheet.Range("A2").Value2 = CStr(index): sheet.Range("B2").Value2 = "text-" & CStr(index)
    Next index
    sheet.Range("B2").Value2 = "=1+1"
    phase = "save_first (" & CStr(Len(paths(0))) & " chars)"
    RecordRowsPhase phase
    firstBook.SaveAs paths(0), xlOpenXMLWorkbook
    phase = "save_second (" & CStr(Len(paths(1))) & " chars)"
    RecordRowsPhase phase
    secondBook.SaveAs paths(1), xlOpenXMLWorkbook
    phase = "consolidate"
    RecordRowsPhase phase
    Set result = NxRunConsolidation(paths, outputPath, False, False, "ROWS", True)
    NxTestHarness.AssertTrue Not result Is Nothing, "Row consolidation returned no result"
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Row consolidation failed: " & result.Recovery
    NxTestHarness.AssertTrue firstBook.Saved And secondBook.Saved, "Row consolidation modified a user-owned input workbook"
    phase = "open_result"
    RecordRowsPhase phase
    Set resultBook = Workbooks.Open(Filename:=outputPath, UpdateLinks:=0, ReadOnly:=True)
    NxTestHarness.AssertTrue resultBook.Worksheets.Count = 1, "Row mode did not produce one worksheet"
    Set sheet = resultBook.Worksheets(1)
    NxTestHarness.AssertTrue sheet.UsedRange.Rows.Count = 5, "Repeated headers were not skipped exactly"
    For index = 1 To 4
        NxTestHarness.AssertTrue CStr(sheet.Cells(index + 1, 1).Value2) = CStr(index), "File/sheet append order changed"
    Next index
    NxTestHarness.AssertTrue CStr(sheet.Range("B5").Value2) = "=1+1" And Not sheet.Range("B5").HasFormula, "Formula-like text was executed in row output"
    resultBook.Close SaveChanges:=False
    Set resultBook = Nothing
    phase = "partial_rows"
    RecordRowsPhase phase
    TestPartialRowsConsolidationUsesApprovedTarget paths(1), secondBook
CleanUp:
    On Error Resume Next
    NxFileOperationFinish
    Application.StatusBar = previousStatus
    If Not resultBook Is Nothing Then resultBook.Close SaveChanges:=False
    If Not firstBook Is Nothing Then firstBook.Close SaveChanges:=False
    If Not secondBook Is Nothing Then secondBook.Close SaveChanges:=False
    DeleteTestFileEventually paths(0): DeleteTestFileEventually paths(1): DeleteTestFileEventually outputPath
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = phase & ": " & CStr(Err.Number) & " / " & Err.Source & " / " & Err.Description
    Resume CleanUp
End Sub

Private Sub TestRowsLongPathRejectedBeforeMutation()
    Dim inputPaths(0 To 0) As String, outputPath As String, result As CNxResult, beforeCount As Long
    beforeCount = Application.Workbooks.Count
    outputPath = Environ$("TEMP") & "\" & String$(219, "x") & ".xlsx"
    inputPaths(0) = Environ$("TEMP") & "\LHexcel-missing.xlsx"
    Set result = NxRunConsolidation(inputPaths, outputPath, False, False, "ROWS", True)
    NxTestHarness.AssertTrue Not result Is Nothing, "Long output path raised outside result contract"
    NxTestHarness.AssertTrue result.Outcome = NxEnvironmentError And Not result.SourceChanged, "Long output path was not rejected safely"
    NxTestHarness.AssertTrue InStr(1, result.Recovery, "경로가 너무 깁니다", vbBinaryCompare) > 0, "Long path rejection omitted actionable Korean guidance"
    NxTestHarness.AssertTrue Application.Workbooks.Count = beforeCount, "Long output path created a result workbook before validation"
End Sub

Private Sub RecordRowsPhase(ByVal phase As String)
    Dim handle As Integer
    handle = FreeFile
    Open Environ$("TEMP") & "\r61-row-phase.log" For Append As #handle
    Print #handle, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & phase
    Close #handle
End Sub

Private Sub TestPartialRowsConsolidationUsesApprovedTarget(ByVal inputPath As String, ByVal sourceBook As Workbook)
    Dim commandObject As New CNxFileFeatureCommand, command As INxFeatureCommand
    Dim router As New CNxExecutionRouter, ticket As CNxExecutionTicket, approval As CNxApproval
    Dim options As Object, paths(0 To 1) As String, outputPath As String, partialPath As String
    Dim result As CNxResult, partialBook As Workbook, dataSheet As Worksheet
    Dim sourceWasSaved As Boolean, detail As String, savedTrace As String, guard As CNxStateGuard
    On Error GoTo Failed
    paths(0) = inputPath
    paths(1) = Environ$("TEMP") & "\LHexcel-rows-missing-" & NxCreateRunUuid() & ".xlsx"
    outputPath = Environ$("TEMP") & "\LHexcel-rows-partial-" & NxCreateRunUuid() & ".xlsx"
    partialPath = Left$(outputPath, Len(outputPath) - 5) & "_부분결과.xlsx"
    sourceBook.Activate
    sourceBook.Worksheets(1).Activate
    sourceBook.Worksheets(1).Range("A1").Select
    sourceWasSaved = sourceBook.Saved
    RecordRowsPhase "partial_guard"
    Set guard = New CNxStateGuard
    guard.Restore
    NxTestHarness.AssertTrue sourceBook.Saved = sourceWasSaved And Len(guard.Recovery) = 0, "No-op state guard modified the saved source workbook"
    Set guard = Nothing
    RecordRowsPhase "partial_guard_pass"
    Set options = CreateObject("Scripting.Dictionary")
    options.Add "consolidation_mode", "ROWS"
    options.Add "skip_headers", True
    commandObject.Configure "NX-FILE-CONSOLIDATE", outputPath, paths, True, secondApproval:=True, workflowOptions:=options
    Set command = commandObject
    Set ticket = router.Prepare(NxFileFeatureDefinition("NX-FILE-CONSOLIDATE"), command)
    savedTrace = "before=" & CStr(sourceWasSaved) & ";prepared=" & CStr(sourceBook.Saved)
    NxTestHarness.AssertTrue ticket.Decision.ResolvedGrade = NxExecutionPlanned, "Partial row fixture did not prepare a planned ticket"
    ' Test-owned approval fixture: no dialog automation; RunApproved reaches the real core result validation.
    Set approval = router.TakePlannedApproval(ticket, True)
    savedTrace = savedTrace & ";approved=" & CStr(sourceBook.Saved)
    Set result = router.RunApproved(ticket, approval)
    RecordRowsPhase "partial_executed"
    savedTrace = savedTrace & ";executed=" & CStr(sourceBook.Saved) & ";sourceChanged=" & CStr(result.SourceChanged)
    NxTestHarness.AssertTrue Not result Is Nothing, "Partial row runner returned no result"
    NxTestHarness.AssertTrue result.Outcome = NxPartialFailure, "Partial row runner did not retain partial outcome: " & result.Recovery
    NxTestHarness.AssertTrue result.Stage = "partial_result", "Partial row runner hit an exception: " & result.Recovery
    NxTestHarness.AssertTrue result.Target = outputPath, "Partial row result no longer identifies the approved output target"
    NxTestHarness.AssertTrue InStr(1, result.Recovery, partialPath, vbBinaryCompare) > 0, "Partial row recovery omitted the actual output path"
    NxTestHarness.AssertTrue InStr(1, result.Recovery, paths(1), vbBinaryCompare) > 0, "Partial row recovery omitted the failed input path"
    NxTestHarness.AssertTrue Len(Dir$(outputPath)) = 0 And Len(Dir$(partialPath)) > 0, "Partial row output was not isolated from the full output path"
    NxTestHarness.AssertTrue sourceBook.Saved = sourceWasSaved And Not result.SourceChanged, "Partial row execution modified the source workbook: " & savedTrace
    RecordRowsPhase "partial_asserts_pass_open"
    Set approval = Nothing: Set ticket = Nothing: Set command = Nothing
    Set commandObject = Nothing: Set router = Nothing: Set result = Nothing
    DoEvents
    RecordRowsPhase "partial_command_released"
    Set partialBook = Workbooks.Open(Filename:=partialPath, UpdateLinks:=0, ReadOnly:=True)
    RecordRowsPhase "partial_opened"
    NxTestHarness.AssertTrue partialBook.Worksheets.Count = 2, "Partial row output omitted data or failure notice"
    NxTestHarness.AssertTrue InStr(1, CStr(partialBook.Worksheets("부분결과_안내").Range("A3").Value2), paths(1), vbBinaryCompare) > 0, "Partial output notice omitted the failed input"
    Set dataSheet = partialBook.Worksheets("통합데이터")
    NxTestHarness.AssertTrue dataSheet.UsedRange.Rows.Count = 3, "Partial output retained the wrong number of rows"
    NxTestHarness.AssertTrue CStr(dataSheet.Range("A2").Value2) = "3" And CStr(dataSheet.Range("A3").Value2) = "4", "Partial output changed input row order"
    NxTestHarness.AssertTrue CStr(dataSheet.Range("B3").Value2) = "=1+1" And Not dataSheet.Range("B3").HasFormula, "Partial output executed formula-like text"
    RecordRowsPhase "partial_contents_pass"
CleanUp:
    On Error Resume Next
    RecordRowsPhase "partial_close"
    If Not partialBook Is Nothing Then partialBook.Close SaveChanges:=False
    RecordRowsPhase "partial_closed"
    DeleteTestFileEventually outputPath
    DeleteTestFileEventually partialPath
    RecordRowsPhase "partial_deleted"
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError "Partial row approval fixture: " & detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Private Sub TestChartJpgHasJpegSignature()
    Dim book As Workbook, chart As ChartObject, result As CNxResult, outputPath As String, detail As String
    Dim handle As Integer, signature(0 To 2) As Byte, opened As Boolean
    On Error GoTo Failed
    outputPath = Environ$("TEMP") & "\LHexcel-chart-jpg-" & NxCreateRunUuid() & ".jpg"
    Set book = Workbooks.Add(xlWBATWorksheet)
    With book.Worksheets(1)
        .Range("A1").Value2 = "Item": .Range("B1").Value2 = "Value"
        .Range("A2").Value2 = "A": .Range("B2").Formula = "=3+4"
        .Range("A3").Value2 = "B": .Range("B3").Value2 = 9
        .Range("F12").Formula = "=40+2"
    End With
    Set chart = book.Worksheets(1).ChartObjects.Add(10, 10, 240, 160)
    chart.Chart.SetSourceData Source:=book.Worksheets(1).Range("A1:B3"), PlotBy:=xlColumns
    book.Activate: chart.Select
    Set result = NxExportChartPng(chart, outputPath)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "JPG chart export failed"
    handle = FreeFile: Open outputPath For Binary Access Read As #handle: opened = True
    Get #handle, 1, signature
    Close #handle: opened = False
    NxTestHarness.AssertTrue signature(0) = &HFF And signature(1) = &HD8 And signature(2) = &HFF, "JPG extension did not contain JPEG bytes"
    NxTestHarness.AssertTrue book.Worksheets(1).ChartObjects.Count = 1, "JPG export removed the original chart"
    TestPictureInsertUsesApprovedWorkbookIdentity book, chart, outputPath
CleanUp:
    On Error Resume Next
    If opened Then Close #handle
    If Not book Is Nothing Then book.Close SaveChanges:=False
    DeleteTestFileEventually outputPath
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Private Sub TestPictureInsertUsesApprovedWorkbookIdentity(ByVal sourceBook As Workbook, ByVal sourceChart As ChartObject, ByVal imagePath As String)
    Dim request As CNxPictureInsertRequest, preview As CNxPictureInsertPreview, journal As New CNxPictureInsertJournal
    Dim commandObject As New CNxPictureInsertCommand, command As INxFeatureCommand
    Dim router As New CNxExecutionRouter, ticket As CNxExecutionTicket, approval As CNxApproval, result As CNxResult
    Dim sourceSheet As Worksheet, target As Range, sourceCell As Range, picture As Shape, files As New Collection
    Dim beforeFormulas As Variant, expectedTarget As String, shapeCount As Long, seriesFormula As String, chartName As String
    Dim chartLeft As Double, chartTop As Double, chartWidth As Double, chartHeight As Double
    NxTestHarness.AssertTrue Len(sourceBook.Path) = 0, "Picture insertion fixture must retain its unsaved workbook identity"
    Set sourceSheet = sourceBook.Worksheets(1)
    sourceBook.Activate: sourceSheet.Activate
    Set target = sourceSheet.Range("F12")
    target.Select
    beforeFormulas = sourceSheet.Range("A1:F12").Formula
    shapeCount = sourceSheet.Shapes.Count: chartName = sourceChart.Name
    NxTestHarness.AssertTrue sourceChart.Chart.SeriesCollection.Count = 1, "JPG chart fixture must retain one data series"
    seriesFormula = sourceChart.Chart.SeriesCollection(1).Formula
    chartLeft = sourceChart.Left: chartTop = sourceChart.Top: chartWidth = sourceChart.Width: chartHeight = sourceChart.Height
    files.Add imagePath
    Set request = NxPictureInsertCreateRequest(target, files, "FIT")
    Set preview = NxPictureInsertBuildPreview(request)
    commandObject.Configure request, preview, journal
    Set command = commandObject
    Set ticket = router.Prepare(NxDrawingFeatureDefinition("NX-DRAW-INSERT-PICTURE"), command)
    NxTestHarness.AssertTrue ticket.Decision.ResolvedGrade = NxExecutionPlanned, "Picture insertion did not retain planned approval"
    ' Test-owned explicit approval reaches the actual dispatcher and result validation.
    Set approval = router.TakePlannedApproval(ticket, True)
    expectedTarget = approval.Context.WorkbookIdentity
    NxTestHarness.AssertTrue Left$(expectedTarget, 8) = "unsaved:" And expectedTarget <> sourceBook.FullName, "Unsaved approval did not use an ephemeral workbook identity"
    Set result = router.RunApproved(ticket, approval)
    NxTestHarness.AssertTrue Not result Is Nothing, "Picture insertion runner returned no result"
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Picture insertion runner failed: " & result.Stage & " / " & result.Recovery
    NxTestHarness.AssertTrue result.Target = expectedTarget, "Picture insertion result did not retain the approved workbook identity"
    NxTestHarness.AssertTrue result.FeatureId = "NX-DRAW-INSERT-PICTURE" And result.Stage = "complete", "Picture insertion result lost its feature or completion stage"
    NxTestHarness.AssertTrue result.SourceChanged And result.MessageKey = "picture_inserted" And Len(result.Recovery) = 0, "Picture insertion result lost its original metadata"
    NxTestHarness.AssertTrue sourceSheet.Shapes.Count = shapeCount + 1, "Picture insertion did not add exactly one shape"
    Set picture = sourceSheet.Shapes(CStr(preview.ShapeNames(1)))
    NxTestHarness.AssertTrue picture.Type = msoPicture And picture.Width > 0 And picture.Height > 0, "JPG insertion did not produce an embedded picture"
    NxTestHarness.AssertTrue sourceSheet.ChartObjects.Count = 1 And sourceChart.Name = chartName, "Picture insertion removed or replaced the existing chart"
    NxTestHarness.AssertTrue sourceChart.Left = chartLeft And sourceChart.Top = chartTop And sourceChart.Width = chartWidth And sourceChart.Height = chartHeight, "Picture insertion moved or resized the existing chart"
    NxTestHarness.AssertTrue sourceChart.Chart.SeriesCollection.Count = 1 And sourceChart.Chart.SeriesCollection(1).Formula = seriesFormula, "Picture insertion changed chart source data"
    For Each sourceCell In sourceSheet.Range("A1:F12")
        NxTestHarness.AssertTrue CStr(sourceCell.Formula) = CStr(beforeFormulas(sourceCell.Row, sourceCell.Column)), "Picture insertion changed an existing cell value or formula"
    Next sourceCell
    NxTestHarness.AssertTrue Len(sourceBook.Path) = 0, "Picture insertion saved the source workbook unexpectedly"
    ' The caller owns the workbook and JPG and closes/deletes both even if an assertion raises.
End Sub

Private Sub DeleteTestFileEventually(ByVal filePath As String)
    Dim attempt As Long, exists As Boolean
    If Len(filePath) = 0 Then Exit Sub
    For attempt = 1 To 20
        On Error Resume Next
        Err.Clear
        exists = (Len(Dir$(filePath)) > 0)
        If exists Then Kill filePath
        Err.Clear
        exists = (Len(Dir$(filePath)) > 0)
        On Error GoTo 0
        If Not exists Then Exit Sub
        DoEvents
    Next attempt
    NxRaiseContractError "Test cleanup could not delete artifact: " & filePath
End Sub
