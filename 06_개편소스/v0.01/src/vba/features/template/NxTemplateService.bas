Attribute VB_Name = "NxTemplateService"
Option Explicit

Private mFailurePoint As String
Private mProgressView As Object
Private mCancelRequested As Boolean
Public Const NX_TEMPLATE_CANCELLED As Long = vbObjectError + 2394

Public Sub NxTemplateBeginProgress(ByVal view As Object)
    Set mProgressView = view
    mCancelRequested = False
End Sub

Public Sub NxTemplateEndProgress()
    Set mProgressView = Nothing
    mCancelRequested = False
End Sub

Public Sub NxTemplateRequestCancel()
    mCancelRequested = True
End Sub

Public Function NxTemplateProgressText(ByVal stage As String, ByVal completed As Long, ByVal total As Long) As String
    If mCancelRequested Then
        NxTemplateProgressText = "취소 요청을 처리하고 임시 결과를 정리하는 중입니다."
    ElseIf total > 1 Then
        NxTemplateProgressText = stage & " · 단계 " & Format$(CDbl(completed) / CDbl(total), "0%") & _
            " (" & Format$(completed, "#,##0") & " / " & Format$(total, "#,##0") & ")"
    Else
        NxTemplateProgressText = stage & " 중입니다."
    End If
End Function

Public Sub NxTemplateProgress(ByVal stage As String, ByVal completed As Long, ByVal total As Long)
    If Not mProgressView Is Nothing Then
        mProgressView.UpdateTemplateProgress stage, completed, total
        DoEvents
    End If
    If mCancelRequested Then Err.Raise NX_TEMPLATE_CANCELLED, "NxTemplateService", "템플릿 작업을 취소했습니다."
End Sub

Public Function NxTemplateRegisterSheet(ByVal request As CNxTemplateRequest) As CNxResult
    Set NxTemplateRegisterSheet = RegisterTemplate(request, NxTemplateStoreRoot())
End Function

Public Function NxTemplateList() As Collection
    Set NxTemplateList = ListTemplates(NxTemplateStoreRoot())
End Function

Public Function NxTemplateLoad(ByVal templateId As String, ByVal targetBook As Workbook, ByVal riskConfirmed As Boolean) As CNxResult
    Set NxTemplateLoad = LoadTemplate(templateId, targetBook, riskConfirmed, NxTemplateStoreRoot())
End Function

Public Function NxTemplateRename(ByVal templateId As String, ByVal displayName As String, ByVal description As String, Optional ByVal category As Variant) As CNxResult
    Set NxTemplateRename = RenameTemplate(templateId, displayName, description, NxTemplateStoreRoot(), category)
End Function

Public Function NxTemplateDeleteToQuarantine(ByVal templateId As String) As CNxResult
    Set NxTemplateDeleteToQuarantine = DeleteTemplate(templateId, NxTemplateStoreRoot())
End Function

Public Function NxTemplateRegisterForTest(ByVal request As CNxTemplateRequest, ByVal root As String) As CNxResult
    Set NxTemplateRegisterForTest = RegisterTemplate(request, root)
End Function

Public Function NxTemplateListForTest(ByVal root As String) As Collection
    Set NxTemplateListForTest = ListTemplates(root)
End Function

Public Function NxTemplateLoadForTest(ByVal templateId As String, ByVal targetBook As Workbook, ByVal riskConfirmed As Boolean, ByVal root As String) As CNxResult
    Set NxTemplateLoadForTest = LoadTemplate(templateId, targetBook, riskConfirmed, root)
End Function

Public Function NxTemplateRenameForTest(ByVal templateId As String, ByVal displayName As String, ByVal description As String, ByVal root As String) As CNxResult
    Set NxTemplateRenameForTest = RenameTemplate(templateId, displayName, description, root)
End Function

Public Function NxTemplateDeleteForTest(ByVal templateId As String, ByVal root As String) As CNxResult
    Set NxTemplateDeleteForTest = DeleteTemplate(templateId, root)
End Function

Public Sub NxTemplateSetFailurePoint(ByVal value As String)
    Select Case value
        Case "after-load-cells", "after-second-load-sheet", "after-quarantine-move"
            mFailurePoint = value
        Case Else
            NxRaiseContractError "Template failure point is outside the test contract"
    End Select
End Sub

Public Sub NxTemplateClearFailurePoint()
    mFailurePoint = vbNullString
End Sub

Private Function RegisterTemplate(ByVal request As CNxTemplateRequest, ByVal root As String) As CNxResult
    Dim records As Collection, record As CNxTemplateRecord, verifyRecord As CNxTemplateRecord, analysis As CNxTemplateAnalysis
    Dim templateId As String, operationId As String, stagingPath As String, finalFolder As String
    Dim packagePath As String, metadataPath As String, packageSha As String, logicalSha As String
    Dim createdUtc As String, transaction As New CNxTemplateTransaction, item As Variant
    Dim fileSystem As Object
    Dim failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    If request Is Nothing Then NxRaiseContractError "Template registration request is required"
    Dim source As Range, riskFlags As New Collection, rowCount As Long, columnCount As Long
    For Each source In request.SourceRanges
        Set analysis = NxTemplateAnalyze(source, request.RetainedForSource(source), request.IncludeHidden, request.ExclusionAccepted)
        If analysis.HasBlockedResource Then NxRaiseContractError analysis.BlockReason
        If analysis.RiskFlags.Count > 0 And Not request.RiskConfirmed Then NxRaiseContractError "Risky template formulas require explicit confirmation"
        For Each item In analysis.RiskFlags
            On Error Resume Next
            riskFlags.Add CStr(item), CStr(item)
            On Error GoTo Failed
        Next item
    Next source
    rowCount = request.SourceRange.Rows.Count
    columnCount = request.SourceRange.Columns.Count
    If request.SourceKind = "workbook" Or request.SourceKind = "file" Then
        rowCount = request.ActualCellCount
        columnCount = 1
    End If
    NxTemplateEnsureStore root
    Set records = NxTemplateListMetadata(root, False)
    For Each item In records
        Set record = item
        If StrComp(record.DisplayName, request.DisplayName, vbTextCompare) = 0 Then NxRaiseContractError "Template display name already exists"
    Next item
    templateId = "template-" & Replace(NxCreateRunUuid(), "-", vbNullString)
    operationId = "operation-" & Replace(NxCreateRunUuid(), "-", vbNullString)
    transaction.Configure root, templateId
    stagingPath = NxTemplateStagingPath(operationId, root)
    finalFolder = NxTemplateTemplateFolderPath(templateId, root)
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FolderExists(stagingPath) Or fileSystem.FolderExists(finalFolder) Then NxRaiseContractError "Template folder identity already exists"
    fileSystem.CreateFolder stagingPath
    transaction.TrackCreatedFolder stagingPath
    transaction.TrackCreatedFolder finalFolder
    packagePath = stagingPath & "\template.xlsx"
    metadataPath = stagingPath & "\metadata.ini"
    logicalSha = NxTemplateRequestLogicalHash(request)
    packageSha = NxTemplateBuildPackage(request, analysis, templateId, packagePath, logicalSha)
    createdUtc = TemplateTimestamp()
    Set record = New CNxTemplateRecord
    record.Configure templateId, request.DisplayName, request.Description, packageSha, logicalSha, request.SourceKind, _
        rowCount, columnCount, JoinCollection(riskFlags, ","), "healthy", createdUtc, createdUtc, request.Category
    record.Seal
    NxTemplateWriteMetadata metadataPath, record
    Set verifyRecord = NxTemplateReadMetadata(metadataPath)
    If Not RecordsEqual(record, verifyRecord) Then NxRaiseContractError "Template staging metadata read-back changed"
    If Not NxTemplateFolderHasExactFiles(stagingPath) Then NxRaiseContractError "Template staging file set is invalid"
    fileSystem.MoveFolder stagingPath, finalFolder
    transaction.Commit
    Set RegisterTemplate = NxCreateResult(request.FeatureId, NxSuccess, "complete", finalFolder & "\template.xlsx", False, vbNullString, "result_success")
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    transaction.Rollback
    If failureNumber = NX_TEMPLATE_CANCELLED Then
        Set RegisterTemplate = NxCreateResult(SafeFeatureId(request, "NX-TPL-REGISTER-SHEET"), NxCancelled, "cancel", root, False, failureDescription, "result_cancelled")
        Exit Function
    End If
    Set RegisterTemplate = NxCreateResult(SafeFeatureId(request, "NX-TPL-REGISTER-SHEET"), NxEnvironmentError, "register", root, False, _
        CStr(failureNumber) & " " & failureDescription, "result_environment_error")
End Function

Private Function ListTemplates(ByVal root As String) As Collection
    Set ListTemplates = NxTemplateListMetadata(root)
End Function

Private Function LoadTemplate(ByVal templateId As String, ByVal targetBook As Workbook, ByVal riskConfirmed As Boolean, ByVal root As String) As CNxResult
    Dim record As CNxTemplateRecord, packageBook As Workbook, target As Worksheet, transaction As New CNxTemplateTransaction
    Dim folderPath As String, packagePath As String, priorSecurity As MsoAutomationSecurity, targetName As String
    Dim failureNumber As Long, failureDescription As String
    Dim packageSheet As Worksheet, priorEvents As Boolean, loadedCount As Long, schema As String, usageWarning As String
    priorSecurity = Application.AutomationSecurity
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    If targetBook Is Nothing Then NxRaiseContractError "Template load target workbook is required"
    If targetBook.IsAddin Or targetBook Is ThisWorkbook Then NxRaiseContractError "일반 통합문서에만 템플릿을 추가할 수 있습니다."
    If targetBook.ProtectStructure Or targetBook.ReadOnly Then NxRaiseContractError "읽기 전용 또는 구조 보호 통합문서에는 템플릿을 추가할 수 없습니다."
    folderPath = NxTemplateTemplateFolderPath(templateId, root)
    Set record = NxTemplateReadRecordFolder(folderPath, False)
    If record Is Nothing Then NxRaiseContractError "Template package integrity validation failed"
    If record.Status <> "healthy" Then NxRaiseContractError "Template package integrity validation failed"
    If (Len(record.RiskFlags) > 0 Or CDbl(record.RowCount) * CDbl(record.ColumnCount) > 10000#) And Not riskConfirmed Then
        Set LoadTemplate = NxCreateResult("NX-TPL-LOAD", NxCancelled, "consent", targetBook.Name, False, "Planned confirmation is required", "result_cancelled")
        Exit Function
    End If
    packagePath = folderPath & "\template.xlsx"
    transaction.Configure root, templateId
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    NxTemplateProgress "패키지 열기", 0, 1
    Set packageBook = Application.Workbooks.Open(packagePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    schema = CStr(packageBook.CustomDocumentProperties("LHexcelTemplateSchema").Value)
    For Each packageSheet In packageBook.Worksheets
        If schema = "1" Then targetName = NextSheetName(targetBook, record.DisplayName) Else targetName = NextSheetName(targetBook, packageSheet.Name)
        Set target = targetBook.Worksheets.Add(After:=targetBook.Worksheets(targetBook.Worksheets.Count))
        transaction.TrackCreatedSheet target
        target.Name = targetName
        RestorePackageToSheet packageSheet, target
        loadedCount = loadedCount + 1
        If mFailurePoint = "after-load-cells" Then NxRaiseContractError "Injected template load failure"
        If mFailurePoint = "after-second-load-sheet" And loadedCount = 2 Then NxRaiseContractError "Injected multi-sheet template load failure"
    Next packageSheet
    packageBook.Close SaveChanges:=False
    Set packageBook = Nothing
    Application.AutomationSecurity = priorSecurity
    Application.EnableEvents = priorEvents
    usageWarning = NxTemplateRecordUse(record, root)
    transaction.Commit
    If Len(usageWarning) > 0 Then Debug.Print usageWarning
    Set LoadTemplate = NxCreateResult("NX-TPL-LOAD", NxSuccess, "complete", target.Name, True, vbNullString, "result_success")
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not packageBook Is Nothing Then packageBook.Close SaveChanges:=False
    Application.AutomationSecurity = priorSecurity
    transaction.Rollback
    Application.EnableEvents = priorEvents
    On Error GoTo 0
    If failureNumber = NX_TEMPLATE_CANCELLED Then
        Set LoadTemplate = NxCreateResult("NX-TPL-LOAD", NxCancelled, "cancel", SafeWorkbookTarget(targetBook), False, failureDescription, "result_cancelled")
        Exit Function
    End If
    Set LoadTemplate = NxCreateResult("NX-TPL-LOAD", NxEnvironmentError, "load", SafeWorkbookTarget(targetBook), False, _
        CStr(failureNumber) & " " & failureDescription, "result_environment_error")
End Function

Public Function NxTemplateOpenCopy(ByVal templateId As String, ByVal riskConfirmed As Boolean) As CNxResult
    Set NxTemplateOpenCopy = OpenTemplateCopy(templateId, riskConfirmed, NxTemplateStoreRoot())
End Function

Public Function NxTemplateOpenCopyForTest(ByVal templateId As String, ByVal riskConfirmed As Boolean, ByVal root As String) As CNxResult
    Set NxTemplateOpenCopyForTest = OpenTemplateCopy(templateId, riskConfirmed, root)
End Function

Private Function OpenTemplateCopy(ByVal templateId As String, ByVal riskConfirmed As Boolean, ByVal root As String) As CNxResult
    Dim copyBook As Workbook, blankSheet As Worksheet, result As CNxResult
    Dim priorAlerts As Boolean, priorEvents As Boolean, failureText As String
    priorAlerts = Application.DisplayAlerts
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    Set copyBook = Application.Workbooks.Add(xlWBATWorksheet)
    Set blankSheet = copyBook.Worksheets(1)
    blankSheet.Name = "nx-" & Left$(Replace(NxCreateRunUuid(), "-", vbNullString), 28)
    Set result = LoadTemplate(templateId, copyBook, riskConfirmed, root)
    If result.Outcome <> NxSuccess Then
        copyBook.Close SaveChanges:=False
        Set copyBook = Nothing
        Set OpenTemplateCopy = result
    Else
        Application.DisplayAlerts = False
        blankSheet.Delete
        Application.DisplayAlerts = priorAlerts
        copyBook.Saved = False
        Set OpenTemplateCopy = NxCreateResult("NX-TPL-LOAD", NxSuccess, "complete", copyBook.Name, False, vbNullString, "result_success")
        copyBook.Activate
    End If
    Application.EnableEvents = priorEvents
    Exit Function
Failed:
    failureText = Err.Description
    On Error Resume Next
    If Not copyBook Is Nothing Then copyBook.Close SaveChanges:=False
    Application.DisplayAlerts = priorAlerts
    Application.EnableEvents = priorEvents
    On Error GoTo 0
    Set OpenTemplateCopy = NxCreateResult("NX-TPL-LOAD", NxEnvironmentError, "copy", templateId, False, failureText, "result_environment_error")
End Function

Public Function NxTemplateBuildCatalog(ByVal targetBook As Workbook) As CNxResult
    Set NxTemplateBuildCatalog = BuildTemplateCatalog(targetBook, NxTemplateStoreRoot())
End Function

Public Function NxTemplateBuildCatalogForTest(ByVal targetBook As Workbook, ByVal root As String) As CNxResult
    Set NxTemplateBuildCatalogForTest = BuildTemplateCatalog(targetBook, root)
End Function

Private Function BuildTemplateCatalog(ByVal targetBook As Workbook, ByVal root As String) As CNxResult
    Dim records As Collection, record As CNxTemplateRecord, sheet As Worksheet, rowIndex As Long
    Dim transaction As New CNxTemplateTransaction, failureText As String, priorEvents As Boolean
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    If targetBook Is Nothing Then NxRaiseContractError "관리표를 추가할 통합문서가 없습니다."
    If targetBook.IsAddin Or targetBook Is ThisWorkbook Then NxRaiseContractError "일반 통합문서만 사용할 수 있습니다."
    If targetBook.ReadOnly Or targetBook.ProtectStructure Then NxRaiseContractError "읽기 전용 또는 구조 보호 통합문서에는 관리표를 추가할 수 없습니다."
    Set records = ListTemplates(root)
    transaction.Configure root, "catalog-" & Replace(NxCreateRunUuid(), "-", vbNullString)
    Application.EnableEvents = False
    Set sheet = targetBook.Worksheets.Add(After:=targetBook.Worksheets(targetBook.Worksheets.Count))
    transaction.TrackCreatedSheet sheet
    sheet.Name = NextSheetName(targetBook, "내엑셀_템플릿_관리")
    sheet.Range("A1:F" & CStr(records.Count + 1)).NumberFormat = "@"
    sheet.Range("A1:F1").Value2 = Array("템플릿명", "설명", "원본 종류", "상태", "수정일", "패키지 경로")
    rowIndex = 2
    For Each record In records
        sheet.Cells(rowIndex, 1).Value2 = record.DisplayName
        sheet.Cells(rowIndex, 2).Value2 = record.Description
        sheet.Cells(rowIndex, 3).Value2 = record.SourceKind
        sheet.Cells(rowIndex, 4).Value2 = record.Status
        sheet.Cells(rowIndex, 5).Value2 = record.ModifiedUtc
        sheet.Cells(rowIndex, 6).Value2 = NxTemplatePackagePath(record.TemplateId, root)
        rowIndex = rowIndex + 1
    Next record
    sheet.Range("A1:F1").Font.Bold = True
    sheet.Columns("A:F").ColumnWidth = 24
    sheet.Columns("B").ColumnWidth = 36
    sheet.Columns("F").ColumnWidth = 60
    sheet.Range("A1:F" & CStr(rowIndex - 1)).WrapText = True
    sheet.Range("A1:F" & CStr(rowIndex - 1)).AutoFilter
    transaction.Commit
    Application.EnableEvents = priorEvents
    Set BuildTemplateCatalog = NxCreateResult("NX-TPL-LOAD", NxSuccess, "complete", sheet.Name, True, vbNullString, "result_success")
    Exit Function
Failed:
    failureText = Err.Description
    transaction.Rollback
    Application.EnableEvents = priorEvents
    Set BuildTemplateCatalog = NxCreateResult("NX-TPL-LOAD", NxEnvironmentError, "catalog", SafeWorkbookTarget(targetBook), False, failureText, "result_environment_error")
End Function

Private Function RenameTemplate(ByVal templateId As String, ByVal displayName As String, ByVal description As String, ByVal root As String, Optional ByVal category As Variant) As CNxResult
    Dim records As Collection, item As Variant, record As CNxTemplateRecord, changed As CNxTemplateRecord
    Dim modifiedUtc As String, failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    If Len(Trim$(displayName)) = 0 Then NxRaiseContractError "Template display name is required"
    Set records = NxTemplateListMetadata(root, False)
    For Each item In records
        Set record = item
        If record.TemplateId <> templateId And StrComp(record.DisplayName, displayName, vbTextCompare) = 0 Then NxRaiseContractError "Template display name already exists"
    Next item
    Set record = NxTemplateReadRecordFolder(NxTemplateTemplateFolderPath(templateId, root), False)
    If record Is Nothing Then NxRaiseContractError "Template ID was not found or is damaged"
    If record.Status <> "healthy" Then NxRaiseContractError "Template ID was not found or is damaged"
    modifiedUtc = TemplateTimestamp()
    If IsMissing(category) Or IsEmpty(category) Then category = record.Category
    Set changed = NxTemplateCopyDetails(record, Trim$(displayName), Trim$(description), CStr(category), record.LastUsedUtc, modifiedUtc)
    NxTemplateReplaceMetadata NxTemplateTemplateFolderPath(templateId, root), changed
    Set RenameTemplate = NxCreateResult("NX-TPL-RENAME", NxSuccess, "complete", templateId, False, vbNullString, "result_success")
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Set RenameTemplate = NxCreateResult("NX-TPL-RENAME", NxEnvironmentError, "rename", templateId, False, CStr(failureNumber) & " " & failureDescription, "result_environment_error")
End Function

Private Function DeleteTemplate(ByVal templateId As String, ByVal root As String) As CNxResult
    Dim record As CNxTemplateRecord
    Dim folderPath As String, quarantineRoot As String, quarantineFolder As String
    Dim fileSystem As Object, moved As Boolean, restored As Boolean, failureNumber As Long, failureDescription As String
    On Error GoTo Failed
    folderPath = NxTemplateTemplateFolderPath(templateId, root)
    Set record = NxTemplateReadRecordFolder(folderPath, True)
    quarantineRoot = NxTemplateQuarantinePath("quarantine-" & Replace(NxCreateRunUuid(), "-", vbNullString), root)
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    fileSystem.CreateFolder quarantineRoot
    quarantineFolder = quarantineRoot & "\" & templateId
    fileSystem.MoveFolder folderPath, quarantineFolder
    moved = True
    If mFailurePoint = "after-quarantine-move" Then NxRaiseContractError "Injected template quarantine failure"
    If Not NxTemplateFolderHasExactFiles(quarantineFolder) Then NxRaiseContractError "Quarantine file set mismatch"
    If record.Status = "healthy" Then
        If NxTemplateFileSha256(quarantineFolder & "\template.xlsx") <> record.PackageSha256 Then NxRaiseContractError "Quarantine package hash mismatch"
    End If
    Set DeleteTemplate = NxCreateResult("NX-TPL-DELETE", NxSuccess, "complete", quarantineFolder, False, vbNullString, "result_success")
    Exit Function
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not fileSystem Is Nothing Then
        If moved Then
            If fileSystem.FolderExists(quarantineFolder) And Not fileSystem.FolderExists(folderPath) Then
                fileSystem.MoveFolder quarantineFolder, folderPath
                restored = fileSystem.FolderExists(folderPath)
            End If
        End If
    End If
    If Not fileSystem Is Nothing Then
        If fileSystem.FolderExists(quarantineRoot) And (Not moved Or restored) Then fileSystem.DeleteFolder quarantineRoot, True
    End If
    On Error GoTo 0
    Set DeleteTemplate = NxCreateResult("NX-TPL-DELETE", NxEnvironmentError, "delete", templateId, False, CStr(failureNumber) & " " & failureDescription, "result_environment_error")
End Function

Private Sub RestorePackageToSheet(ByVal source As Worksheet, ByVal target As Worksheet)
    Dim sourceRange As Range, sourceCell As Range, targetCell As Range, rowIndex As Long, columnIndex As Long
    Dim targetRange As Range
    Dim completed As Long, total As Long
    Set sourceRange = source.UsedRange
    total = CLng(sourceRange.CountLarge)
    NxTemplateProgress "서식 불러오기", 0, total
    Set targetRange = target.Range("A1").Resize(sourceRange.Rows.Count, sourceRange.Columns.Count)
    ' Only the validated, application-generated package reaches this path.
    ' Transfer formats once; write formulas explicitly to retain R1C1 semantics.
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    targetRange.UnMerge
    targetRange.FormatConditions.Delete
    For Each sourceCell In sourceRange.Cells
        If completed Mod 256 = 0 Then NxTemplateProgress "내용 불러오기", completed, total
        Set targetCell = target.Cells(sourceCell.Row - sourceRange.Row + 1, sourceCell.Column - sourceRange.Column + 1)
        If sourceCell.HasFormula Then
            targetCell.FormulaR1C1 = sourceCell.FormulaR1C1
        Else
            If VarType(sourceCell.Value2) = vbString Then targetCell.NumberFormat = "@"
            targetCell.Value2 = sourceCell.Value2
        End If
        If VarType(sourceCell.Value2) = vbString And Not sourceCell.HasFormula Then _
            targetCell.NumberFormat = sourceCell.NumberFormat
        completed = completed + 1
    Next sourceCell
    NxTemplateProgress "내용 불러오기", total, total
    For rowIndex = 1 To sourceRange.Rows.Count
        target.Rows(rowIndex).RowHeight = sourceRange.Rows(rowIndex).RowHeight
    Next rowIndex
    For columnIndex = 1 To sourceRange.Columns.Count
        target.Columns(columnIndex).ColumnWidth = sourceRange.Columns(columnIndex).ColumnWidth
    Next columnIndex
    RestoreMerges sourceRange, target.Range("A1")
    NxTemplateCopyPrintSettings sourceRange, target
End Sub

Private Sub CopyCellFormat(ByVal sourceCell As Range, ByVal targetCell As Range)
    Dim borderIndex As Variant
    On Error Resume Next
    targetCell.NumberFormat = sourceCell.NumberFormat
    targetCell.Font.Name = sourceCell.Font.Name: targetCell.Font.Size = sourceCell.Font.Size
    targetCell.Font.Bold = sourceCell.Font.Bold: targetCell.Font.Italic = sourceCell.Font.Italic
    targetCell.Font.Underline = sourceCell.Font.Underline: targetCell.Font.Color = sourceCell.Font.Color
    targetCell.Interior.Pattern = sourceCell.Interior.Pattern: targetCell.Interior.Color = sourceCell.Interior.Color
    targetCell.HorizontalAlignment = sourceCell.HorizontalAlignment: targetCell.VerticalAlignment = sourceCell.VerticalAlignment
    targetCell.WrapText = sourceCell.WrapText: targetCell.Orientation = sourceCell.Orientation: targetCell.IndentLevel = sourceCell.IndentLevel
    For Each borderIndex In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
        targetCell.Borders(CLng(borderIndex)).LineStyle = sourceCell.Borders(CLng(borderIndex)).LineStyle
        targetCell.Borders(CLng(borderIndex)).Weight = sourceCell.Borders(CLng(borderIndex)).Weight
        targetCell.Borders(CLng(borderIndex)).Color = sourceCell.Borders(CLng(borderIndex)).Color
    Next borderIndex
    On Error GoTo 0
End Sub

Private Sub RestoreMerges(ByVal source As Range, ByVal targetTopLeft As Range)
    Dim cell As Range, mergeArea As Range
    Dim merged As Variant
    merged = source.MergeCells
    If Not IsNull(merged) Then
        If merged = False Then Exit Sub
    End If
    For Each cell In source.Cells
        If cell.MergeCells Then
            Set mergeArea = cell.MergeArea
            If cell.Address = mergeArea.Cells(1, 1).Address Then _
                targetTopLeft.Offset(cell.Row - source.Row, cell.Column - source.Column).Resize(mergeArea.Rows.Count, mergeArea.Columns.Count).Merge
        End If
    Next cell
End Sub

Private Function NextSheetName(ByVal book As Workbook, ByVal displayName As String) As String
    Dim baseName As String, candidate As String, suffix As Long, nextNumber As Long
    baseName = CleanSheetName(displayName)
    If Len(baseName) = 0 Then baseName = "Template"
    candidate = Left$(baseName, 31)
    Do While SheetExists(book, candidate)
        suffix = suffix + 1
        nextNumber = suffix + 1
        candidate = Left$(baseName, 31 - Len(CStr(nextNumber)) - 3) & " (" & CStr(nextNumber) & ")"
    Loop
    NextSheetName = candidate
End Function

Private Function CleanSheetName(ByVal value As String) As String
    Dim forbidden As Variant
    CleanSheetName = Trim$(value)
    For Each forbidden In Array("\", "/", "?", "*", "[", "]", ":")
        CleanSheetName = Replace(CleanSheetName, CStr(forbidden), "-")
    Next forbidden
End Function

Private Function SheetExists(ByVal book As Workbook, ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = book.Worksheets(sheetName)
    SheetExists = Not sheet Is Nothing
    On Error GoTo 0
End Function

Private Function FindRecord(ByVal records As Collection, ByVal templateId As String) As CNxTemplateRecord
    Dim item As Variant, record As CNxTemplateRecord
    For Each item In records
        Set record = item
        If record.TemplateId = templateId Then Set FindRecord = record: Exit Function
    Next item
End Function

Private Function CopyRecord(ByVal source As CNxTemplateRecord, ByVal displayName As String, ByVal description As String, ByVal status As String, ByVal modifiedUtc As String) As CNxTemplateRecord
    Dim result As New CNxTemplateRecord
    result.Configure source.TemplateId, displayName, description, source.PackageSha256, source.LogicalSha256, source.SourceKind, _
        source.RowCount, source.ColumnCount, source.RiskFlags, status, source.CreatedUtc, modifiedUtc, source.Category, source.LastUsedUtc
    result.Seal
    Set CopyRecord = result
End Function

Private Function RecordsEqual(ByVal expected As CNxTemplateRecord, ByVal actual As CNxTemplateRecord) As Boolean
    If expected Is Nothing Or actual Is Nothing Then Exit Function
    RecordsEqual = expected.TemplateId = actual.TemplateId And expected.DisplayName = actual.DisplayName And _
        expected.Description = actual.Description And expected.PackageSha256 = actual.PackageSha256 And _
        expected.LogicalSha256 = actual.LogicalSha256 And expected.SourceKind = actual.SourceKind And _
        expected.RowCount = actual.RowCount And expected.ColumnCount = actual.ColumnCount And _
        expected.RiskFlags = actual.RiskFlags And expected.Status = actual.Status And _
        expected.CreatedUtc = actual.CreatedUtc And expected.ModifiedUtc = actual.ModifiedUtc
End Function

Private Function JoinCollection(ByVal values As Collection, ByVal delimiter As String) As String
    Dim item As Variant
    For Each item In values
        If Len(JoinCollection) > 0 Then JoinCollection = JoinCollection & delimiter
        JoinCollection = JoinCollection & CStr(item)
    Next item
End Function

Private Function TemplateTimestamp() As String
    TemplateTimestamp = Format$(Now, "yyyy-mm-dd\Thh:nn:ss") & "Z"
End Function

Private Function SafeFeatureId(ByVal request As CNxTemplateRequest, ByVal fallback As String) As String
    On Error GoTo UseFallback
    SafeFeatureId = request.FeatureId
    Exit Function
UseFallback:
    SafeFeatureId = fallback
End Function

Private Function SafeWorkbookTarget(ByVal book As Workbook) As String
    On Error GoTo UseFallback
    SafeWorkbookTarget = book.Name
    Exit Function
UseFallback:
    SafeWorkbookTarget = "template-load-target"
End Function
