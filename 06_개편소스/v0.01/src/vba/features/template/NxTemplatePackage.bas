Attribute VB_Name = "NxTemplatePackage"
Option Explicit

Public Function NxTemplateBuildPackage(ByVal request As CNxTemplateRequest, ByVal analysis As CNxTemplateAnalysis, _
    ByVal templateId As String, ByVal packagePath As String, Optional ByVal sourceLogicalSha As String = vbNullString) As String

    Dim book As Workbook, target As Worksheet, priorSecurity As MsoAutomationSecurity
    Dim priorAlerts As Boolean, logicalSha As String
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    Dim source As Range, sheetIndex As Long, schema As String
    Dim ownedPath As Boolean, fileSystem As Object, priorEvents As Boolean
    If request Is Nothing Or analysis Is Nothing Then NxRaiseContractError "Template package requires request and analysis"
    If Not analysis.IsSealed Or Not NxTemplateIdIsValid(templateId) Then NxRaiseContractError "Template package identity is invalid"
    TraceBuildProgress "01-enter"
    logicalSha = sourceLogicalSha
    If Len(logicalSha) = 0 Then logicalSha = NxTemplateRequestLogicalHash(request)
    TraceBuildProgress "02-logical-hash"
    priorSecurity = Application.AutomationSecurity
    priorAlerts = Application.DisplayAlerts
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FileExists(packagePath) Or fileSystem.FolderExists(packagePath) Then NxRaiseContractError "기존 템플릿 파일을 덮어쓸 수 없습니다."
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    TraceBuildProgress "03-workbook-created"
    schema = "1"
    If request.SourceKind = "workbook" Or request.SourceKind = "file" Then schema = "2"
    For Each source In request.SourceRanges
        sheetIndex = sheetIndex + 1
        If sheetIndex = 1 Then
            Set target = book.Worksheets(1)
        Else
            Set target = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
        End If
        If schema = "1" Then target.Name = "Template" Else target.Name = source.Worksheet.Name
        WriteCellsPropertyByProperty source, request.RetainedForSource(source), target.Range("A1")
        WriteDimensionsAndMerges source, target.Range("A1")
    Next source
    TraceBuildProgress "04-cells-written"
    TraceBuildProgress "05-layout-written"
    DeleteNamesLinksAndDocumentIdentity book
    sheetIndex = 0
    For Each source In request.SourceRanges
        sheetIndex = sheetIndex + 1
        NxTemplateCopyPrintSettings source, book.Worksheets(sheetIndex)
    Next source
    TraceBuildProgress "06-identity-stripped"
    AddClosedTemplateProperties book, templateId, logicalSha, schema
    TraceBuildProgress "07-properties-added"
    If fileSystem.FileExists(packagePath) Then NxRaiseContractError "기존 템플릿 파일을 덮어쓸 수 없습니다."
    ownedPath = True
    NxTemplateProgress "양식 파일 저장", 0, 1
    book.SaveAs Filename:=packagePath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False, AddToMru:=False
    TraceBuildProgress "08-package-saved"
    book.Close SaveChanges:=False
    Set book = Nothing
    TraceBuildProgress "09-package-closed"
    NxTemplateProgress "저장 결과 검사", 0, 1
    If Not NxTemplateClosedPackageValid(packagePath, templateId, logicalSha) Then NxRaiseContractError "Closed template package validation failed"
    TraceBuildProgress "10-package-validated"
    NxTemplateBuildPackage = NxTemplateFileSha256(packagePath)
    TraceBuildProgress "11-file-sha"
CleanExit:
    Application.DisplayAlerts = priorAlerts
    Application.AutomationSecurity = priorSecurity
    Application.EnableEvents = priorEvents
    Exit Function
Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    If ownedPath Then
        If Len(Dir$(packagePath)) > 0 Then Kill packagePath
    End If
    Application.DisplayAlerts = priorAlerts
    Application.AutomationSecurity = priorSecurity
    Application.EnableEvents = priorEvents
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Function

Public Function NxTemplateRequestLogicalHash(ByVal request As CNxTemplateRequest) As String
    Dim source As Range, canonical As String
    If request.SourceKind = "sheet_used_range" Then
        NxTemplateRequestLogicalHash = NxTemplateLogicalHash(request.SourceRange, request.RetainedAreas)
        Exit Function
    End If
    canonical = "LHexcelTemplateWorkbook|2|" & CStr(request.SourceRanges.Count) & vbLf
    For Each source In request.SourceRanges
        canonical = canonical & CStr(Len(source.Worksheet.Name)) & ":" & source.Worksheet.Name & "|" & _
            NxTemplateLogicalHash(source, request.RetainedForSource(source)) & vbLf
    Next source
    NxTemplateRequestLogicalHash = NxTemplateTextSha256(canonical)
End Function

Private Sub TraceBuildProgress(ByVal marker As String)
    Dim progressPath As String, handle As Integer
    progressPath = Environ$("LHEXCEL_OWNER_CASE_PROGRESS")
    If Len(progressPath) = 0 Then Exit Sub
    On Error GoTo CleanExit
    handle = FreeFile
    Open progressPath For Append Access Write As #handle
    Print #handle, marker
CleanExit:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    On Error GoTo 0
End Sub

Private Sub WriteCellsPropertyByProperty(ByVal source As Range, ByVal retained As Collection, ByVal targetTopLeft As Range)
    Dim sourceCell As Range, targetCell As Range
    Dim completed As Long, total As Long
    Dim targetRange As Range
    total = CLng(source.CountLarge)
    NxTemplateProgress "기본 서식 작성", 0, total
    Set targetRange = targetTopLeft.Resize(source.Rows.Count, source.Columns.Count)
    source.Copy
    targetRange.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    targetRange.FormatConditions.Delete
    If targetRange.FormatConditions.Count <> 0 Then NxRaiseContractError "조건부서식을 제외하지 못했습니다."
    targetRange.UnMerge
    For Each sourceCell In source.Cells
        If completed Mod 128 = 0 Then NxTemplateProgress "양식 작성", completed, total
        Set targetCell = targetTopLeft.Offset(sourceCell.Row - source.Row, sourceCell.Column - source.Column)
        WriteCellPropertyByProperty sourceCell, targetCell, retained
        completed = completed + 1
    Next sourceCell
    NxTemplateProgress "양식 작성", total, total
End Sub

Private Sub WriteCellPropertyByProperty(ByVal sourceCell As Range, ByVal targetCell As Range, ByVal retained As Collection)
    If sourceCell.HasFormula Then
        targetCell.FormulaR1C1 = sourceCell.FormulaR1C1
    ElseIf IsRetainedCell(sourceCell, retained) Then
        If VarType(sourceCell.Value2) = vbString Then
            targetCell.NumberFormat = "@"
            targetCell.Value2 = sourceCell.Value2
            targetCell.NumberFormat = sourceCell.NumberFormat
        Else
            targetCell.Value2 = sourceCell.Value2
        End If
    Else
        targetCell.ClearContents
    End If
End Sub

Private Sub WriteDimensionsAndMerges(ByVal source As Range, ByVal targetTopLeft As Range)
    Dim index As Long, cell As Range, mergeArea As Range, targetMerge As Range
    Dim merged As Variant
    For index = 1 To source.Rows.Count
        targetTopLeft.Offset(index - 1, 0).EntireRow.RowHeight = source.Rows(index).RowHeight
    Next index
    For index = 1 To source.Columns.Count
        targetTopLeft.Offset(0, index - 1).EntireColumn.ColumnWidth = source.Columns(index).ColumnWidth
    Next index
    merged = source.MergeCells
    If Not IsNull(merged) Then
        If merged = False Then Exit Sub
    End If
    For Each cell In source.Cells
        If cell.MergeCells Then
            Set mergeArea = cell.MergeArea
            If cell.Address = mergeArea.Cells(1, 1).Address Then
                If Not Application.Intersect(mergeArea, source) Is Nothing Then
                    If Application.Intersect(mergeArea, source).Address = mergeArea.Address Then
                        Set targetMerge = targetTopLeft.Offset(cell.Row - source.Row, cell.Column - source.Column).Resize(mergeArea.Rows.Count, mergeArea.Columns.Count)
                        targetMerge.Merge
                    End If
                End If
            End If
        End If
    Next cell
End Sub

Private Sub DeleteNamesLinksAndDocumentIdentity(ByVal book As Workbook)
    Dim index As Long, links As Variant
    On Error Resume Next
    For index = book.Names.Count To 1 Step -1
        book.Names(index).Delete
    Next index
    links = book.LinkSources(xlExcelLinks)
    If IsArray(links) Then
        For index = LBound(links) To UBound(links)
            book.BreakLink Name:=CStr(links(index)), Type:=xlLinkTypeExcelLinks
        Next index
    End If
    book.RemoveDocumentInformation xlRDIAll
    On Error GoTo 0
End Sub

Private Sub AddClosedTemplateProperties(ByVal book As Workbook, ByVal templateId As String, ByVal logicalSha As String, ByVal schema As String)
    AddStringProperty book, "LHexcelTemplateSchema", schema
    AddStringProperty book, "LHexcelTemplateId", templateId
    AddStringProperty book, "LHexcelLogicalSha256", logicalSha
    AddStringProperty book, "LHexcelOrigin", "user-created"
End Sub

Private Sub AddStringProperty(ByVal book As Workbook, ByVal propertyName As String, ByVal propertyValue As String)
    On Error Resume Next
    book.CustomDocumentProperties(propertyName).Delete
    On Error GoTo 0
    book.CustomDocumentProperties.Add Name:=propertyName, LinkToContent:=False, Type:=msoPropertyTypeString, Value:=propertyValue
End Sub

Private Function IsRetainedCell(ByVal cell As Range, ByVal retained As Collection) As Boolean
    Dim area As Variant, overlap As Range
    For Each area In retained
        Set overlap = Nothing
        On Error Resume Next
        Set overlap = Application.Intersect(cell, area)
        On Error GoTo 0
        If Not overlap Is Nothing Then IsRetainedCell = True: Exit Function
    Next area
End Function
