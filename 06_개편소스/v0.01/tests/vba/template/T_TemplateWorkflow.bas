Attribute VB_Name = "T_TemplateWorkflow"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestRegisterRangeRecordsSourceKind"
    names.Add "TestRegisterSheetUsesUsedRange"
    names.Add "TestDuplicateNameNeverOverwrites"
    names.Add "TestListIsLocalFastWithoutPrivacyScan"
    names.Add "TestLoadCreatesNewSheetAtA1"
    names.Add "TestRiskyOrFailedLoadRequiresConsentAndRollsBack"
    names.Add "TestRenameKeepsTemplateIdAndPackageName"
    names.Add "TestDeleteQuarantinesAtomically"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestRegisterRangeRecordsSourceKind": TestRegisterRangeRecordsSourceKind
        Case "TestRegisterSheetUsesUsedRange": TestRegisterSheetUsesUsedRange
        Case "TestDuplicateNameNeverOverwrites": TestDuplicateNameNeverOverwrites
        Case "TestListIsLocalFastWithoutPrivacyScan": TestListIsLocalFastWithoutPrivacyScan
        Case "TestLoadCreatesNewSheetAtA1": TestLoadCreatesNewSheetAtA1
        Case "TestRiskyOrFailedLoadRequiresConsentAndRollsBack": TestRiskyOrFailedLoadRequiresConsentAndRollsBack
        Case "TestRenameKeepsTemplateIdAndPackageName": TestRenameKeepsTemplateIdAndPackageName
        Case "TestDeleteQuarantinesAtomically": TestDeleteQuarantinesAtomically
        Case Else: NxRaiseContractError "Unknown Template workflow test"
    End Select
End Sub

Private Sub TestRegisterRangeRecordsSourceKind()
    Dim root As String, book As Workbook, request As CNxTemplateRequest, result As CNxResult, records As Collection
    root = NewWorkflowRoot()
    Set request = BuildWorkflowRequest(book, "NX-TPL-REGISTER-RANGE", "range", "Range template", False)
    Set result = NxTemplateRegisterForTest(request, root)
    Set records = NxTemplateListForTest(root)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Range template registration failed"
    NxTestHarness.AssertTrue records.Count = 1 And records.Item(1).SourceKind = "range", "Range source kind was not recorded"
    AssertFolderRecord root, records.Item(1)
    CleanupWorkflow root, book
End Sub

Private Sub TestRegisterSheetUsesUsedRange()
    Dim root As String, book As Workbook, request As CNxTemplateRequest, result As CNxResult, records As Collection
    root = NewWorkflowRoot()
    Set request = BuildWorkflowRequest(book, "NX-TPL-REGISTER-SHEET", "sheet_used_range", "Sheet template", False)
    Set result = NxTemplateRegisterForTest(request, root)
    Set records = NxTemplateListForTest(root)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Sheet template registration failed"
    NxTestHarness.AssertTrue records.Item(1).SourceKind = "sheet_used_range" And records.Item(1).RowCount = 2 And records.Item(1).ColumnCount = 2, "Sheet registration did not use UsedRange"
    AssertFolderRecord root, records.Item(1)
    CleanupWorkflow root, book
End Sub

Private Sub TestDuplicateNameNeverOverwrites()
    Dim root As String, bookOne As Workbook, bookTwo As Workbook, firstRequest As CNxTemplateRequest, secondRequest As CNxTemplateRequest
    Dim firstResult As CNxResult, secondResult As CNxResult, beforeSha As String, records As Collection
    root = NewWorkflowRoot()
    Set firstRequest = BuildWorkflowRequest(bookOne, "NX-TPL-REGISTER-RANGE", "range", "Duplicate", False)
    Set secondRequest = BuildWorkflowRequest(bookTwo, "NX-TPL-REGISTER-RANGE", "range", "duplicate", False)
    Set firstResult = NxTemplateRegisterForTest(firstRequest, root)
    Set records = NxTemplateListForTest(root)
    beforeSha = records.Item(1).PackageSha256
    Set secondResult = NxTemplateRegisterForTest(secondRequest, root)
    Set records = NxTemplateListForTest(root)
    NxTestHarness.AssertTrue firstResult.Outcome = NxSuccess And secondResult.Outcome <> NxSuccess, "Duplicate template name was accepted"
    NxTestHarness.AssertTrue records.Count = 1 And records.Item(1).PackageSha256 = beforeSha, "Duplicate registration overwrote the prior package"
    CleanupWorkflow root, bookOne
    On Error Resume Next: bookTwo.Close False: On Error GoTo 0
End Sub

Private Sub TestListIsLocalFastWithoutPrivacyScan()
    Dim root As String, book As Workbook, request As CNxTemplateRequest, records As Collection
    root = NewWorkflowRoot()
    Set request = BuildWorkflowRequest(book, "NX-TPL-REGISTER-RANGE", "range", "Local list", False)
    Call NxTemplateRegisterForTest(request, root)
    Set records = NxTemplateListForTest(root)
    NxTestHarness.AssertTrue records.Count = 1 And records.Item(1).Status = "healthy", "Local template list failed"
    CleanupWorkflow root, book
End Sub

Private Sub TestLoadCreatesNewSheetAtA1()
    Dim root As String, sourceBook As Workbook, targetBook As Workbook, request As CNxTemplateRequest
    Dim records As Collection, result As CNxResult, beforeCount As Long
    root = NewWorkflowRoot()
    Set request = BuildWorkflowRequest(sourceBook, "NX-TPL-REGISTER-RANGE", "range", "Load target", False)
    Call NxTemplateRegisterForTest(request, root)
    Set records = NxTemplateListForTest(root)
    Set targetBook = Application.Workbooks.Add
    beforeCount = targetBook.Worksheets.Count
    Set result = NxTemplateLoadForTest(records.Item(1).TemplateId, targetBook, True, root)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Template load failed"
    NxTestHarness.AssertTrue targetBook.Worksheets.Count = beforeCount + 1, "Load did not create exactly one sheet"
    NxTestHarness.AssertTrue targetBook.Worksheets(beforeCount + 1).Range("A1").Value2 = "kept", "Load target is not A1"
    CleanupWorkflow root, sourceBook
    On Error Resume Next: targetBook.Close False: On Error GoTo 0
End Sub

Private Sub TestRiskyOrFailedLoadRequiresConsentAndRollsBack()
    Dim root As String, sourceBook As Workbook, targetBook As Workbook, request As CNxTemplateRequest
    Dim records As Collection, result As CNxResult, beforeCount As Long
    root = NewWorkflowRoot()
    Set request = BuildWorkflowRequest(sourceBook, "NX-TPL-REGISTER-RANGE", "range", "Risky load", True)
    Call NxTemplateRegisterForTest(request, root)
    Set records = NxTemplateListForTest(root)
    Set targetBook = Application.Workbooks.Add
    beforeCount = targetBook.Worksheets.Count
    Set result = NxTemplateLoadForTest(records.Item(1).TemplateId, targetBook, False, root)
    NxTestHarness.AssertTrue result.Outcome = NxCancelled And targetBook.Worksheets.Count = beforeCount, "Risky load did not require consent"
    NxTemplateSetFailurePoint "after-load-cells"
    Set result = NxTemplateLoadForTest(records.Item(1).TemplateId, targetBook, True, root)
    NxTemplateClearFailurePoint
    NxTestHarness.AssertTrue result.Outcome <> NxSuccess, "Injected load failure succeeded"
    NxTestHarness.AssertTrue targetBook.Worksheets.Count = beforeCount, "Failed load left a created sheet"
    CleanupWorkflow root, sourceBook
    On Error Resume Next: targetBook.Close False: On Error GoTo 0
End Sub

Private Sub TestRenameKeepsTemplateIdAndPackageName()
    Dim root As String, book As Workbook, request As CNxTemplateRequest, records As Collection
    Dim templateId As String, packagePath As String, result As CNxResult
    root = NewWorkflowRoot()
    Set request = BuildWorkflowRequest(book, "NX-TPL-REGISTER-RANGE", "range", "Before rename", False)
    Call NxTemplateRegisterForTest(request, root)
    Set records = NxTemplateListForTest(root)
    templateId = records.Item(1).TemplateId
    packagePath = NxTemplatePackagePath(templateId, root)
    Set result = NxTemplateRenameForTest(templateId, "After rename", "Changed description", root)
    Set records = NxTemplateListForTest(root)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Template rename failed"
    NxTestHarness.AssertTrue records.Item(1).TemplateId = templateId And records.Item(1).DisplayName = "After rename", "Rename changed template identity"
    NxTestHarness.AssertTrue Len(Dir$(packagePath)) > 0, "Rename changed the package filename"
    CleanupWorkflow root, book
End Sub

Private Sub TestDeleteQuarantinesAtomically()
    Dim root As String, book As Workbook, request As CNxTemplateRequest, records As Collection
    Dim templateId As String, folderPath As String, result As CNxResult
    Dim fileSystem As Object
    root = NewWorkflowRoot()
    Set request = BuildWorkflowRequest(book, "NX-TPL-REGISTER-RANGE", "range", "Delete target", False)
    Call NxTemplateRegisterForTest(request, root)
    Set records = NxTemplateListForTest(root)
    templateId = records.Item(1).TemplateId
    folderPath = NxTemplateTemplateFolderPath(templateId, root)
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    NxTemplateSetFailurePoint "after-quarantine-move"
    Set result = NxTemplateDeleteForTest(templateId, root)
    NxTemplateClearFailurePoint
    NxTestHarness.AssertTrue result.Outcome <> NxSuccess And fileSystem.FolderExists(folderPath), "Failed quarantine did not restore the active folder"
    NxTestHarness.AssertTrue NxTemplateFolderHasExactFiles(folderPath), "Failed quarantine did not restore both active files"
    NxTestHarness.AssertTrue NxTemplateListForTest(root).Count = 1, "Failed quarantine changed the active folder set"
    Set result = NxTemplateDeleteForTest(templateId, root)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess And Not fileSystem.FolderExists(folderPath), "Template folder was not quarantined"
    NxTestHarness.AssertTrue NxTemplateFolderHasExactFiles(result.Target) And NxTemplateListForTest(root).Count = 0, "Quarantine folder or active folder set is invalid"
    CleanupWorkflow root, book
End Sub

Private Function BuildWorkflowRequest(ByRef book As Workbook, ByVal featureId As String, ByVal sourceKind As String, ByVal displayName As String, ByVal risky As Boolean) As CNxTemplateRequest
    Dim sheet As Worksheet, source As Range, retained As New Collection
    Dim analysis As CNxTemplateAnalysis, request As New CNxTemplateRequest
    Set book = Application.Workbooks.Add
    Set sheet = book.Worksheets(1)
    sheet.Range("A1").Value2 = "kept"
    sheet.Range("B1").Value2 = "excluded"
    If risky Then sheet.Range("A2").Formula = "=TODAY()" Else sheet.Range("A2").FormulaR1C1 = "=R[-1]C"
    sheet.Range("B2").Value2 = 2
    Set source = sheet.UsedRange
    retained.Add sheet.Range("A1")
    Set analysis = NxTemplateAnalyze(source, retained, False, False)
    request.ConfigureRegistration featureId, sourceKind, source, retained, displayName, "Workflow fixture", False, False, risky
    request.BindAnalysis analysis
    Set BuildWorkflowRequest = request
End Function

Private Function NewWorkflowRoot() As String
    NewWorkflowRoot = Environ$("TEMP") & "\LHexcelWorkflow-" & Replace(NxCreateRunUuid(), "-", vbNullString)
End Function

Private Sub AssertFolderRecord(ByVal root As String, ByVal record As CNxTemplateRecord)
    Dim fileSystem As Object, folderPath As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    folderPath = NxTemplateTemplateFolderPath(record.TemplateId, root)
    NxTestHarness.AssertTrue fileSystem.FolderExists(folderPath), "Template folder missing"
    NxTestHarness.AssertTrue NxTemplateFolderHasExactFiles(folderPath), "Template folder file set mismatch"
    NxTestHarness.AssertTrue Len(Dir$(root & "\index.ini")) = 0, "Global index must not be created"
End Sub

Private Sub CleanupWorkflow(ByVal root As String, ByVal book As Workbook)
    Dim fileSystem As Object
    On Error Resume Next
    NxTemplateClearFailurePoint
    If Not book Is Nothing Then book.Close False
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FolderExists(root) Then fileSystem.DeleteFolder root, True
    On Error GoTo 0
End Sub
