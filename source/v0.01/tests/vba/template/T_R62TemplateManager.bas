Attribute VB_Name = "T_R62TemplateManager"
Option Explicit

Private mRoot As String
Private mSource As Workbook
Private mTarget As Workbook
Private mCopy As Workbook
Private mForm As FNxTemplateManager
Private mRegisterForm As FNxTemplateRegister
Private priorSecurity As MsoAutomationSecurity
Private mPriorAlerts As Boolean
Private mPriorEvents As Boolean

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestR62TemplateWorkbookAndCopyPreserveOriginal"
    names.Add "TestR62TemplateSheetRegistrationUsesWholeSheet"
    names.Add "TestR62TemplateMultisheetFailureRollsBack"
    names.Add "TestR62TemplateCatalogIsLiteralAndCollisionSafe"
    names.Add "TestR62TemplateFileRegistrationClosesOwnedSource"
    names.Add "TestR62TemplateManagerControlsAreActionable"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Dim failureNumber As Long, failureText As String
    priorSecurity = Application.AutomationSecurity
    mPriorAlerts = Application.DisplayAlerts
    mPriorEvents = Application.EnableEvents
    On Error GoTo Failed
    NewFixture
    Select Case name
        Case "TestR62TemplateWorkbookAndCopyPreserveOriginal": TestWorkbookAndCopy
        Case "TestR62TemplateSheetRegistrationUsesWholeSheet": TestSheetRegistration
        Case "TestR62TemplateMultisheetFailureRollsBack": TestLoadRollback
        Case "TestR62TemplateCatalogIsLiteralAndCollisionSafe": TestCatalog
        Case "TestR62TemplateFileRegistrationClosesOwnedSource": TestFileRegistration
        Case "TestR62TemplateManagerControlsAreActionable": TestManagerControls
        Case Else: NxRaiseContractError "Unknown r62 Template manager test"
    End Select
    CleanupFixture
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureText = Err.Description
    CleanupFixture
    Err.Raise failureNumber, "T_R62TemplateManager." & name, failureText
End Sub

Private Sub NewFixture()
    Dim fileSystem As Object, sheet As Worksheet
    mRoot = Environ$("TEMP") & "\NxR62Template-" & Replace(NxCreateRunUuid(), "-", vbNullString)
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FolderExists(mRoot) Then NxRaiseContractError "Fixture path collision"
    fileSystem.CreateFolder mRoot
    Set mSource = Application.Workbooks.Add(xlWBATWorksheet)
    mSource.Worksheets(1).Name = "First"
    mSource.Worksheets(1).Range("A1").Value2 = "original"
    mSource.Worksheets(1).Range("B1").NumberFormat = "@"
    mSource.Worksheets(1).Range("B1").Value2 = "=1+1"
    mSource.Worksheets(1).Range("A2").Formula = "=1+1"
    Set sheet = mSource.Worksheets.Add(After:=mSource.Worksheets(1))
    sheet.Name = "Second"
    sheet.Range("A1").Value2 = "second"
    Set mTarget = Application.Workbooks.Add(xlWBATWorksheet)
    mTarget.Worksheets(1).Name = "First"
    mTarget.Worksheets(1).Range("A1").Value2 = "target-original"
End Sub

Private Function RegisterWorkbook(ByVal displayName As String, Optional ByVal retainAll As Boolean = True) As CNxTemplateRecord
    Dim request As New CNxTemplateRequest, result As CNxResult, records As Collection
    request.ConfigureWorkbookRegistration mSource, "workbook", displayName, "=1+1", retainAll, True, True
    Set result = NxTemplateRegisterForTest(request, mRoot)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "r62 workbook registration failed: " & result.Recovery
    Set records = NxTemplateListForTest(mRoot)
    NxTestHarness.AssertTrue records.Count = 1, "Workbook should register as one record"
    Set RegisterWorkbook = records.Item(1)
End Function

Private Sub TestWorkbookAndCopy()
    Dim record As CNxTemplateRecord, result As CNxResult, request As New CNxTemplateRequest
    Dim packagePath As String, beforeHash As String, beforeCount As Long
    mSource.Worksheets(1).Name = "Sheet1"
    mTarget.Worksheets(1).Name = "Sheet1"
    Set record = RegisterWorkbook("Two sheets")
    packagePath = NxTemplatePackagePath(record.TemplateId, mRoot)
    beforeHash = NxTemplateFileSha256(packagePath)
    NxTestHarness.AssertTrue record.SourceKind = "workbook" And record.RowCount = 5 And record.ColumnCount = 1, "Aggregate cell metadata is invalid"
    NxTestHarness.AssertTrue NxTemplatePackageProperty(packagePath, "LHexcelTemplateSchema") = "2", "Workbook package must use schema 2"
    beforeCount = mTarget.Worksheets.Count
    Set result = NxTemplateLoadForTest(record.TemplateId, mTarget, True, mRoot)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess And mTarget.Worksheets.Count = beforeCount + 2, "Both sheets must load"
    NxTestHarness.AssertTrue mTarget.Worksheets(1).Range("A1").Value2 = "target-original", "Existing target sheet was changed"
    NxTestHarness.AssertTrue mTarget.Worksheets(beforeCount + 1).Name <> "Sheet1", "Sheet collision was not resolved"
    NxTestHarness.AssertTrue Not mTarget.Worksheets(beforeCount + 1).Range("B1").HasFormula, "Literal equals text became a formula"
    Set result = NxTemplateOpenCopyForTest(record.TemplateId, True, mRoot)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Copy-open failed"
    Set mCopy = Application.Workbooks(result.Target)
    NxTestHarness.AssertTrue Len(mCopy.Path) = 0 And Not mCopy.Saved And mCopy.Worksheets.Count = 2, "Copy is not an unsaved two-sheet workbook"
    NxTestHarness.AssertTrue mCopy.Worksheets(1).Name = "Sheet1", "Copy placeholder caused an artificial sheet-name collision"
    mCopy.Worksheets(1).Range("A1").Value2 = "copy-edit"
    NxTestHarness.AssertTrue NxTemplateFileSha256(packagePath) = beforeHash, "Copy editing changed the stored template"
    request.ConfigureWorkbookRegistration mSource, "workbook", "two SHEETS", "", True, True, True
    Set result = NxTemplateRegisterForTest(request, mRoot)
    NxTestHarness.AssertTrue result.Outcome <> NxSuccess And NxTemplateFileSha256(packagePath) = beforeHash, "Duplicate registration overwrote the package"
    NxTestHarness.AssertTrue mSource.Worksheets(1).Range("A1").Value2 = "original", "Registration changed source cells"
End Sub

Private Sub TestSheetRegistration()
    Dim request As CNxTemplateRequest, record As CNxTemplateRecord, result As CNxResult, records As Collection
    Dim preview As String, previewDigest As String, freshDigest As String, excludedShape As Shape
    mSource.Worksheets(1).Range("A2").Formula = "=TODAY()"
    Set excludedShape = mSource.Worksheets(1).Shapes.AddShape(msoShapeRectangle, 1, 1, 10, 10)
    Set request = NxTemplateManagerRequest(mSource, mSource.Worksheets(1), "sheet_used_range", "Preview checks", "", False, True)
    preview = NxTemplateRegistrationPreview(request, previewDigest)
    NxTestHarness.AssertTrue InStr(preview, "[First]") > 0, "Preview omitted source-sheet identity"
    NxTestHarness.AssertTrue InStr(preview, "보존 상수 0개") > 0 And InStr(preview, "제외 상수 2개") > 0, "Preview omitted retained/excluded constants"
    NxTestHarness.AssertTrue InStr(preview, "A2 / VolatileFormula / TODAY") > 0, "Preview omitted detailed formula risk"
    NxTestHarness.AssertTrue InStr(preview, "제외 리소스: shape") > 0, "Preview omitted excluded resource"
    mSource.Worksheets(1).Range("A1").Value2 = "changed excluded constant"
    preview = NxTemplateRegistrationPreview(request, freshDigest)
    NxTestHarness.AssertTrue freshDigest <> previewDigest, "Excluded constant changes must invalidate preview"
    previewDigest = freshDigest
    mSource.Worksheets(1).Range("A1").Font.Size = mSource.Worksheets(1).Range("A1").Font.Size + 1
    preview = NxTemplateRegistrationPreview(request, freshDigest)
    NxTestHarness.AssertTrue freshDigest <> previewDigest, "Exported format changes must invalidate preview"
    previewDigest = freshDigest
    mSource.Worksheets(1).Name = "Renamed"
    preview = NxTemplateRegistrationPreview(request, freshDigest)
    NxTestHarness.AssertTrue freshDigest <> previewDigest, "Sheet rename must invalidate preview"
    mSource.Worksheets(1).Name = "First"
    mSource.Worksheets(1).Range("A2").Formula = "=1+1"
    excludedShape.Delete
    Set request = NxTemplateManagerRequest(mSource, mSource.Worksheets("First"), "workbook", "Order preview", "", False, True)
    preview = NxTemplateRegistrationPreview(request, previewDigest)
    mSource.Worksheets("Second").Move Before:=mSource.Worksheets("First")
    Set request = NxTemplateManagerRequest(mSource, mSource.Worksheets("First"), "workbook", "Order preview", "", False, True)
    preview = NxTemplateRegistrationPreview(request, freshDigest)
    NxTestHarness.AssertTrue freshDigest <> previewDigest, "Workbook sheet order changes must invalidate preview"
    mSource.Worksheets("Second").Move After:=mSource.Worksheets("First")
    Set request = NxTemplateManagerRequest(mSource, mSource.Worksheets(1), "sheet_used_range", "Whole sheet", "", False, True)
    Set result = NxTemplateRegisterForTest(request, mRoot)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Current sheet registration failed"
    Set records = NxTemplateListForTest(mRoot)
    Set record = records.Item(1)
    NxTestHarness.AssertTrue record.RowCount = 2 And record.ColumnCount = 2, "Whole UsedRange was not registered"
    Set result = NxTemplateLoadForTest(record.TemplateId, mTarget, True, mRoot)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Sheet template load failed"
    NxTestHarness.AssertTrue IsEmpty(mTarget.Worksheets(2).Range("A1").Value2), "Default constants-exclusion did not apply"
    NxTestHarness.AssertTrue mTarget.Worksheets(2).Range("A2").HasFormula, "Formula was not retained"
End Sub

Private Sub TestLoadRollback()
    Dim record As CNxTemplateRecord, result As CNxResult, beforeCount As Long, beforeBooks As Long
    Set record = RegisterWorkbook("Rollback")
    beforeCount = mTarget.Worksheets.Count
    beforeBooks = Application.Workbooks.Count
    NxTemplateSetFailurePoint "after-second-load-sheet"
    Set result = NxTemplateLoadForTest(record.TemplateId, mTarget, True, mRoot)
    NxTemplateClearFailurePoint
    NxTestHarness.AssertTrue result.Outcome <> NxSuccess, "Injected failure was ignored"
    NxTestHarness.AssertTrue mTarget.Worksheets.Count = beforeCount, "Rollback left a created sheet"
    NxTestHarness.AssertTrue mTarget.Worksheets(1).Range("A1").Value2 = "target-original", "Rollback changed an original sheet"
    NxTestHarness.AssertTrue Application.Workbooks.Count = beforeBooks, "Failed load left an owned package open"
    NxTestHarness.AssertTrue Application.AutomationSecurity = priorSecurity, "Load did not restore macro security"
End Sub

Private Sub TestCatalog()
    Dim record As CNxTemplateRecord, result As CNxResult, firstName As String
    Set record = RegisterWorkbook("=1+1")
    Set result = NxTemplateBuildCatalogForTest(mTarget, mRoot)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Catalog creation failed"
    firstName = result.Target
    NxTestHarness.AssertTrue Not mTarget.Worksheets(firstName).Range("A2").HasFormula, "Template name formula injection"
    NxTestHarness.AssertTrue Not mTarget.Worksheets(firstName).Range("B2").HasFormula, "Description formula injection"
    mTarget.Worksheets(firstName).Range("A1").Value2 = "keep prior catalog"
    Set result = NxTemplateBuildCatalogForTest(mTarget, mRoot)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess And result.Target <> firstName, "Catalog name collision was not resolved"
    NxTestHarness.AssertTrue mTarget.Worksheets(firstName).Range("A1").Value2 = "keep prior catalog", "Existing catalog was cleared"
    Set result = NxTemplateDeleteForTest(record.TemplateId, mRoot)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess And NxTemplateFolderHasExactFiles(result.Target), "Removal did not preserve recoverable package"
End Sub

Private Sub TestFileRegistration()
    Dim filePath As String, fileSha As String, result As CNxResult, records As Collection, beforeBooks As Long
    Dim preview As String, previewFileSha As String, wrongFileSha As String
    filePath = mRoot & "\source.xlsx"
    mSource.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook, AddToMru:=False
    mSource.Close SaveChanges:=False
    Set mSource = Nothing
    fileSha = NxTemplateFileSha256(filePath)
    beforeBooks = Application.Workbooks.Count
    preview = NxTemplateFileRegistrationPreview(filePath, mTarget, "File template", "", True, previewFileSha)
    NxTestHarness.AssertTrue previewFileSha = fileSha, "File preview hash did not bind exact original bytes"
    NxTestHarness.AssertTrue InStr(preview, "[First]") > 0 And InStr(preview, "[Second]") > 0, "File preview omitted per-sheet analysis"
    NxTestHarness.AssertTrue InStr(preview, "보존 상수 2개") > 0, "File preview omitted retained constants"
    NxTestHarness.AssertTrue Application.Workbooks.Count = beforeBooks, "File preview left its hidden source open"
    NxTestHarness.AssertTrue Application.AutomationSecurity = priorSecurity, "File preview did not restore macro security"
    NxTestHarness.AssertTrue Application.EnableEvents = mPriorEvents, "File preview did not restore events"
    If Left$(fileSha, 1) = "0" Then wrongFileSha = "1" & Mid$(fileSha, 2) Else wrongFileSha = "0" & Mid$(fileSha, 2)
    Set result = NxTemplateRegisterFileForTest(filePath, mTarget, "Changed preview", True, True, mRoot, wrongFileSha)
    NxTestHarness.AssertTrue result.Outcome <> NxSuccess, "Changed file preview was accepted"
    NxTestHarness.AssertTrue Application.Workbooks.Count = beforeBooks, "Rejected file preview left a workbook open"
    Set result = NxTemplateRegisterFileForTest(filePath, mTarget, "File template", True, True, mRoot, previewFileSha)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "File registration failed: " & result.Recovery
    Set records = NxTemplateListForTest(mRoot)
    NxTestHarness.AssertTrue records.Count = 1 And records.Item(1).SourceKind = "file", "File source kind was not recorded"
    NxTestHarness.AssertTrue Application.Workbooks.Count = beforeBooks, "File registration did not close its owned source"
    NxTestHarness.AssertTrue NxTemplateFileSha256(filePath) = fileSha, "File registration changed the original file"
    NxTestHarness.AssertTrue Application.AutomationSecurity = priorSecurity, "File registration did not restore macro security"
    Set result = NxTemplateRegisterFileForTest(filePath, mTarget, "No consent", True, False, mRoot)
    NxTestHarness.AssertTrue result.Outcome <> NxSuccess, "File registration bypassed explicit consent"
    ' Replace only this fixture-owned file to exercise a real change after preview.
    NxTestHarness.AssertTrue StrComp(CreateObject("Scripting.FileSystemObject").GetParentFolderName(filePath), mRoot, vbTextCompare) = 0, "Fixture replacement escaped its owned root"
    Set mSource = Application.Workbooks.Add(xlWBATWorksheet)
    mSource.Worksheets(1).Range("A1").Value2 = "changed after preview"
    Application.DisplayAlerts = False
    mSource.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook, AddToMru:=False
    Application.DisplayAlerts = mPriorAlerts
    mSource.Close SaveChanges:=False
    Set mSource = Nothing
    NxTestHarness.AssertTrue NxTemplateFileSha256(filePath) <> previewFileSha, "Replacement fixture bytes did not change"
    Set result = NxTemplateRegisterFileForTest(filePath, mTarget, "Replaced file", True, True, mRoot, previewFileSha)
    NxTestHarness.AssertTrue result.Outcome <> NxSuccess, "Real file replacement bypassed reviewed-byte verification"
    NxTestHarness.AssertTrue Application.Workbooks.Count = beforeBooks, "Real file replacement rejection leaked a workbook"
    Set records = NxTemplateListForTest(mRoot)
    NxTestHarness.AssertTrue records.Count = 1, "Rejected replacement created a template record"
    ' Opening xltx for inspection must retain the owned snapshot path, not create an unnamed workbook.
    filePath = mRoot & "\source.xltx"
    Set mSource = Application.Workbooks.Add(xlWBATWorksheet)
    mSource.Worksheets(1).Range("A1").Value2 = "xltx source"
    mSource.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLTemplate, AddToMru:=False
    mSource.Close SaveChanges:=False
    Set mSource = Nothing
    fileSha = NxTemplateFileSha256(filePath)
    preview = NxTemplateFileRegistrationPreview(filePath, mTarget, "Xltx template", "", True, previewFileSha)
    NxTestHarness.AssertTrue previewFileSha = fileSha, "xltx preview lost the original-file hash"
    NxTestHarness.AssertTrue Application.Workbooks.Count = beforeBooks, "xltx preview leaked an unnamed workbook"
    Set result = NxTemplateRegisterFileForTest(filePath, mTarget, "Xltx template", True, True, mRoot, previewFileSha)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "xltx registration failed: " & result.Recovery
    NxTestHarness.AssertTrue NxTemplateFileSha256(filePath) = fileSha, "xltx source changed during registration"
    NxTestHarness.AssertTrue Application.Workbooks.Count = beforeBooks, "xltx registration leaked its source"
End Sub

Private Sub TestManagerControls()
    Dim controlName As Variant
    Set mForm = New FNxTemplateManager
    Load mForm
    For Each controlName In Array("lstTemplates", "cmdRegisterFile", "cmdRegisterWorkbook", "cmdRegisterSheet", _
        "cmdUse", "cmdOpenCopy", "cmdDelete", "cmdRefresh", "cmdCatalog", "cmdOpenFolder", "cmdCancel")
        NxTestHarness.AssertTrue mForm.Controls(CStr(controlName)).Visible, "Manager control missing or hidden"
    Next controlName
    NxTestHarness.AssertTrue mForm.Controls.Count = 16, "Manager must preserve donor control composition"
    NxTestHarness.AssertTrue mForm.lstTemplates.ColumnCount = 5, "Manager must preserve donor list columns"
    NxTestHarness.AssertTrue Abs(mForm.lstTemplates.Width - 392) <= 0.051 And Abs(mForm.lstTemplates.Height - 220) <= 0.051, _
        "Manager list dimensions changed: " & CStr(mForm.lstTemplates.Width) & " x " & CStr(mForm.lstTemplates.Height)
    Set mRegisterForm = New FNxTemplateRegister
    Load mRegisterForm
    mRegisterForm.BindSource mTarget, mTarget.Worksheets(1), "sheet_used_range"
    NxTestHarness.AssertTrue mRegisterForm.cboContents.ListCount = 2 And mRegisterForm.cboContents.ListIndex = 0, "Registration defaults must exclude constants"
    NxTestHarness.AssertTrue Not mRegisterForm.chkConfirmRisk.Value, "Registration consent must default off"
    NxTestHarness.AssertTrue Not mRegisterForm.chkConfirmRisk.Enabled, "Registration consent must stay disabled until preview"
    NxTestHarness.AssertTrue mRegisterForm.txtPreview.TabStop And mRegisterForm.txtPreview.ScrollBars = fmScrollBarsVertical, "Detailed preview must be keyboard-scrollable"
    NxTestHarness.AssertTrue mRegisterForm.txtName.Value = mTarget.Worksheets(1).Name, "Current-sheet registration lost the donor name default"
End Sub

Private Sub CleanupFixture()
    Dim fileSystem As Object, resolved As String, tempRoot As String
    On Error Resume Next
    NxTemplateClearFailurePoint
    If Not mForm Is Nothing Then Unload mForm
    Set mForm = Nothing
    If Not mRegisterForm Is Nothing Then Unload mRegisterForm
    Set mRegisterForm = Nothing
    If Not mCopy Is Nothing Then mCopy.Close SaveChanges:=False
    If Not mTarget Is Nothing Then mTarget.Close SaveChanges:=False
    If Not mSource Is Nothing Then mSource.Close SaveChanges:=False
    Set mCopy = Nothing
    Set mTarget = Nothing
    Set mSource = Nothing
    Application.AutomationSecurity = priorSecurity
    Application.DisplayAlerts = mPriorAlerts
    Application.EnableEvents = mPriorEvents
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    resolved = fileSystem.GetAbsolutePathName(mRoot)
    tempRoot = fileSystem.GetAbsolutePathName(Environ$("TEMP"))
    If StrComp(fileSystem.GetParentFolderName(resolved), tempRoot, vbTextCompare) = 0 Then
        If Left$(fileSystem.GetFileName(resolved), 14) = "NxR62Template-" Then
            If fileSystem.FolderExists(resolved) Then fileSystem.DeleteFolder resolved, True
        End If
    End If
    mRoot = vbNullString
    On Error GoTo 0
End Sub
