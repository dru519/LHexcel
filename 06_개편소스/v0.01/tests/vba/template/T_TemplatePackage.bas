Attribute VB_Name = "T_TemplatePackage"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestPackageIsMacroFreeSingleSheet"
    names.Add "TestPackageMetadataUsesClosedSchema"
    names.Add "TestClosedPackageShaMatchesMetadata"
    names.Add "TestLogicalHashRoundTrips"
    names.Add "TestPackageRejectsExternalWorkbookFeatures"
    names.Add "TestTempPackageReopensReadOnlyWithoutLinks"
    names.Add "TestMetadataReplacementRestoresPriorFile"
    names.Add "TestTamperedPackageIsListedDamaged"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestPackageIsMacroFreeSingleSheet": TestPackageIsMacroFreeSingleSheet
        Case "TestPackageMetadataUsesClosedSchema": TestPackageMetadataUsesClosedSchema
        Case "TestClosedPackageShaMatchesMetadata": TestClosedPackageShaMatchesMetadata
        Case "TestLogicalHashRoundTrips": TestLogicalHashRoundTrips
        Case "TestPackageRejectsExternalWorkbookFeatures": TestPackageRejectsExternalWorkbookFeatures
        Case "TestTempPackageReopensReadOnlyWithoutLinks": TestTempPackageReopensReadOnlyWithoutLinks
        Case "TestMetadataReplacementRestoresPriorFile": TestMetadataReplacementRestoresPriorFile
        Case "TestTamperedPackageIsListedDamaged": TestTamperedPackageIsListedDamaged
        Case Else: NxRaiseContractError "Unknown Template package test"
    End Select
End Sub

Private Sub TestPackageIsMacroFreeSingleSheet()
    Dim sourceBook As Workbook, packageBook As Workbook, packagePath As String
    BuildPackageFixture sourceBook, packagePath
    Set packageBook = Application.Workbooks.Open(packagePath, 0, True)
    NxTestHarness.AssertTrue packageBook.Worksheets.Count = 1, "Template package must contain one sheet"
    NxTestHarness.AssertTrue packageBook.Worksheets(1).Name = "Template", "Template package sheet name mismatch"
    NxTestHarness.AssertTrue Not packageBook.HasVBProject, "Template package contains a VBA project"
    packageBook.Close False
    CleanupPackageFixture sourceBook, packagePath
End Sub

Private Sub TestPackageMetadataUsesClosedSchema()
    Dim sourceBook As Workbook, packagePath As String, metadataPath As String
    Dim record As CNxTemplateRecord, loaded As CNxTemplateRecord
    BuildPackageFixture sourceBook, packagePath
    metadataPath = ParentPath(packagePath) & "\metadata.ini"
    Set record = FixtureRecord("fixture-template", NxTemplateFileSha256(packagePath), _
        NxTemplatePackageProperty(packagePath, "LHexcelLogicalSha256"), "한글 템플릿", "설명")
    NxTemplateWriteMetadata metadataPath, record
    Set loaded = NxTemplateReadMetadata(metadataPath)
    NxTestHarness.AssertTrue loaded.DisplayName = "한글 템플릿" And loaded.Description = "설명", "Closed metadata text did not round-trip"
    NxTestHarness.AssertTrue NxTemplateFolderHasExactFiles(ParentPath(packagePath)), "Closed metadata file set mismatch"
    CleanupPackageFixture sourceBook, packagePath
End Sub

Private Sub TestClosedPackageShaMatchesMetadata()
    Dim sourceBook As Workbook, packagePath As String, packageSha As String, logicalSha As String
    Dim record As CNxTemplateRecord, loaded As CNxTemplateRecord
    BuildPackageFixture sourceBook, packagePath
    TracePackageCaseProgress "12-package-returned"
    packageSha = NxTemplateFileSha256(packagePath)
    TracePackageCaseProgress "13-package-sha"
    logicalSha = NxTemplatePackageProperty(packagePath, "LHexcelLogicalSha256")
    TracePackageCaseProgress "14-logical-property"
    Set record = FixtureRecord("fixture-template", packageSha, logicalSha)
    TracePackageCaseProgress "15-record-created"
    NxTemplateWriteMetadata ParentPath(packagePath) & "\metadata.ini", record
    TracePackageCaseProgress "16-metadata-written"
    Set loaded = NxTemplateReadRecordFolder(ParentPath(packagePath), False)
    TracePackageCaseProgress "17-folder-read"
    NxTestHarness.AssertTrue loaded.PackageSha256 = NxTemplateFileSha256(packagePath), "Closed package SHA does not match metadata"
    TracePackageCaseProgress "18-sha-confirmed"
    CleanupPackageFixture sourceBook, packagePath
    TracePackageCaseProgress "19-cleanup"
End Sub

Private Sub TracePackageCaseProgress(ByVal marker As String)
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

Private Sub TestLogicalHashRoundTrips()
    Dim sourceBook As Workbook, packagePath As String, propertySha As String, logicalSha As String, missingProperty As String
    Dim invalidHighRejected As Boolean, invalidLowRejected As Boolean, beforeBookCount As Long
    Dim priorSecurity As MsoAutomationSecurity
    TracePackageCaseProgress "20-hash-vectors-enter"
    NxTestHarness.AssertTrue NxTemplateTextSha256(vbNullString) = _
        "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855", "Empty UTF-8 hash mismatch"
    NxTestHarness.AssertTrue NxTemplateTextSha256("abc") = _
        "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", "ASCII UTF-8 hash mismatch"
    NxTestHarness.AssertTrue NxTemplateTextSha256("한글") = _
        "bd87f9bb68b67d2fa1cb82b6751820e946d5b1316d25d5fd96512fb4be44a2a8", "Korean UTF-8 hash mismatch"
    NxTestHarness.AssertTrue NxTemplateTextSha256("A" & vbNullChar & "B") = _
        "76fe3925c7167317f2df68454339f5ec3650e4062178b4f2be219b105a507907", "Embedded NUL UTF-8 hash mismatch"
    NxTestHarness.AssertTrue NxTemplateTextSha256(ChrW$(&HD83D) & ChrW$(&HDE00)) = _
        "f0443a342c5ef54783a111b51ba56c938e474c32324d90c3a60c9c8e3a37e2d9", "Non-BMP UTF-8 hash mismatch"
    On Error Resume Next
    Call NxTemplateTextSha256(ChrW$(&HD800))
    invalidHighRejected = (Err.Number <> 0)
    Err.Clear
    Call NxTemplateTextSha256(ChrW$(&HDC00))
    invalidLowRejected = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    NxTestHarness.AssertTrue invalidHighRejected And invalidLowRejected, "Invalid UTF-16 surrogate was accepted"
    TracePackageCaseProgress "21-hash-vectors-passed"
    BuildPackageFixture sourceBook, packagePath
    TracePackageCaseProgress "22-package-returned"
    priorSecurity = Application.AutomationSecurity
    beforeBookCount = Application.Workbooks.Count
    propertySha = NxTemplatePackageProperty(packagePath, "LHexcelLogicalSha256")
    TracePackageCaseProgress "23-logical-property"
    NxTestHarness.AssertTrue Application.Workbooks.Count = beforeBookCount And _
        Application.AutomationSecurity = priorSecurity, "Package property read did not restore its owner state"
    TracePackageCaseProgress "24-property-cleanup-verified"
    missingProperty = NxTemplatePackageProperty(packagePath, "LHexcelMissingProperty")
    NxTestHarness.AssertTrue Len(missingProperty) = 0 And Application.Workbooks.Count = beforeBookCount And _
        Application.AutomationSecurity = priorSecurity, "Failed package property read did not restore its owner state"
    TracePackageCaseProgress "25-missing-property-cleanup-verified"
    NxTestHarness.AssertTrue Not NxTemplateClosedPackageValid(packagePath, "wrong-template") And _
        Application.Workbooks.Count = beforeBookCount And Application.AutomationSecurity = priorSecurity, _
        "Rejected package validation did not restore its owner state"
    TracePackageCaseProgress "26-validation-cleanup-verified"
    logicalSha = NxTemplateLogicalHash(sourceBook.Worksheets(1).Range("A1:B2"), OneRetained(sourceBook.Worksheets(1).Range("A1")))
    TracePackageCaseProgress "27-logical-recomputed"
    NxTestHarness.AssertTrue propertySha = logicalSha, "Logical hash did not round-trip"
    TracePackageCaseProgress "28-roundtrip-confirmed"
    CleanupPackageFixture sourceBook, packagePath
    TracePackageCaseProgress "29-cleanup"
End Sub

Private Sub TestPackageRejectsExternalWorkbookFeatures()
    Dim sourceBook As Workbook, packageBook As Workbook, packagePath As String
    BuildPackageFixture sourceBook, packagePath
    NxTestHarness.AssertTrue NxTemplateClosedPackageValid(packagePath, "fixture-template"), "Clean package rejected"
    Set packageBook = Application.Workbooks.Open(packagePath, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    packageBook.Names.Add Name:="LHexcelExternalFixture", RefersTo:="='C:\missing\[external.xlsx]Sheet1'!$A$1"
    packageBook.Close SaveChanges:=True
    Set packageBook = Nothing
    NxTestHarness.AssertTrue Not NxTemplateClosedPackageValid(packagePath, "fixture-template"), "External defined name package accepted"
    CleanupPackageFixture sourceBook, packagePath
End Sub

Private Sub TestTempPackageReopensReadOnlyWithoutLinks()
    Dim sourceBook As Workbook, packagePath As String
    BuildPackageFixture sourceBook, packagePath
    NxTestHarness.AssertTrue NxTemplateClosedPackageValid(packagePath, "fixture-template"), "Closed package validation failed"
    CleanupPackageFixture sourceBook, packagePath
End Sub

Private Sub TestMetadataReplacementRestoresPriorFile()
    Dim sourceBook As Workbook, packagePath As String, metadataPath As String, priorSha As String
    Dim record As CNxTemplateRecord, changed As CNxTemplateRecord, failed As Boolean
    BuildPackageFixture sourceBook, packagePath
    metadataPath = ParentPath(packagePath) & "\metadata.ini"
    Set record = FixtureRecord("fixture-template", NxTemplateFileSha256(packagePath), NxTemplatePackageProperty(packagePath, "LHexcelLogicalSha256"))
    NxTemplateWriteMetadata metadataPath, record
    priorSha = NxTemplateFileSha256(metadataPath)
    Set changed = FixtureRecord("fixture-template", record.PackageSha256, record.LogicalSha256, "Changed", "Changed")
    On Error Resume Next
    NxTemplateReplaceMetadataForTest ParentPath(packagePath), changed, True
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Injected metadata replacement did not fail"
    NxTestHarness.AssertTrue NxTemplateFileSha256(metadataPath) = priorSha, "Metadata rollback changed prior bytes"
    NxTestHarness.AssertTrue Len(Dir$(ParentPath(packagePath) & "\metadata.new")) = 0, "Metadata new residue remains"
    NxTestHarness.AssertTrue Len(Dir$(ParentPath(packagePath) & "\metadata.bak")) = 0, "Metadata backup residue remains"
    CleanupPackageFixture sourceBook, packagePath
End Sub

Private Sub TestTamperedPackageIsListedDamaged()
    Dim sourceBook As Workbook, packagePath As String, fileNumber As Integer
    Dim record As CNxTemplateRecord, records As Collection
    BuildPackageFixture sourceBook, packagePath
    Set record = FixtureRecord("fixture-template", NxTemplateFileSha256(packagePath), NxTemplatePackageProperty(packagePath, "LHexcelLogicalSha256"))
    NxTemplateWriteMetadata ParentPath(packagePath) & "\metadata.ini", record
    fileNumber = FreeFile
    Open packagePath For Append As #fileNumber
    Print #fileNumber, "tampered"
    Close #fileNumber
    Set records = NxTemplateListMetadata(PackageRoot(packagePath))
    NxTestHarness.AssertTrue records.Count = 1 And records.Item(1).Status = "damaged", "Tampered package was not listed damaged"
    CleanupPackageFixture sourceBook, packagePath
End Sub

Private Sub BuildPackageFixture(ByRef sourceBook As Workbook, ByRef packagePath As String)
    Dim sheet As Worksheet, request As New CNxTemplateRequest
    Dim analysis As CNxTemplateAnalysis, retained As Collection, root As String, folderPath As String
    Dim fileSystem As Object
    Set sourceBook = Application.Workbooks.Add
    Set sheet = sourceBook.Worksheets(1)
    sheet.Range("A1").Value2 = "kept"
    sheet.Range("B1").Value2 = "excluded"
    sheet.Range("A2").FormulaR1C1 = "=R[-1]C"
    sheet.Range("B2").NumberFormat = "0.00"
    Set retained = OneRetained(sheet.Range("A1"))
    Set analysis = NxTemplateAnalyze(sheet.Range("A1:B2"), retained, False, False)
    request.ConfigureRegistration "NX-TPL-REGISTER-RANGE", "range", sheet.Range("A1:B2"), retained, "Fixture", "Package fixture", False, False, False
    request.BindAnalysis analysis
    root = NewPackageRoot()
    NxTemplateEnsureStore root
    folderPath = NxTemplateTemplateFolderPath("fixture-template", root)
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    fileSystem.CreateFolder folderPath
    packagePath = folderPath & "\template.xlsx"
    Call NxTemplateBuildPackage(request, analysis, "fixture-template", packagePath)
End Sub

Private Function FixtureRecord(ByVal templateId As String, ByVal packageSha As String, ByVal logicalSha As String, _
    Optional ByVal displayName As String = "Fixture", Optional ByVal description As String = "Package fixture") As CNxTemplateRecord
    Dim record As New CNxTemplateRecord
    record.Configure templateId, displayName, description, packageSha, logicalSha, "range", 2, 2, "", "healthy", "2026-07-17T00:00:00Z", "2026-07-17T00:00:00Z"
    record.Seal
    Set FixtureRecord = record
End Function

Private Function OneRetained(ByVal target As Range) As Collection
    Dim result As New Collection
    result.Add target
    Set OneRetained = result
End Function

Private Function NewPackageRoot() As String
    NewPackageRoot = Environ$("TEMP") & "\LHexcelTemplate-" & Replace(NxCreateRunUuid(), "-", vbNullString)
End Function

Private Function ParentPath(ByVal path As String) As String
    ParentPath = Left$(path, InStrRev(path, "\") - 1)
End Function

Private Function PackageRoot(ByVal packagePath As String) As String
    PackageRoot = ParentPath(ParentPath(ParentPath(packagePath)))
End Function

Private Sub CleanupPackageFixture(ByVal sourceBook As Workbook, ByVal packagePath As String)
    On Error Resume Next
    sourceBook.Close False
    DeleteFolderTree PackageRoot(packagePath)
    On Error GoTo 0
End Sub

Private Sub DeleteFolderTree(ByVal path As String)
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FolderExists(path) Then fileSystem.DeleteFolder path, True
End Sub
