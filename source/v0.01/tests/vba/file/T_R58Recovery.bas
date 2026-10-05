Attribute VB_Name = "T_R58Recovery"
Option Explicit

' Direct native entrypoint: every case creates its own files/workbook and does
' not require a class instance to be passed through the Windows BAT runner.
Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestFileChainTemporaryMoveFailure"
    names.Add "TestFileChainTargetMoveFailure"
    names.Add "TestFileRollbackBlockedByReplacement"
    names.Add "TestFileTargetReplacementIdentityFailsClosed"
    names.Add "TestSheetChainRollbackUsesObjects"
    names.Add "TestLockedPdfCleanupReportsResidual"
    names.Add "TestRenameInjectionRejectsInvalidStage"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestFileChainTemporaryMoveFailure": TestFileChainTemporaryMoveFailure
        Case "TestFileChainTargetMoveFailure": TestFileChainTargetMoveFailure
        Case "TestFileRollbackBlockedByReplacement": TestFileRollbackBlockedByReplacement
        Case "TestFileTargetReplacementIdentityFailsClosed": TestFileTargetReplacementIdentityFailsClosed
        Case "TestSheetChainRollbackUsesObjects": TestSheetChainRollbackUsesObjects
        Case "TestLockedPdfCleanupReportsResidual": TestLockedPdfCleanupReportsResidual
        Case "TestRenameInjectionRejectsInvalidStage": TestRenameInjectionRejectsInvalidStage
        Case Else: NxRaiseContractError "Unknown r58 recovery test"
    End Select
End Sub

Private Sub TestFileChainTemporaryMoveFailure()
    Dim folderPath As String, aPath As String, aaPath As String, preview As Collection, failed As Boolean, detail As String
    NxBatchRenameTestClearMoveFailure
    folderPath = R58TestFolder(): aPath = folderPath & "\a.txt": aaPath = folderPath & "\aa.txt"
    On Error GoTo Failed
    R58WriteText aPath, "a-original": R58WriteText aaPath, "aa-original"
    Set preview = R58FileChainPreview(aPath, aaPath, folderPath & "\aaa.txt")
    NxBatchRenameTestInjectMoveFailure "original-to-temporary", 2
    On Error Resume Next
    Call NxBatchRenameCommit(preview)
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    R58AssertTrue failed, "Injected second temporary move did not fail"
    R58AssertTrue R58ReadText(aPath) = "a-original" And R58ReadText(aaPath) = "aa-original", "Temporary failure moved a chain item to the wrong original path"
Cleanup:
    NxBatchRenameTestClearMoveFailure
    R58DeleteTree folderPath
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanupWithError
CleanupWithError:
    NxBatchRenameTestClearMoveFailure
    R58DeleteTree folderPath
    On Error GoTo 0
    NxRaiseContractError detail
End Sub

Private Sub TestFileChainTargetMoveFailure()
    Dim folderPath As String, aPath As String, aaPath As String, preview As Collection, failed As Boolean, detail As String
    NxBatchRenameTestClearMoveFailure
    folderPath = R58TestFolder(): aPath = folderPath & "\a.txt": aaPath = folderPath & "\aa.txt"
    On Error GoTo Failed
    R58WriteText aPath, "a-original": R58WriteText aaPath, "aa-original"
    Set preview = R58FileChainPreview(aPath, aaPath, folderPath & "\aaa.txt")
    NxBatchRenameTestInjectMoveFailure "temporary-to-target", 2
    On Error Resume Next
    Call NxBatchRenameCommit(preview)
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    R58AssertTrue failed, "Injected second target move did not fail"
    R58AssertTrue R58ReadText(aPath) = "a-original" And R58ReadText(aaPath) = "aa-original", "Target partial failure did not restore the chain"
Cleanup:
    NxBatchRenameTestClearMoveFailure
    R58DeleteTree folderPath
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanupWithError
CleanupWithError:
    NxBatchRenameTestClearMoveFailure
    R58DeleteTree folderPath
    On Error GoTo 0
    NxRaiseContractError detail
End Sub

Private Sub TestFileRollbackBlockedByReplacement()
    Dim folderPath As String, originalPath As String, temporaryPath As String, targetPath As String
    Dim journal As New CNxBatchRenameJournal, failed As Boolean, detail As String
    NxBatchRenameTestClearMoveFailure
    folderPath = R58TestFolder(): originalPath = folderPath & "\a.txt": temporaryPath = folderPath & "\a.tmp": targetPath = folderPath & "\aa.txt"
    On Error GoTo Failed
    R58WriteText originalPath, "owned-original"
    journal.Add originalPath, temporaryPath, targetPath: journal.Seal: journal.MoveOriginalToTemporary 1
    R58WriteText originalPath, "replacement"
    On Error Resume Next
    journal.Rollback
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    R58AssertTrue failed, "Rollback overwrote a replacement original"
    R58AssertTrue R58ReadText(originalPath) = "replacement", "Replacement original was changed"
    R58AssertTrue R58ReadText(temporaryPath) = "owned-original", "Owned temporary residual was not reported/preserved"
Cleanup:
    NxBatchRenameTestClearMoveFailure
    R58DeleteTree folderPath
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanupWithError
CleanupWithError:
    NxBatchRenameTestClearMoveFailure
    R58DeleteTree folderPath
    On Error GoTo 0
    NxRaiseContractError detail
End Sub

Private Sub TestFileTargetReplacementIdentityFailsClosed()
    Dim folderPath As String, originalPath As String, temporaryPath As String, targetPath As String
    Dim journal As New CNxBatchRenameJournal, failed As Boolean, detail As String
    NxBatchRenameTestClearMoveFailure
    folderPath = R58TestFolder(): originalPath = folderPath & "\a.txt": temporaryPath = folderPath & "\a.tmp": targetPath = folderPath & "\aa.txt"
    On Error GoTo Failed
    R58WriteText originalPath, "owned-original"
    journal.Add originalPath, temporaryPath, targetPath: journal.Seal
    journal.MoveOriginalToTemporary 1: journal.MoveTemporaryToTarget 1
    Kill targetPath: R58WriteText targetPath, "replacement-target"
    On Error Resume Next
    journal.Rollback
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    R58AssertTrue failed, "Rollback accepted a replacement target"
    R58AssertTrue R58ReadText(targetPath) = "replacement-target", "Replacement target was changed"
Cleanup:
    NxBatchRenameTestClearMoveFailure
    R58DeleteTree folderPath
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanupWithError
CleanupWithError:
    NxBatchRenameTestClearMoveFailure
    R58DeleteTree folderPath
    On Error GoTo 0
    NxRaiseContractError detail
End Sub

Private Sub TestSheetChainRollbackUsesObjects()
    Dim book As Workbook, firstSheet As Worksheet, secondSheet As Worksheet, preview As New Collection
    Dim failed As Boolean, hasDuplicate As Boolean, hasCollision As Boolean, detail As String
    NxBatchRenameTestClearMoveFailure
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set firstSheet = book.Worksheets(1): firstSheet.Name = "a"
    Set secondSheet = book.Worksheets.Add(After:=firstSheet): secondSheet.Name = "aa"
    Set preview = NxSheetBatchRenameBuildPreview(book, vbNullString, vbNullString, vbNullString, vbNullString, hasDuplicate, hasCollision)
    Set preview = NxSheetBatchRenameEditPreview(book, preview, 1, "aa", hasDuplicate, hasCollision)
    Set preview = NxSheetBatchRenameEditPreview(book, preview, 2, "aaa", hasDuplicate, hasCollision)
    NxBatchRenameTestInjectMoveFailure "sheet-temporary-to-target", 2
    On Error Resume Next
    Call NxSheetBatchRenameCommit(book, preview)
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    R58AssertTrue failed, "Injected second sheet target move did not fail"
    R58AssertTrue firstSheet.Name = "a" And secondSheet.Name = "aa", "Sheet rollback did not restore retained worksheet objects"
Cleanup:
    NxBatchRenameTestClearMoveFailure
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    NxBatchRenameTestClearMoveFailure
    NxRaiseContractError detail
End Sub

Private Sub TestLockedPdfCleanupReportsResidual()
    Dim folderPath As String, pdfPath As String, handle As Integer, recovery As String, detail As String, ownedIdentity As String
    folderPath = R58TestFolder(): pdfPath = folderPath & "\locked.pdf"
    On Error GoTo Failed
    R58WriteText pdfPath, "command-owned test PDF"
    ownedIdentity = NxPdfTestCaptureOwnedIdentity(pdfPath)
    handle = FreeFile
    Open pdfPath For Binary Access Read Write Lock Read Write As #handle
    recovery = NxPdfTestCleanupCapturedPath(pdfPath, ownedIdentity)
    Close #handle: handle = 0
    R58AssertTrue InStr(1, recovery, "residual=" & pdfPath, vbBinaryCompare) > 0, "Locked PDF cleanup did not report the exact residual path"
    R58DeleteTree folderPath
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next
    If handle <> 0 Then Close #handle
    On Error GoTo 0
    R58DeleteTree folderPath
    NxRaiseContractError detail
End Sub

Private Sub TestRenameInjectionRejectsInvalidStage()
    Dim failed As Boolean
    NxBatchRenameTestClearMoveFailure
    On Error Resume Next
    NxBatchRenameTestInjectMoveFailure "unknown-stage", 1
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    NxBatchRenameTestClearMoveFailure
    R58AssertTrue failed, "Rename failure injection accepted an unknown stage"
End Sub

Private Function R58FileChainPreview(ByVal aPath As String, ByVal aaPath As String, ByVal aaaPath As String) As Collection
    Dim preview As New Collection, row As Variant
    row = Array(aPath, aaPath, "준비"): preview.Add row
    row = Array(aaPath, aaaPath, "준비"): preview.Add row
    Set R58FileChainPreview = preview
End Function

Private Function R58TestFolder() As String
    R58TestFolder = Environ$("TEMP") & "\LHexcel-r58-recovery-" & NxCreateRunUuid()
    MkDir R58TestFolder
End Function

Private Sub R58WriteText(ByVal candidate As String, ByVal value As String)
    Dim handle As Integer
    handle = FreeFile
    Open candidate For Output As #handle
    Print #handle, value
    Close #handle
End Sub

Private Function R58ReadText(ByVal candidate As String) As String
    Dim handle As Integer
    handle = FreeFile
    Open candidate For Input As #handle
    Line Input #handle, R58ReadText
    Close #handle
End Function

Private Sub R58DeleteTree(ByVal folderPath As String)
    Dim candidate As String
    On Error Resume Next
    candidate = Dir$(folderPath & "\*", vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    Do While Len(candidate) > 0
        Kill folderPath & "\" & candidate
        candidate = Dir$()
    Loop
    RmDir folderPath
    On Error GoTo 0
End Sub

Private Sub R58AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then NxRaiseContractError message
End Sub
