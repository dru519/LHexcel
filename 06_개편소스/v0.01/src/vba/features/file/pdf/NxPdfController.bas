Attribute VB_Name = "NxPdfController"
Option Explicit

#Const NX_R58_TEST_BUILD = False

#If VBA7 Then
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal fileNamePointer As LongPtr, ByVal desiredAccess As Long, _
    ByVal shareMode As Long, ByVal securityAttributes As LongPtr, ByVal creationDisposition As Long, _
    ByVal flagsAndAttributes As Long, ByVal templateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFileInformationByHandle Lib "kernel32" (ByVal fileHandle As LongPtr, _
    ByRef information As NX_PDF_BY_HANDLE_FILE_INFORMATION) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal objectHandle As LongPtr) As Long
#Else
Private Declare Function CreateFileW Lib "kernel32" (ByVal fileNamePointer As Long, ByVal desiredAccess As Long, _
    ByVal shareMode As Long, ByVal securityAttributes As Long, ByVal creationDisposition As Long, _
    ByVal flagsAndAttributes As Long, ByVal templateFile As Long) As Long
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal fileHandle As Long, _
    ByRef information As NX_PDF_BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal objectHandle As Long) As Long
#End If

Private Type NX_PDF_FILETIME
    lowDateTime As Long
    highDateTime As Long
End Type

Private Type NX_PDF_BY_HANDLE_FILE_INFORMATION
    fileAttributes As Long
    creationTime As NX_PDF_FILETIME
    lastAccessTime As NX_PDF_FILETIME
    lastWriteTime As NX_PDF_FILETIME
    volumeSerialNumber As Long
    fileSizeHigh As Long
    fileSizeLow As Long
    numberOfLinks As Long
    fileIndexHigh As Long
    fileIndexLow As Long
End Type

Private Const NX_PDF_OPEN_EXISTING As Long = 3
Private Const NX_PDF_FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const NX_PDF_FILE_SHARE_ALL As Long = &H7
Private Const NX_PDF_INVALID_HANDLE_VALUE As Long = -1

Public Function NxPdfExecute(ByVal featureId As String, ByVal outputFolder As String, _
    ByVal baseName As String, ByVal qualityText As String) As CNxResult

    Dim commandOwnedPaths As New Collection, outputPath As String, sheet As Worksheet
    Dim qualityValue As Long, errorNumber As Long, errorDescription As String, recoveryDescription As String
    On Error GoTo Failed
    NxPdfValidateFeature featureId
    qualityValue = NxPdfQualityValue(qualityText)

    If featureId = "NX-FILE-PDF-SETTINGS" Then
        NxPdfSaveSettings outputFolder, baseName, qualityText
        Set NxPdfExecute = NxCreateResult(featureId, NxSuccess, "complete", NxPdfSettingsPath(), _
            False, vbNullString, "pdf_settings_saved")
        Exit Function
    End If
    If ActiveWorkbook Is Nothing Then NxRaiseContractError "PDF로 저장할 통합문서가 없습니다."

    Select Case featureId
        Case "NX-FILE-PDF-CURRENT-SHEET"
            If Not TypeOf ActiveSheet Is Worksheet Then NxRaiseContractError "현재 워크시트를 찾을 수 없습니다."
            outputPath = NxPdfOutputPath(outputFolder, baseName)
            NxPdfExportWorksheet ActiveSheet, outputPath, qualityValue, commandOwnedPaths
        Case "NX-FILE-PDF-EACH-SHEET"
            For Each sheet In ActiveWorkbook.Worksheets
                If sheet.Visible = xlSheetVisible Then
                    outputPath = NxPdfOutputPath(outputFolder, baseName, sheet.Name)
                    NxPdfExportWorksheet sheet, outputPath, qualityValue, commandOwnedPaths
                End If
            Next sheet
            If commandOwnedPaths.Count = 0 Then NxRaiseContractError "PDF로 저장할 표시 워크시트가 없습니다."
        Case "NX-FILE-PDF-SELECTED-COMBINED"
            outputPath = NxPdfOutputPath(outputFolder, baseName)
            NxPdfExportSelectedSheets outputPath, qualityValue, commandOwnedPaths
        Case "NX-FILE-PDF-ALL-COMBINED"
            outputPath = NxPdfOutputPath(outputFolder, baseName)
            NxPdfExportWorkbook ActiveWorkbook, outputPath, qualityValue, commandOwnedPaths
    End Select
    Set NxPdfExecute = NxCreateResult(featureId, NxSuccess, "complete", outputFolder, _
        False, vbNullString, "pdf_export_complete")
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    recoveryDescription = NxPdfRollbackCreated(commandOwnedPaths)
    If Len(recoveryDescription) > 0 Then errorDescription = errorDescription & " | PDF 복구 실패: " & recoveryDescription
    Err.Raise errorNumber, "LHexcel.File.Pdf", errorDescription
End Function

Private Sub NxPdfExportWorksheet(ByVal sheet As Worksheet, ByVal outputPath As String, ByVal qualityValue As Long, _
    ByVal commandOwnedPaths As Collection)
    Dim temporaryPath As String
    NxPdfRequireVacantOutput outputPath
    temporaryPath = NxPdfTemporaryPath(outputPath)
    sheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=temporaryPath, Quality:=qualityValue, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    NxPdfRegisterOwnedPath commandOwnedPaths, temporaryPath
    NxPdfAtomicPublishOwned temporaryPath, outputPath, commandOwnedPaths
End Sub

Private Sub NxPdfExportWorkbook(ByVal book As Workbook, ByVal outputPath As String, ByVal qualityValue As Long, _
    ByVal commandOwnedPaths As Collection)
    Dim temporaryPath As String
    NxPdfRequireVacantOutput outputPath
    temporaryPath = NxPdfTemporaryPath(outputPath)
    book.ExportAsFixedFormat Type:=xlTypePDF, Filename:=temporaryPath, Quality:=qualityValue, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    NxPdfRegisterOwnedPath commandOwnedPaths, temporaryPath
    NxPdfAtomicPublishOwned temporaryPath, outputPath, commandOwnedPaths
End Sub

Private Sub NxPdfExportSelectedSheets(ByVal outputPath As String, ByVal qualityValue As Long, _
    ByVal commandOwnedPaths As Collection)
    Dim sourceBook As Workbook, temporaryBook As Workbook, selectedSheet As Object
    Dim names() As Variant, index As Long, temporaryPath As String
    Dim errorNumber As Long, errorDescription As String, recoveryDescription As String
    Set sourceBook = ActiveWorkbook
    If ActiveWindow Is Nothing Then NxRaiseContractError "선택한 워크시트를 확인할 수 없습니다."
    If ActiveWindow.SelectedSheets.Count = 0 Then NxRaiseContractError "PDF로 저장할 워크시트를 선택하세요."
    ReDim names(1 To ActiveWindow.SelectedSheets.Count)
    For Each selectedSheet In ActiveWindow.SelectedSheets
        If TypeName(selectedSheet) <> "Worksheet" Then NxRaiseContractError "워크시트만 PDF로 통합할 수 있습니다."
        index = index + 1
        names(index) = selectedSheet.Name
    Next selectedSheet

    NxPdfRequireVacantOutput outputPath
    temporaryPath = NxPdfTemporaryPath(outputPath)
    On Error GoTo Failed
    sourceBook.Worksheets(names).Copy
    Set temporaryBook = ActiveWorkbook
    temporaryBook.ExportAsFixedFormat Type:=xlTypePDF, Filename:=temporaryPath, Quality:=qualityValue, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    NxPdfRegisterOwnedPath commandOwnedPaths, temporaryPath
    temporaryBook.Close SaveChanges:=False
    Set temporaryBook = Nothing
    sourceBook.Activate
    NxPdfAtomicPublishOwned temporaryPath, outputPath, commandOwnedPaths
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    On Error Resume Next
    If Not temporaryBook Is Nothing Then temporaryBook.Close SaveChanges:=False
    If Err.Number <> 0 Then recoveryDescription = "선택 PDF 임시 통합문서|" & CStr(Err.Number) & "|" & Err.Description
    Err.Clear
    If Len(temporaryPath) > 0 And NxPdfPathExists(temporaryPath) Then _
        recoveryDescription = NxPdfJoinFailure(recoveryDescription, "unverified-temporary-residual=" & temporaryPath)
    If Not sourceBook Is Nothing Then sourceBook.Activate
    If Err.Number <> 0 Then recoveryDescription = NxPdfJoinFailure(recoveryDescription, "원본 통합문서 활성화|" & CStr(Err.Number) & "|" & Err.Description)
    Err.Clear
    On Error GoTo 0
    If Len(recoveryDescription) > 0 Then errorDescription = errorDescription & " | " & recoveryDescription
    Err.Raise errorNumber, "LHexcel.File.Pdf", errorDescription
End Sub

Private Function NxPdfTemporaryPath(ByVal outputPath As String) As String
    NxPdfTemporaryPath = Left$(outputPath, Len(outputPath) - 4) & ".nx-" & _
        Replace$(NxCreateRunUuid(), "-", vbNullString) & ".tmp.pdf"
End Function

Public Sub NxPdfAtomicPublish(ByVal temporaryPath As String, ByVal outputPath As String)
    NxPdfRequireVacantOutput outputPath
    If Not NxPdfPathExists(temporaryPath) Then NxRaiseContractError "PDF 임시 파일을 찾을 수 없습니다."
    Name temporaryPath As outputPath
End Sub

Private Sub NxPdfAtomicPublishOwned(ByVal temporaryPath As String, ByVal outputPath As String, _
    ByVal commandOwnedPaths As Collection)
    NxPdfAtomicPublish temporaryPath, outputPath
    NxPdfRegisterOwnedPath commandOwnedPaths, outputPath
End Sub

Private Sub NxPdfRegisterOwnedPath(ByVal commandOwnedPaths As Collection, ByVal candidate As String)
    If commandOwnedPaths Is Nothing Then NxRaiseContractError "PDF command ownership journal is required"
    commandOwnedPaths.Add Array(candidate, NxPdfFileIdentity(candidate))
End Sub

' Test-only entrypoints are compiled only in the isolated native test copy.
' Product.xlam leaves NX_R58_TEST_BUILD False, so it has no caller-controlled
' capture/delete surface.
#If NX_R58_TEST_BUILD Then
Public Function NxPdfTestCaptureOwnedIdentity(ByVal candidate As String) As String
    NxPdfTestCaptureOwnedIdentity = NxPdfFileIdentity(candidate)
End Function

Public Function NxPdfTestCleanupCapturedPath(ByVal candidate As String, ByVal expectedIdentity As String) As String
    Dim commandOwnedPaths As New Collection
    If Len(expectedIdentity) = 0 Then NxRaiseContractError "PDF test cleanup requires a captured identity"
    commandOwnedPaths.Add Array(candidate, expectedIdentity)
    NxPdfTestCleanupCapturedPath = NxPdfRollbackCreated(commandOwnedPaths)
End Function
#End If

' Deletes only files still carrying the identity recorded immediately after this
' command created them. A replacement is reported rather than deleted.
Private Function NxPdfRollbackCreated(ByVal commandOwnedPaths As Collection) As String
    Dim index As Long, record As Variant, candidate As String, expectedIdentity As String
    Dim actualIdentity As String, failures As String
    If commandOwnedPaths Is Nothing Then Exit Function
    For index = commandOwnedPaths.Count To 1 Step -1
        record = commandOwnedPaths.Item(index)
        candidate = CStr(record(0))
        expectedIdentity = CStr(record(1))
        On Error GoTo CleanupFailed
        If NxPdfPathExists(candidate) Then
            actualIdentity = NxPdfFileIdentity(candidate)
            If StrComp(actualIdentity, expectedIdentity, vbBinaryCompare) <> 0 Then
                failures = NxPdfJoinFailure(failures, "replacement=" & candidate)
            Else
                Kill candidate
                If NxPdfPathExists(candidate) Then _
                    NxRaiseContractError "PDF 명령 소유 파일이 삭제되지 않았습니다: " & candidate
            End If
        End If
ContinueCleanup:
        On Error GoTo 0
    Next index
    NxPdfRollbackCreated = failures
    Exit Function
CleanupFailed:
    failures = NxPdfJoinFailure(failures, "residual=" & candidate & "|" & CStr(Err.Number) & "|" & Err.Description)
    Err.Clear
    Resume ContinueCleanup
End Function

Private Function NxPdfFileIdentity(ByVal candidate As String) As String
#If VBA7 Then
    Dim fileHandle As LongPtr
#Else
    Dim fileHandle As Long
#End If
    Dim information As NX_PDF_BY_HANDLE_FILE_INFORMATION, lastError As Long
    fileHandle = CreateFileW(StrPtr(candidate), 0, NX_PDF_FILE_SHARE_ALL, 0, NX_PDF_OPEN_EXISTING, NX_PDF_FILE_ATTRIBUTE_NORMAL, 0)
    If fileHandle = NX_PDF_INVALID_HANDLE_VALUE Then
        lastError = Err.LastDllError
        NxRaiseContractError "PDF 파일 정체성을 확인할 수 없습니다: " & candidate & " (Win32 " & CStr(lastError) & ")"
    End If
    If GetFileInformationByHandle(fileHandle, information) = 0 Then
        lastError = Err.LastDllError
        CloseHandle fileHandle
        NxRaiseContractError "PDF 파일 정체성을 확인할 수 없습니다: " & candidate & " (Win32 " & CStr(lastError) & ")"
    End If
    CloseHandle fileHandle
    NxPdfFileIdentity = Hex$(information.volumeSerialNumber) & ":" & Hex$(information.fileIndexHigh) & ":" & Hex$(information.fileIndexLow)
End Function

Private Function NxPdfPathExists(ByVal candidate As String) As Boolean
    On Error GoTo Missing
    NxPdfPathExists = (Len(Dir$(candidate, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
    Exit Function
Missing:
    NxRaiseContractError "PDF 경로 존재 여부를 확인할 수 없습니다: " & candidate
End Function

Private Function NxPdfJoinFailure(ByVal aggregate As String, ByVal detail As String) As String
    If Len(aggregate) > 0 Then aggregate = aggregate & "; "
    NxPdfJoinFailure = aggregate & detail
End Function
