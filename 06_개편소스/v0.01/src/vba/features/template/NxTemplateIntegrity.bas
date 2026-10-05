Attribute VB_Name = "NxTemplateIntegrity"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" (ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, ByRef pbOutput As Any, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, ByRef pbHashObject As Any, ByVal cbHashObject As Long, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByRef pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByRef pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As LongPtr) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32.dll" (ByVal codePage As Long, ByVal flags As Long, ByVal wideText As LongPtr, ByVal wideLength As Long, ByVal utf8Buffer As LongPtr, ByVal utf8Length As Long, ByVal defaultCharacter As LongPtr, ByVal usedDefaultCharacter As LongPtr) As Long
#Else
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGetProperty Lib "bcrypt.dll" (ByVal hObject As Long, ByVal pszProperty As Long, ByRef pbOutput As Any, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByRef phHash As Long, ByRef pbHashObject As Any, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As Long, ByRef pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As Long, ByRef pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" (ByVal codePage As Long, ByVal flags As Long, ByVal wideText As Long, ByVal wideLength As Long, ByVal utf8Buffer As Long, ByVal utf8Length As Long, ByVal defaultCharacter As Long, ByVal usedDefaultCharacter As Long) As Long
#End If

Private Const BCRYPT_OBJECT_LENGTH As String = "ObjectLength"
Private Const BCRYPT_HASH_LENGTH As String = "HashDigestLength"
Private Const STATUS_SUCCESS As Long = 0
Private Const FILE_HASH_CHUNK_SIZE As Long = 65536
Private Const CP_UTF8 As Long = 65001
Private Const WC_ERR_INVALID_CHARS As Long = &H80&

Public Function NxTemplateFileSha256(ByVal path As String) As String
#If VBA7 Then
    Dim algorithm As LongPtr, hashHandle As LongPtr
#Else
    Dim algorithm As Long, hashHandle As Long
#End If
    Dim handle As Integer, remaining As Long, chunkLength As Long
    Dim objectLength As Long, hashLength As Long, resultLength As Long, index As Long
    Dim hashObject() As Byte, buffer() As Byte, digest() As Byte
    On Error GoTo Failed
    If Len(Dir$(path)) = 0 Then NxRaiseContractError "Template package does not exist"
    handle = FreeFile
    Open path For Binary Access Read Lock Write As #handle
    If BCryptOpenAlgorithmProvider(algorithm, StrPtr("SHA256"), 0, 0) <> STATUS_SUCCESS Then GoTo Failed
    If BCryptGetProperty(algorithm, StrPtr(BCRYPT_OBJECT_LENGTH), objectLength, 4, resultLength, 0) <> STATUS_SUCCESS Then GoTo Failed
    If BCryptGetProperty(algorithm, StrPtr(BCRYPT_HASH_LENGTH), hashLength, 4, resultLength, 0) <> STATUS_SUCCESS Then GoTo Failed
    ReDim hashObject(0 To objectLength - 1)
    ReDim digest(0 To hashLength - 1)
    If BCryptCreateHash(algorithm, hashHandle, hashObject(0), objectLength, 0, 0, 0) <> STATUS_SUCCESS Then GoTo Failed
    remaining = LOF(handle)
    Do While remaining > 0
        chunkLength = remaining
        If chunkLength > FILE_HASH_CHUNK_SIZE Then chunkLength = FILE_HASH_CHUNK_SIZE
        ReDim buffer(0 To chunkLength - 1)
        Get #handle, , buffer
        If BCryptHashData(hashHandle, buffer(0), chunkLength, 0) <> STATUS_SUCCESS Then GoTo Failed
        remaining = remaining - chunkLength
    Loop
    If BCryptFinishHash(hashHandle, digest(0), hashLength, 0) <> STATUS_SUCCESS Then GoTo Failed
    For index = 0 To hashLength - 1
        NxTemplateFileSha256 = NxTemplateFileSha256 & LCase$(Right$("0" & Hex$(digest(index)), 2))
    Next index
CleanExit:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    If hashHandle <> 0 Then Call BCryptDestroyHash(hashHandle)
    If algorithm <> 0 Then Call BCryptCloseAlgorithmProvider(algorithm, 0)
    On Error GoTo 0
    If Len(NxTemplateFileSha256) <> 64 Then NxRaiseContractError "Windows CNG file SHA-256 failed"
    Exit Function
Failed:
    NxTemplateFileSha256 = vbNullString
    Resume CleanExit
End Function

Public Function NxTemplateTextSha256(ByVal value As String) As String
    Dim bytes() As Byte, byteCount As Long, converted As Long
    If Len(value) = 0 Then
        NxTemplateTextSha256 = Sha256Bytes(bytes)
        Exit Function
    End If
    byteCount = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, StrPtr(value), Len(value), 0, 0, 0, 0)
    If byteCount <= 0 Then NxRaiseContractError "Template text is not valid UTF-16"
    ReDim bytes(0 To byteCount - 1)
    converted = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, StrPtr(value), Len(value), VarPtr(bytes(0)), byteCount, 0, 0)
    If converted <> byteCount Then NxRaiseContractError "Template UTF-8 conversion failed"
    NxTemplateTextSha256 = Sha256Bytes(bytes)
End Function

Public Function NxTemplateLogicalHash(ByVal source As Range, ByVal retained As Collection) As String
    Dim canonical As String, cell As Range, rowIndex As Long, columnIndex As Long
    Dim cellTexts() As String, cellIndex As Long
    If source Is Nothing Or retained Is Nothing Then NxRaiseContractError "Logical hash requires source and retained areas"
    canonical = "LHexcelTemplateLogical|1|" & CStr(source.Rows.Count) & "|" & CStr(source.Columns.Count) & vbLf
    For rowIndex = 1 To source.Rows.Count
        canonical = canonical & "row|" & CStr(rowIndex) & "|" & CanonicalNumber(source.Rows(rowIndex).RowHeight) & vbLf
    Next rowIndex
    For columnIndex = 1 To source.Columns.Count
        canonical = canonical & "column|" & CStr(columnIndex) & "|" & CanonicalNumber(source.Columns(columnIndex).ColumnWidth) & vbLf
    Next columnIndex
    ReDim cellTexts(0 To CLng(source.CountLarge) - 1)
    For Each cell In source.Cells
        If cellIndex Mod 256 = 0 Then NxTemplateProgress "내용 확인", cellIndex, CLng(source.CountLarge)
        cellTexts(cellIndex) = CellCanonical(source, cell, retained)
        cellIndex = cellIndex + 1
    Next cell
    canonical = canonical & Join(cellTexts, vbNullString)
    NxTemplateLogicalHash = NxTemplateTextSha256(canonical)
End Function

Public Function NxTemplatePackageProperty(ByVal packagePath As String, ByVal propertyName As String) As String
    Dim book As Workbook, priorSecurity As MsoAutomationSecurity, propertyValue As String
    Dim securityCaptured As Boolean, operationSucceeded As Boolean, cleanupFailed As Boolean
    Dim priorEvents As Boolean
    On Error GoTo CleanExit
    If TemplatePackageAlreadyOpen(packagePath) Then Exit Function
    priorSecurity = Application.AutomationSecurity
    priorEvents = Application.EnableEvents
    securityCaptured = True
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    Set book = Application.Workbooks.Open(packagePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    propertyValue = CStr(book.CustomDocumentProperties(propertyName).Value)
    operationSucceeded = True
CleanExit:
    On Error Resume Next
    If Not book Is Nothing Then
        Err.Clear
        book.Close SaveChanges:=False
        If Err.Number <> 0 Then cleanupFailed = True
        Err.Clear
        Set book = Nothing
    End If
    If securityCaptured Then
        Err.Clear
        Application.AutomationSecurity = priorSecurity
        If Err.Number <> 0 Then cleanupFailed = True
        Err.Clear
        Application.EnableEvents = priorEvents
        If Err.Number <> 0 Then cleanupFailed = True
    End If
    Err.Clear
    On Error GoTo 0
    If operationSucceeded And Not cleanupFailed Then NxTemplatePackageProperty = propertyValue
End Function

Public Function NxTemplateClosedPackageValid(ByVal packagePath As String, ByVal templateId As String, _
    Optional ByVal expectedLogicalSha As String = vbNullString) As Boolean
    Dim book As Workbook, priorSecurity As MsoAutomationSecurity, logicalSha As String
    Dim securityCaptured As Boolean, validationPassed As Boolean, cleanupFailed As Boolean
    Dim schema As String, sheet As Worksheet, totalCells As Double, priorEvents As Boolean
    On Error GoTo CleanExit
    If Len(Dir$(packagePath)) = 0 Or Not NxTemplateIdIsValid(templateId) Then Exit Function
    If TemplatePackageAlreadyOpen(packagePath) Then Exit Function
    priorSecurity = Application.AutomationSecurity
    priorEvents = Application.EnableEvents
    securityCaptured = True
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    Set book = Application.Workbooks.Open(packagePath, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, AddToMru:=False)
    schema = CStr(book.CustomDocumentProperties("LHexcelTemplateSchema").Value)
    Select Case schema
        Case "1"
            If book.Worksheets.Count <> 1 Then GoTo CleanExit
            If book.Worksheets(1).Name <> "Template" Then GoTo CleanExit
        Case "2"
            If book.Worksheets.Count = 0 Then GoTo CleanExit
        Case Else
            GoTo CleanExit
    End Select
    If book.Sheets.Count <> book.Worksheets.Count Then GoTo CleanExit
    For Each sheet In book.Worksheets
        If sheet.Visible <> xlSheetVisible Then GoTo CleanExit
        If sheet.Shapes.Count <> 0 Or sheet.QueryTables.Count <> 0 Or sheet.PivotTables.Count <> 0 Then GoTo CleanExit
        totalCells = totalCells + CDbl(sheet.UsedRange.CountLarge)
        If totalCells > 100000# Then GoTo CleanExit
    Next sheet
    If book.HasVBProject Then GoTo CleanExit
    If book.Connections.Count <> 0 Then GoTo CleanExit
    If Not NxTemplatePrintNamesOnly(book) Then GoTo CleanExit
    If HasWorkbookLinks(book, xlExcelLinks) Or HasWorkbookLinks(book, xlOLELinks) Then GoTo CleanExit
    If CStr(book.CustomDocumentProperties("LHexcelTemplateId").Value) <> templateId Then GoTo CleanExit
    If CStr(book.CustomDocumentProperties("LHexcelOrigin").Value) <> "user-created" Then GoTo CleanExit
    logicalSha = CStr(book.CustomDocumentProperties("LHexcelLogicalSha256").Value)
    If Not IsLowerSha256(logicalSha) Then GoTo CleanExit
    If Len(expectedLogicalSha) > 0 And logicalSha <> expectedLogicalSha Then GoTo CleanExit
    validationPassed = True
CleanExit:
    On Error Resume Next
    If Not book Is Nothing Then
        Err.Clear
        book.Close SaveChanges:=False
        If Err.Number <> 0 Then cleanupFailed = True
        Err.Clear
        Set book = Nothing
    End If
    If securityCaptured Then
        Err.Clear
        Application.AutomationSecurity = priorSecurity
        If Err.Number <> 0 Then cleanupFailed = True
        Err.Clear
        Application.EnableEvents = priorEvents
        If Err.Number <> 0 Then cleanupFailed = True
    End If
    Err.Clear
    On Error GoTo 0
    NxTemplateClosedPackageValid = validationPassed And Not cleanupFailed
End Function

Private Function TemplatePackageAlreadyOpen(ByVal packagePath As String) As Boolean
    Dim candidate As Workbook, fileName As String
    On Error GoTo FailClosed
    fileName = CreateObject("Scripting.FileSystemObject").GetFileName(packagePath)
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, packagePath, vbTextCompare) = 0 Or StrComp(candidate.Name, fileName, vbTextCompare) = 0 Then
            TemplatePackageAlreadyOpen = True
            Exit Function
        End If
    Next candidate
    Exit Function
FailClosed:
    TemplatePackageAlreadyOpen = True
End Function

Public Function NxTemplateRecordIntegrityStatus(ByVal record As CNxTemplateRecord, ByVal packagePath As String) As String
    If record Is Nothing Then NxRaiseContractError "Template record is required"
    If Len(Dir$(packagePath)) = 0 Then NxTemplateRecordIntegrityStatus = "damaged": Exit Function
    If NxTemplateFileSha256(packagePath) <> record.PackageSha256 Then NxTemplateRecordIntegrityStatus = "damaged": Exit Function
    If Not NxTemplateClosedPackageValid(packagePath, record.TemplateId, record.LogicalSha256) Then NxTemplateRecordIntegrityStatus = "damaged": Exit Function
    NxTemplateRecordIntegrityStatus = "healthy"
End Function

Private Function CellCanonical(ByVal source As Range, ByVal cell As Range, ByVal retained As Collection) As String
    Dim relativeAddress As String, valueText As String, kind As String, mergeText As String
    relativeAddress = "R" & CStr(cell.Row - source.Row + 1) & "C" & CStr(cell.Column - source.Column + 1)
    If cell.HasFormula Then
        kind = "formula"
        valueText = CStr(cell.FormulaR1C1)
    ElseIf IsRetainedCell(cell, retained) Then
        kind = "constant:" & CStr(VarType(cell.Value2))
        valueText = CStr(cell.Value2)
    Else
        kind = "blank"
        valueText = vbNullString
    End If
    If cell.MergeCells Then
        If cell.Address = cell.MergeArea.Cells(1, 1).Address Then mergeText = CStr(cell.MergeArea.Rows.Count) & "x" & CStr(cell.MergeArea.Columns.Count)
    End If
    CellCanonical = "cell|" & relativeAddress & "|" & EscapeCanonical(kind) & "|" & EscapeCanonical(valueText) & "|" & _
        EscapeCanonical(cell.NumberFormat) & "|" & CStr(cell.Font.Bold) & "|" & CStr(cell.Font.Italic) & "|" & _
        CStr(cell.Font.Color) & "|" & CStr(cell.Interior.Color) & "|" & CStr(cell.HorizontalAlignment) & "|" & _
        CStr(cell.VerticalAlignment) & "|" & EscapeCanonical(mergeText) & vbLf
End Function

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

Private Function EscapeCanonical(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, "|", "\p")
    value = Replace(value, vbCr, "\r")
    EscapeCanonical = Replace(value, vbLf, "\n")
End Function

Private Function CanonicalNumber(ByVal value As Double) As String
    CanonicalNumber = Replace(Format$(value, "0.###############"), Application.International(xlDecimalSeparator), ".")
End Function

Private Function HasWorkbookLinks(ByVal book As Workbook, ByVal linkType As XlLinkType) As Boolean
    Dim links As Variant, lowerBound As Long, upperBound As Long
    On Error GoTo NoLinks
    links = book.LinkSources(linkType)
    If IsEmpty(links) Then Exit Function
    lowerBound = LBound(links)
    upperBound = UBound(links)
    HasWorkbookLinks = (upperBound >= lowerBound)
NoLinks:
End Function

Private Function IsLowerSha256(ByVal value As String) As Boolean
    Dim index As Long
    If Len(value) <> 64 Then Exit Function
    For index = 1 To Len(value)
        If InStr(1, "0123456789abcdef", Mid$(value, index, 1), vbBinaryCompare) = 0 Then Exit Function
    Next index
    IsLowerSha256 = True
End Function

Private Function Sha256Bytes(ByRef bytes() As Byte) As String
#If VBA7 Then
    Dim algorithm As LongPtr, hashHandle As LongPtr
#Else
    Dim algorithm As Long, hashHandle As Long
#End If
    Dim objectLength As Long, hashLength As Long, resultLength As Long, byteCount As Long
    Dim hashObject() As Byte, digest() As Byte, index As Long
    On Error GoTo Failed
    If BCryptOpenAlgorithmProvider(algorithm, StrPtr("SHA256"), 0, 0) <> STATUS_SUCCESS Then GoTo Failed
    If BCryptGetProperty(algorithm, StrPtr(BCRYPT_OBJECT_LENGTH), objectLength, 4, resultLength, 0) <> STATUS_SUCCESS Then GoTo Failed
    If BCryptGetProperty(algorithm, StrPtr(BCRYPT_HASH_LENGTH), hashLength, 4, resultLength, 0) <> STATUS_SUCCESS Then GoTo Failed
    ReDim hashObject(0 To objectLength - 1)
    ReDim digest(0 To hashLength - 1)
    If BCryptCreateHash(algorithm, hashHandle, hashObject(0), objectLength, 0, 0, 0) <> STATUS_SUCCESS Then GoTo Failed
    byteCount = ByteArrayLength(bytes)
    If byteCount > 0 Then
        If BCryptHashData(hashHandle, bytes(LBound(bytes)), byteCount, 0) <> STATUS_SUCCESS Then GoTo Failed
    End If
    If BCryptFinishHash(hashHandle, digest(0), hashLength, 0) <> STATUS_SUCCESS Then GoTo Failed
    For index = 0 To hashLength - 1
        Sha256Bytes = Sha256Bytes & LCase$(Right$("0" & Hex$(digest(index)), 2))
    Next index
CleanExit:
    If hashHandle <> 0 Then BCryptDestroyHash hashHandle
    If algorithm <> 0 Then BCryptCloseAlgorithmProvider algorithm, 0
    Exit Function
Failed:
    If hashHandle <> 0 Then BCryptDestroyHash hashHandle
    If algorithm <> 0 Then BCryptCloseAlgorithmProvider algorithm, 0
    NxRaiseContractError "Windows CNG SHA-256 failed"
End Function

Private Function ByteArrayLength(ByRef bytes() As Byte) As Long
    On Error GoTo EmptyArray
    ByteArrayLength = UBound(bytes) - LBound(bytes) + 1
EmptyArray:
End Function
