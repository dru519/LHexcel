Attribute VB_Name = "NxSafeWorkbook"
Option Explicit

Private mSafeWorkbookSnapshotFolders As Object

' Safe closed-file policy: a command-owned snapshot is structurally inspected and opened.
' Legacy .xls/.xlsb, encrypted packages, malformed ZIPs, and XLM macro sheets fail closed.
' VBA project macros remain allowed only with ForceDisable.
Public Function NxSafeWorkbookOpen(ByVal inputPath As String, ByRef openedByUs As Boolean, Optional ByVal strictTemplate As Boolean = False) As Workbook
    Dim existingBook As Workbook, openedBook As Workbook, snapshotPath As String, snapshotFolder As String
    Dim priorSecurity As MsoAutomationSecurity, priorEvents As Boolean, restoreIssue As String, issue As String
    Dim stateCaptured As Boolean, errorNumber As Long, errorDescription As String
    openedByUs = False
    Set existingBook = NxSafeWorkbookFindOpen(inputPath)
    If Not existingBook Is Nothing Then
        If strictTemplate Then NxRaiseContractError "열려 있는 파일은 현재 문서 등록을 사용하세요."
        Set NxSafeWorkbookOpen = existingBook: Exit Function
    End If
    snapshotPath = NxSafeWorkbookPrepareSnapshot(inputPath, snapshotFolder, issue, strictTemplate)
    If Len(issue) > 0 Then NxRaiseContractError issue
    On Error GoTo Failed
    priorSecurity = Application.AutomationSecurity: priorEvents = Application.EnableEvents: stateCaptured = True
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    Set openedBook = Application.Workbooks.Open(Filename:=snapshotPath, UpdateLinks:=0, ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, AddToMru:=False, Notify:=False, Editable:=strictTemplate)
    restoreIssue = NxSafeWorkbookRestoreApplicationState(priorSecurity, priorEvents, stateCaptured)
    If Len(restoreIssue) > 0 Then
        openedBook.Close SaveChanges:=False
        NxSafeWorkbookDeleteTemporaryFolder snapshotFolder
        Err.Raise 5, , "안전 열기 후 Excel 상태를 복원하지 못했습니다: " & restoreIssue
    End If
    NxSafeWorkbookTrackSnapshot openedBook, snapshotFolder
    openedByUs = True
    Set NxSafeWorkbookOpen = openedBook
    Exit Function
Failed:
    errorNumber = Err.Number: errorDescription = Err.Description
    restoreIssue = NxSafeWorkbookRestoreApplicationState(priorSecurity, priorEvents, stateCaptured)
    On Error Resume Next
    If Not openedBook Is Nothing Then openedBook.Close SaveChanges:=False
    On Error GoTo 0
    NxSafeWorkbookDeleteTemporaryFolder snapshotFolder
    If Len(restoreIssue) > 0 Then errorDescription = errorDescription & " / 상태 복원 실패: " & restoreIssue
    Err.Raise errorNumber, "LHexcel.File.SafeOpen", errorDescription
End Function

Public Sub NxSafeWorkbookClose(ByVal book As Workbook, ByVal openedByUs As Boolean)
    Dim snapshotFolder As String, snapshotKey As String
    If book Is Nothing Or Not openedByUs Then Exit Sub
    snapshotKey = NxSafeWorkbookAbsolutePath(book.FullName)
    snapshotFolder = NxSafeWorkbookSnapshotFolder(book)
    book.Close SaveChanges:=False
    NxSafeWorkbookForgetSnapshot snapshotKey
    NxSafeWorkbookDeleteTemporaryFolder snapshotFolder
End Sub

Public Function NxSafeWorkbookInputIssue(ByVal inputPath As String) As String
    Dim snapshotPath As String, snapshotFolder As String, issue As String
    snapshotPath = NxSafeWorkbookPrepareSnapshot(inputPath, snapshotFolder, issue)
    NxSafeWorkbookDeleteTemporaryFolder snapshotFolder
    NxSafeWorkbookInputIssue = issue
End Function

Private Function NxSafeWorkbookPrepareSnapshot(ByVal inputPath As String, ByRef snapshotFolder As String, ByRef issue As String, Optional ByVal strictTemplate As Boolean = False) As String
    Dim fso As Object, extension As String, snapshotPath As String, inspectionZip As String
    Dim validationStage As String, errorNumber As Long
    On Error GoTo Rejected
    validationStage = "input.validate"
    issue = vbNullString: snapshotFolder = vbNullString
    If Len(Trim$(inputPath)) = 0 Then issue = "입력 통합문서를 찾을 수 없습니다.": Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(inputPath) Then issue = "입력 통합문서를 찾을 수 없습니다: " & inputPath: Exit Function
    If CDbl(fso.GetFile(inputPath).Size) > 104857600# Then issue = "입력 통합문서는 100 MiB를 넘을 수 없습니다: " & inputPath: Exit Function
    extension = LCase$(fso.GetExtensionName(inputPath))
    If strictTemplate Then
        If extension <> "xlsx" And extension <> "xltx" Then
            issue = "파일 템플릿은 .xlsx 또는 .xltx만 등록할 수 있습니다.": Exit Function
        End If
    End If
    Select Case extension
        Case "xlsx", "xlsm"
        Case "xltx"
            If Not strictTemplate Then issue = "템플릿 파일은 템플릿 등록에서 선택하세요.": Exit Function
        Case "xls", "xlsb"
            issue = "안전 열기는 .xlsx 또는 .xlsm만 지원합니다. .xls/.xlsb는 신뢰할 수 있는 환경에서 .xlsx/.xlsm으로 변환한 뒤 다시 선택하세요."
            Exit Function
        Case Else
            issue = "안전 열기는 구조 검사 가능한 .xlsx 또는 .xlsm만 지원합니다: " & inputPath
            Exit Function
    End Select
    validationStage = "snapshot.mkdir"
    snapshotFolder = Environ$("TEMP") & "\LHexcel-safe-open-" & Replace$(NxCreateRunUuid(), "-", vbNullString)
    MkDir snapshotFolder
    snapshotPath = snapshotFolder & "\source-" & Replace$(NxCreateRunUuid(), "-", vbNullString) & "." & extension
    validationStage = "snapshot.copy"
    If fso.FileExists(snapshotPath) Or fso.FolderExists(snapshotPath) Then Err.Raise 58, , "Safe snapshot destination already exists"
    ' VBA FileCopy rejects a source already open by navigation's write-denial lock.
    ' A shared-reader copy preserves that lock and never overwrites a destination.
    fso.CopyFile inputPath, snapshotPath, False
    validationStage = "inspection.copy"
    inspectionZip = snapshotFolder & "\inspection.zip"
    FileCopy snapshotPath, inspectionZip
    validationStage = "package.inspect"
    issue = NxSafeWorkbookPackageIssue(inspectionZip, snapshotFolder, strictTemplate)
    If Len(issue) > 0 Then GoTo CleanUp
    NxSafeWorkbookPrepareSnapshot = snapshotPath
    Exit Function
CleanUp:
    NxSafeWorkbookDeleteTemporaryFolder snapshotFolder
    snapshotFolder = vbNullString
    Exit Function
Rejected:
    errorNumber = Err.Number
    issue = "입력 통합문서를 안전하게 검사할 수 없습니다: " & inputPath & _
        " [stage=" & validationStage & "; error=" & CStr(errorNumber) & "]"
    Resume CleanUp
End Function

Public Function NxSafeWorkbookFindOpen(ByVal inputPath As String) As Workbook
    Dim book As Workbook, normalizedPath As String
    normalizedPath = NxSafeWorkbookAbsolutePath(inputPath)
    For Each book In Application.Workbooks
        If StrComp(NxSafeWorkbookAbsolutePath(book.FullName), normalizedPath, vbTextCompare) = 0 Then
            Set NxSafeWorkbookFindOpen = book
            Exit Function
        End If
    Next book
End Function

Private Function NxSafeWorkbookAbsolutePath(ByVal filePath As String) As String
    On Error GoTo RawPath
    NxSafeWorkbookAbsolutePath = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(filePath)
    Exit Function
RawPath:
    NxSafeWorkbookAbsolutePath = filePath
End Function

Private Function NxSafeWorkbookRestoreApplicationState(ByVal priorSecurity As MsoAutomationSecurity, _
    ByVal priorEvents As Boolean, ByVal stateCaptured As Boolean) As String
    Dim restoreIssue As String
    If Not stateCaptured Then Exit Function
    On Error Resume Next
    Err.Clear
    Application.AutomationSecurity = priorSecurity
    If Err.Number <> 0 Then restoreIssue = "AutomationSecurity: " & Err.Description
    Err.Clear
    Application.EnableEvents = priorEvents
    If Err.Number <> 0 Then restoreIssue = restoreIssue & IIf(Len(restoreIssue) > 0, "; ", vbNullString) & "EnableEvents: " & Err.Description
    Err.Clear
    On Error GoTo 0
    NxSafeWorkbookRestoreApplicationState = restoreIssue
End Function

Private Sub NxSafeWorkbookTrackSnapshot(ByVal book As Workbook, ByVal snapshotFolder As String)
    If mSafeWorkbookSnapshotFolders Is Nothing Then Set mSafeWorkbookSnapshotFolders = CreateObject("Scripting.Dictionary")
    mSafeWorkbookSnapshotFolders(NxSafeWorkbookAbsolutePath(book.FullName)) = snapshotFolder
End Sub

Private Function NxSafeWorkbookSnapshotFolder(ByVal book As Workbook) As String
    If mSafeWorkbookSnapshotFolders Is Nothing Then Exit Function
    If mSafeWorkbookSnapshotFolders.Exists(NxSafeWorkbookAbsolutePath(book.FullName)) Then _
        NxSafeWorkbookSnapshotFolder = mSafeWorkbookSnapshotFolders(NxSafeWorkbookAbsolutePath(book.FullName))
End Function

Private Sub NxSafeWorkbookForgetSnapshot(ByVal snapshotKey As String)
    If mSafeWorkbookSnapshotFolders Is Nothing Then Exit Sub
    If mSafeWorkbookSnapshotFolders.Exists(snapshotKey) Then mSafeWorkbookSnapshotFolders.Remove snapshotKey
End Sub

Private Function NxSafeWorkbookPackageIssue(ByVal inspectionZip As String, ByVal snapshotFolder As String, Optional ByVal strictTemplate As Boolean = False) As String
    Dim shellApp As Object, zipFolder As Object, xlItem As Object, xlFolder As Object, contentTypes As Object, relationships As Object
    Dim entryCount As Long, tooManyEntries As Boolean, contentTypesPath As String, relationshipsPath As String
    On Error GoTo Rejected
    Set shellApp = CreateObject("Shell.Application")
    Set zipFolder = shellApp.NameSpace(CVar(inspectionZip))
    If zipFolder Is Nothing Then GoTo Rejected
    If strictTemplate Then
        If NxSafeWorkbookTemplateHasActiveParts(zipFolder, entryCount) Then
            NxSafeWorkbookPackageIssue = "외부 연결·매크로·실행 개체가 포함된 파일은 템플릿으로 등록할 수 없습니다."
            Exit Function
        End If
        entryCount = 0
    End If
    If NxSafeWorkbookFolderHasMacroSheet(zipFolder, entryCount, tooManyEntries) Then
        NxSafeWorkbookPackageIssue = "XLM 매크로 시트가 포함된 통합문서는 안전 열기에서 거부됩니다."
        Exit Function
    End If
    If tooManyEntries Then
        NxSafeWorkbookPackageIssue = "입력 통합문서의 ZIP 항목 수가 안전 검사 한도를 넘습니다."
        Exit Function
    End If
    Set xlItem = zipFolder.ParseName("xl")
    If xlItem Is Nothing Then GoTo Rejected
    Set xlFolder = xlItem.GetFolder
    contentTypesPath = NxSafeWorkbookExtractSmallZipEntry(shellApp, zipFolder, "[Content_Types].xml", snapshotFolder & "\content")
    relationshipsPath = NxSafeWorkbookExtractSmallZipEntry(shellApp, xlFolder, "_rels\workbook.xml.rels", snapshotFolder & "\relationships")
    Set contentTypes = NxSafeWorkbookLoadXml(contentTypesPath, "Types")
    Set relationships = NxSafeWorkbookLoadXml(relationshipsPath, "Relationships")
    If strictTemplate Then
        If NxSafeWorkbookTemplateXmlHasActiveParts(contentTypes) Or NxSafeWorkbookTemplateXmlHasActiveParts(relationships) Then
            NxSafeWorkbookPackageIssue = "외부 연결 또는 실행 개체가 선언된 파일은 템플릿으로 등록할 수 없습니다."
            Exit Function
        End If
    End If
    If NxSafeWorkbookXmlHasMacroSheet(contentTypes) Or NxSafeWorkbookXmlHasMacroSheet(relationships) Then _
        NxSafeWorkbookPackageIssue = "XLM 매크로 시트 관계가 포함된 통합문서는 안전 열기에서 거부됩니다."
    Exit Function
Rejected:
    NxSafeWorkbookPackageIssue = "암호화되었거나 구조를 검사할 수 없는 통합문서는 안전 열기에서 거부됩니다. 신뢰할 수 있는 환경에서 .xlsx/.xlsm으로 변환한 뒤 다시 선택하세요."
End Function

Private Function NxSafeWorkbookTemplateHasActiveParts(ByVal folder As Object, ByRef entryCount As Long) As Boolean
    Dim item As Object, token As Variant, itemName As String
    For Each item In folder.Items
        entryCount = entryCount + 1
        If entryCount > 10000 Then NxSafeWorkbookTemplateHasActiveParts = True: Exit Function
        itemName = LCase$(item.Name)
        For Each token In Array("vbaproject", "macrosheet", "activex", "embeddings", "externallinks", "connections", "querytables", "pivotcache", "customui")
            If InStr(1, itemName, CStr(token), vbBinaryCompare) > 0 Then NxSafeWorkbookTemplateHasActiveParts = True: Exit Function
        Next token
        If item.IsFolder Then
            If NxSafeWorkbookTemplateHasActiveParts(item.GetFolder, entryCount) Then NxSafeWorkbookTemplateHasActiveParts = True: Exit Function
        End If
    Next item
End Function

Private Function NxSafeWorkbookTemplateXmlHasActiveParts(ByVal document As Object) As Boolean
    Dim element As Object, xmlAttribute As Object, token As Variant, value As String
    For Each element In document.DocumentElement.ChildNodes
        If element.nodeType = 1 Then
            For Each xmlAttribute In element.Attributes
                value = LCase$(CStr(xmlAttribute.Text))
                For Each token In Array("vbaproject", "macrosheet", "activex", "oleobject", "externallink", "connection", "querytable", "pivotcache", "customui")
                    If InStr(1, value, CStr(token), vbBinaryCompare) > 0 Then NxSafeWorkbookTemplateXmlHasActiveParts = True: Exit Function
                Next token
            Next xmlAttribute
        End If
    Next element
End Function

Private Function NxSafeWorkbookFolderHasMacroSheet(ByVal folder As Object, ByRef entryCount As Long, _
    ByRef tooManyEntries As Boolean) As Boolean
    Dim item As Object, itemName As String
    For Each item In folder.Items
        entryCount = entryCount + 1
        If entryCount > 10000 Then tooManyEntries = True: Exit Function
        itemName = LCase$(item.Name)
        If InStr(1, itemName, "macrosheet", vbTextCompare) > 0 Or InStr(1, itemName, "intlmacrosheet", vbTextCompare) > 0 Then
            NxSafeWorkbookFolderHasMacroSheet = True
            Exit Function
        End If
        If item.IsFolder Then
            If NxSafeWorkbookFolderHasMacroSheet(item.GetFolder, entryCount, tooManyEntries) Then
                NxSafeWorkbookFolderHasMacroSheet = True
                Exit Function
            End If
            If tooManyEntries Then Exit Function
        End If
    Next item
End Function

Private Function NxSafeWorkbookExtractSmallZipEntry(ByVal shellApp As Object, ByVal sourceFolder As Object, _
    ByVal itemPath As String, ByVal temporaryFolder As String) As String
    Dim item As Object, destinationFolder As Object, extractedPath As String, startedAt As Date, expectedSize As Double
    Set item = NxSafeWorkbookZipItem(sourceFolder, itemPath)
    If item Is Nothing Then Err.Raise 5, , "Required OOXML package entry is missing"
    expectedSize = CDbl(item.Size)
    If expectedSize <= 0 Or expectedSize > 2097152# Then Err.Raise 5, , "OOXML inspection entry exceeds 2 MiB"
    If Len(Dir$(temporaryFolder, vbDirectory)) = 0 Then MkDir temporaryFolder
    Set destinationFolder = shellApp.NameSpace(CVar(temporaryFolder))
    destinationFolder.CopyHere item, 16
    extractedPath = temporaryFolder & "\" & item.Name
    startedAt = Now
    Do
        If Len(Dir$(extractedPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then
            If CDbl(FileLen(extractedPath)) = expectedSize Then Exit Do
        End If
        If DateDiff("s", startedAt, Now) >= 10 Then Err.Raise 5, , "Timed out extracting OOXML inspection entry"
        DoEvents
    Loop
    NxSafeWorkbookExtractSmallZipEntry = extractedPath
End Function

Private Function NxSafeWorkbookZipItem(ByVal folder As Object, ByVal itemPath As String) As Object
    Dim parts() As String, index As Long, item As Object, currentFolder As Object
    itemPath = Replace$(itemPath, "/", "\")
    parts = Split(itemPath, "\")
    Set currentFolder = folder
    For index = LBound(parts) To UBound(parts)
        Set item = currentFolder.ParseName(parts(index))
        If item Is Nothing Then Exit Function
        If index < UBound(parts) Then
            If Not item.IsFolder Then Exit Function
            Set currentFolder = item.GetFolder
        End If
    Next index
    Set NxSafeWorkbookZipItem = item
End Function

Private Function NxSafeWorkbookLoadXml(ByVal filePath As String, ByVal expectedRoot As String) As Object
    Dim document As Object
    Set document = CreateObject("MSXML2.DOMDocument.6.0")
    document.Async = False: document.ResolveExternals = False: document.ValidateOnParse = False
    document.setProperty "ProhibitDTD", True
    If Not document.Load(filePath) Then Err.Raise 5, , "OOXML XML parse failed"
    If document.DocumentElement Is Nothing Then Err.Raise 5, , "OOXML XML root is missing"
    If StrComp(document.DocumentElement.baseName, expectedRoot, vbBinaryCompare) <> 0 Then Err.Raise 5, , "Unexpected OOXML XML root"
    Set NxSafeWorkbookLoadXml = document
End Function

Private Function NxSafeWorkbookXmlHasMacroSheet(ByVal document As Object) As Boolean
    Dim node As Object, attributeValue As String
    For Each node In document.DocumentElement.ChildNodes
        If node.NodeType = 1 Then
            attributeValue = LCase$(NxSafeWorkbookXmlAttribute(node, "ContentType") & "|" & NxSafeWorkbookXmlAttribute(node, "Type"))
            If InStr(1, attributeValue, "macrosheet", vbBinaryCompare) > 0 Then NxSafeWorkbookXmlHasMacroSheet = True: Exit Function
        End If
    Next node
End Function

Private Function NxSafeWorkbookXmlAttribute(ByVal node As Object, ByVal attributeName As String) As String
    Dim attributeNode As Object
    Set attributeNode = node.Attributes.getNamedItem(attributeName)
    If Not attributeNode Is Nothing Then NxSafeWorkbookXmlAttribute = CStr(attributeNode.Text)
End Function

Private Sub NxSafeWorkbookDeleteTemporaryFolder(ByVal folderPath As String)
    If Len(folderPath) = 0 Then Exit Sub
    If InStr(1, folderPath, "\LHexcel-safe-open-", vbTextCompare) = 0 Then Exit Sub
    On Error Resume Next
    If CreateObject("Scripting.FileSystemObject").FolderExists(folderPath) Then _
        CreateObject("Scripting.FileSystemObject").DeleteFolder folderPath, True
    On Error GoTo 0
End Sub
