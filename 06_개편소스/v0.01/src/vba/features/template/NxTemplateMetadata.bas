Attribute VB_Name = "NxTemplateMetadata"
Option Explicit

Private Const METADATA_LINE_COUNT As Long = 15

Public Sub NxTemplateWriteMetadata(ByVal path As String, ByVal record As CNxTemplateRecord)
    Dim handle As Integer
    If record Is Nothing Then NxRaiseContractError "Template metadata record is required"
    handle = FreeFile
    On Error GoTo Failed
    Open path For Output Access Write As #handle
    Print #handle, "schema_version=2"
    Print #handle, "template_id=" & record.TemplateId
    Print #handle, "display_name=" & EncodeText(record.DisplayName)
    Print #handle, "description=" & EncodeText(record.Description)
    Print #handle, "package_sha256=" & record.PackageSha256
    Print #handle, "logical_sha256=" & record.LogicalSha256
    Print #handle, "source_kind=" & record.SourceKind
    Print #handle, "row_count=" & CStr(record.RowCount)
    Print #handle, "column_count=" & CStr(record.ColumnCount)
    Print #handle, "risk_flags=" & EncodeText(record.RiskFlags)
    Print #handle, "status=" & record.Status
    Print #handle, "created_utc=" & record.CreatedUtc
    Print #handle, "modified_utc=" & record.ModifiedUtc
    Print #handle, "category=" & EncodeText(record.Category)
    Print #handle, "last_used_utc=" & record.LastUsedUtc
    Close #handle
    Exit Sub
Failed:
    On Error Resume Next
    Close #handle
    On Error GoTo 0
    NxRaiseContractError "Template metadata write failed"
End Sub

Public Function NxTemplateReadMetadata(ByVal path As String) As CNxTemplateRecord
    Dim lines As Variant, values(0 To METADATA_LINE_COUNT - 1) As String
    Dim index As Long, separator As Long, key As String
    Dim record As New CNxTemplateRecord
    lines = ReadMetadataLines(path)
    For index = 0 To UBound(lines)
        separator = InStr(1, CStr(lines(index)), "=", vbBinaryCompare)
        If separator < 1 Then NxRaiseContractError "Template metadata field is malformed"
        key = Left$(CStr(lines(index)), separator - 1)
        If key <> ExpectedKey(index) Then NxRaiseContractError "Template metadata schema is not closed"
        values(index) = Mid$(CStr(lines(index)), separator + 1)
    Next index
    If Not ((values(0) = "1" And UBound(lines) = 12) Or (values(0) = "2" And UBound(lines) = 14)) Then NxRaiseContractError "Template metadata schema version is invalid"
    If Not PositiveIntegerText(values(7)) Or Not PositiveIntegerText(values(8)) Then NxRaiseContractError "Template metadata dimensions are invalid"
    If Not ClosedTimestamp(values(11)) Or Not ClosedTimestamp(values(12)) Then NxRaiseContractError "Template metadata timestamp is invalid"
    If Len(values(14)) > 0 Then
        If Not ClosedTimestamp(values(14)) Then NxRaiseContractError "Template last-used timestamp is invalid"
    End If
    record.Configure values(1), DecodeText(values(2)), DecodeText(values(3)), values(4), values(5), values(6), _
        CLng(values(7)), CLng(values(8)), DecodeText(values(9)), values(10), values(11), values(12), DecodeText(values(13)), values(14)
    record.Seal
    Set NxTemplateReadMetadata = record
End Function

Public Function NxTemplateListMetadata(ByVal root As String, Optional ByVal verifyPackage As Boolean = True) As Collection
    Dim result As New Collection, fileSystem As Object, packages As Object, folder As Object
    Dim record As CNxTemplateRecord
    Dim names() As String, count As Long, index As Long, inner As Long, swap As String
    NxTemplateEnsureStore root
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set packages = fileSystem.GetFolder(NxTemplatePackagesRoot(root))
    For Each folder In packages.SubFolders
        If NxTemplateIdIsValid(folder.Name) Then
            count = count + 1
            ReDim Preserve names(1 To count)
            names(count) = folder.Name
        End If
    Next folder
    For index = 2 To count
        swap = names(index)
        inner = index - 1
        Do While inner >= 1
            If StrComp(names(inner), swap, vbBinaryCompare) <= 0 Then Exit Do
            names(inner + 1) = names(inner)
            inner = inner - 1
        Loop
        names(inner + 1) = swap
    Next index
    For index = 1 To count
        Set record = NxTemplateReadRecordFolder(NxTemplateTemplateFolderPath(names(index), root), True, verifyPackage)
        result.Add record
    Next index
    Set NxTemplateListMetadata = result
End Function

Public Function NxTemplateReadRecordFolder(ByVal folderPath As String, Optional ByVal allowPlaceholder As Boolean = True, Optional ByVal verifyPackage As Boolean = True) As CNxTemplateRecord
    Dim fileSystem As Object, templateId As String, record As CNxTemplateRecord, status As String
    On Error GoTo Damaged
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then NxRaiseContractError "Template folder is missing"
    templateId = fileSystem.GetFileName(folderPath)
    If Not NxTemplateIdIsValid(templateId) Then NxRaiseContractError "Template folder ID is invalid"
    If Not NxTemplateFolderHasExactFiles(folderPath) Then NxRaiseContractError "Template folder file set is invalid"
    Set record = NxTemplateReadMetadata(folderPath & "\metadata.ini")
    If record.TemplateId <> templateId Then NxRaiseContractError "Template metadata ID does not match its folder"
    If record.Status = "damaged" Then
        status = "damaged"
    ElseIf Not verifyPackage Then
        status = record.Status
    Else
        status = NxTemplateRecordIntegrityStatus(record, folderPath & "\template.xlsx")
    End If
    If status = record.Status Then
        Set NxTemplateReadRecordFolder = record
    Else
        Set NxTemplateReadRecordFolder = CopyWithStatus(record, status)
    End If
    Exit Function
Damaged:
    If allowPlaceholder And Len(templateId) > 0 And NxTemplateIdIsValid(templateId) Then
        Set NxTemplateReadRecordFolder = DamagedPlaceholder(templateId)
        Err.Clear
        Exit Function
    End If
    NxRaiseContractError "Template folder record is invalid"
End Function

Public Sub NxTemplateReplaceMetadata(ByVal folderPath As String, ByVal record As CNxTemplateRecord)
    ReplaceMetadataCore folderPath, record, False
End Sub

Public Sub NxTemplateReplaceMetadataForTest(ByVal folderPath As String, ByVal record As CNxTemplateRecord, ByVal failAfterBackup As Boolean)
    ReplaceMetadataCore folderPath, record, failAfterBackup
End Sub

Public Function NxTemplateFolderHasExactFiles(ByVal folderPath As String) As Boolean
    Dim fileSystem As Object, folder As Object
    On Error GoTo NotExact
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then Exit Function
    Set folder = fileSystem.GetFolder(folderPath)
    If folder.SubFolders.Count <> 0 Or folder.Files.Count <> 2 Then Exit Function
    If Not fileSystem.FileExists(folderPath & "\template.xlsx") Then Exit Function
    If Not fileSystem.FileExists(folderPath & "\metadata.ini") Then Exit Function
    NxTemplateFolderHasExactFiles = True
NotExact:
End Function

Private Sub ReplaceMetadataCore(ByVal folderPath As String, ByVal record As CNxTemplateRecord, ByVal failAfterBackup As Boolean)
    Dim fileSystem As Object, currentPath As String, newPath As String, backupPath As String
    Dim replacementStage As Long
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    currentPath = folderPath & "\metadata.ini"
    newPath = folderPath & "\metadata.new"
    backupPath = folderPath & "\metadata.bak"
    On Error GoTo Failed
    If Not NxTemplateFolderHasExactFiles(folderPath) Then NxRaiseContractError "Template folder is not replaceable"
    If fileSystem.FileExists(newPath) Or fileSystem.FileExists(backupPath) Then NxRaiseContractError "Template metadata residue blocks replacement"
    NxTemplateAssertLocalFolder folderPath
    NxTemplateBackupMetadata folderPath
    NxTemplateWriteMetadata newPath, record
    AssertSameRecord record, NxTemplateReadMetadata(newPath)
    Name currentPath As backupPath
    replacementStage = 1
    If failAfterBackup Then NxRaiseContractError "Injected metadata replacement failure"
    Name newPath As currentPath
    replacementStage = 2
    AssertSameRecord record, NxTemplateReadMetadata(currentPath)
    Kill backupPath
    replacementStage = 3
    If Not NxTemplateFolderHasExactFiles(folderPath) Then NxRaiseContractError "Template metadata replacement left residue"
    Exit Sub
Failed:
    On Error Resume Next
    If (replacementStage = 1 Or replacementStage = 2) And fileSystem.FileExists(backupPath) Then
        If fileSystem.FileExists(currentPath) Then fileSystem.DeleteFile currentPath, True
        Name backupPath As currentPath
    End If
    If fileSystem.FileExists(newPath) Then fileSystem.DeleteFile newPath, True
    If fileSystem.FileExists(backupPath) And fileSystem.FileExists(currentPath) Then fileSystem.DeleteFile backupPath, True
    On Error GoTo 0
    NxRaiseContractError "Template metadata replacement failed"
End Sub

Private Function ReadMetadataLines(ByVal path As String) As Variant
    Dim handle As Integer, text As String, lines As Variant
    On Error GoTo Failed
    handle = FreeFile
    Open path For Binary Access Read As #handle
    If LOF(handle) <= 0 Or LOF(handle) > 1048576 Then GoTo Failed
    text = Space$(LOF(handle))
    Get #handle, , text
    Close #handle
    If Right$(text, 2) <> vbCrLf Then NxRaiseContractError "Template metadata line ending is invalid"
    text = Left$(text, Len(text) - 2)
    If InStr(text, vbCr & vbCr) > 0 Or InStr(Replace(text, vbCrLf, vbNullString), vbCr) > 0 Then NxRaiseContractError "Template metadata contains invalid line endings"
    lines = Split(text, vbCrLf)
    If UBound(lines) <> 12 And UBound(lines) <> 14 Then NxRaiseContractError "Template metadata field count is invalid"
    ReadMetadataLines = lines
    Exit Function
Failed:
    On Error Resume Next
    Close #handle
    On Error GoTo 0
    NxRaiseContractError "Template metadata read failed"
End Function

Private Function ExpectedKey(ByVal lineIndex As Long) As String
    ExpectedKey = Split("schema_version|template_id|display_name|description|package_sha256|logical_sha256|source_kind|row_count|column_count|risk_flags|status|created_utc|modified_utc|category|last_used_utc", "|")(lineIndex)
End Function

Private Function EncodeText(ByVal value As String) As String
    Dim index As Long, codeUnit As Long
    For index = 1 To Len(value)
        codeUnit = AscW(Mid$(value, index, 1))
        If codeUnit < 0 Then codeUnit = codeUnit + 65536
        EncodeText = EncodeText & "%" & Right$("0000" & UCase$(Hex$(codeUnit)), 4)
    Next index
End Function

Private Function DecodeText(ByVal value As String) As String
    Dim index As Long, hexText As String, codeUnit As Long
    If Len(value) Mod 5 <> 0 Then NxRaiseContractError "Template metadata text encoding is invalid"
    For index = 1 To Len(value) Step 5
        If Mid$(value, index, 1) <> "%" Then NxRaiseContractError "Template metadata text encoding is invalid"
        hexText = Mid$(value, index + 1, 4)
        If Not UpperHexText(hexText) Then NxRaiseContractError "Template metadata text encoding is invalid"
        codeUnit = CLng("&H" & hexText)
        If codeUnit > 32767 Then codeUnit = codeUnit - 65536
        DecodeText = DecodeText & ChrW$(codeUnit)
    Next index
End Function

Private Function UpperHexText(ByVal value As String) As Boolean
    Dim index As Long
    If Len(value) <> 4 Then Exit Function
    For index = 1 To Len(value)
        If InStr(1, "0123456789ABCDEF", Mid$(value, index, 1), vbBinaryCompare) = 0 Then Exit Function
    Next index
    UpperHexText = True
End Function

Private Function PositiveIntegerText(ByVal value As String) As Boolean
    Dim index As Long
    If Len(value) = 0 Or Len(value) > 6 Or Left$(value, 1) = "0" Then Exit Function
    For index = 1 To Len(value)
        If InStr(1, "0123456789", Mid$(value, index, 1), vbBinaryCompare) = 0 Then Exit Function
    Next index
    On Error GoTo Invalid
    PositiveIntegerText = (CLng(value) > 0)
Invalid:
End Function

Private Function ClosedTimestamp(ByVal value As String) As Boolean
    Dim index As Long, character As String
    If Len(value) <> 20 Then Exit Function
    If Mid$(value, 5, 1) <> "-" Or Mid$(value, 8, 1) <> "-" Or Mid$(value, 11, 1) <> "T" Then Exit Function
    If Mid$(value, 14, 1) <> ":" Or Mid$(value, 17, 1) <> ":" Or Right$(value, 1) <> "Z" Then Exit Function
    For index = 1 To Len(value)
        Select Case index
            Case 5, 8, 11, 14, 17, 20
            Case Else
                character = Mid$(value, index, 1)
                If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then Exit Function
        End Select
    Next index
    ClosedTimestamp = True
End Function

Private Sub AssertSameRecord(ByVal expected As CNxTemplateRecord, ByVal actual As CNxTemplateRecord)
    If expected Is Nothing Or actual Is Nothing Then NxRaiseContractError "Template metadata read-back changed"
    If expected.TemplateId <> actual.TemplateId Or expected.DisplayName <> actual.DisplayName Or _
       expected.Description <> actual.Description Or expected.PackageSha256 <> actual.PackageSha256 Or _
       expected.LogicalSha256 <> actual.LogicalSha256 Or expected.SourceKind <> actual.SourceKind Or _
       expected.RowCount <> actual.RowCount Or expected.ColumnCount <> actual.ColumnCount Or _
       expected.RiskFlags <> actual.RiskFlags Or expected.Status <> actual.Status Or _
       expected.CreatedUtc <> actual.CreatedUtc Or expected.ModifiedUtc <> actual.ModifiedUtc Or _
       expected.Category <> actual.Category Or expected.LastUsedUtc <> actual.LastUsedUtc Then _
        NxRaiseContractError "Template metadata read-back changed"
End Sub

Private Function CopyWithStatus(ByVal source As CNxTemplateRecord, ByVal status As String) As CNxTemplateRecord
    Dim result As New CNxTemplateRecord
    result.Configure source.TemplateId, source.DisplayName, source.Description, source.PackageSha256, source.LogicalSha256, _
        source.SourceKind, source.RowCount, source.ColumnCount, source.RiskFlags, status, source.CreatedUtc, source.ModifiedUtc, source.Category, source.LastUsedUtc
    result.Seal
    Set CopyWithStatus = result
End Function

Private Function DamagedPlaceholder(ByVal templateId As String) As CNxTemplateRecord
    Dim result As New CNxTemplateRecord, timestamp As String
    timestamp = Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")
    result.Configure templateId, "손상된 템플릿 (" & templateId & ")", "invalid-folder-record", String$(64, "0"), _
        String$(64, "0"), "range", 1, 1, vbNullString, "damaged", timestamp, timestamp
    result.Seal
    Set DamagedPlaceholder = result
End Function
