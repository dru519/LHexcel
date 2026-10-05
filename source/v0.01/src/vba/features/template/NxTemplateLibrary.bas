Attribute VB_Name = "NxTemplateLibrary"
Option Explicit

Public Function NxTemplateCopyDetails(ByVal source As CNxTemplateRecord, ByVal title As String, ByVal description As String, _
    ByVal category As String, ByVal lastUsed As String, ByVal modified As String) As CNxTemplateRecord
    Dim result As New CNxTemplateRecord
    result.Configure source.TemplateId, title, description, source.PackageSha256, source.LogicalSha256, source.SourceKind, _
        source.RowCount, source.ColumnCount, source.RiskFlags, source.Status, source.CreatedUtc, modified, category, lastUsed
    result.Seal
    Set NxTemplateCopyDetails = result
End Function

Public Function NxTemplateLibraryRecords(ByVal query As String, ByVal category As String, ByVal sortMode As Long, _
    ByVal deleted As Boolean, ByRef keys As Collection, Optional ByVal root As String = vbNullString) As Collection
    Dim all As Collection, result As New Collection, record As CNxTemplateRecord, fs As Object, batch As Object, folder As Object
    Dim records() As CNxTemplateRecord, ids() As String, count As Long, i As Long, j As Long, key As String, temp As CNxTemplateRecord
    If Len(root) = 0 Then root = NxTemplateStoreRoot()
    NxTemplateEnsureStore root
    Set keys = New Collection
    Set all = New Collection
    Set fs = CreateObject("Scripting.FileSystemObject")
    If deleted Then
        For Each batch In fs.GetFolder(root & "\Quarantine").SubFolders
            If NxTemplateIdIsValid(batch.Name) And (batch.Attributes And 1024) = 0 Then
                For Each folder In batch.SubFolders
                    If NxTemplateIdIsValid(folder.Name) And (folder.Attributes And 1024) = 0 Then
                        Set record = NxTemplateReadRecordFolder(folder.Path, True, False)
                        all.Add record
                        keys.Add batch.Name
                    End If
                Next folder
            End If
        Next batch
    Else
        Set all = NxTemplateListMetadata(root, False)
        Set all = NxTemplateApplyOrder(all, root)
        For Each record In all: keys.Add vbNullString: Next record
    End If
    Set NxTemplateLibraryRecords = NxTemplateFilterRecords(all, keys, query, category, sortMode, keys)
End Function

Public Function NxTemplateFilterRecords(ByVal all As Collection, ByVal inputKeys As Collection, ByVal query As String, _
    ByVal category As String, ByVal sortMode As Long, ByRef keys As Collection) As Collection
    Dim result As New Collection, record As CNxTemplateRecord, temp As CNxTemplateRecord
    Dim records() As CNxTemplateRecord, ids() As String, count As Long, i As Long, j As Long, key As String
    If all.Count > 0 Then
        ReDim records(1 To all.Count): ReDim ids(1 To all.Count)
    End If
    For i = 1 To all.Count
        Set record = all(i)
        If (Len(query) = 0 Or InStr(1, record.DisplayName & " " & record.Description & " " & record.Category, query, vbTextCompare) > 0) And _
            (Len(category) = 0 Or StrComp(record.Category, category, vbTextCompare) = 0) Then
            count = count + 1: Set records(count) = record: ids(count) = inputKeys(i)
        End If
    Next i
    For i = 2 To count
        If sortMode = 0 Then Exit For
        Set temp = records(i): key = ids(i): j = i - 1
        Do While j >= 1
            If Not ComesBefore(temp, records(j), sortMode) Then Exit Do
            Set records(j + 1) = records(j): ids(j + 1) = ids(j): j = j - 1
        Loop
        Set records(j + 1) = temp: ids(j + 1) = key
    Next i
    Set keys = New Collection
    For i = 1 To count: result.Add records(i): keys.Add ids(i): Next i
    Set NxTemplateFilterRecords = result
End Function

Private Function NxTemplateApplyOrder(ByVal records As Collection, ByVal root As String) As Collection
    Dim result As New Collection, lookup As Object, seen As Object, record As CNxTemplateRecord
    Dim path As String, handle As Integer, id As String
    Set lookup = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    For Each record In records: Set lookup(record.TemplateId) = record: Next record
    path = root & "\order.txt"
    If Len(Dir$(path)) > 0 Then
        On Error GoTo Failed
        handle = FreeFile: Open path For Input As #handle
        If LOF(handle) > 1048576 Then NxRaiseContractError "템플릿 순서 파일이 너무 큽니다."
        Do While Not EOF(handle)
            Line Input #handle, id
            If Not NxTemplateIdIsValid(id) Then NxRaiseContractError "템플릿 순서 파일을 확인하세요."
            If lookup.Exists(id) And Not seen.Exists(id) Then result.Add lookup(id): seen(id) = True
        Loop
        Close #handle: handle = 0
    End If
    For Each record In records
        If Not seen.Exists(record.TemplateId) Then result.Add record
    Next record
    Set NxTemplateApplyOrder = result
    Exit Function
Failed:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    On Error GoTo 0
    NxRaiseContractError "템플릿 순서 파일을 읽지 못했습니다."
End Function

Public Function NxTemplateMove(ByVal templateId As String, ByVal direction As Long, Optional ByVal root As String = vbNullString) As CNxResult
    Dim records As Collection, keys As Collection, i As Long, selected As Long, adjacent As Long, handle As Integer
    Dim path As String, temporary As String, backup As String, fs As Object, movedOld As Boolean, committed As Boolean, failure As String
    On Error GoTo Failed
    If direction <> -1 And direction <> 1 Then NxRaiseContractError "이동 방향이 잘못되었습니다."
    If Len(root) = 0 Then root = NxTemplateStoreRoot()
    NxTemplateAssertLocalFolder root
    Set records = NxTemplateLibraryRecords("", "", 0, False, keys, root)
    For i = 1 To records.Count
        If records(i).TemplateId = templateId Then selected = i: Exit For
    Next i
    If selected = 0 Then NxRaiseContractError "선택한 템플릿이 없습니다."
    adjacent = selected + direction
    If adjacent >= 1 And adjacent <= records.Count Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        path = root & "\order.txt"
        temporary = root & "\order-" & Replace(NxCreateRunUuid(), "-", "") & ".tmp"
        backup = root & "\order-" & Replace(NxCreateRunUuid(), "-", "") & ".bak"
        handle = FreeFile: Open temporary For Output As #handle
        For i = 1 To records.Count
            If i = selected Then
                Print #handle, records(adjacent).TemplateId
            ElseIf i = adjacent Then
                Print #handle, records(selected).TemplateId
            Else
                Print #handle, records(i).TemplateId
            End If
        Next i
        Close #handle: handle = 0
        If fs.FileExists(path) Then Name path As backup: movedOld = True
        Name temporary As path
        committed = True
    End If
    Set NxTemplateMove = NxCreateResult("NX-TPL-RENAME", NxSuccess, "complete", templateId, False, vbNullString, "result_success")
    Exit Function
Failed:
    failure = Err.Description
    On Error Resume Next
    If handle <> 0 Then Close #handle
    If movedOld And Not committed Then Name backup As path
    If Len(temporary) > 0 Then If fs.FileExists(temporary) Then fs.DeleteFile temporary
    On Error GoTo 0
    Set NxTemplateMove = NxCreateResult("NX-TPL-RENAME", NxEnvironmentError, "move", templateId, False, failure, "result_environment_error")
End Function

Private Function ComesBefore(ByVal a As CNxTemplateRecord, ByVal b As CNxTemplateRecord, ByVal mode As Long) As Boolean
    Dim leftKey As String, rightKey As String
    Select Case mode
        Case 1: leftKey = a.ModifiedUtc: rightKey = b.ModifiedUtc
        Case 2: leftKey = a.LastUsedUtc: rightKey = b.LastUsedUtc
        Case Else: ComesBefore = StrComp(a.DisplayName, b.DisplayName, vbTextCompare) < 0: Exit Function
    End Select
    If leftKey = rightKey Then
        ComesBefore = StrComp(a.DisplayName, b.DisplayName, vbTextCompare) < 0
    Else
        ComesBefore = leftKey > rightKey
    End If
End Function

Public Sub NxTemplateBackupMetadata(ByVal folderPath As String)
    Dim fs As Object, root As String, backups As String, destination As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    root = fs.GetParentFolderName(fs.GetParentFolderName(folderPath))
    backups = root & "\Backups"
    If Not fs.FolderExists(backups) Then fs.CreateFolder backups
    NxTemplateAssertLocalFolder backups
    destination = backups & "\" & fs.GetFileName(folderPath) & "-" & Replace(NxCreateRunUuid(), "-", "") & ".ini"
    fs.CopyFile folderPath & "\metadata.ini", destination, False
End Sub

Public Function NxTemplateRecordUse(ByVal record As CNxTemplateRecord, ByVal root As String) As String
    Dim changed As CNxTemplateRecord
    On Error GoTo Failed
    Set changed = NxTemplateCopyDetails(record, record.DisplayName, record.Description, record.Category, _
        Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z"), record.ModifiedUtc)
    NxTemplateReplaceMetadata NxTemplateTemplateFolderPath(record.TemplateId, root), changed
    Exit Function
Failed:
    NxTemplateRecordUse = "최근 사용 기록 저장 실패: " & Err.Description
End Function

Public Sub NxTemplateAssertLocalFolder(ByVal folderPath As String)
    Dim fs As Object, cursor As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    cursor = fs.GetAbsolutePathName(folderPath)
    Do While Len(cursor) > 0
        If fs.FolderExists(cursor) Then
            If (fs.GetFolder(cursor).Attributes And 1024) <> 0 Then NxRaiseContractError "연결 폴더는 템플릿 작업에 사용할 수 없습니다."
        End If
        cursor = fs.GetParentFolderName(cursor)
    Loop
End Sub

Public Function NxTemplateRestore(ByVal templateId As String, ByVal quarantineId As String, _
    Optional ByVal root As String = vbNullString, Optional ByVal failAfterMove As Boolean = False) As CNxResult
    Dim source As String, destination As String, fs As Object, record As CNxTemplateRecord, item As CNxTemplateRecord
    Dim moved As Boolean, failure As String, restored As Boolean
    On Error GoTo Failed
    If Len(root) = 0 Then root = NxTemplateStoreRoot()
    destination = NxTemplateTemplateFolderPath(templateId, root)
    source = NxTemplateQuarantinePath(quarantineId, root) & "\" & templateId
    NxTemplateAssertLocalFolder source
    NxTemplateAssertLocalFolder destination
    Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(destination) Then NxRaiseContractError "같은 템플릿이 이미 있습니다."
    Set record = NxTemplateReadRecordFolder(source, False)
    If record.Status <> "healthy" Then NxRaiseContractError "손상된 템플릿은 복원할 수 없습니다."
    For Each item In NxTemplateListMetadata(root, False)
        If StrComp(item.DisplayName, record.DisplayName, vbTextCompare) = 0 Then NxRaiseContractError "같은 이름의 템플릿을 먼저 변경하세요."
    Next item
    fs.MoveFolder source, destination: moved = True
    If failAfterMove Then NxRaiseContractError "Injected restore failure"
    Set item = NxTemplateReadRecordFolder(destination, False)
    If item.Status <> "healthy" Then NxRaiseContractError "복원 후 무결성 검사를 통과하지 못했습니다."
    Set NxTemplateRestore = NxCreateResult("NX-TPL-DELETE", NxSuccess, "complete", templateId, False, vbNullString, "result_success")
    Exit Function
Failed:
    failure = Err.Description
    If moved Then
        On Error Resume Next
        fs.MoveFolder destination, source
        restored = fs.FolderExists(source) And Not fs.FolderExists(destination)
        On Error GoTo 0
        If Not restored Then failure = failure & " / 자동 복구 실패: " & destination
    End If
    Set NxTemplateRestore = NxCreateResult("NX-TPL-DELETE", NxEnvironmentError, "restore", templateId, False, failure, "result_environment_error")
End Function
