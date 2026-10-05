Attribute VB_Name = "NxBatchRenameEngine"
Option Explicit

#Const NX_R58_TEST_BUILD = False

Private Const NX_BATCH_RENAME_LIMIT As Long = 1000
Private mLastFileJournal As CNxBatchRenameJournal
Private mLastSheetJournal As Collection
Private mLastSheetBook As Workbook
#If NX_R58_TEST_BUILD Then
Private mTestFailMoveStage As String
Private mTestFailMoveIndex As Long

' Native VBA suites use this deterministic, one-shot seam to exercise the
' recovery contract without replacing production file operations.
Public Sub NxBatchRenameTestInjectMoveFailure(ByVal stageName As String, ByVal itemIndex As Long)
    If Len(Trim$(stageName)) = 0 Or itemIndex < 1 Then NxRaiseContractError "Rename test move failure requires a stage and positive index"
    Select Case LCase$(Trim$(stageName))
        Case "original-to-temporary", "temporary-to-target", "sheet-original-to-temporary", "sheet-temporary-to-target"
        Case Else
            NxRaiseContractError "Rename test move failure stage is invalid: " & stageName
    End Select
    mTestFailMoveStage = stageName
    mTestFailMoveIndex = itemIndex
End Sub

Public Sub NxBatchRenameTestClearMoveFailure()
    mTestFailMoveStage = vbNullString
    mTestFailMoveIndex = 0
End Sub
#End If

Public Function NxBatchRenameBuildPreview(ByVal pathsText As String, ByVal findText As String, _
    ByVal replaceText As String, ByVal prefixText As String, ByVal suffixText As String, _
    ByRef HasDuplicateTarget As Boolean, ByRef HasExistingCollision As Boolean) As Collection

    Dim preview As New Collection, lines As Variant, rawPath As Variant
    Dim originalPath As String, targetPath As String, statusText As String
    Dim originalSet As Object, targetSet As Object, movingSet As Object, row As Variant

    Set originalSet = CreateObject("Scripting.Dictionary")
    Set targetSet = CreateObject("Scripting.Dictionary")
    Set movingSet = CreateObject("Scripting.Dictionary")
    originalSet.CompareMode = vbTextCompare
    targetSet.CompareMode = vbTextCompare
    movingSet.CompareMode = vbTextCompare
    lines = Split(Replace$(Replace$(pathsText, vbCrLf, vbLf), vbCr, vbLf), vbLf)

    For Each rawPath In lines
        originalPath = Trim$(Replace$(CStr(rawPath), Chr$(34), vbNullString))
        If Len(originalPath) > 0 Then
            If preview.Count >= NX_BATCH_RENAME_LIMIT Then NxRaiseContractError "이름 변경은 한 번에 1,000개까지 가능합니다."
            If originalSet.Exists(originalPath) Then
                statusText = "오류: 중복 입력"
            ElseIf Not NxBatchRenameFileExists(originalPath) Then
                statusText = "오류: 파일 없음"
            ElseIf (GetAttr(originalPath) And vbDirectory) <> 0 Then
                statusText = "오류: 폴더 제외"
            Else
                originalSet.Add originalPath, True
                targetPath = NxBatchRenameTargetPath(originalPath, findText, replaceText, prefixText, suffixText)
                statusText = IIf(StrComp(originalPath, targetPath, vbTextCompare) = 0, "변경 없음", "준비")
            End If
            row = Array(originalPath, targetPath, statusText)
            preview.Add row
        End If
    Next rawPath
    If preview.Count = 0 Then NxRaiseContractError "이름을 변경할 파일 경로를 입력하세요."

    For Each row In preview
        If CStr(row(2)) = "준비" Then
            movingSet.Add CStr(row(0)), True
            If targetSet.Exists(CStr(row(1))) Then
                HasDuplicateTarget = True
            Else
                targetSet.Add CStr(row(1)), True
            End If
        End If
    Next row
    For Each row In preview
        If CStr(row(2)) = "준비" Then
            If NxBatchRenameFileExists(CStr(row(1))) And Not movingSet.Exists(CStr(row(1))) Then _
                HasExistingCollision = True
        End If
    Next row
    Set NxBatchRenameBuildPreview = NxBatchRenameRevalidatePreview(preview, HasDuplicateTarget, HasExistingCollision)
End Function

Public Function NxBatchRenameRevalidatePreview(ByVal preview As Collection, ByRef hasDuplicate As Boolean, _
    ByRef hasCollision As Boolean) As Collection
    Dim checked As New Collection, row As Variant, sources As Object, targets As Object, moving As Object
    Dim reader As New CNxBatchRenameJournal, originalPath As String, targetPath As String, statusText As String
    Dim fso As Object, key As String
    If preview Is Nothing Then NxRaiseContractError "이름 변경 미리보기가 필요합니다."
    If preview.Count = 0 Or preview.Count > NX_BATCH_RENAME_LIMIT Then NxRaiseContractError "이름 변경 대상은 1~1,000개여야 합니다."
    hasDuplicate = False: hasCollision = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set sources = CreateObject("Scripting.Dictionary"): sources.CompareMode = vbTextCompare
    Set targets = CreateObject("Scripting.Dictionary"): targets.CompareMode = vbTextCompare
    Set moving = CreateObject("Scripting.Dictionary"): moving.CompareMode = vbTextCompare
    For Each row In preview
        If Not IsArray(row) Then NxRaiseContractError "이름 변경 미리보기를 다시 만드세요."
        If LBound(row) <> 0 Or UBound(row) < 2 Then NxRaiseContractError "이름 변경 미리보기를 다시 만드세요."
        originalPath = CStr(row(0)): targetPath = CStr(row(1))
        statusText = vbNullString
        If sources.Exists(originalPath) Then
            statusText = "오류: 중복 입력"
        Else
            sources.Add originalPath, True
        End If
        If Not fso.FileExists(originalPath) Then
            statusText = "오류: 파일 없음"
        ElseIf InStrRev(originalPath, "\") = 0 Or Len(targetPath) > 240 Then
            statusText = "오류: 경로 규칙"
        ElseIf StrComp(fso.GetAbsolutePathName(originalPath), originalPath, vbTextCompare) <> 0 Then
            statusText = "오류: 절대 경로 필요"
        ElseIf StrComp(fso.GetParentFolderName(originalPath), fso.GetParentFolderName(targetPath), vbTextCompare) <> 0 Then
            statusText = "오류: 다른 폴더로 이동 불가"
        ElseIf Not NxFileRenameNameIsValid(NxRenameLeafName(targetPath)) Then
            statusText = "오류: 파일명 규칙"
        ElseIf UBound(row) >= 3 Then
            If StrComp(CStr(row(3)), reader.FileIdentity(originalPath), vbBinaryCompare) <> 0 Then statusText = "오류: 원본 정체성 변경"
        End If
        If Len(statusText) = 0 Then
            If StrComp(originalPath, targetPath, vbBinaryCompare) = 0 Then
                statusText = "변경 없음"
            Else
                statusText = "준비"
                moving.Add originalPath, True
            End If
        End If
        row(2) = statusText
        checked.Add row
        key = LCase$(targetPath)
        If targets.Exists(key) Then targets.Item(key) = CLng(targets.Item(key)) + 1 Else targets.Add key, 1
    Next row
    Set NxBatchRenameRevalidatePreview = New Collection
    For Each row In checked
        If CLng(targets.Item(LCase$(CStr(row(1))))) > 1 Then
            row(2) = "오류: 중복 이름": hasDuplicate = True
        ElseIf CStr(row(2)) = "준비" Then
            If (fso.FileExists(CStr(row(1))) Or fso.FolderExists(CStr(row(1)))) And Not moving.Exists(CStr(row(1))) Then
                row(2) = "오류: 기존 파일 충돌": hasCollision = True
            End If
        End If
        NxBatchRenameRevalidatePreview.Add row
    Next row
End Function

Public Sub NxBatchRenameRequireReady(ByVal preview As Collection)
    Dim checked As Collection, row As Variant, hasDuplicate As Boolean, hasCollision As Boolean, changes As Long
    Set checked = NxBatchRenameRevalidatePreview(preview, hasDuplicate, hasCollision)
    For Each row In checked
        If Left$(CStr(row(2)), 3) = "오류:" Then NxRaiseContractError CStr(row(0)) & " · " & CStr(row(2))
        If CStr(row(2)) = "준비" Then changes = changes + 1
    Next row
    If changes = 0 Then NxRaiseContractError "변경할 파일 이름이 없습니다."
End Sub

Public Function NxBatchRenameCommit(ByVal preview As Collection) As Long
    Dim journal As New CNxBatchRenameJournal, row As Variant, index As Long
    Dim temporaryPath As String, errorNumber As Long, errorDescription As String, recoveryDescription As String
    Dim movingSet As Object, hasDuplicate As Boolean, hasCollision As Boolean
    NxBatchRenameRequireReady preview
    Set preview = NxBatchRenameRevalidatePreview(preview, hasDuplicate, hasCollision)
    Set movingSet = CreateObject("Scripting.Dictionary")
    movingSet.CompareMode = vbTextCompare
    For Each row In preview
        If CStr(row(2)) = "준비" Then movingSet.Add CStr(row(0)), True
    Next row

    For Each row In preview
        If CStr(row(2)) = "준비" Then
            If Not NxBatchRenameFileExists(CStr(row(0))) Then NxRaiseContractError "미리보기 이후 원본 파일이 변경되었습니다."
            If NxBatchRenameFileExists(CStr(row(1))) And Not movingSet.Exists(CStr(row(1))) Then _
                NxRaiseContractError "미리보기 이후 대상 파일 충돌이 발생했습니다."
            temporaryPath = Left$(CStr(row(0)), InStrRev(CStr(row(0)), "\")) & ".nx-" & Replace$(NxCreateRunUuid(), "-", vbNullString) & ".temporary"
            If Len(temporaryPath) > 240 Then NxRaiseContractError "임시 이름을 안전하게 만들기에는 폴더 경로가 너무 깁니다."
            If NxBatchRenameFileExists(temporaryPath) Then NxRaiseContractError "임시 이름 충돌이 발생했습니다."
            journal.Add CStr(row(0)), temporaryPath, CStr(row(1))
        End If
    Next row
    If journal.Count = 0 Then NxRaiseContractError "변경할 파일 이름이 없습니다."
    journal.Seal

    On Error GoTo Failed
    For index = 1 To journal.Count
        NxBatchRenameMaybeInjectMoveFailure "original-to-temporary", index
        journal.MoveOriginalToTemporary index
    Next index
    For index = 1 To journal.Count
        NxBatchRenameMaybeInjectMoveFailure "temporary-to-target", index
        journal.MoveTemporaryToTarget index
    Next index
    NxBatchRenameCommit = journal.Count
    Set mLastFileJournal = journal
    Set mLastSheetJournal = Nothing
    Set mLastSheetBook = Nothing
    Exit Function
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    On Error Resume Next
    journal.Rollback
    If Err.Number <> 0 Then recoveryDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If Len(recoveryDescription) > 0 Then
        ' Retain exact identities/stages for a safe same-session recovery retry.
        Set mLastFileJournal = journal
        Set mLastSheetJournal = Nothing
        Set mLastSheetBook = Nothing
        errorDescription = errorDescription & " | " & recoveryDescription
    End If
    Err.Raise errorNumber, "LHexcel.File.Rename", errorDescription
End Function

Public Function NxSheetBatchRenameBuildPreview(ByVal book As Workbook, ByVal findText As String, _
    ByVal replaceText As String, ByVal prefixText As String, ByVal suffixText As String, _
    ByRef HasDuplicateTarget As Boolean, ByRef HasExistingCollision As Boolean) As Collection

    Dim preview As New Collection, sheet As Worksheet, targetName As String, row As Variant, structure As String
    If book Is Nothing Then NxRaiseContractError "열린 통합문서가 없습니다."
    structure = NxSheetBatchRenameStructure(book)
    For Each sheet In book.Worksheets
        targetName = sheet.Name
        If Len(findText) > 0 Then targetName = Replace$(targetName, findText, replaceText, 1, -1, vbTextCompare)
        targetName = prefixText & targetName & suffixText
        row = Array(sheet.Name, targetName, vbNullString, Empty, book.FullName, structure)
        Set row(3) = sheet
        preview.Add row
    Next sheet
    Set NxSheetBatchRenameBuildPreview = NxSheetBatchRenameRevalidatePreview(book, preview, HasDuplicateTarget, HasExistingCollision)
End Function

Public Function NxSheetBatchRenameEditPreview(ByVal book As Workbook, ByVal preview As Collection, _
    ByVal rowIndex As Long, ByVal plannedName As String, ByRef HasDuplicateTarget As Boolean, _
    ByRef HasExistingCollision As Boolean) As Collection

    Dim edited As New Collection, row As Variant, index As Long
    NxSheetBatchRenameValidateSnapshot book, preview
    If rowIndex < 1 Or rowIndex > preview.Count Then NxRaiseContractError "수정할 시트를 선택하세요."
    ' Build a new draft; a previously confirmed collection is never changed.
    For Each row In preview
        index = index + 1
        If index = rowIndex Then row(1) = plannedName
        edited.Add row
    Next row
    Set NxSheetBatchRenameEditPreview = NxSheetBatchRenameRevalidatePreview(book, edited, HasDuplicateTarget, HasExistingCollision)
End Function

Public Function NxSheetBatchRenameRevalidatePreview(ByVal book As Workbook, ByVal preview As Collection, _
    ByRef HasDuplicateTarget As Boolean, ByRef HasExistingCollision As Boolean) As Collection

    Dim checked As New Collection, row As Variant, targetSet As Object, movingSet As Object, existing As Object
    NxSheetBatchRenameValidateSnapshot book, preview
    HasDuplicateTarget = False
    HasExistingCollision = False
    Set targetSet = CreateObject("Scripting.Dictionary")
    Set movingSet = CreateObject("Scripting.Dictionary")
    targetSet.CompareMode = vbTextCompare
    movingSet.CompareMode = vbTextCompare
    For Each row In preview
        If targetSet.Exists(CStr(row(1))) Then
            targetSet.Item(CStr(row(1))) = CLng(targetSet.Item(CStr(row(1)))) + 1
        Else
            targetSet.Add CStr(row(1)), 1
        End If
        If NxSheetRenameNameIsValid(CStr(row(1))) And StrComp(CStr(row(0)), CStr(row(1)), vbBinaryCompare) <> 0 Then _
            movingSet.Add CStr(row(0)), True
    Next row
    For Each row In preview
        If Not NxSheetRenameNameIsValid(CStr(row(1))) Then
            row(2) = "오류: 시트명 규칙"
        ElseIf CLng(targetSet.Item(CStr(row(1)))) > 1 Then
            HasDuplicateTarget = True
            row(2) = "오류: 중복 이름"
        ElseIf StrComp(CStr(row(0)), CStr(row(1)), vbBinaryCompare) = 0 Then
            row(2) = "변경 없음"
        Else
            Set existing = NxBatchRenameSheet(book, CStr(row(1)))
            If Not existing Is Nothing Then
                If Not movingSet.Exists(existing.Name) Then
                    row(2) = "오류: 기존 시트 충돌"
                    HasExistingCollision = True
                Else
                    row(2) = "준비"
                End If
            Else
                row(2) = "준비"
            End If
        End If
        checked.Add row
    Next row
    Set NxSheetBatchRenameRevalidatePreview = checked
End Function

Public Sub NxSheetBatchRenameRequireReady(ByVal book As Workbook, ByVal preview As Collection)
    Dim checked As Collection, row As Variant, hasDuplicate As Boolean, hasCollision As Boolean, changes As Long
    Set checked = NxSheetBatchRenameRevalidatePreview(book, preview, hasDuplicate, hasCollision)
    If book.ProtectStructure Then NxRaiseContractError "통합문서 구조가 보호되어 있습니다."
    For Each row In checked
        If Left$(CStr(row(2)), 3) = "오류:" Then NxRaiseContractError CStr(row(0)) & " · " & CStr(row(2))
        If CStr(row(2)) = "준비" Then changes = changes + 1
    Next row
    If changes = 0 Then NxRaiseContractError "변경할 시트 이름이 없습니다."
End Sub

Private Sub NxSheetBatchRenameValidateSnapshot(ByVal book As Workbook, ByVal preview As Collection)
    Dim row As Variant, sheet As Worksheet, index As Long, structure As String, seen As Object, existing As Object
    If book Is Nothing Then NxRaiseContractError "열린 통합문서가 없습니다."
    If preview Is Nothing Then NxRaiseContractError "시트 이름 변경 미리보기가 필요합니다."
    If preview.Count = 0 Or preview.Count <> book.Worksheets.Count Then NxRaiseContractError "미리보기 이후 시트 구성이 변경되었습니다."
    structure = NxSheetBatchRenameStructure(book)
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For Each row In preview
        If Not IsArray(row) Then NxRaiseContractError "시트 이름 변경 미리보기를 다시 만드세요."
        If LBound(row) <> 0 Or UBound(row) <> 5 Then NxRaiseContractError "시트 이름 변경 미리보기를 다시 만드세요."
        Set sheet = row(3)
        If sheet Is Nothing Then NxRaiseContractError "미리보기 이후 시트 정체성이 변경되었습니다."
        If Not sheet.Parent Is book Then NxRaiseContractError "미리보기 이후 통합문서가 변경되었습니다."
        index = index + 1
        Set existing = NxBatchRenameSheet(book, CStr(row(0)))
        If Not sheet Is existing Then NxRaiseContractError "미리보기 이후 시트 정체성이 변경되었습니다."
        If seen.Exists(CStr(row(0))) Then NxRaiseContractError "중복 시트가 미리보기에 포함되었습니다."
        seen.Add CStr(row(0)), True
        If StrComp(CStr(row(4)), book.FullName, vbBinaryCompare) <> 0 Then NxRaiseContractError "미리보기 이후 통합문서 경로가 변경되었습니다."
        If StrComp(CStr(row(5)), structure, vbBinaryCompare) <> 0 Then NxRaiseContractError "미리보기 이후 시트 구성이 변경되었습니다."
        If StrComp(CStr(row(0)), sheet.Name, vbBinaryCompare) <> 0 Then NxRaiseContractError "미리보기 이후 시트 이름이 변경되었습니다."
    Next row
End Sub

Private Function NxSheetBatchRenameStructure(ByVal book As Workbook) As String
    Dim sheet As Object, result As String
    ' Chart sheets share the same name namespace and participate in drift checks.
    For Each sheet In book.Sheets
        result = result & TypeName(sheet) & ":" & CStr(Len(sheet.Name)) & ":" & sheet.Name & ":" & CStr(sheet.Visible) & ";"
    Next sheet
    NxSheetBatchRenameStructure = result
End Function

Public Function NxSheetBatchRenameCommit(ByVal book As Workbook, ByVal preview As Collection) As Long
    Dim journal As New Collection, row As Variant, mapping As Collection, sheet As Worksheet
    Dim temporaryName As String, index As Long, errorNumber As Long, errorDescription As String, recoveryDescription As String
    Dim hasDuplicate As Boolean, hasCollision As Boolean
    NxSheetBatchRenameRequireReady book, preview
    Set preview = NxSheetBatchRenameRevalidatePreview(book, preview, hasDuplicate, hasCollision)
    For Each row In preview
        If CStr(row(2)) = "준비" Then
            Set sheet = row(3)
            temporaryName = NxSheetBatchRenameTemporaryName(book)
            Set mapping = New Collection
            ' Retain the worksheet object itself. Recovery never reselects a
            ' sheet by a guessed temporary/target name.
            mapping.Add sheet
            mapping.Add CStr(row(0))
            mapping.Add temporaryName
            mapping.Add CStr(row(1))
            mapping.Add 0
            journal.Add mapping
        End If
    Next row
    If journal.Count = 0 Then NxRaiseContractError "변경할 시트 이름이 없습니다."
    On Error GoTo Failed
    For index = 1 To journal.Count
        Set mapping = journal(index)
        Set sheet = mapping.Item(1)
        If StrComp(sheet.Name, CStr(mapping.Item(2)), vbBinaryCompare) <> 0 Then _
            NxRaiseContractError "미리보기 이후 시트 정체성이 변경되었습니다: " & CStr(mapping.Item(2))
        NxBatchRenameMaybeInjectMoveFailure "sheet-original-to-temporary", index
        sheet.Name = CStr(mapping.Item(3))
        NxSheetBatchRenameSetStage mapping, 1
    Next index
    For index = 1 To journal.Count
        Set mapping = journal(index)
        Set sheet = mapping.Item(1)
        If StrComp(sheet.Name, CStr(mapping.Item(3)), vbBinaryCompare) <> 0 Then _
            NxRaiseContractError "임시 시트 정체성이 변경되었습니다: " & CStr(mapping.Item(3))
        NxBatchRenameMaybeInjectMoveFailure "sheet-temporary-to-target", index
        sheet.Name = CStr(mapping.Item(4))
        NxSheetBatchRenameSetStage mapping, 2
    Next index
    NxSheetBatchRenameCommit = journal.Count
    Set mLastSheetJournal = journal
    Set mLastSheetBook = book
    Set mLastFileJournal = Nothing
    Exit Function
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    On Error Resume Next
    NxSheetBatchRenameRollback book, journal
    If Err.Number <> 0 Then recoveryDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    If Len(recoveryDescription) > 0 Then
        Set mLastSheetJournal = journal
        Set mLastSheetBook = book
        Set mLastFileJournal = Nothing
        errorDescription = errorDescription & " | " & recoveryDescription
    End If
    Err.Raise errorNumber, "LHexcel.Sheet.Rename", errorDescription
End Function

Public Function NxBatchRenameCanUndo(ByVal fileMode As Boolean, Optional ByVal book As Workbook) As Boolean
    If fileMode Then
        NxBatchRenameCanUndo = Not mLastFileJournal Is Nothing
    ElseIf Not mLastSheetJournal Is Nothing Then
        NxBatchRenameCanUndo = (book Is mLastSheetBook)
    End If
End Function

Public Function NxBatchRenameUndoJournal(ByVal fileMode As Boolean, Optional ByVal book As Workbook) As Object
    If Not NxBatchRenameCanUndo(fileMode, book) Then NxRaiseContractError "이 Excel 세션에서 되돌릴 이름 변경이 없습니다."
    If fileMode Then
        Set NxBatchRenameUndoJournal = mLastFileJournal
    Else
        Set NxBatchRenameUndoJournal = mLastSheetJournal
    End If
End Function

Public Function NxBatchRenameUndoTarget(ByVal fileMode As Boolean, Optional ByVal book As Workbook) As String
    Dim ignoredCount As Long, path As String
    ignoredCount = NxBatchRenameUndoLast(fileMode, book, True)
    If fileMode Then
        path = mLastFileJournal.OriginalPath(1)
        NxBatchRenameUndoTarget = Left$(path, InStrRev(path, Application.PathSeparator) - 1)
    Else
        NxBatchRenameUndoTarget = book.FullName
    End If
End Function

Public Function NxBatchRenameUndoLast(ByVal fileMode As Boolean, Optional ByVal book As Workbook, Optional ByVal validateOnly As Boolean = False) As Long
    Dim mapping As Collection, sheet As Worksheet, moving As Object, existing As Object, candidate As Variant
    If Not NxBatchRenameCanUndo(fileMode, book) Then NxRaiseContractError "이 Excel 세션에서 되돌릴 직전 이름 변경이 없습니다."
    If fileMode Then
        mLastFileJournal.RequireRollbackSafe
        NxBatchRenameUndoLast = mLastFileJournal.Count
        If validateOnly Then Exit Function
        mLastFileJournal.Rollback
        Set mLastFileJournal = Nothing
    Else
        If book.ProtectStructure Then NxRaiseContractError "통합문서 구조 보호를 확인하세요."
        Set moving = CreateObject("Scripting.Dictionary"): moving.CompareMode = vbTextCompare
        For Each mapping In mLastSheetJournal
            Set sheet = mapping.Item(1)
            If Not sheet.Parent Is book Then NxRaiseContractError "되돌릴 시트 정체성이 변경되었습니다."
            Select Case CLng(mapping.Item(5))
                Case 0: candidate = mapping.Item(2)
                Case 1: candidate = mapping.Item(3)
                Case 2: candidate = mapping.Item(4)
            End Select
            If StrComp(sheet.Name, CStr(candidate), vbBinaryCompare) <> 0 Then NxRaiseContractError "되돌릴 시트 이름이 변경되었습니다."
            Set moving.Item(sheet.Name) = sheet
        Next mapping
        For Each mapping In mLastSheetJournal
            For Each candidate In Array(CStr(mapping.Item(2)), CStr(mapping.Item(3)))
                Set existing = NxBatchRenameSheet(book, CStr(candidate))
                If Not existing Is Nothing Then
                    If Not moving.Exists(existing.Name) Then NxRaiseContractError "되돌리기 이름이 다른 시트에 사용 중입니다: " & CStr(candidate)
                End If
            Next candidate
        Next mapping
        NxBatchRenameUndoLast = mLastSheetJournal.Count
        If validateOnly Then Exit Function
        NxSheetBatchRenameRollback book, mLastSheetJournal
        Set mLastSheetJournal = Nothing
        Set mLastSheetBook = Nothing
    End If
End Function

Private Sub NxBatchRenameMaybeInjectMoveFailure(ByVal stageName As String, ByVal itemIndex As Long)
#If NX_R58_TEST_BUILD Then
    If mTestFailMoveIndex <> itemIndex Then Exit Sub
    If StrComp(mTestFailMoveStage, stageName, vbTextCompare) <> 0 Then Exit Sub
    NxBatchRenameTestClearMoveFailure
    Err.Raise vbObjectError + 5850, "LHexcel.File.Rename.Test", "Injected rename move failure at " & stageName & " item " & CStr(itemIndex)
#End If
End Sub

Private Function NxSheetBatchRenameTemporaryName(ByVal book As Workbook) As String
    Dim attempt As Long, candidate As String, existing As Object
    For attempt = 1 To 20
        candidate = "NX_" & Left$(Replace$(NxCreateRunUuid(), "-", vbNullString), 20)
        Set existing = NxBatchRenameSheet(book, candidate)
        If existing Is Nothing Then
            NxSheetBatchRenameTemporaryName = candidate
            Exit Function
        End If
    Next attempt
    NxRaiseContractError "임시 시트 이름 충돌이 반복되었습니다."
End Function

Private Sub NxSheetBatchRenameSetStage(ByVal mapping As Collection, ByVal stageValue As Long)
    ' Preserve the existing fifth item until the replacement has been inserted.
    mapping.Add stageValue, , 5
    mapping.Remove 6
End Sub

Private Sub NxSheetBatchRenameRollback(ByVal book As Workbook, ByVal journal As Collection)
    Dim index As Long, mapping As Collection, sheet As Worksheet, failures As String
    If journal Is Nothing Then Exit Sub

    ' First vacate every target through the retained object references.
    For index = journal.Count To 1 Step -1
        Set mapping = journal.Item(index)
        If CLng(mapping.Item(5)) = 2 Then
            On Error GoTo TargetRollbackFailed
            Set sheet = mapping.Item(1)
            If StrComp(sheet.Name, CStr(mapping.Item(4)), vbBinaryCompare) <> 0 Then _
                NxRaiseContractError "시트 target 정체성이 변경되었습니다: " & CStr(mapping.Item(4))
            If Not NxSheetBatchRenameNameVacant(book, CStr(mapping.Item(3)), sheet) Then _
                NxRaiseContractError "시트 임시 복구 이름이 비어 있지 않습니다: " & CStr(mapping.Item(3))
            sheet.Name = CStr(mapping.Item(3))
            NxSheetBatchRenameSetStage mapping, 1
        End If
ContinueTargetRollback:
        On Error GoTo 0
    Next index

    For index = journal.Count To 1 Step -1
        Set mapping = journal.Item(index)
        If CLng(mapping.Item(5)) = 1 Then
            On Error GoTo OriginalRollbackFailed
            Set sheet = mapping.Item(1)
            If StrComp(sheet.Name, CStr(mapping.Item(3)), vbBinaryCompare) <> 0 Then _
                NxRaiseContractError "시트 temporary 정체성이 변경되었습니다: " & CStr(mapping.Item(3))
            If Not NxSheetBatchRenameNameVacant(book, CStr(mapping.Item(2)), sheet) Then _
                NxRaiseContractError "시트 원래 복구 이름이 비어 있지 않습니다: " & CStr(mapping.Item(2))
            sheet.Name = CStr(mapping.Item(2))
            NxSheetBatchRenameSetStage mapping, 0
        End If
ContinueOriginalRollback:
        On Error GoTo 0
    Next index

    If Len(failures) > 0 Then NxRaiseContractError "시트 이름 변경 복구 실패: " & failures
    Exit Sub

TargetRollbackFailed:
    failures = NxSheetBatchRenameJoinFailure(failures, "target-to-temporary|residual=" & CStr(mapping.Item(4)) & "|" & CStr(Err.Number) & "|" & Err.Description)
    Err.Clear
    Resume ContinueTargetRollback
OriginalRollbackFailed:
    failures = NxSheetBatchRenameJoinFailure(failures, "temporary-to-original|residual=" & CStr(mapping.Item(3)) & "|" & CStr(Err.Number) & "|" & Err.Description)
    Err.Clear
    Resume ContinueOriginalRollback
End Sub

Private Function NxSheetBatchRenameNameVacant(ByVal book As Workbook, ByVal candidate As String, ByVal expectedSheet As Worksheet) As Boolean
    Dim existing As Object
    Set existing = NxBatchRenameSheet(book, candidate)
    If existing Is Nothing Then
        NxSheetBatchRenameNameVacant = True
    Else
        NxSheetBatchRenameNameVacant = (existing Is expectedSheet)
    End If
End Function

Private Function NxSheetBatchRenameJoinFailure(ByVal aggregate As String, ByVal detail As String) As String
    If Len(aggregate) > 0 Then aggregate = aggregate & "; "
    NxSheetBatchRenameJoinFailure = aggregate & detail
End Function

Private Function NxBatchRenameTargetPath(ByVal originalPath As String, ByVal findText As String, _
    ByVal replaceText As String, ByVal prefixText As String, ByVal suffixText As String) As String
    Dim separator As Long, folderPath As String, fileName As String, extensionAt As Long
    Dim stem As String, extensionText As String, targetName As String
    separator = InStrRev(originalPath, Application.PathSeparator)
    If separator = 0 Then NxRaiseContractError "파일 경로를 확인하세요."
    folderPath = Left$(originalPath, separator)
    fileName = Mid$(originalPath, separator + 1)
    extensionAt = InStrRev(fileName, ".")
    If extensionAt > 1 Then
        stem = Left$(fileName, extensionAt - 1)
        extensionText = Mid$(fileName, extensionAt)
    Else
        stem = fileName
    End If
    targetName = prefixText & Replace$(stem, findText, replaceText, 1, -1, vbTextCompare) & suffixText & extensionText
    If Not NxFileRenameNameIsValid(targetName) Then NxRaiseContractError "변경할 파일 이름에 사용할 수 없는 문자가 있습니다."
    NxBatchRenameTargetPath = folderPath & targetName
End Function

Private Function NxFileRenameNameIsValid(ByVal value As String) As Boolean
    Dim forbidden As Variant, token As Variant, stem As String, index As Long, character As Long
    If Len(Trim$(value)) = 0 Or Len(value) > 240 Or Right$(value, 1) = "." Or Right$(value, 1) = " " Then Exit Function
    forbidden = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    For Each token In forbidden
        If InStr(1, value, CStr(token), vbBinaryCompare) > 0 Then Exit Function
    Next token
    For index = 1 To Len(value)
        character = AscW(Mid$(value, index, 1))
        If character >= 0 And character < 32 Then Exit Function
    Next index
    stem = UCase$(Split(value, ".")(0))
    Select Case stem
        Case "CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$": Exit Function
    End Select
    If Len(stem) = 4 Then
        If Left$(stem, 3) = "COM" Or Left$(stem, 3) = "LPT" Then
            If InStr(1, "123456789" & ChrW$(185) & ChrW$(178) & ChrW$(179), Right$(stem, 1), vbBinaryCompare) > 0 Then Exit Function
        End If
    End If
    NxFileRenameNameIsValid = True
End Function

Private Function NxSheetRenameNameIsValid(ByVal value As String) As Boolean
    Dim forbidden As Variant, token As Variant, index As Long
    If Len(Trim$(value)) = 0 Or Len(value) > 31 Then Exit Function
    If Left$(value, 1) = "'" Or Right$(value, 1) = "'" Then Exit Function
    If StrComp(value, "History", vbTextCompare) = 0 Then Exit Function
    For index = 1 To Len(value)
        If AscW(Mid$(value, index, 1)) >= 0 And AscW(Mid$(value, index, 1)) < 32 Then Exit Function
    Next index
    forbidden = Array(":", "\", "/", "?", "*", "[", "]")
    For Each token In forbidden
        If InStr(1, value, CStr(token), vbBinaryCompare) > 0 Then Exit Function
    Next token
    NxSheetRenameNameIsValid = True
End Function

Private Function NxBatchRenameFileExists(ByVal candidate As String) As Boolean
    On Error Resume Next
    NxBatchRenameFileExists = (Len(Dir$(candidate, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function NxBatchRenameSheet(ByVal book As Workbook, ByVal sheetName As String) As Object
    On Error Resume Next
    Set NxBatchRenameSheet = book.Sheets(sheetName)
    On Error GoTo 0
End Function
