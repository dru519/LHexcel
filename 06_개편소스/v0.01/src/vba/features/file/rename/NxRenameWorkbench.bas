Attribute VB_Name = "NxRenameWorkbench"
Option Explicit

Public Function NxRenameApplyRules(ByVal originalName As String, ByVal rules As Collection, _
    ByVal itemIndex As Long, ByVal fileMode As Boolean) As String
    Dim stem As String, extension As String, dotAt As Long, rule As CNxRenameRule
    Dim numberValue As Double, numberText As String, compareMode As VbCompareMethod, replaceCount As Long
    If itemIndex < 1 Then NxRaiseContractError "항목 순서가 올바르지 않습니다."
    stem = originalName
    If fileMode Then
        dotAt = InStrRev(stem, ".")
        If dotAt > 1 Then
            extension = Mid$(stem, dotAt)
            stem = Left$(stem, dotAt - 1)
        End If
    End If
    For Each rule In rules
        rule.Validate fileMode
        Select Case rule.Kind
            Case "REPLACE"
                compareMode = IIf(rule.CaseSensitive, vbBinaryCompare, vbTextCompare)
                replaceCount = IIf(rule.FirstOnly, 1, -1)
                stem = Replace$(stem, rule.Text, rule.Replacement, 1, replaceCount, compareMode)
            Case "PREFIX": stem = rule.Text & stem
            Case "SUFFIX": stem = stem & rule.Text
            Case "NUMBER"
                numberValue = CDbl(rule.Start) + CDbl(itemIndex - 1) * CDbl(rule.StepValue)
                If numberValue < -2147483648# Or numberValue > 2147483647# Then NxRaiseContractError "번호가 허용 범위를 넘습니다."
                numberText = CStr(CLng(numberValue))
                If rule.Digits > 0 Then numberText = Format$(CLng(numberValue), String$(rule.Digits, "0"))
                If rule.Position = "FRONT" Then
                    stem = numberText & rule.Text & stem
                Else
                    stem = stem & rule.Text & numberText
                End If
            Case "DELETE"
                If rule.Start >= Len(stem) Then
                    stem = vbNullString
                ElseIf rule.Position = "FRONT" Then
                    stem = Mid$(stem, rule.Start + 1)
                Else
                    stem = Left$(stem, Len(stem) - rule.Start)
                End If
            Case "FULL": stem = rule.Text
            Case "CASE"
                Select Case rule.Position
                    Case "LOWER": stem = LCase$(stem)
                    Case "UPPER": stem = UCase$(stem)
                    Case "PROPER": stem = StrConv(LCase$(stem), vbProperCase)
                End Select
            Case "EXTENSION"
                extension = Trim$(rule.Text)
                Do While Left$(extension, 1) = ".": extension = Mid$(extension, 2): Loop
                If Len(extension) > 0 Then extension = "." & extension
        End Select
    Next rule
    NxRenameApplyRules = stem & extension
End Function

Public Function NxRenameLeafName(ByVal fullPath As String) As String
    NxRenameLeafName = Mid$(fullPath, InStrRev(fullPath, "\") + 1)
End Function

Public Function NxRenameFileRulesPreview(ByVal paths As Collection, ByVal rules As Collection, ByVal overrides As Object, _
    ByRef hasDuplicate As Boolean, ByRef hasCollision As Boolean) As Collection
    Dim preview As New Collection, path As Variant, targetName As String, index As Long
    Dim identityReader As New CNxBatchRenameJournal, fso As Object, identity As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each path In paths
        index = index + 1
        targetName = NxRenameApplyRules(NxRenameLeafName(CStr(path)), rules, index, True)
        If overrides.Exists(CStr(path)) Then targetName = CStr(overrides.Item(CStr(path)))
        identity = vbNullString
        If fso.FileExists(CStr(path)) Then identity = identityReader.FileIdentity(CStr(path))
        preview.Add Array(CStr(path), Left$(CStr(path), InStrRev(CStr(path), "\")) & targetName, vbNullString, identity)
    Next path
    Set NxRenameFileRulesPreview = NxBatchRenameRevalidatePreview(preview, hasDuplicate, hasCollision)
End Function

Public Function NxRenameSheetRulesPreview(ByVal book As Workbook, ByVal order As Collection, ByVal rules As Collection, _
    ByVal overrides As Object, ByRef hasDuplicate As Boolean, ByRef hasCollision As Boolean) As Collection
    Dim original As Collection, preview As New Collection, row As Variant, name As Variant, index As Long
    Set original = NxSheetBatchRenameBuildPreview(book, vbNullString, vbNullString, vbNullString, vbNullString, hasDuplicate, hasCollision)
    For Each name In order
        index = index + 1
        For Each row In original
            If StrComp(CStr(row(0)), CStr(name), vbBinaryCompare) = 0 Then
                row(1) = NxRenameApplyRules(CStr(row(0)), rules, index, False)
                If overrides.Exists(CStr(name)) Then row(1) = CStr(overrides.Item(CStr(name)))
                preview.Add row
                Exit For
            End If
        Next row
    Next name
    Set NxRenameSheetRulesPreview = NxSheetBatchRenameRevalidatePreview(book, preview, hasDuplicate, hasCollision)
End Function

Public Function NxRenameAddFile(ByVal paths As Collection, ByVal candidate As String) As Boolean
    Dim fso As Object, existing As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    candidate = fso.GetAbsolutePathName(candidate)
    If Not fso.FileExists(candidate) Then NxRaiseContractError "파일이 없습니다: " & candidate
    For Each existing In paths
        If StrComp(CStr(existing), candidate, vbTextCompare) = 0 Then Exit Function
    Next existing
    If paths.Count >= 1000 Then NxRaiseContractError "이름 변경은 한 번에 1,000개까지 가능합니다."
    paths.Add candidate
    NxRenameAddFile = True
End Function

Public Function NxRenameAddFolder(ByVal paths As Collection, ByVal folderPath As String, ByVal recursive As Boolean) As Long
    Dim fso As Object, beforeCount As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    beforeCount = paths.Count
    NxRenameCollectFolder paths, fso.GetFolder(folderPath), recursive
    NxRenameAddFolder = paths.Count - beforeCount
End Function

Private Sub NxRenameCollectFolder(ByVal paths As Collection, ByVal folder As Object, ByVal recursive As Boolean)
    Dim item As Object, ignored As Boolean
    ' Do not recurse through junctions/symlinks, including a chosen root junction.
    If (CLng(folder.Attributes) And 1024) <> 0 Then NxRaiseContractError "연결 폴더는 자동 탐색하지 않습니다: " & folder.Path
    For Each item In folder.Files
        ignored = NxRenameAddFile(paths, CStr(item.Path))
    Next item
    If recursive Then
        For Each item In folder.SubFolders
            If (CLng(item.Attributes) And 1024) = 0 Then NxRenameCollectFolder paths, item, True
        Next item
    End If
End Sub

Public Sub NxRenameMoveItem(ByVal items As Collection, ByVal index As Long, ByVal offset As Long)
    Dim value As Variant, target As Long
    target = index + offset
    If index < 1 Or index > items.Count Or target < 1 Or target > items.Count Then Exit Sub
    If IsObject(items(index)) Then Set value = items(index) Else value = items(index)
    items.Remove index
    If target > items.Count Then items.Add value Else items.Add value, , target
End Sub
