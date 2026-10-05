Attribute VB_Name = "NxTemplateManager"
Option Explicit

Public Sub NxTemplateValidateUserBook(ByVal book As Workbook)
    Dim candidate As Workbook, found As Boolean
    If book Is Nothing Then NxRaiseContractError "일반 Excel 통합문서를 열어 주세요."
    For Each candidate In Application.Workbooks
        If candidate Is book Then found = True: Exit For
    Next candidate
    If Not found Then NxRaiseContractError "대상 통합문서가 닫혔습니다. 템플릿 관리를 다시 여세요."
    If book.IsAddin Or book Is ThisWorkbook Then NxRaiseContractError "추가기능이 아닌 일반 통합문서를 선택하세요."
End Sub

Public Function NxTemplateManagerRequest(ByVal sourceBook As Workbook, ByVal sourceSheet As Worksheet, _
    ByVal sourceKind As String, ByVal displayName As String, ByVal description As String, _
    ByVal retainAll As Boolean, ByVal consent As Boolean, Optional ByVal category As String = vbNullString) As CNxTemplateRequest

    Dim request As New CNxTemplateRequest, retained As New Collection, source As Range, analysis As CNxTemplateAnalysis
    NxTemplateValidateUserBook sourceBook
    If Not consent Then NxRaiseContractError "저장 내용과 제외 리소스를 확인하고 동의해 주세요."
    If sourceKind = "sheet_used_range" Then
        If sourceSheet Is Nothing Then NxRaiseContractError "등록할 워크시트를 선택하세요."
        If Not sourceSheet.Parent Is sourceBook Then NxRaiseContractError "등록 시트의 원본 통합문서가 달라졌습니다."
        Set source = sourceSheet.UsedRange
        If retainAll Then retained.Add source
        Set analysis = NxTemplateAnalyze(source, retained, True, True)
        request.ConfigureRegistration "NX-TPL-REGISTER-SHEET", "sheet_used_range", source, retained, _
            displayName, description, True, True, True
        request.BindAnalysis analysis
    Else
        request.ConfigureWorkbookRegistration sourceBook, sourceKind, displayName, description, retainAll, True, True
    End If
    request.Category = category
    Set NxTemplateManagerRequest = request
End Function

Public Function NxTemplatePickRegistrationFile() As String
    Dim picker As FileDialog
    Set picker = Application.FileDialog(msoFileDialogFilePicker)
    picker.Title = "등록할 템플릿 파일 선택"
    picker.AllowMultiSelect = False
    picker.Filters.Clear
    picker.Filters.Add "매크로 없는 Excel 파일", "*.xlsx;*.xltx"
    If picker.Show <> -1 Then Exit Function
    If picker.SelectedItems.Count <> 1 Then Exit Function
    NxTemplatePickRegistrationFile = NxTemplateRegistrationFilePath(CStr(picker.SelectedItems(1)))
End Function

Public Function NxTemplateRegistrationFilePath(ByVal filePath As String) As String
    Dim fileSystem As Object, normalized As String, candidate As Workbook
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Len(Trim$(filePath)) = 0 Or InStr(filePath, vbNullChar) > 0 Then NxRaiseContractError "등록 파일 경로가 올바르지 않습니다."
    normalized = fileSystem.GetAbsolutePathName(filePath)
    If Not fileSystem.FileExists(normalized) Then NxRaiseContractError "등록 파일을 찾지 못했습니다."
    Select Case LCase$(fileSystem.GetExtensionName(normalized))
        Case "xlsx", "xltx"
        Case Else: NxRaiseContractError ".xlsx 또는 .xltx 파일만 등록할 수 있습니다."
    End Select
    If Left$(fileSystem.GetFileName(normalized), 2) = "~$" Then NxRaiseContractError "Excel 임시 잠금 파일은 등록할 수 없습니다."
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, normalized, vbTextCompare) = 0 Then _
            NxRaiseContractError "이미 열린 파일은 현재파일 등록을 사용하세요. 열린 원본을 닫거나 다시 열지 않습니다."
        If StrComp(candidate.Name, fileSystem.GetFileName(normalized), vbTextCompare) = 0 Then _
            NxRaiseContractError "같은 이름의 통합문서가 이미 열려 있습니다. 현재파일 등록을 사용하거나 등록 파일 이름을 확인하세요."
    Next candidate
    NxTemplateRegistrationFilePath = normalized
End Function

Public Function NxTemplateRegisterFileFromManager(ByVal filePath As String, ByVal targetBook As Workbook, _
    ByVal displayName As String, ByVal description As String, ByVal retainAll As Boolean, ByVal consent As Boolean, _
    Optional ByVal expectedFileSha As String = vbNullString, Optional ByVal category As String = vbNullString) As CNxResult
    Set NxTemplateRegisterFileFromManager = RegisterTemplateFile(filePath, targetBook, displayName, description, retainAll, consent, vbNullString, expectedFileSha, category)
End Function

Public Function NxTemplateRegisterFileForTest(ByVal filePath As String, ByVal targetBook As Workbook, _
    ByVal displayName As String, ByVal retainAll As Boolean, ByVal consent As Boolean, ByVal root As String, _
    Optional ByVal expectedFileSha As String = vbNullString) As CNxResult
    If Len(root) = 0 Then NxRaiseContractError "Template file fixture root is required"
    Set NxTemplateRegisterFileForTest = RegisterTemplateFile(filePath, targetBook, displayName, "File fixture", retainAll, consent, root, expectedFileSha)
End Function

Private Function RegisterTemplateFile(ByVal filePath As String, ByVal targetBook As Workbook, _
    ByVal displayName As String, ByVal description As String, ByVal retainAll As Boolean, ByVal consent As Boolean, _
    ByVal root As String, ByVal expectedFileSha As String, Optional ByVal category As String = vbNullString) As CNxResult

    Dim sourceBook As Workbook, sourceSheet As Worksheet, request As CNxTemplateRequest, priorSecurity As MsoAutomationSecurity
    Dim priorEvents As Boolean, failureText As String, result As CNxResult, openedByUs As Boolean
    priorSecurity = Application.AutomationSecurity
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    NxTemplateValidateUserBook targetBook
    If Not consent Then NxRaiseContractError "저장 내용과 제외 리소스를 확인하고 동의해 주세요."
    filePath = NxTemplateRegistrationFilePath(filePath)
    If Len(expectedFileSha) > 0 Then NxTemplateAssertFilePreview filePath, sourceBook, expectedFileSha
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    Set sourceBook = NxSafeWorkbookOpen(filePath, openedByUs, strictTemplate:=True)
    If Len(expectedFileSha) > 0 Then NxTemplateAssertFilePreview filePath, sourceBook, expectedFileSha
    Set request = NxTemplateManagerRequest(sourceBook, sourceSheet, "file", displayName, description, retainAll, consent, category)
    targetBook.Activate
    If Len(root) = 0 Then
        Set result = NxTemplateRunRequestFromManager(request, targetBook)
    Else
        Set result = NxTemplateRegisterForTest(request, root)
    End If
    NxSafeWorkbookClose sourceBook, openedByUs
    Set sourceBook = Nothing
    Application.AutomationSecurity = priorSecurity
    Application.EnableEvents = priorEvents
    Set RegisterTemplateFile = result
    Exit Function
Failed:
    failureText = Err.Description
    On Error Resume Next
    NxSafeWorkbookClose sourceBook, openedByUs
    Application.AutomationSecurity = priorSecurity
    Application.EnableEvents = priorEvents
    targetBook.Activate
    On Error GoTo 0
    Set RegisterTemplateFile = NxCreateResult("NX-TPL-REGISTER-SHEET", NxEnvironmentError, "file-register", filePath, False, failureText, "result_environment_error")
End Function

Public Function NxTemplateRegistrationPreview(ByVal request As CNxTemplateRequest, ByRef previewDigest As String) As String
    Dim source As Range, retained As Collection, allConstants As Collection, analysis As CNxTemplateAnalysis
    Dim detail As Variant, sheetText As String, digestText As String, book As Workbook, sheet As Worksheet
    previewDigest = vbNullString
    If request Is Nothing Then NxRaiseContractError "등록 미리보기 원본이 없습니다."
    Set book = request.SourceRange.Worksheet.Parent
    NxTemplateValidateUserBook book
    digestText = request.SourceKind & "|" & book.FullName & vbLf
    If request.SourceKind <> "file" Then
        digestText = digestText & NxContextFactory.UnsavedWorkbookSessionToken(book) & vbLf
    End If
    digestText = digestText & request.DisplayName & vbLf & request.Description & vbLf
    NxTemplateRegistrationPreview = CStr(request.SourceRanges.Count) & "개 시트 / " & CStr(request.ActualCellCount) & "개 사용 셀" & vbCrLf
    NxTemplateRegistrationPreview = NxTemplateRegistrationPreview & _
        "보존: 값·수식(선택 내용), 기본 셀 서식, 병합, 행 높이·열 너비" & vbCrLf & _
        "보존: 용지·방향·여백·배율·인쇄 영역·반복 제목행/열 (등록 범위 기준)" & vbCrLf & _
        "제외: 그림·도형, 하이퍼링크, 조건부서식, 머리글·바닥글" & vbCrLf
    If request.SourceKind <> "sheet_used_range" Then
        For Each sheet In book.Worksheets
            digestText = digestText & "sheet|" & sheet.CodeName & "|" & sheet.Name & "|" & CStr(sheet.Index) & "|" & CStr(sheet.Visible) & vbLf
            If sheet.Visible <> xlSheetVisible Then _
                NxTemplateRegistrationPreview = NxTemplateRegistrationPreview & "제외 시트: " & sheet.Name & " (숨김)" & vbCrLf
        Next sheet
    End If
    For Each source In request.SourceRanges
        Set retained = request.RetainedForSource(source)
        Set analysis = NxTemplateAnalyze(source, retained, request.IncludeHidden, True)
        sheetText = "[" & source.Worksheet.Name & "]" & vbCrLf & _
            NxTemplateAnalysisPreviewText(analysis)
        If analysis.RiskDetails.Count = 0 Then sheetText = sheetText & vbCrLf & "위험 수식 상세: 없음"
        For Each detail In analysis.RiskDetails
            sheetText = sheetText & vbCrLf & "위험 수식 상세: " & Replace(CStr(detail), "|", " / ")
        Next detail
        If Not analysis.HasExcludedDisplayResource Then sheetText = sheetText & vbCrLf & "제외 리소스: 없음"
        sheetText = sheetText & vbCrLf & "도형 " & CStr(source.Worksheet.Shapes.Count) & _
            "개 / 하이퍼링크 " & CStr(source.Hyperlinks.Count) & "개 / 조건부서식 " & CStr(source.FormatConditions.Count) & "개 (제외)"
        NxTemplateRegistrationPreview = NxTemplateRegistrationPreview & vbCrLf & sheetText & vbCrLf
        Set allConstants = New Collection
        allConstants.Add source
        digestText = digestText & source.Worksheet.CodeName & "|" & source.Worksheet.Name & "|" & _
            CStr(source.Worksheet.Index) & "|" & CStr(source.Worksheet.Visible) & "|" & source.Address & "|" & _
            NxTemplateLogicalHash(source, allConstants) & "|" & TemplatePreviewExtraFormatHash(source) & "|" & _
            NxTemplatePrintFingerprint(source.Worksheet) & vbLf & sheetText & vbLf
    Next source
    previewDigest = NxTemplateTextSha256(digestText)
End Function

Private Function TemplatePreviewExtraFormatHash(ByVal source As Range) As String
    ' Bind the remaining fields written by NxTemplatePackage, beyond its existing logical hash.
    Dim cell As Range, edge As Variant, value As String
    For Each cell In source.Cells
        value = value & cell.Address & "|" & CStr(cell.Font.Name) & "|" & CStr(cell.Font.Size) & "|" & _
            CStr(cell.Font.Underline) & "|" & CStr(cell.Interior.Pattern) & "|" & CStr(cell.WrapText) & "|" & _
            CStr(cell.Orientation) & "|" & CStr(cell.IndentLevel) & vbLf
        For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
            value = value & CStr(cell.Borders(CLng(edge)).LineStyle) & "|" & CStr(cell.Borders(CLng(edge)).Weight) & "|" & _
                CStr(cell.Borders(CLng(edge)).Color) & vbLf
        Next edge
    Next cell
    TemplatePreviewExtraFormatHash = NxTemplateTextSha256(value)
End Function

Public Function NxTemplateFileRegistrationPreview(ByVal filePath As String, ByVal targetBook As Workbook, _
    ByVal displayName As String, ByVal description As String, ByVal retainAll As Boolean, ByRef fileSha As String) As String
    Dim sourceBook As Workbook, sourceSheet As Worksheet, request As CNxTemplateRequest, openedByUs As Boolean
    Dim expectedSha As String, unusedDigest As String, previewText As String, failureNumber As Long, failureText As String
    Dim priorSecurity As MsoAutomationSecurity, priorEvents As Boolean
    priorSecurity = Application.AutomationSecurity
    priorEvents = Application.EnableEvents
    fileSha = vbNullString
    On Error GoTo Failed
    NxTemplateValidateUserBook targetBook
    filePath = NxTemplateRegistrationFilePath(filePath)
    expectedSha = NxTemplateFileSha256(filePath)
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    Set sourceBook = NxSafeWorkbookOpen(filePath, openedByUs, strictTemplate:=True)
    NxTemplateAssertFilePreview filePath, sourceBook, expectedSha
    Set request = NxTemplateManagerRequest(sourceBook, sourceSheet, "file", displayName, description, retainAll, True)
    previewText = NxTemplateRegistrationPreview(request, unusedDigest)
    NxTemplateAssertFilePreview filePath, sourceBook, expectedSha
    NxSafeWorkbookClose sourceBook, openedByUs
    Set sourceBook = Nothing
    targetBook.Activate
    Application.AutomationSecurity = priorSecurity
    Application.EnableEvents = priorEvents
    fileSha = expectedSha
    NxTemplateFileRegistrationPreview = "파일: " & filePath & vbCrLf & previewText
    Exit Function
Failed:
    failureNumber = Err.Number
    failureText = Err.Description
    On Error Resume Next
    NxSafeWorkbookClose sourceBook, openedByUs
    targetBook.Activate
    Application.AutomationSecurity = priorSecurity
    Application.EnableEvents = priorEvents
    On Error GoTo 0
    fileSha = vbNullString
    Err.Raise failureNumber, "NxTemplateFileRegistrationPreview", failureText
End Function

Public Sub NxTemplateAssertFilePreview(ByVal filePath As String, ByVal snapshotBook As Workbook, ByVal expectedSha As String)
    If Len(expectedSha) <> 64 Then NxRaiseContractError "파일 미리보기 검증값이 없습니다. 미리보기를 다시 확인하세요."
    If NxTemplateFileSha256(filePath) <> expectedSha Then NxRaiseContractError "미리보기 이후 원본 파일이 변경되었습니다. 미리보기와 동의를 다시 확인하세요."
    If Not snapshotBook Is Nothing Then
        If NxTemplateFileSha256(snapshotBook.FullName) <> expectedSha Then _
            NxRaiseContractError "열린 파일 스냅샷이 미리보기 원본과 다릅니다. 미리보기를 다시 확인하세요."
    End If
End Sub

Public Function NxTemplateManagerKindLabel(ByVal sourceKind As String) As String
    Select Case sourceKind
        Case "sheet_used_range": NxTemplateManagerKindLabel = "현재시트"
        Case "workbook": NxTemplateManagerKindLabel = "현재파일"
        Case "file": NxTemplateManagerKindLabel = "특정파일"
        Case Else: NxTemplateManagerKindLabel = "기존 템플릿"
    End Select
End Function

Public Sub NxTemplateManagerOpenFolder()
    Dim folderPath As String, fileSystem As Object
    NxTemplateEnsureStore
    folderPath = NxTemplateStoreRoot()
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folderPath) Then NxRaiseContractError "템플릿 저장 폴더를 찾지 못했습니다."
    CreateObject("Shell.Application").Explore folderPath
End Sub

Public Function NxTemplateManagerActionConsent(ByVal templateId As String) As Boolean
    Dim records As Collection, record As CNxTemplateRecord
    Set records = New Collection
    records.Add NxTemplateReadRecordFolder(NxTemplateTemplateFolderPath(templateId), False, False)
    For Each record In records
        If record.TemplateId = templateId Then
            If record.Status <> "healthy" Then NxRaiseContractError "손상되었거나 사용 중인 템플릿은 열 수 없습니다."
            If Len(record.RiskFlags) > 0 Or CDbl(record.RowCount) * CDbl(record.ColumnCount) > 10000# Then
                NxTemplateManagerActionConsent = (MsgBox("템플릿: " & record.DisplayName & vbCrLf & _
                    "사용 셀: " & CStr(CDbl(record.RowCount) * CDbl(record.ColumnCount)) & vbCrLf & _
                    "위험 수식: " & record.RiskFlags & vbCrLf & "내용을 확인하고 계속할까요?", _
                    vbYesNo + vbQuestion + vbDefaultButton2, "내엑셀 - 템플릿 사용 확인") = vbYes)
            Else
                NxTemplateManagerActionConsent = True
            End If
            Exit Function
        End If
    Next record
    NxRaiseContractError "선택한 템플릿을 찾지 못했습니다."
End Function
