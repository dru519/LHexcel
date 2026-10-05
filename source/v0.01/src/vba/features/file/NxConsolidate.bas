Attribute VB_Name = "NxConsolidate"
Option Explicit

Public Function NxConsolidationPlan(ByVal inputPaths As Variant, _
    Optional ByVal allowInvalidInputs As Boolean = False, _
    Optional ByRef excludedSummary As String = "") As Collection
    Dim result As New Collection, index As Long, path As String, issue As String
    Dim seen As Object
    If Not IsArray(inputPaths) Then NxRaiseContractError "Consolidation input files are required"
    ' Files are opened and closed sequentially; count alone does not bound memory.
    excludedSummary = vbNullString
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    For index = LBound(inputPaths) To UBound(inputPaths)
        path = Trim$(CStr(inputPaths(index)))
        issue = NxConsolidationInputIssue(path, seen)
        If Len(issue) > 0 Then
            If Not allowInvalidInputs Then NxRaiseContractError issue
            NxAppendConsolidationIssue excludedSummary, path, issue
        Else
            result.Add path
        End If
    Next index
    Set NxConsolidationPlan = result
End Function

Public Function NxRunConsolidation(ByVal inputPaths As Variant, ByVal resultPath As String, _
    Optional ByVal allowPartial As Boolean = False, Optional ByVal secondApproval As Boolean = False, _
    Optional ByVal mode As String = "SHEETS", Optional ByVal skipHeaders As Boolean = False, _
    Optional ByVal createContents As Boolean = False) As CNxResult
    Dim resultBook As Workbook, paths As Collection, usedNames As Object, path As Variant
    Dim failureList As String, failure As String, temporaryPath As String, partialPath As String
    Dim created As Long, current As Long, rowCreated As Long
    Dim nextRow As Long, expectedColumns As Long, importedCells As Double
    Dim operationStarted As Boolean, cancelledByUser As Boolean
    On Error GoTo Failed
    If mode <> "SHEETS" And mode <> "ROWS" And mode <> "BOTH" Then NxRaiseContractError "파일 통합 방식을 확인하세요."
    NxConsolidationValidateOutput resultPath
    Set paths = NxConsolidationPlan(inputPaths, allowPartial, failureList)
    If Len(Dir$(resultPath)) > 0 Then NxRaiseContractError "Result path already exists; automatic overwrite is forbidden"
    If paths.Count = 0 Then
        Set NxRunConsolidation = NxCreateResult("NX-FILE-CONSOLIDATE", NxEnvironmentError, "preflight", resultPath, False, _
            IIf(Len(failureList) > 0, failureList, "No valid input files were found"), "result_environment_error")
        Exit Function
    End If
    Set resultBook = Workbooks.Add(xlWBATWorksheet)
    nextRow = 1
    If mode <> "SHEETS" Then resultBook.Worksheets(1).Name = "통합데이터"
    Set usedNames = CreateObject("Scripting.Dictionary")
    usedNames.CompareMode = vbTextCompare
    If mode = "BOTH" Then usedNames.Add "통합데이터", True
    If createContents And mode <> "ROWS" Then usedNames.Add "목차", True
    NxFileOperationBegin paths.Count, "내엑셀 파일 통합"
    operationStarted = True
    For Each path In paths
        current = current + 1
        NxFileOperationReport current - 1, "입력 확인: " & CStr(path)
        If NxFileOperationCancelled() Then cancelledByUser = True: Exit For
        If mode = "ROWS" Then
            failure = NxAppendConsolidationRows(resultBook.Worksheets(1), CStr(path), created, nextRow, expectedColumns, importedCells, skipHeaders)
        ElseIf mode = "BOTH" Then
            failure = NxAppendConsolidationBoth(resultBook, CStr(path), usedNames, created, rowCreated, nextRow, expectedColumns, importedCells, skipHeaders)
        Else
            failure = NxAppendConsolidationFile(resultBook, CStr(path), usedNames, created)
        End If
        If Len(failure) > 0 Then
            NxAppendConsolidationIssue failureList, CStr(path), failure
            NxFileOperationReport current, "건너뜀: " & CStr(path)
        Else
            NxFileOperationReport current, "완료: " & CStr(path)
        End If
        If NxFileOperationCancelled() Then cancelledByUser = True: Exit For
    Next path

    If cancelledByUser Then
        NxAppendConsolidationIssue failureList, "사용자 취소", "사용자가 파일 통합을 취소했습니다."
        If created = 0 Or Not allowPartial Or Not secondApproval Then
            resultBook.Close SaveChanges:=False
            NxFinishConsolidationOperation operationStarted
            Set NxRunConsolidation = NxCreateResult("NX-FILE-CONSOLIDATE", NxCancelled, "cancelled", resultPath, False, _
                IIf(allowPartial And Not secondApproval, _
                    "사용자가 파일 통합을 취소했습니다. 현재까지의 결과를 저장하려면 부분결과 2차 승인이 필요합니다.", _
                    "사용자가 파일 통합을 취소했습니다."), "result_cancelled")
            Exit Function
        End If
    End If

    If created = 0 Then
        resultBook.Close SaveChanges:=False
        NxFinishConsolidationOperation operationStarted
        Set NxRunConsolidation = NxCreateResult("NX-FILE-CONSOLIDATE", NxEnvironmentError, "consolidate", resultPath, False, _
            IIf(Len(failureList) > 0, failureList, "No visible sheets were included"), "result_environment_error")
        Exit Function
    End If
    If mode <> "SHEETS" And nextRow > 1 And expectedColumns > 0 Then _
        NxDrawStyleGeneratedTable resultBook, resultBook.Worksheets(1).Cells(1, 1).Resize(nextRow - 1, expectedColumns), IIf(skipHeaders, 1, 0)
    If createContents And mode <> "ROWS" Then NxConsolidationAddContents resultBook, IIf(mode = "BOTH", 2, 1)
    If Len(failureList) > 0 Then
        If Not allowPartial Or Not secondApproval Then
            resultBook.Close SaveChanges:=False
            NxFinishConsolidationOperation operationStarted
            Set NxRunConsolidation = NxCreateResult("NX-FILE-CONSOLIDATE", NxCancelled, "consent", resultPath, False, _
                "실패 파일이 있어 부분결과를 만들려면 second approval(2차 승인)이 필요합니다.", "result_cancelled")
            Exit Function
        End If
        partialPath = NxPartialResultPath(resultPath)
        If Len(Dir$(partialPath)) > 0 Then NxRaiseContractError "Partial result path already exists; automatic overwrite is forbidden"
        NxWritePartialNotice resultBook, failureList
        temporaryPath = NxConsolidationTemporaryPath(partialPath)
        resultBook.SaveAs temporaryPath, xlOpenXMLWorkbook
        resultBook.Close SaveChanges:=False
        Name temporaryPath As partialPath
        If Len(Dir$(partialPath)) = 0 Then NxRaiseContractError "Partial consolidation result was not materialized"
        NxFinishConsolidationOperation operationStarted
        Set NxRunConsolidation = NxCreateResult("NX-FILE-CONSOLIDATE", NxPartialFailure, "partial_result", resultPath, False, _
            failureList & vbCrLf & "partial_output=" & partialPath, "result_partial_failure")
        Exit Function
    End If

    If created = 0 Then NxRaiseContractError "No visible sheets were included"
    temporaryPath = NxConsolidationTemporaryPath(resultPath)
    resultBook.SaveAs temporaryPath, xlOpenXMLWorkbook
    resultBook.Close SaveChanges:=False
    Name temporaryPath As resultPath
    If Len(Dir$(resultPath)) = 0 Then NxRaiseContractError "Consolidation result was not materialized"
    NxFinishConsolidationOperation operationStarted
    Set NxRunConsolidation = NxCreateResult("NX-FILE-CONSOLIDATE", NxSuccess, "complete", resultPath, False, vbNullString, "result_success")
    Exit Function
Failed:
    failure = Err.Description
    NxFinishConsolidationOperation operationStarted
    On Error Resume Next
    If Not resultBook Is Nothing Then resultBook.Close SaveChanges:=False
    If Len(temporaryPath) > 0 Then
        If CreateObject("Scripting.FileSystemObject").FileExists(temporaryPath) Then Kill temporaryPath
    End If
    On Error GoTo 0
    Set NxRunConsolidation = NxCreateResult("NX-FILE-CONSOLIDATE", NxEnvironmentError, "consolidate", resultPath, False, failure, "result_environment_error")
End Function

Public Sub NxConsolidationValidateOutput(ByVal outputPath As String)
    Dim fso As Object, temporaryPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    outputPath = fso.GetAbsolutePathName(outputPath)
    If Len(outputPath) > 218 Or Len(NxPartialResultPath(outputPath)) > 218 Then _
        NxRaiseContractError "파일 통합 저장 경로가 너무 깁니다. 더 짧은 폴더나 파일 이름을 선택하세요."
    If LCase$(fso.GetExtensionName(outputPath)) <> "xlsx" Then NxRaiseContractError "파일 통합 결과는 .xlsx로 저장하세요."
    If Not fso.FolderExists(fso.GetParentFolderName(outputPath)) Then NxRaiseContractError "파일 통합 저장 폴더를 찾을 수 없습니다."
    If fso.FileExists(outputPath) Or fso.FolderExists(outputPath) Then NxRaiseContractError "같은 이름의 파일 또는 폴더가 있습니다. 다른 이름으로 저장하세요."
    temporaryPath = NxConsolidationTemporaryPath(outputPath)
End Sub

Private Function NxConsolidationTemporaryPath(ByVal outputPath As String) As String
    Dim fso As Object, candidate As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    candidate = fso.BuildPath(fso.GetParentFolderName(outputPath), _
        "nx-" & Replace$(NxCreateRunUuid(), "-", vbNullString) & ".tmp.xlsx")
    candidate = fso.GetAbsolutePathName(candidate)
    If Len(candidate) > 218 Then NxRaiseContractError "파일 통합 임시 저장 경로가 너무 깁니다. 더 짧은 저장 폴더를 선택하세요."
    If fso.FileExists(candidate) Or fso.FolderExists(candidate) Then NxRaiseContractError "임시 저장 경로가 이미 사용 중입니다. 다시 실행하세요."
    NxConsolidationTemporaryPath = candidate
End Function

Private Function NxAppendConsolidationRows(ByVal target As Worksheet, ByVal inputPath As String, _
    ByRef created As Long, ByRef nextRow As Long, ByRef expectedColumns As Long, ByRef importedCells As Double, ByVal skipHeaders As Boolean) As String
    Dim book As Workbook, sheet As Worksheet, source As Range, block As Range, owned As Boolean
    Dim startRow As Long, startCreated As Long, startColumns As Long, startCells As Double
    Dim lastTouchedRow As Long, touchedColumns As Long, offset As Long, rowCount As Long, columnIndex As Long
    Dim detail As String, rollbackFailed As Boolean
    Dim sourceHeaders As Variant, targetHeaders As Variant, sourceHeader As Variant, targetHeader As Variant
    On Error GoTo Failed
    startRow = nextRow: startCreated = created: startColumns = expectedColumns: startCells = importedCells
    Set book = NxSafeWorkbookOpen(inputPath, owned)
    For Each sheet In book.Worksheets
        If sheet.Visible = xlSheetVisible Then
            Set source = NxConsolidationContentRange(sheet)
            If Not source Is Nothing Then
                If expectedColumns = 0 Then expectedColumns = source.Columns.Count
                If source.Columns.Count <> expectedColumns Then NxRaiseContractError "열 수가 다른 시트는 이어붙일 수 없습니다: " & sheet.Name
                offset = 0
                If skipHeaders And created > 0 Then
                    sourceHeaders = source.Rows(1).Value2
                    targetHeaders = target.Cells(1, 1).Resize(1, expectedColumns).Value2
                    For columnIndex = 1 To expectedColumns
                        If expectedColumns = 1 Then
                            sourceHeader = sourceHeaders: targetHeader = targetHeaders
                        Else
                            sourceHeader = sourceHeaders(1, columnIndex): targetHeader = targetHeaders(1, columnIndex)
                        End If
                        If IsError(sourceHeader) Or IsError(targetHeader) Then NxRaiseContractError "제목행에 오류 값이 있습니다: " & sheet.Name
                        If CStr(sourceHeader) <> CStr(targetHeader) Then NxRaiseContractError "제목행이 다른 시트는 제목행 생략 옵션으로 통합할 수 없습니다: " & sheet.Name
                    Next columnIndex
                    offset = 1
                End If
                rowCount = source.Rows.Count - offset
                If rowCount > 0 Then
                    If CDbl(nextRow) + rowCount - 1 > target.Rows.Count Then NxRaiseContractError "이어붙인 결과가 Excel 행 한도를 초과합니다."
                    If importedCells + CDbl(rowCount) * expectedColumns > 500000 Then NxRaiseContractError "이어붙이기는 전체 500,000셀까지 지원합니다."
                    Set block = source.Cells(1 + offset, 1).Resize(rowCount, expectedColumns)
                    ' Track the intended write before assigning, so a partially failed write is removed.
                    lastTouchedRow = nextRow + rowCount - 1: touchedColumns = expectedColumns
                    With target.Cells(nextRow, 1).Resize(rowCount, expectedColumns)
                        .NumberFormat = "@"
                        .Value2 = block.Value2
                    End With
                    nextRow = lastTouchedRow + 1
                    importedCells = importedCells + CDbl(rowCount) * expectedColumns
                    created = created + 1
                End If
            End If
        End If
    Next sheet
    NxSafeWorkbookClose book, owned
    Exit Function
Failed:
    detail = Err.Description: Err.Clear
    On Error Resume Next
    If lastTouchedRow >= startRow And touchedColumns > 0 Then target.Cells(startRow, 1).Resize(lastTouchedRow - startRow + 1, touchedColumns).Clear
    If Err.Number <> 0 Then rollbackFailed = True: Err.Clear
    If Not book Is Nothing Then NxSafeWorkbookClose book, owned
    If Err.Number <> 0 Then rollbackFailed = True: Err.Clear
    On Error GoTo 0
    created = startCreated: nextRow = startRow: expectedColumns = startColumns: importedCells = startCells
    If rollbackFailed Then NxRaiseContractError "부분 입력 복구에 실패하여 결과 전체를 폐기합니다: " & detail
    NxAppendConsolidationRows = detail
End Function

Private Function NxConsolidationContentRange(ByVal sheet As Worksheet) As Range
    Dim used As Range, lastRow As Range, lastColumn As Range
    Set used = sheet.UsedRange
    ' Keep the original top/left layout. Trailing formatting is not row data.
    ' LookIn formulas retains formula cells even when their result is blank.
    Set lastRow = used.Find(What:="*", After:=used.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    If lastRow Is Nothing Then Exit Function
    Set lastColumn = used.Find(What:="*", After:=used.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, _
        MatchCase:=False, MatchByte:=False, SearchFormat:=False)
    If lastColumn Is Nothing Then NxRaiseContractError "통합할 데이터의 마지막 열을 확인할 수 없습니다."
    Set NxConsolidationContentRange = sheet.Range(used.Cells(1, 1), sheet.Cells(lastRow.Row, lastColumn.Column))
End Function

Private Function NxConsolidationInputIssue(ByVal inputPath As String, ByVal seen As Object) As String
    Dim key As String, existingBook As Workbook
    On Error GoTo Unreadable
    If Len(inputPath) = 0 Then NxConsolidationInputIssue = "Input file is not readable: " & inputPath: Exit Function
    key = LCase$(inputPath)
    If seen.Exists(key) Then NxConsolidationInputIssue = "Duplicate input file: " & inputPath: Exit Function
    seen.Add key, True
    Set existingBook = NxSafeWorkbookFindOpen(inputPath)
    If existingBook Is Nothing Then NxConsolidationInputIssue = NxSafeWorkbookInputIssue(inputPath)
    Exit Function
Unreadable:
    NxConsolidationInputIssue = "Input file is not readable: " & inputPath
End Function

Private Sub NxAppendConsolidationIssue(ByRef issueList As String, ByVal inputPath As String, ByVal detail As String)
    If Len(issueList) > 0 Then issueList = issueList & vbCrLf
    issueList = issueList & inputPath & " :: " & detail
End Sub

Private Sub NxFinishConsolidationOperation(ByRef operationStarted As Boolean)
    If Not operationStarted Then Exit Sub
    NxFileOperationFinish
    operationStarted = False
End Sub

Private Function NxAppendConsolidationFile(ByVal resultBook As Workbook, ByVal inputPath As String, _
    ByVal usedNames As Object, ByRef created As Long, Optional ByVal keepFirst As Boolean = False) As String
    Dim sourceBook As Workbook, source As Worksheet, target As Worksheet, detail As String
    Dim startSheetCount As Long, startCreated As Long, index As Long
    Dim sourceOwned As Boolean
    On Error GoTo Failed
    startSheetCount = resultBook.Worksheets.Count
    startCreated = created
    Set sourceBook = NxSafeWorkbookOpen(inputPath, sourceOwned)
    For Each source In sourceBook.Worksheets
        If source.Visible = xlSheetVisible Then
            If created = 0 And Not keepFirst Then Set target = resultBook.Worksheets(1) Else Set target = resultBook.Worksheets.Add(After:=resultBook.Worksheets(resultBook.Worksheets.Count))
            NxRebuildVisibleSheet source, target, NxSafeSheetName(source.Name, usedNames)
            created = created + 1
            Set target = Nothing
        End If
    Next source
    NxSafeWorkbookClose sourceBook, sourceOwned
    Set sourceBook = Nothing
    Exit Function
Failed:
    detail = Err.Description
    On Error Resume Next
    If startCreated = 0 And Not keepFirst Then
        For index = resultBook.Worksheets.Count To 2 Step -1
            resultBook.Worksheets(index).Delete
        Next index
        resultBook.Worksheets(1).Cells.Clear
    Else
        For index = resultBook.Worksheets.Count To startSheetCount + 1 Step -1
            resultBook.Worksheets(index).Delete
        Next index
    End If
    created = startCreated
    If Not sourceBook Is Nothing Then NxSafeWorkbookClose sourceBook, sourceOwned
    On Error GoTo 0
    NxAppendConsolidationFile = detail
End Function

Private Function NxAppendConsolidationBoth(ByVal resultBook As Workbook, ByVal inputPath As String, _
    ByVal usedNames As Object, ByRef created As Long, ByRef rowCreated As Long, _
    ByRef nextRow As Long, ByRef expectedColumns As Long, ByRef importedCells As Double, ByVal skipHeaders As Boolean) As String
    Dim startRow As Long, startCreated As Long, startColumns As Long, startCells As Double, issue As String
    startRow = nextRow: startCreated = rowCreated: startColumns = expectedColumns: startCells = importedCells
    issue = NxAppendConsolidationRows(resultBook.Worksheets(1), inputPath, rowCreated, nextRow, expectedColumns, importedCells, skipHeaders)
    If Len(issue) = 0 Then
        issue = NxAppendConsolidationFile(resultBook, inputPath, usedNames, created, True)
        If Len(issue) > 0 Then
            ' Both outputs must contain the same accepted files, including partial results.
            If nextRow > startRow Then resultBook.Worksheets(1).Cells(startRow, 1).Resize(nextRow - startRow, expectedColumns).Clear
            nextRow = startRow: rowCreated = startCreated: expectedColumns = startColumns: importedCells = startCells
        End If
    End If
    NxAppendConsolidationBoth = issue
End Function

Private Sub NxConsolidationAddContents(ByVal resultBook As Workbook, ByVal firstDataSheet As Long)
    Dim targets As New Collection, sheet As Worksheet, contents As Worksheet, index As Long
    For index = firstDataSheet To resultBook.Worksheets.Count
        targets.Add resultBook.Worksheets(index)
    Next index
    Set contents = resultBook.Worksheets.Add(Before:=resultBook.Worksheets(1))
    contents.Name = "목차"
    contents.Range("A1").Value2 = "시트 목차"
    contents.Range("A3").Value2 = "번호": contents.Range("B3").Value2 = "시트"
    NxDrawStyleGeneratedTable resultBook, contents.Range("A3").Resize(targets.Count + 1, 2), 1
    For index = 1 To targets.Count
        NxFileOperationReport index - 1, "목차 연결: " & CStr(index) & " / " & CStr(targets.Count)
        If NxFileOperationCancelled() Then NxRaiseContractError "목차 생성을 취소했습니다."
        Set sheet = targets(index)
        sheet.Rows(1).Insert Shift:=xlDown
        sheet.Rows(1).RowHeight = 20
        sheet.Hyperlinks.Add Anchor:=sheet.Range("A1"), Address:=vbNullString, _
            SubAddress:="'목차'!B" & CStr(index + 3), TextToDisplay:="목차로 이동"
        contents.Cells(index + 3, 1).Value2 = index
        contents.Hyperlinks.Add Anchor:=contents.Cells(index + 3, 2), Address:=vbNullString, _
            SubAddress:="'" & Replace$(sheet.Name, "'", "''") & "'!A2", TextToDisplay:=sheet.Name
    Next index
    contents.Columns(1).ColumnWidth = 8: contents.Columns(2).ColumnWidth = 38
    contents.Activate: contents.Range("A1").Select
End Sub

Private Sub NxWritePartialNotice(ByVal resultBook As Workbook, ByVal failureList As String)
    Dim notice As Worksheet, baseName As String, suffix As Long
    Set notice = resultBook.Worksheets.Add(Before:=resultBook.Worksheets(1))
    baseName = "부분결과_안내"
    Do
        On Error Resume Next
        Err.Clear
        notice.Name = baseName
        If Err.Number = 0 Then Exit Do
        Err.Clear
        suffix = suffix + 1
        baseName = "부분결과_안내_" & CStr(suffix)
        On Error GoTo 0
    Loop
    On Error GoTo 0
    notice.Range("A1").Value = "내엑셀 파일 통합 부분결과"
    notice.Range("A2").Value = "실패 파일 목록"
    notice.Range("A3").Value = failureList
    notice.Columns("A:A").ColumnWidth = 110
    notice.Rows("3:3").EntireRow.AutoFit
End Sub

Private Function NxPartialResultPath(ByVal resultPath As String) As String
    Dim dot As Long
    dot = InStrRev(resultPath, ".")
    If dot > InStrRev(resultPath, Application.PathSeparator) Then
        NxPartialResultPath = Left$(resultPath, dot - 1) & "_부분결과" & Mid$(resultPath, dot)
    Else
        NxPartialResultPath = resultPath & "_부분결과.xlsx"
    End If
End Function
