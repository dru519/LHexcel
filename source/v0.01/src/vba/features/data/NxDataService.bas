Attribute VB_Name = "NxDataService"
Option Explicit

Private Const NX_DATA_RESULT_COLUMNS As Long = 6

Public Function NxDataRunAnalysis(ByVal request As CNxDataRequest) As CNxResult
    Dim commandObject As New CNxDataFeatureCommand, command As INxFeatureCommand
    Dim result As CNxResult, failureNumber As Long, detail As String, cleanupDetail As String
    Dim output As Worksheet
    On Error GoTo Failed
    If request Is Nothing Then NxRaiseContractError "미리보기 요청이 없습니다."
    request.RequireFreshAnalysis
    commandObject.Configure request
    Set command = commandObject
    Set result = NxDataRunCommand(command, "미리보기에서 확인한 분석 결과를 " & request.OutputMode & " 방식으로 만듭니다.")
    If result Is Nothing Then NxRaiseContractError "분석 실행 결과가 없습니다."
    If result.Outcome <> NxSuccess Then
        commandObject.DiscardAnalysisOutput
    Else
        ' The shared state guard has restored the source selection; show the owned result now.
        Set output = commandObject.AnalysisOutput
        If output Is Nothing Then NxRaiseContractError "생성한 분석 결과를 찾을 수 없습니다."
        output.Parent.Activate
        output.Activate
        output.Range("A1").Select
    End If
    Set NxDataRunAnalysis = result
    Exit Function
Failed:
    failureNumber = Err.Number: detail = Err.Description: Err.Clear
    On Error Resume Next
    commandObject.DiscardAnalysisOutput
    If Err.Number <> 0 Then cleanupDetail = " | 생성된 분석 결과를 수동으로 정리하세요: " & Err.Description
    Err.Clear
    On Error GoTo 0
    If Len(cleanupDetail) > 0 Then
        Set NxDataRunAnalysis = NxCreateResult(request.FeatureId, NxPartialFailure, "analysis_cleanup", NxDataSafeTarget(request), _
            (request.OutputMode = "새 시트 만들기"), detail & cleanupDetail, "result_partial_failure")
    Else
        Err.Raise failureNumber, "NxDataRunAnalysis", detail
    End If
End Function

Public Function NxDataAnalysisPreview(ByVal request As CNxDataRequest) As String
    Dim values As Variant, rows As Long, columns As Long, rowIndex As Long, columnIndex As Long, line As String
    values = NxDataAnalysisValues(request)
    Call NxDataAnalysisOutputCells(request.FeatureId, values)
    rows = UBound(values, 1): columns = UBound(values, 2)
    NxDataAnalysisPreview = "결과 " & CStr(rows) & "행 × " & CStr(columns) & "열 · " & request.OutputMode
    For rowIndex = 1 To rows
        If rowIndex > 8 Then NxDataAnalysisPreview = NxDataAnalysisPreview & vbCrLf & "… 이하 생략": Exit For
        line = ""
        For columnIndex = 1 To columns
            If columnIndex > 6 Then line = line & " | …": Exit For
            If columnIndex > 1 Then line = line & " | "
            line = line & Left$(NxDataAnalysisCleanText(values(rowIndex, columnIndex), False), 56)
        Next columnIndex
        NxDataAnalysisPreview = NxDataAnalysisPreview & vbCrLf & line
    Next rowIndex
    request.RequireFreshAnalysis
End Function

Public Function NxDataCreateAnalysisResult(ByVal request As CNxDataRequest, ByVal values As Variant, ByRef createdOutput As Worksheet) As CNxResult
    Dim sourceBook As Workbook, outputBook As Workbook, outputSheet As Worksheet, output As Range
    Dim newBook As Boolean, outputRows As Long, outputColumns As Long, rowIndex As Long, columnIndex As Long
    Dim formulas As Range, sourceChanged As Boolean, failureNumber As Long, detail As String
    Dim priorAlerts As Boolean, priorEvents As Boolean, cleanupFailed As Boolean, created As Boolean, stateCaptured As Boolean
    On Error GoTo Failed
    If request Is Nothing Then NxRaiseContractError "분석 요청이 필요합니다."
    request.RequireFreshAnalysis
    Call NxDataAnalysisOutputCells(request.FeatureId, values)
    Set sourceBook = request.SourceRange.Worksheet.Parent
    newBook = (request.OutputMode = "새창으로 보기")
    If Not newBook Then
        If sourceBook.ReadOnly Or sourceBook.ProtectStructure Then NxRaiseContractError "현재 통합문서에는 결과 시트를 만들 수 없습니다. 새창으로 보기를 선택하세요."
    End If
    priorAlerts = Application.DisplayAlerts: priorEvents = Application.EnableEvents: stateCaptured = True
    Application.EnableEvents = False
    If newBook Then
        Set outputBook = Application.Workbooks.Add(xlWBATWorksheet)
        created = True
        Set outputSheet = outputBook.Worksheets(1)
    Else
        Set outputBook = sourceBook
        Set outputSheet = outputBook.Worksheets.Add(After:=outputBook.Worksheets(outputBook.Worksheets.Count))
        created = True: sourceChanged = True
    End If
    Set createdOutput = outputSheet
    If request.FeatureId = "NX-DATA-UNIQUE-COUNT" Then
        outputSheet.Name = NxDataNextSheetName(outputBook, "값별 개수·비율")
    Else
        outputSheet.Name = NxDataNextSheetName(outputBook, "두 목록 대조")
    End If
    outputRows = UBound(values, 1): outputColumns = UBound(values, 2)
    Set output = outputSheet.Range("A1").Resize(outputRows, outputColumns)
    ' Set literal formats before writing any user text; never evaluate formula-like source values.
    output.NumberFormat = "@"
    output.Value2 = values
    If newBook Then NxDrawStyleGeneratedTable outputBook, output, IIf(request.FeatureId = "NX-DATA-UNIQUE-COUNT" Or (request.HasTitleRow And Not request.RowDirection), 1, 0)
    If request.FeatureId = "NX-DATA-UNIQUE-COUNT" Then
        columnIndex = outputColumns - IIf(request.IncludeRatio, 1, 0)
        output.Columns(columnIndex).NumberFormat = "0"
        If request.IncludeRatio Then output.Columns(outputColumns).NumberFormat = "0.0%"
        output.Rows(1).Font.Bold = True
    ElseIf request.HasTitleRow Then
        If request.RowDirection Then output.Columns(1).Font.Bold = True Else output.Rows(1).Font.Bold = True
    End If
    On Error Resume Next
    Set formulas = output.SpecialCells(xlCellTypeFormulas)
    Err.Clear
    On Error GoTo Failed
    If Not formulas Is Nothing Then NxRaiseContractError "분석 결과에서 실행 가능한 수식이 발견되었습니다."
    output.Columns.AutoFit
    For columnIndex = 1 To outputColumns
        If output.Columns(columnIndex).ColumnWidth > 50 Then output.Columns(columnIndex).ColumnWidth = 50
    Next columnIndex
    Application.EnableEvents = priorEvents
    Set NxDataCreateAnalysisResult = NxCreateResult(request.FeatureId, NxSuccess, "complete", sourceBook.Name, sourceChanged, vbNullString, "result_success")
    Exit Function
Failed:
    failureNumber = Err.Number: detail = Err.Description: Err.Clear
    On Error Resume Next
    If created Then
        Application.DisplayAlerts = False
        If newBook Then
            ' The exact created object is the only workbook eligible for compensation.
            If Not outputBook Is sourceBook And Not outputBook Is ThisWorkbook Then outputBook.Close SaveChanges:=False
        Else
            If Not outputSheet Is request.SourceRange.Worksheet Then outputSheet.Delete
        End If
        cleanupFailed = (Err.Number <> 0)
        If cleanupFailed Then detail = detail & " | 생성 결과 정리 실패: " & Err.Description
        Err.Clear
        Application.DisplayAlerts = priorAlerts
        If Not cleanupFailed Then Set createdOutput = Nothing
    End If
    If stateCaptured Then Application.EnableEvents = priorEvents
    On Error GoTo 0
    If cleanupFailed Then
        Set NxDataCreateAnalysisResult = NxCreateResult(request.FeatureId, NxPartialFailure, "analysis_result", NxDataSafeTarget(request), _
            sourceChanged, detail, "result_partial_failure")
    Else
        Set NxDataCreateAnalysisResult = NxCreateResult(request.FeatureId, NxEnvironmentError, "analysis_result", NxDataSafeTarget(request), _
            False, CStr(failureNumber) & " " & detail, "result_environment_error")
    End If
End Function

Public Function NxDataGroups(ByVal request As CNxDataRequest) As Collection
    Dim values As Variant
    Dim sourceRows As Variant
    If request Is Nothing Then NxRaiseContractError "Data grouping requires a request"
    NxDataExtractKeyValues request, values, sourceRows
    Set NxDataGroups = NxDataBuildGroupsAtRows(values, sourceRows, request.TrimText, request.CaseSensitive)
End Function

Public Function NxDataDuplicateOutputCellCount(ByVal groups As Collection) As Long
    Dim group As CNxDataGroup
    Dim outputRows As Double
    Dim maximumCells As Long
    Dim policy As New CNxLimitsPolicy
    If groups Is Nothing Then NxRaiseContractError "Duplicate output requires sealed groups"

    outputRows = 1#
    For Each group In groups
        If Not group.IsSealed Then NxRaiseContractError "Duplicate output requires sealed groups"
        If group.Count > 1 Then outputRows = outputRows + CDbl(group.Count)
    Next group
    maximumCells = policy.MaximumCellsFor("NX-DATA-DUPLICATE-LIST")
    If outputRows * NX_DATA_RESULT_COLUMNS > CDbl(maximumCells) Then NxRaiseContractError "Duplicate output exceeds the 200,000-cell safety limit"
    NxDataDuplicateOutputCellCount = CLng(outputRows * NX_DATA_RESULT_COLUMNS)
End Function

Public Function NxDataCreateDuplicateResult(ByVal request As CNxDataRequest, ByVal groups As Collection) As CNxResult
    Dim source As Range
    Dim targetBook As Workbook
    Dim targetSheet As Worksheet
    Dim output As Range
    Dim formulaCells As Range
    Dim matrix As Variant
    Dim group As CNxDataGroup
    Dim sourceRows As Collection
    Dim sourceRow As Variant
    Dim occurrence As Long
    Dim outputCellCount As Long
    Dim outputRowCount As Long
    Dim outputRow As Long
    Dim groupIndex As Long
    Dim alertsBefore As Boolean
    Dim createdSheetName As String
    Dim residualSheet As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed

    If request Is Nothing Then NxRaiseContractError "Duplicate result requires request and groups"
    If groups Is Nothing Then NxRaiseContractError "Duplicate result requires request and groups"
    If request.FeatureId <> "NX-DATA-DUPLICATE-LIST" Then NxRaiseContractError "Duplicate result feature identity mismatch"
    Set source = request.SourceRange
    Set targetBook = source.Parent.Parent
    If targetBook.ReadOnly Then NxRaiseContractError "Duplicate result cannot be created in a read-only workbook"

    outputCellCount = NxDataDuplicateOutputCellCount(groups)
    outputRowCount = outputCellCount \ NX_DATA_RESULT_COLUMNS
    ReDim matrix(1 To outputRowCount, 1 To NX_DATA_RESULT_COLUMNS)
    matrix(1, 1) = "중복 그룹 ID"
    matrix(1, 2) = "정규화 비교값"
    matrix(1, 3) = "원본 값"
    matrix(1, 4) = "발생 건수"
    matrix(1, 5) = "원본 시트"
    matrix(1, 6) = "원본 행"

    outputRow = 1
    For Each group In groups
        If group.Count > 1 Then
            groupIndex = groupIndex + 1
            Set sourceRows = group.SourceRows
            occurrence = 0
            For Each sourceRow In sourceRows
                occurrence = occurrence + 1
                outputRow = outputRow + 1
                matrix(outputRow, 1) = "D" & Format$(groupIndex, "000000")
                matrix(outputRow, 2) = group.Key
                matrix(outputRow, 3) = group.OriginalValue(occurrence)
                matrix(outputRow, 4) = group.Count
                matrix(outputRow, 5) = source.Parent.Name
                matrix(outputRow, 6) = CLng(sourceRow)
            Next sourceRow
        End If
    Next group

    Set targetSheet = targetBook.Worksheets.Add(After:=targetBook.Worksheets(targetBook.Worksheets.Count))
    targetSheet.Name = NxDataNextSheetName(targetBook, "중복목록")
    createdSheetName = targetSheet.Name
    Set output = targetSheet.Range("A1").Resize(outputRowCount, NX_DATA_RESULT_COLUMNS)
    NxDataPrepareLiteralFormats output, groups
    output.Value = matrix

    On Error Resume Next
    Set formulaCells = output.SpecialCells(xlCellTypeFormulas)
    Err.Clear
    On Error GoTo Failed
    If Not formulaCells Is Nothing Then NxRaiseContractError "Duplicate result rejected an executable formula value"
    output.Columns.AutoFit
    Set NxDataCreateDuplicateResult = NxCreateResult(request.FeatureId, NxSuccess, "complete", targetSheet.Name, True, vbNullString, "result_success")
    Exit Function

Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If Not targetSheet Is Nothing Then
        alertsBefore = Application.DisplayAlerts
        Application.DisplayAlerts = False
        targetSheet.Delete
        Application.DisplayAlerts = alertsBefore
    End If
    If Not targetBook Is Nothing And Len(createdSheetName) > 0 Then residualSheet = NxDataSheetExists(targetBook, createdSheetName)
    On Error GoTo 0
    If residualSheet Then
        Set NxDataCreateDuplicateResult = NxCreateResult("NX-DATA-DUPLICATE-LIST", NxPartialFailure, "duplicate_result", NxDataSafeTarget(request), True, _
            CStr(failureNumber) & " " & failureDescription & " | Delete the command-created sheet: " & createdSheetName, "result_partial_failure")
    Else
        Set NxDataCreateDuplicateResult = NxCreateResult("NX-DATA-DUPLICATE-LIST", NxEnvironmentError, "duplicate_result", NxDataSafeTarget(request), False, _
            CStr(failureNumber) & " " & failureDescription, "result_environment_error")
    End If
End Function

Public Function NxDataRunForTest(ByVal featureId As String, ByVal source As Range, ByVal keyColumn As Long, _
    ByVal trimText As Boolean, ByVal caseSensitive As Boolean, ByVal includeHidden As Boolean, _
    Optional ByVal hasTitleRow As Boolean = True) As CNxResult

    Dim request As New CNxDataRequest
    Dim commandObject As New CNxDataFeatureCommand
    Dim command As INxFeatureCommand
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim approval As CNxApproval

    request.Configure featureId, source, hasTitleRow, keyColumn, trimText, caseSensitive, includeHidden
    commandObject.Configure request
    Set command = commandObject
    Set ticket = router.Prepare(NxDataFeatureDefinition(featureId), command)
    Select Case ticket.Decision.ResolvedGrade
        Case NxExecutionFast
            Set NxDataRunForTest = router.ExecuteFast(ticket)
        Case NxExecutionGuarded
            Set NxDataRunForTest = router.ExecuteGuarded(ticket, True)
        Case NxExecutionPlanned
            Set NxDataRunForTest = NxDataRunPlannedFrameForTest(command, ticket, router, "계획된 데이터 작업을 실행합니다.")
        Case Else
            NxRaiseContractError "Data execution grade is invalid"
    End Select
End Function

Private Sub NxDataExtractKeyValues(ByVal request As CNxDataRequest, ByRef values As Variant, ByRef sourceRows As Variant)
    Dim source As Range
    Dim keyRange As Range
    Dim raw As Variant
    Dim relativeRow As Long
    Dim firstDataRow As Long
    Dim includedCount As Long
    Dim outputIndex As Long

    Set source = request.SourceRange
    firstDataRow = IIf(request.HasTitleRow, 2, 1)
    For relativeRow = firstDataRow To source.Rows.Count
        If request.IncludeHidden Or Not CBool(source.Rows(relativeRow).EntireRow.Hidden) Then includedCount = includedCount + 1
    Next relativeRow
    If includedCount = 0 Then NxRaiseContractError "Data source has no included rows"

    Set keyRange = source.Columns(request.KeyColumn)
    raw = keyRange.Value
    ReDim values(0 To includedCount - 1)
    ReDim sourceRows(0 To includedCount - 1)
    For relativeRow = firstDataRow To source.Rows.Count
        If request.IncludeHidden Or Not CBool(source.Rows(relativeRow).EntireRow.Hidden) Then
            values(outputIndex) = NxDataRawValue(raw, relativeRow, source.Rows.Count)
            sourceRows(outputIndex) = source.Row + relativeRow - 1
            outputIndex = outputIndex + 1
        End If
        If relativeRow Mod 1024 = 0 Then DoEvents
    Next relativeRow
End Sub

Private Function NxDataRawValue(ByRef raw As Variant, ByVal relativeRow As Long, ByVal rowCount As Long) As Variant
    If rowCount = 1 Then
        NxDataRawValue = raw
    Else
        NxDataRawValue = raw(relativeRow, 1)
    End If
End Function

Private Sub NxDataPrepareLiteralFormats(ByVal output As Range, ByVal groups As Collection)
    Dim group As CNxDataGroup
    Dim outputRow As Long
    Dim occurrence As Long
    output.Rows(1).NumberFormat = "@"
    output.Columns(1).NumberFormat = "@"
    output.Columns(2).NumberFormat = "@"
    output.Columns(5).NumberFormat = "@"
    outputRow = 1
    For Each group In groups
        If group.Count > 1 Then
            For occurrence = 1 To group.Count
                outputRow = outputRow + 1
                If VarType(group.OriginalValue(occurrence)) = vbString Then
                    output.Cells(outputRow, 3).NumberFormat = "@"
                ElseIf VarType(group.OriginalValue(occurrence)) = vbDate Then
                    output.Cells(outputRow, 3).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                End If
            Next occurrence
        End If
    Next group
End Sub

Private Function NxDataNextSheetName(ByVal workbook As Workbook, ByVal baseName As String) As String
    Dim suffix As Long
    Dim candidate As String
    candidate = baseName
    suffix = 1
    Do While NxDataSheetExists(workbook, candidate)
        suffix = suffix + 1
        candidate = baseName & " (" & CStr(suffix) & ")"
    Loop
    NxDataNextSheetName = candidate
End Function

Private Function NxDataSheetExists(ByVal workbook As Workbook, ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = workbook.Worksheets(sheetName)
    NxDataSheetExists = Not sheet Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

Private Function NxDataSafeTarget(ByVal request As CNxDataRequest) As String
    On Error Resume Next
    NxDataSafeTarget = request.SourceRange.Parent.Parent.Name
    If Len(NxDataSafeTarget) = 0 Then NxDataSafeTarget = "data_result"
    Err.Clear
    On Error GoTo 0
End Function
