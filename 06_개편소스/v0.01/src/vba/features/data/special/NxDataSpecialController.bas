Attribute VB_Name = "NxDataSpecialController"
Option Explicit

Public Sub NxDataRunSpecialFromRibbon(ByVal featureId As String)
    ' The native form owns the CNxRangeSelectionSession and preview lifetime.
    Dim form As Object
    If featureId = NX_FEATURE_DATA_AGE Then Set form = New FNxAgeCalculator Else Set form = New FNxDataSpecial
    If featureId = NX_FEATURE_DATA_PRIVACY_SCAN Then
        NxDataRunWorkbookPrivacyScan
        Exit Sub
    End If
    form.BindFeature featureId
    form.Show vbModal
End Sub

Public Function NxDataSpecialPreview(ByVal featureId As String, ByVal source As Range, ByVal options As CNxDataSpecialOptions) As String
    Dim cell As Range, sample As String, result As Variant, n As Long, errors As Long, formulas As Long, processed As Long
    Dim blankCount As Long, outputSummary As String
    Dim raw As Variant, value As Variant, formulaState As Variant
    Dim rowCount As Long, columnCount As Long, r As Long, c As Long, hasFormula As Boolean
    Set source = NxDataContentRange(source)
    NxDataSpecialValidateSource source, featureId
    raw = source.Value2: formulaState = source.HasFormula
    rowCount = source.Rows.Count: columnCount = source.Columns.Count
    For r = 1 To rowCount
      For c = 1 To columnCount
        value = NxDataSpecialMatrixValue(raw, rowCount, columnCount, r, c)
        If NxDataValueIsBlank(value) Then
            blankCount = blankCount + 1
            GoTo NextPreviewCell
        End If
        If options.HasHeader And ((rowCount = 1 And c = 1) Or _
            (rowCount > 1 And r = 1)) Then
            result = value
            If IsError(result) Then result = vbNullString
        Else
            result = options.TransformValue(featureId, value)
        End If
        If NxDataValueIsBlank(result) Then errors = errors + 1
        If IsNull(formulaState) Then
            hasFormula = CBool(source.Cells(r, c).HasFormula)
        Else
            hasFormula = CBool(formulaState)
        End If
        If hasFormula Then formulas = formulas + 1
        If n < 10 Then
            n = n + 1
            Set cell = source.Cells(r, c)
            sample = sample & vbCrLf & cell.Address(False, False) & ": " & NxDataSpecialDisplay(value) & " → " & NxDataSpecialDisplay(result)
        End If
        processed = processed + 1
        If processed Mod 1024 = 0 Then DoEvents
NextPreviewCell:
      Next c
    Next r
    outputSummary = "결과: " & options.OutputMode
    If featureId = NX_FEATURE_DATA_AGE Then outputSummary = outputSummary & " · 기준일: " & Format$(options.ReferenceDate, "yyyy-mm-dd")
    NxDataSpecialPreview = "원본: " & source.Worksheet.Name & "!" & source.Address & vbCrLf & _
        outputSummary & vbCrLf & _
        "처리 " & processed & "셀 · 빈 셀 제외 " & blankCount & "개 · 빈칸 처리 " & errors & "개 · 수식→값 " & formulas & "개" & vbCrLf & _
        "앞 " & n & "개 셀 전 → 후" & sample
End Function

Private Function NxDataSpecialDisplay(ByVal value As Variant) As String
    If IsError(value) Then NxDataSpecialDisplay = "[셀 오류]": Exit Function
    If IsEmpty(value) Or IsNull(value) Then NxDataSpecialDisplay = "[빈 셀]": Exit Function
    NxDataSpecialDisplay = Left$(Replace(Replace(CStr(value), vbCr, "↵"), vbLf, "↵"), 80)
End Function

Public Function NxDataSpecialRunOptions(ByVal featureId As String, ByVal source As Range, ByVal options As CNxDataSpecialOptions, _
    Optional ByVal forTest As Boolean = False) As CNxResult
    Dim output As Range, commandObject As New CNxDataSpecialCommand, command As INxFeatureCommand
    Dim result As CNxResult, router As New CNxExecutionRouter, ticket As CNxExecutionTicket
    Dim report As Worksheet, errorNumber As Long, detail As String
    On Error GoTo Failed
    Set source = NxDataContentRange(source)
    NxDataSpecialValidateSource source, featureId
    If options Is Nothing Then NxRaiseContractError "옵션을 확인하세요."
    Select Case options.OutputMode
        Case "원본 변경": Set output = source
        Case "빈 결과 열/행"
            If source.Rows.Count = 1 And source.Columns.Count > 1 Then
                If source.Row = source.Worksheet.Rows.Count Then NxRaiseContractError "아래에 결과 행이 부족합니다."
                Set output = source.Offset(1, 0)
            Else
                Set output = NxDataSpecialOutputRange(source)
            End If
            NxDataSpecialRequireBlankOutput output
        Case "오른쪽 새 열 삽입", "새 시트 만들기", "새 통합문서": Set output = source
        Case Else: NxRaiseContractError "결과 위치를 확인하세요."
    End Select
    output.Worksheet.Activate: output.Select
    commandObject.ConfigureOptions featureId, source, output, options
    Set command = commandObject
    If forTest Then
        Set ticket = router.Prepare(NxDataFeatureDefinition(featureId), command)
        Select Case ticket.Decision.ResolvedGrade
            Case NxExecutionFast: Set result = router.ExecuteFast(ticket)
            Case NxExecutionGuarded: Set result = router.ExecuteGuarded(ticket, True)
            Case NxExecutionPlanned: Set result = NxDataRunPlannedFrameForTest(command, ticket, router, options.OutputMode)
        End Select
    Else
        Set result = NxDataRunCommand(command, "변환 결과를 " & options.OutputMode & " 방식으로 적용합니다.")
    End If
    Set report = commandObject.ResultSheet
    If result Is Nothing Then NxRaiseContractError "변환 결과가 없습니다."
    If result.Outcome <> NxSuccess Then commandObject.DiscardTransformResult: Set report = Nothing
    source.Worksheet.Activate: source.Select
    If Not report Is Nothing Then report.Activate
    Set NxDataSpecialRunOptions = result
    Exit Function
Failed:
    errorNumber = Err.Number: detail = Err.Description: Err.Clear
    On Error Resume Next
    commandObject.DiscardTransformResult
    If Err.Number <> 0 Then detail = detail & " | 결과 정리 실패: " & Err.Description: Err.Clear
    If Not source Is Nothing Then source.Worksheet.Activate: source.Select
    On Error GoTo 0
    Err.Raise errorNumber, "NxDataSpecialRunOptions", detail
End Function

Private Sub NxDataRunWorkbookPrivacyScan()
    Static running As Boolean
    Dim sourceBook As Workbook, originalSheet As Object, originalSelection As Object, anchor As Worksheet, sheet As Worksheet
    Dim commandObject As New CNxDataSpecialCommand, command As INxFeatureCommand, result As CNxResult, report As Worksheet
    Dim priorEvents As Boolean, errorNumber As Long, detail As String, restoreFailed As Boolean
    Dim priorCancel As XlEnableCancelKey, priorStatus As Variant
    If running Then Exit Sub
    running = True
    priorEvents = Application.EnableEvents
    priorCancel = Application.EnableCancelKey: priorStatus = Application.StatusBar
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    NxPrivacyBeginCancellation
    Application.StatusBar = "개인정보 검사 중 · Esc를 누르면 중단합니다."
    Set sourceBook = ActiveWorkbook: Set originalSheet = ActiveSheet: Set originalSelection = Application.Selection
    If sourceBook Is Nothing Then NxRaiseContractError "점검할 엑셀 파일을 먼저 여세요."
    If TypeName(originalSheet) = "Worksheet" Then Set anchor = originalSheet
    If anchor Is Nothing Then
        For Each sheet In sourceBook.Worksheets
            If sheet.Visible = xlSheetVisible Then Set anchor = sheet: Exit For
        Next sheet
    End If
    If anchor Is Nothing Then NxRaiseContractError "워크시트 하나를 표시한 뒤 점검을 실행하세요. 숨김 시트도 함께 검사합니다."
    ' Only the execution-context anchor is selected; the actual scan is workbook-wide.
    ' This avoids treating a selected whole column as a 1,048,576-cell write journal.
    Application.EnableEvents = False
    anchor.Activate: anchor.Cells(1, 1).Select
    Application.EnableEvents = priorEvents
    commandObject.ConfigureWorkbookScan sourceBook
    Set command = commandObject
    Set result = NxDataRunCommand(command, "숨김 시트까지 전체 셀 값을 점검하고 의심 값과 위치를 별도 통합문서에 표시합니다. 원본은 바꾸지 않습니다.")
    Set report = commandObject.ScanReport
    If result Is Nothing Then NxRaiseContractError "개인정보 점검 결과를 확인할 수 없습니다."
    If result.Outcome <> NxSuccess And result.Outcome <> NxPartialFailure And result.Outcome <> NxCancelled Then NxRaiseContractError result.Recovery
    If result.Outcome = NxCancelled Then
        commandObject.DiscardScanReport
        Set report = Nothing
    End If
CleanUp:
    On Error Resume Next
    NxPrivacyEndCancellation
    Application.EnableCancelKey = xlDisabled
    Application.EnableEvents = False
    If errorNumber <> 0 Then
        commandObject.DiscardScanReport
        If Err.Number <> 0 Then detail = detail & " | 보고서 정리 실패: " & Err.Description: Err.Clear
        Set report = Nothing
    End If
    If Not originalSheet Is Nothing Then originalSheet.Activate
    If Not originalSelection Is Nothing Then originalSelection.Select
    If Err.Number <> 0 Then restoreFailed = True: Err.Clear
    If restoreFailed Then
        If errorNumber = 0 Then errorNumber = NX_CONTRACT_ERROR: detail = "점검 후 원래 선택을 복원하지 못했습니다. 선택 상태를 확인하세요."
        commandObject.DiscardScanReport
        If Err.Number <> 0 Then detail = detail & " | 보고서 정리 실패: " & Err.Description: Err.Clear
        Set report = Nothing
    End If
    Application.EnableEvents = priorEvents
    If IsEmpty(priorStatus) Or VarType(priorStatus) = vbBoolean Then
        Application.StatusBar = vbNullString
    Else
        Application.StatusBar = priorStatus
    End If
    Application.EnableCancelKey = priorCancel
    running = False
    On Error GoTo 0
    If errorNumber = 18 Then Exit Sub
    If errorNumber <> 0 Then Err.Raise errorNumber, "NxPrivacyScan", detail
    On Error GoTo Failed
    If Not report Is Nothing Then report.Activate
    Exit Sub
Failed:
    errorNumber = Err.Number: detail = Err.Description: Err.Clear
    Resume CleanUp
End Sub

Public Function NxDataSpecialSourceMatches(ByVal source As Range, ByVal snapshot As Variant, _
    ByVal rowCount As Long, ByVal columnCount As Long, Optional ByVal compareFormulas As Boolean = False) As Boolean

    Dim current As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    If source Is Nothing Then Exit Function
    If source.Rows.Count <> rowCount Or source.Columns.Count <> columnCount Then Exit Function
    If compareFormulas Then current = source.Formula Else current = source.Value2
    For rowIndex = 1 To rowCount
        For columnIndex = 1 To columnCount
            If Not NxDataSpecialValuesMatch(NxDataSpecialMatrixValue(current, rowCount, columnCount, rowIndex, columnIndex), _
                NxDataSpecialMatrixValue(snapshot, rowCount, columnCount, rowIndex, columnIndex)) Then Exit Function
        Next columnIndex
    Next rowIndex
    NxDataSpecialSourceMatches = True
End Function

Private Sub NxDataSpecialValidateSource(ByVal source As Range, ByVal featureId As String)
    If source Is Nothing Then NxRaiseContractError "변환할 범위를 선택하세요."
    If source.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 선택하세요."
    If source.Worksheet.Parent Is ThisWorkbook Then NxRaiseContractError "추가기능 파일의 시트는 처리할 수 없습니다."
    If source.Worksheet.ProtectContents Then NxRaiseContractError "보호된 시트에서는 실행할 수 없습니다."
    If source.Worksheet.Parent.ReadOnly Then NxRaiseContractError "읽기 전용 파일은 먼저 편집 가능한 사본으로 여세요."
    If source.Worksheet.Parent.ProtectStructure Then NxRaiseContractError "통합문서 구조 보호를 해제한 뒤 실행하세요."
    If IsNull(source.MergeCells) Then NxRaiseContractError "병합된 셀은 먼저 정규화해 주세요."
    If source.MergeCells Then NxRaiseContractError "병합된 셀은 먼저 정규화해 주세요."
    If featureId = NX_FEATURE_DATA_PRIVACY_SCAN Then
        If source.CountLarge > 49999 Then NxRaiseContractError "개인정보 점검은 한 번에 49,999셀까지 지원합니다."
    ElseIf source.CountLarge > 200000 Then
        NxRaiseContractError "데이터 변환은 한 번에 200,000셀까지 지원합니다."
    End If
End Sub

Private Function NxDataSpecialOutputRange(ByVal source As Range) As Range
    If source.Column + source.Columns.Count * 2 - 1 > source.Worksheet.Columns.Count Then _
        NxRaiseContractError "오른쪽에 결과 범위를 만들 열이 부족합니다."
    Set NxDataSpecialOutputRange = source.Offset(0, source.Columns.Count)
End Function

Private Sub NxDataSpecialRequireBlankOutput(ByVal output As Range)
    Dim cell As Range
    For Each cell In output.Cells
        If cell.MergeCells Then NxRaiseContractError "결과 범위에 병합된 셀이 있습니다."
        If cell.HasFormula Or Not IsEmpty(cell.Value2) Then NxRaiseContractError "인접 결과 범위가 비어 있지 않습니다. 빈 범위를 확보한 뒤 다시 실행하세요."
    Next cell
End Sub

Private Function NxDataSpecialMatrixValue(ByRef values As Variant, ByVal rowCount As Long, ByVal columnCount As Long, _
    ByVal rowIndex As Long, ByVal columnIndex As Long) As Variant
    If rowCount = 1 And columnCount = 1 Then
        NxDataSpecialMatrixValue = values
    Else
        NxDataSpecialMatrixValue = values(rowIndex, columnIndex)
    End If
End Function

Private Function NxDataSpecialValuesMatch(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If IsError(leftValue) Or IsError(rightValue) Then
        If IsError(leftValue) And IsError(rightValue) Then NxDataSpecialValuesMatch = (NxDataKey(leftValue, False, True) = NxDataKey(rightValue, False, True))
    ElseIf IsEmpty(leftValue) Or IsEmpty(rightValue) Then
        NxDataSpecialValuesMatch = (IsEmpty(leftValue) And IsEmpty(rightValue))
    ElseIf IsNull(leftValue) Or IsNull(rightValue) Then
        NxDataSpecialValuesMatch = (IsNull(leftValue) And IsNull(rightValue))
    Else
        NxDataSpecialValuesMatch = (NxDataKey(leftValue, False, True) = NxDataKey(rightValue, False, True))
    End If
End Function
