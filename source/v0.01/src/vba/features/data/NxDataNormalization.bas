Attribute VB_Name = "NxDataNormalization"
Option Explicit

Public Function NxDataNormalizePreset(ByVal value As String) As String
    Select Case Trim$(value)
        Case "기본 정리", "숫자 변환", "텍스트 변환", "날짜 정리", "표 구조 정리"
            NxDataNormalizePreset = Trim$(value)
        Case Else
            NxRaiseContractError "데이터 정규화 프리셋을 선택하세요."
    End Select
End Function
Public Function NxDataNormalizeDateFormat(ByVal value As String) As String
    Select Case Trim$(value)
        Case "yyyy-mm-dd", "yyyy-mm-dd hh:mm", "yyyy년 m월 d일", "yyyy.mm.dd.", "yyyy.mm.dd", "yyyy.m.d", "yyyy/mm/dd", "yyyymmdd", "yymmdd", "yy.mm.dd", "yy-mm-dd"
            NxDataNormalizeDateFormat = Trim$(value)
        Case Else
            NxRaiseContractError "날짜 표시 형식이 닫힌 선택지에 없습니다."
    End Select
End Function

Public Function NxDataNormalizeOutputMode(ByVal value As String) As String
    Select Case Trim$(value)
        Case "오른쪽 새 열 삽입", "새 시트 만들기", "원본 변경", "새 통합문서"
            NxDataNormalizeOutputMode = Trim$(value)
        Case Else
            NxRaiseContractError "결과 위치를 선택하세요."
    End Select
End Function

Public Function NxDataNormalizationDetail(ByVal preset As String, ByVal dateFormat As String, _
    ByVal outputMode As String) As String

    Dim detail As String
    preset = NxDataNormalizePreset(preset)
    dateFormat = NxDataNormalizeDateFormat(dateFormat)
    outputMode = NxDataNormalizeOutputMode(outputMode)
    Select Case preset
        Case "기본 정리"
            detail = "대상: 다운로드·복사한 텍스트. 예: '  홍  길동' → '홍 길동'." & vbCrLf & _
                "처리: 앞뒤·중복 공백과 제어문자, 특수 대시를 정리합니다." & vbCrLf & _
                "주의: 의도한 줄바꿈·들여쓰기도 일반 공백으로 바뀝니다."
        Case "숫자 변환"
            detail = "대상: '1,234원'처럼 문자로 들어온 금액. 합계·정렬에 사용할 숫자로 바꿉니다." & vbCrLf & _
                "처리: 쉼표·원화·괄호 음수·백분율을 해석하고 변환할 수 없는 값은 빈칸으로 둡니다." & vbCrLf & _
                "주의: 코드·우편번호의 앞자리 0은 사라질 수 있습니다."
        Case "텍스트 변환"
            detail = "대상: 코드·계정·품번 등 문자로 다룰 값. 현재 값을 텍스트로 고정합니다." & vbCrLf & _
                "처리: 수식처럼 보이는 문자열도 실행하지 않습니다." & vbCrLf & _
                "주의: 이미 숫자로 저장되어 사라진 앞자리 0은 복원하지 않습니다. 수식은 값이 됩니다."
        Case "날짜 정리"
            detail = "입력: 240617, 20240617, 2024.06.17, 2024-6-17, 주민번호형 생년월일." & vbCrLf & _
                "출력: " & dateFormat & ". 두 자리 연도는 현재 연도를 넘으면 이전 세기이며, 주민번호는 7번째 숫자로 세기를 구분합니다." & vbCrLf & _
                "예: 240617 → 2024-06-17. 주민번호 뒷자리가 전부 가려졌으면 YYMMDD 기준을 사용합니다. 존재하지 않는 날짜는 빈칸으로 둡니다."
        Case "표 구조 정리"
            detail = "대상: 병합 제목이 있는 보고서형 표. 병합 대표값을 각 셀에 펼칩니다." & vbCrLf & _
                "처리: 새 시트 결과는 병합 없이 생성하고 텍스트 공백을 정리합니다." & vbCrLf & _
                "주의: 빈 셀을 위 행 값으로 임의 보충하지 않습니다. 원본 변경은 병합 해제가 아니므로 새 시트를 권장합니다."
    End Select
    If outputMode = "원본 변경" Then
        detail = detail & " 선택 범위를 직접 변경하며 실패 시 되돌립니다."
    Else
        detail = detail & " 원본을 보존하고 " & outputMode & "에 결과를 만듭니다."
    End If
    NxDataNormalizationDetail = detail
End Function

Public Function NxDataNormalizationPreview(ByVal source As Range, ByVal preset As String, ByVal dateFormat As String, _
    Optional ByVal outputMode As String = "새 시트 만들기") As String
    Dim request As New CNxDataNormalizeRequest
    Dim matrix As Variant, raw As Variant, beforeValue As Variant, afterValue As Variant
    Dim rowIndex As Long, columnIndex As Long, sampleCount As Long
    Dim changedCount As Long, formulaCount As Long, errorCount As Long, blankCount As Long
    Dim rowCount As Long, columnCount As Long
    Dim samples As String, cell As Range, isChanged As Boolean
    Dim formulaState As Variant, cellHasFormula As Boolean
    Set source = NxDataContentRange(source)
    source.Worksheet.Activate: source.Select
    request.Configure source, preset, dateFormat, outputMode
    raw = source.Value2
    formulaState = source.HasFormula
    rowCount = source.Rows.Count: columnCount = source.Columns.Count
    matrix = BuildNormalizationMatrix(source, request.Preset)
    For rowIndex = 1 To rowCount
        For columnIndex = 1 To columnCount
            beforeValue = MatrixValue(raw, rowIndex, columnIndex, rowCount, columnCount)
            afterValue = matrix(rowIndex, columnIndex)
            If NxDataValueIsBlank(beforeValue) Then
                blankCount = blankCount + 1
                GoTo NextPreviewCell
            End If
            isChanged = (NxDataKey(beforeValue, False, True) <> NxDataKey(afterValue, False, True))
            If IsNull(formulaState) Then
                cellHasFormula = CBool(source.Cells(rowIndex, columnIndex).HasFormula)
            Else
                cellHasFormula = CBool(formulaState)
            End If
            If cellHasFormula Then formulaCount = formulaCount + 1: isChanged = True
            If isChanged Then changedCount = changedCount + 1
            If NxDataValueIsBlank(afterValue) Then errorCount = errorCount + 1
            If sampleCount < 10 Then
                Set cell = source.Cells(rowIndex, columnIndex)
                sampleCount = sampleCount + 1
                samples = samples & vbCrLf & cell.Address(False, False) & ": " & _
                    NormalizationPreviewText(beforeValue) & " → " & NormalizationPreviewText(afterValue, request.DateFormat) & _
                    IIf(cellHasFormula, " (수식→값)", vbNullString)
            End If
NextPreviewCell:
        Next columnIndex
        If rowIndex Mod 128 = 0 Then DoEvents
    Next rowIndex
    NxDataNormalizationPreview = "원본: " & source.Worksheet.Name & "!" & source.Address(True, True, xlA1, False) & vbCrLf & _
        "정리 방식: " & request.Preset & " · 결과 위치: " & request.OutputMode & vbCrLf & _
        "값 변경 " & changedCount & "개 · 빈 셀 제외 " & blankCount & "개 · 빈칸 처리 " & errorCount & "개" & vbCrLf & _
        "수식→값 " & formulaCount & "개 · 표시 형식: " & NormalizationNumberFormat(request.Preset, request.DateFormat) & vbCrLf & _
        "앞 " & sampleCount & "개 셀의 전 → 후 (표시 형식은 선택한 방식으로 적용)" & samples
End Function

Private Function NormalizationPreviewText(ByVal value As Variant, Optional ByVal dateFormat As String = "yyyy-mm-dd") As String
    Dim rendered As String
    If IsError(value) Then
        rendered = "[셀 오류]"
    ElseIf IsEmpty(value) Or IsNull(value) Then
        rendered = "[빈 셀]"
    ElseIf VarType(value) = vbDate Then
        rendered = Format$(CDate(value), dateFormat)
    ElseIf VarType(value) = vbString Then
        rendered = Chr$(34) & Replace(Replace(Replace(CStr(value), vbCr, "↵"), vbLf, "↵"), vbTab, "⇥") & Chr$(34)
    Else
        rendered = CStr(value)
    End If
    If Len(rendered) > 70 Then rendered = Left$(rendered, 67) & "..."
    NormalizationPreviewText = rendered
End Function

Public Function NxDataRunNormalizationForTest(ByVal source As Range, ByVal preset As String, ByVal dateFormat As String, _
    Optional ByVal outputMode As String = "새 시트 만들기") As CNxResult
    Dim request As New CNxDataNormalizeRequest
    Dim commandObject As New CNxDataNormalizeCommand
    Dim command As INxFeatureCommand
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim result As CNxResult, failure As Long, detail As String
    On Error GoTo Failed
    request.Configure source, preset, dateFormat, outputMode
    commandObject.Configure request
    Set command = commandObject
    Set ticket = router.Prepare(NxDataFeatureDefinition("NX-DATA-NORMALIZE"), command)
    Select Case ticket.Decision.ResolvedGrade
        Case NxExecutionFast
            Set result = router.ExecuteFast(ticket)
        Case NxExecutionGuarded
            Set result = router.ExecuteGuarded(ticket, True)
        Case NxExecutionPlanned
            Set result = NxDataRunPlannedFrameForTest(command, ticket, router, _
                "정규화 결과를 " & outputMode & " 방식으로 적용합니다.")
        Case Else
            NxRaiseContractError "Data normalization execution grade is invalid"
    End Select
    commandObject.CompleteOutput result
    Set NxDataRunNormalizationForTest = result
    Exit Function
Failed:
    failure = Err.Number: detail = Err.Description: Err.Clear
    On Error Resume Next
    commandObject.DiscardOutput
    If Err.Number <> 0 Then detail = detail & " | 정규화 결과 정리 실패: " & Err.Description
    On Error GoTo 0
    Err.Raise failure, "NxDataRunNormalizationForTest", detail
End Function

' Excel.Application.Run cannot marshal a private VBA class back to an external
' automation client. Keep the typed result function for in-project tests and
' expose a no-return wrapper for native PowerShell/COM acceptance.
Public Sub NxDataRunNormalizationForNativeTest(ByVal source As Range, ByVal preset As String, _
    ByVal dateFormat As String, Optional ByVal outputMode As String = "새 시트 만들기")
    Dim result As CNxResult
    Set result = NxDataRunNormalizationForTest(source, preset, dateFormat, outputMode)
    If result Is Nothing Then NxRaiseContractError "Data normalization returned no result"
End Sub

Public Function NxDataCreateNormalizationResult(ByVal request As CNxDataNormalizeRequest, _
    ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan, ByVal destination As CNxDataOutput) As CNxResult
    Dim source As Range, output As Range, matrix As Variant
    Dim failure As Long, detail As String
    On Error GoTo Failed
    Set source = request.SourceRange
    If context Is Nothing Or approvedPlan Is Nothing Then NxRaiseContractError "실행 계획을 확인하세요."
    matrix = BuildNormalizationMatrix(source, request.Preset)
    If request.OutputMode = "원본 변경" Then
        ApplyNormalizationToSource source, matrix, request.Preset, request.DateFormat, context, approvedPlan
    Else
        Set output = destination.CreateOutput("NX-DATA-NORMALIZE", approvedPlan)
        PrepareNormalizationLiteralFormats output, matrix, request.Preset, request.DateFormat
        output.Value = matrix
        ApplyNormalizationFormat output, request.Preset, request.DateFormat
        If output.Columns.Count <= 50 And output.Rows.Count <= 10000 Then output.Columns.AutoFit
    End If
    Set NxDataCreateNormalizationResult = NxCreateResult("NX-DATA-NORMALIZE", NxSuccess, "complete", _
        context.WorkbookIdentity, request.OutputMode <> "새 통합문서", vbNullString, "result_success")
    Exit Function
Failed:
    failure = Err.Number: detail = Err.Description: Err.Clear
    destination.Rollback
    Err.Raise failure, "NxDataCreateNormalizationResult", detail
End Function

Private Sub ApplyNormalizationToSource(ByVal source As Range, ByRef matrix As Variant, _
    ByVal preset As String, ByVal dateFormat As String, ByVal context As CNxExecutionContext, _
    ByVal approvedPlan As CNxPlan)

    Dim addresses As Variant
    Dim values As Variant
    BuildNormalizationVectors source, matrix, addresses, values
    NxApplyApprovedCellNumberFormats approvedPlan, context, 0, 0, source.Rows.Count, source.Columns.Count, _
        NormalizationNumberFormat(preset, dateFormat)
    NxApplyApprovedCellValueVector approvedPlan, context, addresses, values
End Sub

Private Sub BuildNormalizationVectors(ByVal source As Range, ByRef matrix As Variant, _
    ByRef addresses As Variant, ByRef values As Variant)

    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim itemIndex As Long
    Dim rowCount As Long, columnCount As Long
    rowCount = source.Rows.Count: columnCount = source.Columns.Count
    ReDim addresses(0 To CLng(source.CountLarge) - 1)
    ReDim values(0 To CLng(source.CountLarge) - 1)
    For rowIndex = 1 To rowCount
        For columnIndex = 1 To columnCount
            addresses(itemIndex) = source.Cells(rowIndex, columnIndex).Address(True, True, xlA1, False)
            values(itemIndex) = matrix(rowIndex, columnIndex)
            itemIndex = itemIndex + 1
        Next columnIndex
    Next rowIndex
End Sub

Private Function NormalizationNumberFormat(ByVal preset As String, ByVal dateFormat As String) As String
    Select Case preset
        Case "텍스트 변환": NormalizationNumberFormat = "@"
        Case "날짜 정리": NormalizationNumberFormat = dateFormat
        Case Else: NormalizationNumberFormat = "General"
    End Select
End Function

Public Function NxDataNextNormalizationSheetName(ByVal workbook As Workbook, ByVal baseName As String) As String
    Dim suffix As Long
    Dim candidate As String
    candidate = baseName
    suffix = 1
    Do While NxDataNormalizationSheetExists(workbook, candidate)
        suffix = suffix + 1
        candidate = baseName & " (" & CStr(suffix) & ")"
    Loop
    NxDataNextNormalizationSheetName = candidate
End Function

Private Function BuildNormalizationMatrix(ByVal source As Range, ByVal preset As String) As Variant
    Dim raw As Variant
    Dim matrix As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim value As Variant
    Dim rowCount As Long, columnCount As Long
    Dim mergeState As Variant, inspectMerges As Boolean
    rowCount = source.Rows.Count: columnCount = source.Columns.Count
    If preset = "표 구조 정리" Then
        mergeState = source.MergeCells
        If IsNull(mergeState) Then inspectMerges = True Else inspectMerges = CBool(mergeState)
    End If
    raw = source.Value2
    ReDim matrix(1 To rowCount, 1 To columnCount)
    For rowIndex = 1 To rowCount
        For columnIndex = 1 To columnCount
            value = MatrixValue(raw, rowIndex, columnIndex, rowCount, columnCount)
            If inspectMerges Then
                value = TableStructureValue(source.Cells(rowIndex, columnIndex), value)
            End If
            matrix(rowIndex, columnIndex) = NormalizeValue(value, preset)
            If ((rowIndex - 1) * columnCount + columnIndex) Mod 1024 = 0 Then DoEvents
        Next columnIndex
    Next rowIndex
    BuildNormalizationMatrix = matrix
End Function

Private Function MatrixValue(ByRef raw As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long, _
    ByVal rowCount As Long, ByVal columnCount As Long) As Variant
    If rowCount = 1 And columnCount = 1 Then
        MatrixValue = raw
    Else
        MatrixValue = raw(rowIndex, columnIndex)
    End If
End Function

Private Function TableStructureValue(ByVal sourceCell As Range, ByVal fallbackValue As Variant) As Variant
    If CBool(sourceCell.MergeCells) Then
        TableStructureValue = sourceCell.MergeArea.Cells(1, 1).Value2
    Else
        TableStructureValue = fallbackValue
    End If
End Function

Private Function NormalizeValue(ByVal value As Variant, ByVal preset As String) As Variant
    If IsError(value) Then
        NormalizeValue = vbNullString
        Exit Function
    End If
    If NxDataValueIsBlank(value) Then NormalizeValue = vbNullString: Exit Function
    Select Case preset
        Case "기본 정리", "표 구조 정리"
            If VarType(value) = vbString Then NormalizeValue = NormalizeText(CStr(value)) Else NormalizeValue = value
        Case "숫자 변환"
            NormalizeValue = NormalizeNumber(value)
        Case "텍스트 변환"
            If IsEmpty(value) Then NormalizeValue = vbNullString Else NormalizeValue = CStr(value)
        Case "날짜 정리"
            NormalizeValue = NormalizeDate(value)
        Case Else
            NxRaiseContractError "Data normalization preset is outside the closed contract"
    End Select
End Function

Private Function NormalizeText(ByVal value As String) As String
    Dim characterCode As Long
    value = Replace(value, ChrW$(160), " ")
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, vbTab, " ")
    For characterCode = 0 To 31
        value = Replace(value, Chr$(characterCode), " ")
    Next characterCode
    value = Replace(value, ChrW$(&H2010), "-")
    value = Replace(value, ChrW$(&H2011), "-")
    value = Replace(value, ChrW$(&H2012), "-")
    value = Replace(value, ChrW$(&H2013), "-")
    value = Replace(value, ChrW$(&H2014), "-")
    value = Replace(value, ChrW$(&H2212), "-")
    Do While InStr(1, value, "  ", vbBinaryCompare) > 0
        value = Replace(value, "  ", " ")
    Loop
    NormalizeText = Trim$(value)
End Function

Private Function NormalizeNumber(ByVal value As Variant) As Variant
    Dim text As String
    Dim negative As Boolean
    Dim percent As Boolean
    If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then
        NormalizeNumber = vbNullString
        Exit Function
    End If
    If IsNumeric(value) Then
        NormalizeNumber = CDbl(value)
        Exit Function
    End If
    text = NormalizeText(CStr(value))
    negative = (Len(text) >= 2 And Left$(text, 1) = "(" And Right$(text, 1) = ")")
    If negative Then text = Mid$(text, 2, Len(text) - 2)
    percent = (Right$(text, 1) = "%")
    If percent Then text = Left$(text, Len(text) - 1)
    text = Replace(text, ",", vbNullString)
    text = Replace(text, "원", vbNullString)
    text = Replace(text, "₩", vbNullString)
    text = Replace(text, "￦", vbNullString)
    text = Replace(text, " ", vbNullString)
    If Not IsNumeric(text) Then
        NormalizeNumber = vbNullString
        Exit Function
    End If
    NormalizeNumber = CDbl(text)
    If negative Then NormalizeNumber = -CDbl(NormalizeNumber)
    If percent Then NormalizeNumber = CDbl(NormalizeNumber) / 100#
End Function

Private Function NormalizeDate(ByVal value As Variant) As Variant
    Dim parsed As Date
    If IsEmpty(value) Or Len(Trim$(CStr(value))) = 0 Then
        NormalizeDate = vbNullString
        Exit Function
    End If
    If NxDateTryParseValue(value, Date, parsed) Then
        NormalizeDate = parsed
        Exit Function
    End If
    NormalizeDate = vbNullString
End Function

Private Sub PrepareNormalizationLiteralFormats(ByVal output As Range, ByRef matrix As Variant, _
    ByVal preset As String, ByVal dateFormat As String)
    Dim rowIndex As Long, columnIndex As Long, rowCount As Long, columnCount As Long, runStart As Long
    ApplyNormalizationFormat output, preset, dateFormat
    If preset = "텍스트 변환" Then Exit Sub
    rowCount = output.Rows.Count: columnCount = output.Columns.Count
    ' A Date written into @ becomes locale-dependent text. Keep date/number cells
    ' typed, and use text format only for strings that must never execute as formulas.
    For rowIndex = 1 To rowCount
        runStart = 0
        For columnIndex = 1 To columnCount
            If VarType(matrix(rowIndex, columnIndex)) = vbString Then
                If runStart = 0 Then runStart = columnIndex
            ElseIf runStart > 0 Then
                output.Cells(rowIndex, runStart).Resize(1, columnIndex - runStart).NumberFormat = "@"
                runStart = 0
            End If
        Next columnIndex
        If runStart > 0 Then output.Cells(rowIndex, runStart).Resize(1, columnCount - runStart + 1).NumberFormat = "@"
    Next rowIndex
End Sub

Private Sub ApplyNormalizationFormat(ByVal output As Range, ByVal preset As String, ByVal dateFormat As String)
    Select Case preset
        Case "텍스트 변환"
            output.NumberFormat = "@"
        Case "날짜 정리"
            output.NumberFormat = dateFormat
        Case Else
            output.NumberFormat = "General"
    End Select
End Sub

Private Function NxDataNormalizationSheetExists(ByVal workbook As Workbook, ByVal sheetName As String) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = workbook.Worksheets(sheetName)
    NxDataNormalizationSheetExists = Not sheet Is Nothing
    Err.Clear
    On Error GoTo 0
End Function

Private Function SafeNormalizationTarget(ByVal request As CNxDataNormalizeRequest) As String
    On Error Resume Next
    SafeNormalizationTarget = request.SourceRange.Parent.Parent.Name
    If Len(SafeNormalizationTarget) = 0 Then SafeNormalizationTarget = "data_normalization"
    Err.Clear
    On Error GoTo 0
End Function
