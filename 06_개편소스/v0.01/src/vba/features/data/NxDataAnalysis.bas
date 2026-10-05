Attribute VB_Name = "NxDataAnalysis"
Option Explicit

Public Function NxDataAnalysisRangeMatrix(ByVal source As Range, ByVal formulas As Boolean) As Variant
    Dim values As Variant, singleCellValues(1 To 1, 1 To 1) As Variant
    If formulas Then values = source.Formula Else values = source.Value2
    If source.CountLarge = 1 Then singleCellValues(1, 1) = values: NxDataAnalysisRangeMatrix = singleCellValues Else NxDataAnalysisRangeMatrix = values
End Function

Public Function NxDataAnalysisMatrixMatches(ByVal left As Variant, ByVal right As Variant) As Boolean
    Dim rowIndex As Long, columnIndex As Long
    If UBound(left, 1) <> UBound(right, 1) Or UBound(left, 2) <> UBound(right, 2) Then Exit Function
    For rowIndex = 1 To UBound(left, 1)
        For columnIndex = 1 To UBound(left, 2)
            If NxDataKey(left(rowIndex, columnIndex), False, True) <> NxDataKey(right(rowIndex, columnIndex), False, True) Then Exit Function
        Next columnIndex
    Next rowIndex
    NxDataAnalysisMatrixMatches = True
End Function

Public Function NxDataAnalysisHiddenAxes(ByVal source As Range, ByVal rowDirection As Boolean) As String
    Dim index As Long, flags As String, count As Long, hidden As Boolean
    If rowDirection Then count = source.Columns.Count Else count = source.Rows.Count
    flags = String$(count, "0")
    For index = 1 To count
        If rowDirection Then hidden = CBool(source.Columns(index).EntireColumn.Hidden) Else hidden = CBool(source.Rows(index).EntireRow.Hidden)
        If hidden Then Mid$(flags, index, 1) = "1"
    Next index
    NxDataAnalysisHiddenAxes = flags
End Function

Public Function NxDataAnalysisCleanText(ByVal value As Variant, ByVal cleanSpaces As Boolean) As String
    Dim text As String, index As Long, code As Long, character As String, cleaned As String
    If IsError(value) Then NxDataAnalysisCleanText = "error": Exit Function
    text = CStr(value)
    If Not cleanSpaces Then NxDataAnalysisCleanText = text: Exit Function
    text = Replace$(text, ChrW$(&H2013), "-")
    text = Replace$(text, ChrW$(&H2014), "-")
    text = Replace$(text, ChrW$(&HFF0D), "-")
    text = Replace$(text, ChrW$(&H2011), "-")
    For index = 1 To Len(text)
        character = Mid$(text, index, 1): code = AscW(character)
        If code < 0 Then code = code + 65536
        If code = 10 Or code = 13 Then
            cleaned = cleaned & " "
        ElseIf code >= 32 Then
            cleaned = cleaned & character
        End If
    Next index
    cleaned = Trim$(cleaned)
    Do While InStr(1, cleaned, "  ", vbBinaryCompare) > 0
        cleaned = Replace$(cleaned, "  ", " ")
    Loop
    NxDataAnalysisCleanText = cleaned
End Function

Public Function NxDataAnalysisValues(ByVal request As CNxDataRequest) As Variant
    If request Is Nothing Then NxRaiseContractError "분석 요청이 필요합니다."
    request.RequireFreshAnalysis
    If request.FeatureId = "NX-DATA-UNIQUE-COUNT" Then
        NxDataAnalysisValues = NxDataAnalysisUniqueMatrix(request.SnapshotValues, request.RowDirection, _
            request.WholeRecord, request.HasTitleRow, request.SkipBlank, request.TrimText, request.IncludeRatio, _
            request.SortMode, request.HiddenAxes, request.IncludeHidden)
    Else
        NxDataAnalysisValues = NxDataAnalysisCompareMatrix(request.SnapshotValues, request.RowDirection, _
            request.HasTitleRow, request.KeyColumn, request.CompareColumn, Not request.CaseSensitive, _
            request.DenseCompare, request.HiddenAxes, request.IncludeHidden)
    End If
End Function

Public Function NxDataAnalysisUniqueMatrix(ByVal values As Variant, ByVal rowDirection As Boolean, _
    ByVal wholeRecord As Boolean, ByVal hasHeader As Boolean, ByVal skipBlank As Boolean, _
    ByVal cleanSpaces As Boolean, ByVal includeRatio As Boolean, ByVal sortMode As String, _
    Optional ByVal hiddenAxes As String = "", Optional ByVal includeHidden As Boolean = True) As Variant
    Dim itemCount As Long, width As Long, axisCount As Long, firstData As Long
    Dim item As Long, part As Long, groupCount As Long, groupIndex As Long, totalCount As Long
    Dim key As String, valueText As String, sortText As String, hasValue As Boolean
    Dim byKey As Object, firstItems() As Long, counts() As Long, sortTexts() As String, order() As Long
    Dim result As Variant, outputColumns As Long, outputRow As Long, sourceItem As Long
    AnalysisDimensions values, rowDirection, itemCount, axisCount
    width = IIf(wholeRecord, axisCount, 1): firstData = IIf(hasHeader, 2, 1)
    If firstData > itemCount Then NxRaiseContractError "집계할 데이터가 없습니다."
    If sortMode <> "값 오름차순" And sortMode <> "값 내림차순" And sortMode <> "원본 순서" Then NxRaiseContractError "집계 정렬 방식을 확인하세요."
    ReDim firstItems(1 To itemCount): ReDim counts(1 To itemCount)
    ReDim sortTexts(1 To itemCount): ReDim order(1 To itemCount)
    Set byKey = CreateObject("Scripting.Dictionary")
    byKey.CompareMode = vbBinaryCompare
    For item = firstData To itemCount
        If AnalysisIncluded(item, hiddenAxes, includeHidden) Then
            key = "": sortText = "": hasValue = False
            For part = 1 To width
                valueText = NxDataAnalysisCleanText(AnalysisValueAt(values, item, part, rowDirection), cleanSpaces)
                If Len(Trim$(valueText)) > 0 Then hasValue = True
                ' Length framing avoids collisions with real "(빈값)" text or embedded separators.
                key = key & CStr(Len(valueText)) & ":" & valueText
                If Len(valueText) = 0 Then valueText = "(빈값)"
                If part > 1 Then sortText = sortText & Chr$(30)
                sortText = sortText & valueText
            Next part
            If hasValue Or Not skipBlank Then
                If byKey.Exists(key) Then
                    groupIndex = CLng(byKey(key))
                Else
                    groupCount = groupCount + 1: groupIndex = groupCount
                    byKey.Add key, groupIndex: firstItems(groupIndex) = item
                    sortTexts(groupIndex) = sortText: order(groupIndex) = groupIndex
                End If
                counts(groupIndex) = counts(groupIndex) + 1: totalCount = totalCount + 1
            End If
        End If
    Next item
    If groupCount = 0 Then NxRaiseContractError "집계할 값이 없습니다."
    outputColumns = width + 1 + IIf(includeRatio, 1, 0)
    AnalysisRequireOutputSize groupCount + 1, outputColumns
    If sortMode <> "원본 순서" And groupCount > 1 Then AnalysisSort order, sortTexts, 1, groupCount, (sortMode = "값 내림차순")
    ReDim result(1 To groupCount + 1, 1 To outputColumns)
    For part = 1 To width
        valueText = ""
        If hasHeader Then valueText = NxDataAnalysisCleanText(AnalysisValueAt(values, 1, part, rowDirection), False)
        If Len(Trim$(valueText)) = 0 Then
            If axisCount = 1 Then valueText = "고유값" Else valueText = "값" & CStr(part)
        End If
        result(1, part) = valueText
    Next part
    result(1, width + 1) = "건수"
    If includeRatio Then result(1, width + 2) = "비율"
    For outputRow = 1 To groupCount
        groupIndex = order(outputRow): sourceItem = firstItems(groupIndex)
        For part = 1 To width
            valueText = NxDataAnalysisCleanText(AnalysisValueAt(values, sourceItem, part, rowDirection), cleanSpaces)
            If Len(valueText) = 0 Then valueText = "(빈값)"
            result(outputRow + 1, part) = valueText
        Next part
        result(outputRow + 1, width + 1) = counts(groupIndex)
        If includeRatio Then result(outputRow + 1, width + 2) = CDbl(counts(groupIndex)) / CDbl(totalCount)
    Next outputRow
    NxDataAnalysisUniqueMatrix = result
End Function

Public Function NxDataAnalysisCompareMatrix(ByVal values As Variant, ByVal rowDirection As Boolean, _
    ByVal hasHeader As Boolean, ByVal leftIndex As Long, ByVal rightIndex As Long, _
    ByVal ignoreCase As Boolean, ByVal denseCompare As Boolean, _
    Optional ByVal hiddenAxes As String = "", Optional ByVal includeHidden As Boolean = True) As Variant
    Dim itemCount As Long, axisCount As Long, item As Long, firstData As Long
    Dim leftValues As Object, rightValues As Object, leftKey As String, rightKey As String
    Dim result As Variant, leftName As String, rightName As String, leftStatus As String, rightStatus As String
    AnalysisDimensions values, rowDirection, itemCount, axisCount
    If leftIndex < 1 Or rightIndex < 1 Or leftIndex > axisCount Or rightIndex > axisCount Then NxRaiseContractError "기준 또는 대조 위치가 범위를 벗어났습니다."
    If leftIndex = rightIndex Then NxRaiseContractError "기준과 대조 위치는 서로 달라야 합니다."
    firstData = IIf(hasHeader, 2, 1)
    If firstData > itemCount Then NxRaiseContractError "대조할 데이터가 없습니다."
    Set leftValues = CreateObject("Scripting.Dictionary"): Set rightValues = CreateObject("Scripting.Dictionary")
    leftValues.CompareMode = vbBinaryCompare: rightValues.CompareMode = vbBinaryCompare
    For item = firstData To itemCount
        If AnalysisIncluded(item, hiddenAxes, includeHidden) Then
            leftKey = AnalysisCompareKey(AnalysisValueAt(values, item, leftIndex, rowDirection), ignoreCase, denseCompare)
            rightKey = AnalysisCompareKey(AnalysisValueAt(values, item, rightIndex, rowDirection), ignoreCase, denseCompare)
            If Len(leftKey) > 0 Then leftValues(leftKey) = True
            If Len(rightKey) > 0 Then rightValues(rightKey) = True
        End If
    Next item
    If rowDirection Then
        AnalysisRequireOutputSize 2, itemCount
        ReDim result(1 To 2, 1 To itemCount)
    Else
        AnalysisRequireOutputSize itemCount, 4
        ReDim result(1 To itemCount, 1 To 4)
        For item = 1 To itemCount
            result(item, 1) = AnalysisValueAt(values, item, leftIndex, False)
            result(item, 2) = AnalysisValueAt(values, item, rightIndex, False)
        Next item
    End If
    If hasHeader Then
        leftName = NxDataAnalysisCleanText(AnalysisValueAt(values, 1, leftIndex, rowDirection), False)
        rightName = NxDataAnalysisCleanText(AnalysisValueAt(values, 1, rightIndex, rowDirection), False)
        If Len(Trim$(leftName)) = 0 Then leftName = "기준"
        If Len(Trim$(rightName)) = 0 Then rightName = "대조"
        If rowDirection Then
            result(1, 1) = leftName & " 기준 " & rightName & " 중복"
            result(2, 1) = rightName & " 기준 " & leftName & " 중복"
        Else
            result(1, 3) = leftName & " 기준 " & rightName & " 중복"
            result(1, 4) = rightName & " 기준 " & leftName & " 중복"
        End If
    End If
    For item = firstData To itemCount
        leftStatus = "": rightStatus = ""
        If AnalysisIncluded(item, hiddenAxes, includeHidden) Then
            leftKey = AnalysisCompareKey(AnalysisValueAt(values, item, leftIndex, rowDirection), ignoreCase, denseCompare)
            rightKey = AnalysisCompareKey(AnalysisValueAt(values, item, rightIndex, rowDirection), ignoreCase, denseCompare)
            If Len(leftKey) > 0 Then leftStatus = IIf(rightValues.Exists(leftKey), "중복", "고유")
            If Len(rightKey) > 0 Then rightStatus = IIf(leftValues.Exists(rightKey), "중복", "고유")
            If IsError(AnalysisValueAt(values, item, leftIndex, rowDirection)) Then leftStatus = "error"
            If IsError(AnalysisValueAt(values, item, rightIndex, rowDirection)) Then rightStatus = "error"
        End If
        If rowDirection Then
            result(1, item) = leftStatus: result(2, item) = rightStatus
        Else
            result(item, 3) = leftStatus: result(item, 4) = rightStatus
        End If
    Next item
    NxDataAnalysisCompareMatrix = result
End Function

Public Function NxDataAnalysisOutputCells(ByVal featureId As String, ByVal values As Variant) As Long
    Dim cells As Double, policy As New CNxLimitsPolicy
    cells = CDbl(UBound(values, 1)) * CDbl(UBound(values, 2))
    If cells < 1 Or cells > CDbl(policy.MaximumCellsFor(featureId)) Then NxRaiseContractError "분석 결과가 안전 셀 한도를 초과합니다."
    NxDataAnalysisOutputCells = CLng(cells)
End Function

Private Sub AnalysisDimensions(ByVal values As Variant, ByVal rowDirection As Boolean, ByRef itemCount As Long, ByRef axisCount As Long)
    If Not IsArray(values) Then NxRaiseContractError "분석 입력은 2차원 배열이어야 합니다."
    If rowDirection Then itemCount = UBound(values, 2): axisCount = UBound(values, 1) Else itemCount = UBound(values, 1): axisCount = UBound(values, 2)
    If LBound(values, 1) <> 1 Or LBound(values, 2) <> 1 Then NxRaiseContractError "분석 배열은 1부터 시작해야 합니다."
    If itemCount < 1 Or axisCount < 1 Or CDbl(itemCount) * CDbl(axisCount) > 200000# Then NxRaiseContractError "분석 입력이 안전 한도를 초과합니다."
End Sub

Private Sub AnalysisRequireOutputSize(ByVal rows As Long, ByVal columns As Long)
    If rows < 1 Or columns < 1 Or rows > 1048576 Or columns > 16384 Then NxRaiseContractError "분석 결과가 Excel 시트 크기를 초과합니다."
    If CDbl(rows) * CDbl(columns) > 200000# Then NxRaiseContractError "분석 결과가 200,000개 셀을 초과합니다."
End Sub

Private Function AnalysisValueAt(ByRef values As Variant, ByVal item As Long, ByVal part As Long, ByVal rowDirection As Boolean) As Variant
    If rowDirection Then AnalysisValueAt = values(part, item) Else AnalysisValueAt = values(item, part)
End Function

Private Function AnalysisIncluded(ByVal item As Long, ByVal hiddenAxes As String, ByVal includeHidden As Boolean) As Boolean
    AnalysisIncluded = includeHidden Or Mid$(hiddenAxes, item, 1) <> "1"
End Function

Private Function AnalysisCompareKey(ByVal value As Variant, ByVal ignoreCase As Boolean, ByVal denseCompare As Boolean) As String
    Dim text As String
    If IsError(value) Then Exit Function
    text = NxDataAnalysisCleanText(value, True)
    If denseCompare Then text = Replace$(Replace$(text, " ", ""), "-", "")
    If ignoreCase Then text = LCase$(text)
    AnalysisCompareKey = text
End Function

Private Sub AnalysisSort(ByRef order() As Long, ByRef texts() As String, ByVal low As Long, ByVal high As Long, ByVal descending As Boolean)
    Dim left As Long, right As Long, pivot As String, swap As Long, direction As Long
    left = low: right = high: pivot = texts(order((low + high) \ 2))
    direction = IIf(descending, -1, 1)
    Do While left <= right
        Do While StrComp(texts(order(left)), pivot, vbTextCompare) * direction < 0: left = left + 1: Loop
        Do While StrComp(texts(order(right)), pivot, vbTextCompare) * direction > 0: right = right - 1: Loop
        If left <= right Then
            swap = order(left): order(left) = order(right): order(right) = swap
            left = left + 1: right = right - 1
        End If
    Loop
    If low < right Then AnalysisSort order, texts, low, right, descending
    If left < high Then AnalysisSort order, texts, left, high, descending
End Sub
