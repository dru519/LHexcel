Attribute VB_Name = "NxAiSerializer"
Option Explicit

Private Const NX_AI_MAX_AREAS As Long = 20
Private Const NX_AI_MAX_CELLS As Long = 50000

Public Function NxAiSerializeRanges(ByVal inputValue As Object) As String
    Dim areas As New Collection
    CollectRanges inputValue, areas

    Dim areaCount As Long
    areaCount = areas.Count
    If areaCount = 0 Then RaiseInputError "선택 범위가 없습니다."
    If areaCount > NX_AI_MAX_AREAS Then RaiseInputError "AI 입력 영역은 최대 20개입니다."

    Dim ordered() As Range
    ReDim ordered(1 To areaCount)
    Dim index As Long
    For index = 1 To areaCount
        Set ordered(index) = areas.Item(index)
    Next index
    SortRanges ordered, areaCount

    Dim ownerBook As Workbook
    Set ownerBook = ordered(1).Worksheet.Parent
    Dim areaBook As Workbook
    Dim cellCount As Double
    For index = 1 To areaCount
        Set areaBook = ordered(index).Worksheet.Parent
        If Not areaBook Is ownerBook Then RaiseInputError "서로 다른 통합문서의 범위는 함께 처리할 수 없습니다."
        cellCount = cellCount + ordered(index).Cells.CountLarge
        If cellCount > NX_AI_MAX_CELLS Then RaiseInputError "AI 입력 셀은 최대 50,000개입니다."
    Next index

    Dim emitted As Object
    Set emitted = CreateObject("Scripting.Dictionary")
    emitted.CompareMode = vbBinaryCompare

    Dim output As String
    For index = 1 To areaCount
        If index > 1 Then output = output & vbLf
        output = output & SerializeArea(ordered(index), index, emitted)
    Next index
    NxAiSerializeRanges = output
End Function

Public Function NxAiValidateSelectionKind(ByVal rangeCount As Long, ByVal pictureCount As Long, ByVal featureId As String) As String
    If rangeCount < 0 Or pictureCount < 0 Then RaiseInputError "선택 개수는 음수일 수 없습니다."
    If rangeCount > 0 And pictureCount > 0 Then RaiseInputError "범위와 그림을 동시에 선택할 수 없습니다."
    If rangeCount > 0 Then
        NxAiValidateSelectionKind = "RANGE"
        Exit Function
    End If
    If pictureCount = 1 Then
        If featureId <> "NX-AI-IMAGE" Then RaiseInputError "단일 그림은 이미지 활용 모드에서만 사용할 수 있습니다."
        NxAiValidateSelectionKind = "PICTURE"
        Exit Function
    End If
    If pictureCount > 1 Then RaiseInputError "그림은 하나만 선택할 수 있습니다."
    RaiseInputError "선택된 범위 또는 그림이 없습니다."
End Function

Private Sub CollectRanges(ByVal inputValue As Object, ByVal target As Collection)
    If inputValue Is Nothing Then RaiseInputError "선택 범위가 없습니다."
    Dim item As Variant
    If TypeOf inputValue Is Excel.Range Then
        For Each item In inputValue.Areas
            target.Add item
        Next item
        Exit Sub
    End If
    If TypeName(inputValue) = "Collection" Then
        For Each item In inputValue
            If Not IsObject(item) Then RaiseInputError "범위가 아닌 입력이 포함되어 있습니다."
            If Not TypeOf item Is Excel.Range Then RaiseInputError "범위가 아닌 입력이 포함되어 있습니다."
            target.Add item
        Next item
        Exit Sub
    End If
    RaiseInputError "Excel 범위만 직렬화할 수 있습니다."
End Sub

Private Sub SortRanges(ByRef values() As Range, ByVal count As Long)
    Dim index As Long, cursor As Long
    Dim current As Range
    For index = 2 To count
        Set current = values(index)
        cursor = index - 1
        Do While cursor >= 1
            If CompareRanges(values(cursor), current) <= 0 Then Exit Do
            Set values(cursor + 1) = values(cursor)
            cursor = cursor - 1
        Loop
        Set values(cursor + 1) = current
    Next index
End Sub

Private Function CompareRanges(ByVal leftValue As Range, ByVal rightValue As Range) As Long
    Dim comparison As Long
    comparison = StrComp(leftValue.Worksheet.CodeName, rightValue.Worksheet.CodeName, vbBinaryCompare)
    If comparison <> 0 Then CompareRanges = comparison: Exit Function
    If leftValue.Row <> rightValue.Row Then CompareRanges = Sgn(leftValue.Row - rightValue.Row): Exit Function
    If leftValue.Column <> rightValue.Column Then CompareRanges = Sgn(leftValue.Column - rightValue.Column): Exit Function
    CompareRanges = StrComp(leftValue.Address(True, True, xlA1, False), rightValue.Address(True, True, xlA1, False), vbBinaryCompare)
End Function

Private Function SerializeArea(ByVal area As Range, ByVal areaNumber As Long, ByVal emitted As Object) As String
    Dim sheet As Worksheet
    Set sheet = area.Worksheet
    Dim output As String
    output = "영역 " & CStr(areaNumber) & "|시트=" & EscapeText(sheet.Name) & "|주소=" & area.Address(True, True, xlA1, False) & "|행=" & CStr(area.Rows.Count) & "|열=" & CStr(area.Columns.Count) & vbLf

    Dim rowIndex As Long, columnIndex As Long
    Dim cell As Range
    Dim key As String, rowText As String, valueText As String
    For rowIndex = 1 To area.Rows.Count
        rowText = vbNullString
        For columnIndex = 1 To area.Columns.Count
            Set cell = area.Cells(rowIndex, columnIndex)
            key = sheet.CodeName & "|" & CStr(cell.Row) & "|" & CStr(cell.Column)
            If emitted.Exists(key) Then
                valueText = "<겹침 제외>"
            Else
                emitted.Add key, True
                valueText = SerializeCell(cell)
            End If
            If columnIndex > 1 Then rowText = rowText & vbTab
            rowText = rowText & valueText
        Next columnIndex
        output = output & rowText
        If rowIndex < area.Rows.Count Then output = output & vbLf
    Next rowIndex
    SerializeArea = output
End Function

Private Function SerializeCell(ByVal cell As Range) As String
    If cell.MergeCells Then
        Dim merged As Range
        Set merged = cell.MergeArea
        If cell.Row <> merged.Row Or cell.Column <> merged.Column Then
            SerializeCell = "<병합 셀>"
        Else
            SerializeCell = SerializeValue(cell.Value2) & "[병합=" & merged.Address(True, True, xlA1, False) & "]"
        End If
        Exit Function
    End If
    SerializeCell = SerializeValue(cell.Value2)
End Function

Private Function SerializeValue(ByVal value As Variant) As String
    If IsError(value) Then
        SerializeValue = "<오류>"
    ElseIf IsEmpty(value) Or Len(CStr(value)) = 0 Then
        SerializeValue = "<빈 셀>"
    Else
        SerializeValue = EscapeText(CStr(value))
    End If
End Function

Private Function EscapeText(ByVal value As String) As String
    value = Replace(value, vbCrLf, "\n")
    value = Replace(value, vbCr, "\n")
    value = Replace(value, vbLf, "\n")
    EscapeText = Replace(value, vbTab, "\t")
End Function

Private Sub RaiseInputError(ByVal message As String)
    Err.Raise vbObjectError + 880, "NxAiSerializer", message
End Sub
