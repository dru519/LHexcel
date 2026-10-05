Attribute VB_Name = "NxHwpxSerializer"
Option Explicit

Private Const NX_HWPX_MAX_CELLS As Long = 16384

Public Function NxHwpxSerialize(ByVal source As Range, ByVal titleRow As Boolean, _
    ByVal includeHidden As Boolean, ByVal sourceMode As String) As String

    Dim rowMap As Object, rowIndex As Long, outputRow As Long
    Dim columnIndex As Long, cell As Range, rowsJson As String
    ValidateSource source, sourceMode
    Set rowMap = CreateObject("Scripting.Dictionary")
    For rowIndex = 1 To source.Rows.Count
        If includeHidden Or Not CBool(source.Rows(rowIndex).EntireRow.Hidden) Then
            outputRow = outputRow + 1
            rowMap.Add CStr(rowIndex), outputRow
        End If
    Next rowIndex
    If outputRow = 0 Then NxRaiseContractError "전송할 표시 행이 없습니다."
    If CDbl(outputRow) * CDbl(source.Columns.Count) > NX_HWPX_MAX_CELLS Then NxRaiseContractError "아래한글 표는 최대 50,000셀입니다."

    Dim firstCell As Boolean: firstCell = True
    For rowIndex = 1 To source.Rows.Count
        If rowMap.Exists(CStr(rowIndex)) Then
            For columnIndex = 1 To source.Columns.Count
                Set cell = source.Cells(rowIndex, columnIndex)
                If Not firstCell Then rowsJson = rowsJson & ","
                rowsJson = rowsJson & SerializeCell(cell, CLng(rowMap(CStr(rowIndex))), columnIndex, sourceMode)
                firstCell = False
            Next columnIndex
        End If
    Next rowIndex

    NxHwpxSerialize = "{""schema_version"":1,""rows"":" & CStr(outputRow) & _
        ",""columns"":" & CStr(source.Columns.Count) & ",""title_row"":" & JsonBoolean(titleRow) & _
        ",""source_mode"":""" & sourceMode & """,""cells"":[" & rowsJson & "]" & _
        ",""merges"":" & SerializeMerges(source, rowMap) & _
        ",""column_widths"":" & SerializeColumnWidths(source) & _
        ",""row_heights"":" & SerializeRowHeights(source, rowMap) & "}"
End Function

Private Sub ValidateSource(ByVal source As Range, ByVal sourceMode As String)
    If source Is Nothing Then NxRaiseContractError "아래한글로 보낼 범위를 선택하세요."
    If source.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 전송할 수 있습니다."
    If CDbl(source.CountLarge) > NX_HWPX_MAX_CELLS Then NxRaiseContractError "아래한글 표는 최대 16,384셀입니다."
    If sourceMode <> "display" And sourceMode <> "value" Then NxRaiseContractError "전송 값 형식이 잘못되었습니다."
End Sub

Private Function SerializeCell(ByVal cell As Range, ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal sourceMode As String) As String
    Dim text As String, fillRgb As String, fontRgb As String
    If sourceMode = "display" Then
        text = CStr(cell.Text)
    ElseIf IsError(cell.Value2) Then
        text = CStr(cell.Text)
    ElseIf IsEmpty(cell.Value2) Then
        text = vbNullString
    Else
        text = CStr(cell.Value2)
    End If
    fillRgb = ColorToRgb(cell.Interior.Color, "FFFFFF")
    fontRgb = ColorToRgb(cell.Font.Color, "000000")
    SerializeCell = "{""row"":" & CStr(rowIndex) & ",""column"":" & CStr(columnIndex) & _
        ",""text"":""" & JsonEscape(text) & """,""fill_rgb"":""" & fillRgb & _
        """,""font_bold"":" & JsonBoolean(CBool(cell.Font.Bold)) & ",""font_rgb"":""" & fontRgb & _
        """,""horizontal"":""" & HorizontalWire(cell.HorizontalAlignment) & _
        """,""vertical"":""" & VerticalWire(cell.VerticalAlignment) & _
        """,""borders"":""" & BorderWire(cell) & """}"
End Function

Private Function SerializeMerges(ByVal source As Range, ByVal rowMap As Object) As String
    Dim seen As Object, cell As Range, merged As Range, overlap As Range, key As String, output As String
    Dim relativeRow As Long, relativeColumn As Long, rowOffset As Long
    Set seen = CreateObject("Scripting.Dictionary")
    For Each cell In source.Cells
        If cell.MergeCells Then
            Set merged = cell.MergeArea
            key = merged.Address(True, True, xlA1, True)
            If Not seen.Exists(key) Then
                seen.Add key, True
                Set overlap = Application.Intersect(merged, source)
                If overlap Is Nothing Then NxRaiseContractError "선택 범위 경계를 넘는 병합 셀은 전송할 수 없습니다."
                If overlap.Address <> merged.Address Then NxRaiseContractError "선택 범위 경계를 넘는 병합 셀은 전송할 수 없습니다."
                relativeRow = merged.Row - source.Row + 1
                relativeColumn = merged.Column - source.Column + 1
                For rowOffset = 0 To merged.Rows.Count - 1
                    If Not rowMap.Exists(CStr(relativeRow + rowOffset)) Then NxRaiseContractError "숨김 행과 겹치는 병합 셀은 숨김 행 포함으로 다시 실행하세요."
                Next rowOffset
                If Len(output) > 0 Then output = output & ","
                output = output & "{""row"":" & CStr(rowMap(CStr(relativeRow))) & ",""column"":" & CStr(relativeColumn) & _
                    ",""row_span"":" & CStr(merged.Rows.Count) & ",""column_span"":" & CStr(merged.Columns.Count) & "}"
            End If
        End If
    Next cell
    SerializeMerges = "[" & output & "]"
End Function

Private Function SerializeColumnWidths(ByVal source As Range) As String
    Dim index As Long, output As String
    For index = 1 To source.Columns.Count
        If index > 1 Then output = output & ","
        output = output & InvariantNumber(CDbl(source.Columns(index).ColumnWidth))
    Next index
    SerializeColumnWidths = "[" & output & "]"
End Function

Private Function SerializeRowHeights(ByVal source As Range, ByVal rowMap As Object) As String
    Dim index As Long, output As String
    For index = 1 To source.Rows.Count
        If rowMap.Exists(CStr(index)) Then
            If Len(output) > 0 Then output = output & ","
            output = output & InvariantNumber(CDbl(source.Rows(index).RowHeight))
        End If
    Next index
    SerializeRowHeights = "[" & output & "]"
End Function

Private Function HorizontalWire(ByVal alignment As Variant) As String
    If IsNull(alignment) Then HorizontalWire = "left": Exit Function
    Select Case CLng(alignment)
        Case xlCenter, xlCenterAcrossSelection: HorizontalWire = "center"
        Case xlRight: HorizontalWire = "right"
        Case xlJustify, xlDistributed: HorizontalWire = "justify"
        Case Else: HorizontalWire = "left"
    End Select
End Function

Private Function VerticalWire(ByVal alignment As Variant) As String
    If IsNull(alignment) Then VerticalWire = "center": Exit Function
    Select Case CLng(alignment)
        Case xlTop: VerticalWire = "top"
        Case xlBottom: VerticalWire = "bottom"
        Case Else: VerticalWire = "center"
    End Select
End Function

Private Function BorderWire(ByVal cell As Range) As String
    Dim side As Variant
    For Each side In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
        If cell.Borders(CLng(side)).LineStyle <> xlLineStyleNone Then BorderWire = "all": Exit Function
    Next side
    BorderWire = "none"
End Function

Private Function ColorToRgb(ByVal colorValue As Variant, ByVal fallback As String) As String
    If IsNull(colorValue) Or Not IsNumeric(colorValue) Then ColorToRgb = fallback: Exit Function
    Dim value As Long: value = CLng(colorValue)
    If value < 0 Then ColorToRgb = fallback: Exit Function
    ColorToRgb = Right$("0" & Hex$(value Mod 256), 2) & Right$("0" & Hex$((value \ 256) Mod 256), 2) & Right$("0" & Hex$((value \ 65536) Mod 256), 2)
End Function

Private Function InvariantNumber(ByVal value As Double) As String
    InvariantNumber = Replace$(Trim$(CStr(value)), Application.International(xlDecimalSeparator), ".")
End Function

Private Function JsonBoolean(ByVal value As Boolean) As String
    If value Then JsonBoolean = "true" Else JsonBoolean = "false"
End Function

Private Function JsonEscape(ByVal value As String) As String
    Dim index As Long, character As String, code As Long
    For index = 1 To Len(value)
        character = Mid$(value, index, 1): code = AscW(character)
        Select Case character
            Case Chr$(34): JsonEscape = JsonEscape & "\" & Chr$(34)
            Case "\": JsonEscape = JsonEscape & "\\"
            Case vbCr: JsonEscape = JsonEscape & "\r"
            Case vbLf: JsonEscape = JsonEscape & "\n"
            Case vbTab: JsonEscape = JsonEscape & "\t"
            Case Else
                If code >= 0 And code < 32 Then JsonEscape = JsonEscape & "\u" & Right$("0000" & Hex$(code), 4) Else JsonEscape = JsonEscape & character
        End Select
    Next index
End Function


