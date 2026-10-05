Attribute VB_Name = "NxCalculatorCell"
Option Explicit

Public Function NxCalculatorTryReadActiveNumber(ByRef numericValue As Double, ByRef statusText As String) As Boolean
    Dim target As Range, raw As Variant
    On Error GoTo Failed
    If Application.ActiveWorkbook Is Nothing Then statusText = "활성 통합문서가 없습니다.": Exit Function
    If TypeName(Application.ActiveCell) <> "Range" Then statusText = "숫자를 가져올 셀을 선택하세요.": Exit Function
    Set target = Application.ActiveCell
    If target.Areas.Count <> 1 Or target.Cells.CountLarge <> 1 Then statusText = "활성 셀 한 개만 사용할 수 있습니다.": Exit Function
    raw = target.Value2
    If IsError(raw) Or IsEmpty(raw) Then statusText = "빈 셀이나 오류 셀은 가져올 수 없습니다.": Exit Function
    If VarType(raw) = vbBoolean Or VarType(raw) = vbString Then statusText = "숫자 결과가 있는 셀만 가져올 수 있습니다.": Exit Function
    If VarType(target.Value) = vbDate Or NxCalculatorIsDateFormatted(CStr(target.NumberFormat)) Then statusText = "날짜·시간 값은 가져올 수 없습니다.": Exit Function
    If Not IsNumeric(raw) Then statusText = "숫자 결과가 있는 셀만 가져올 수 있습니다.": Exit Function
    numericValue = CDbl(raw)
    statusText = "활성 셀의 숫자값을 가져왔습니다."
    NxCalculatorTryReadActiveNumber = True
    Exit Function
Failed:
    statusText = "활성 셀 값을 확인할 수 없습니다."
    NxCalculatorTryReadActiveNumber = False
End Function

Public Function NxCalculatorPrepareActiveWrite(ByVal resultValue As Double) As CNxCalculatorCellWrite
    Dim target As Range, prepared As New CNxCalculatorCellWrite
    If Application.ActiveWorkbook Is Nothing Or TypeName(Application.ActiveCell) <> "Range" Then NxRaiseContractError "결과를 입력할 활성 셀을 선택하세요."
    Set target = Application.ActiveCell
    prepared.Capture target, resultValue
    Set NxCalculatorPrepareActiveWrite = prepared
End Function

Public Function NxCalculatorCellIdentity(ByVal target As Range) As String
    Dim book As Workbook, sheet As Worksheet, bookIdentity As String
    If target Is Nothing Then Exit Function
    If target.Areas.Count <> 1 Or target.Cells.CountLarge <> 1 Then Exit Function
    Set sheet = target.Worksheet
    Set book = sheet.Parent
    If Len(book.Path) > 0 Then bookIdentity = book.FullName Else bookIdentity = book.Name
    NxCalculatorCellIdentity = bookIdentity & "|" & sheet.CodeName & "|" & target.Address(False, False, xlA1)
End Function

Public Function NxCalculatorActiveCellIdentity() As String
    Dim target As Range
    If Application.ActiveWorkbook Is Nothing Or TypeName(Application.ActiveCell) <> "Range" Then Exit Function
    Set target = Application.ActiveCell
    NxCalculatorActiveCellIdentity = NxCalculatorCellIdentity(target)
End Function

Public Function NxCalculatorCellFingerprint(ByVal target As Range) As String
    Dim raw As Variant, payload As String
    If target Is Nothing Then NxRaiseContractError "셀 지문 대상을 확인할 수 없습니다."
    raw = target.Value2
    If IsError(raw) Then
        payload = "ERROR:" & CStr(target.Text)
    ElseIf IsEmpty(raw) Then
        payload = "EMPTY"
    Else
        payload = TypeName(raw) & ":" & CStr(raw)
    End If
    NxCalculatorCellFingerprint = CStr(target.HasFormula) & "|" & CStr(VarType(raw)) & "|" & payload & "|" & CStr(target.Formula)
End Function

Public Sub NxCalculatorRequireWritableTarget(ByVal target As Range)
    If target Is Nothing Then NxRaiseContractError "결과를 입력할 셀을 확인할 수 없습니다."
    If target.Areas.Count <> 1 Or target.Cells.CountLarge <> 1 Then NxRaiseContractError "한 셀에만 결과를 입력할 수 있습니다."
    If CBool(target.MergeCells) Then NxRaiseContractError "병합된 셀에는 결과를 입력할 수 없습니다."
    If target.Worksheet.Parent.ReadOnly Then NxRaiseContractError "읽기 전용 통합문서에는 결과를 입력할 수 없습니다."
    If target.Worksheet.ProtectContents And target.Locked Then NxRaiseContractError "보호된 잠금 셀에는 결과를 입력할 수 없습니다."
End Sub

Public Function NxCalculatorIsDateFormatted(ByVal numberFormat As String) As Boolean
    Dim cleaned As String, index As Long, character As String, inQuote As Boolean, inBracket As Boolean
    numberFormat = LCase$(numberFormat)
    For index = 1 To Len(numberFormat)
        character = Mid$(numberFormat, index, 1)
        If inQuote Then
            If character = """" Then inQuote = False
        ElseIf inBracket Then
            If character = "]" Then inBracket = False
        ElseIf character = """" Then
            inQuote = True
        ElseIf character = "[" Then
            inBracket = True
        ElseIf character = "\" Then
            index = index + 1
        Else
            cleaned = cleaned & character
        End If
    Next index
    If InStr(cleaned, "y") > 0 Or InStr(cleaned, "d") > 0 Or InStr(cleaned, "h") > 0 Or InStr(cleaned, "s") > 0 Then
        NxCalculatorIsDateFormatted = True
    ElseIf InStr(cleaned, "m") > 0 And (InStr(cleaned, "/") > 0 Or InStr(cleaned, "-") > 0 Or InStr(cleaned, ":") > 0) Then
        NxCalculatorIsDateFormatted = True
    End If
End Function
