Attribute VB_Name = "NxSymbolCell"
Option Explicit

Private Const NX_SYMBOL_MAX_WRITE_CELLS As Long = 10000

Public Function NxSymbolSelectionIdentity(ByVal target As Range) As String
    Dim book As Workbook
    Dim bookIdentity As String
    Dim sheet As Worksheet

    If target Is Nothing Then Exit Function
    If target.Areas.Count <> 1 Then Exit Function
    Set sheet = target.Worksheet
    Set book = sheet.Parent
    If Len(book.Path) > 0 Then bookIdentity = book.FullName Else bookIdentity = book.Name
    NxSymbolSelectionIdentity = bookIdentity & "|" & sheet.CodeName & "|" & _
        target.Address(False, False, xlA1)
End Function

Public Function NxSymbolCellIdentity(ByVal target As Range) As String
    If target Is Nothing Then Exit Function
    If target.Areas.Count <> 1 Or target.Cells.CountLarge <> 1 Then Exit Function
    NxSymbolCellIdentity = NxSymbolSelectionIdentity(target)
End Function

Public Function NxSymbolActiveSelectionIdentity() As String
    If Application.ActiveWorkbook Is Nothing Or TypeName(Application.Selection) <> "Range" Then Exit Function
    NxSymbolActiveSelectionIdentity = NxSymbolSelectionIdentity(Application.Selection)
End Function

Public Function NxSymbolActiveCellIdentity() As String
    NxSymbolActiveCellIdentity = NxSymbolActiveSelectionIdentity()
End Function

Public Function NxSymbolActiveTargetLabel() As String
    Dim target As Range
    If Application.ActiveWorkbook Is Nothing Or TypeName(Application.Selection) <> "Range" Then
        NxSymbolActiveTargetLabel = "대상 셀 없음"
        Exit Function
    End If
    Set target = Application.Selection
    NxSymbolActiveTargetLabel = target.Worksheet.Parent.Name & " / " & target.Worksheet.Name & " / " & _
        target.Address(False, False, xlA1)
End Function

Public Function NxSymbolPrepareActiveWrite(ByVal bufferText As String, ByVal expectedIdentity As String) As CNxSymbolCellWrite
    Dim prepared As New CNxSymbolCellWrite
    prepared.Capture bufferText, expectedIdentity
    Set NxSymbolPrepareActiveWrite = prepared
End Function

Public Function NxSymbolCellFingerprint(ByVal target As Range) As String
    Dim payload As String
    Dim raw As Variant

    If target Is Nothing Then NxRaiseContractError "셀 지문 대상을 확인할 수 없습니다."
    If target.Cells.CountLarge <> 1 Then NxRaiseContractError "셀 지문은 한 셀씩 계산해야 합니다."
    raw = target.Value2
    If IsError(raw) Then
        payload = "ERROR:" & CStr(target.Text)
    ElseIf IsEmpty(raw) Then
        payload = "EMPTY"
    Else
        payload = TypeName(raw) & ":" & CStr(raw)
    End If
    NxSymbolCellFingerprint = CStr(target.HasFormula) & "|" & CStr(VarType(raw)) & "|" & payload & "|" & CStr(target.Formula)
End Function

Public Sub NxSymbolRequireWritableTarget(ByVal target As Range)
    Dim targetCell As Range

    If target Is Nothing Then NxRaiseContractError "결과를 입력할 셀 범위를 확인할 수 없습니다."
    If target.Areas.Count <> 1 Then NxRaiseContractError "서로 떨어진 여러 범위에는 기호를 입력할 수 없습니다."
    If target.Cells.CountLarge < 1 Or target.Cells.CountLarge > NX_SYMBOL_MAX_WRITE_CELLS Then _
        NxRaiseContractError "한 번에 입력할 수 있는 셀은 10,000개까지입니다."
    If target.Worksheet.Parent.ReadOnly Then NxRaiseContractError "읽기 전용 통합문서에는 기호를 입력할 수 없습니다."

    For Each targetCell In target.Cells
        If CBool(targetCell.MergeCells) Then NxRaiseContractError "병합된 셀에는 기호를 입력할 수 없습니다."
        If target.Worksheet.ProtectContents And CBool(targetCell.Locked) Then _
            NxRaiseContractError "보호된 잠금 셀에는 기호를 입력할 수 없습니다."
    Next targetCell
End Sub

Public Function NxSymbolIsDateFormatted(ByVal numberFormat As String) As Boolean
    Dim character As String
    Dim cleaned As String
    Dim inBracket As Boolean
    Dim index As Long
    Dim inQuote As Boolean

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
        NxSymbolIsDateFormatted = True
    ElseIf InStr(cleaned, "m") > 0 And (InStr(cleaned, "/") > 0 Or InStr(cleaned, "-") > 0 Or InStr(cleaned, ":") > 0) Then
        NxSymbolIsDateFormatted = True
    End If
End Function
