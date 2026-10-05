Attribute VB_Name = "NxFunctionWrap"
Option Explicit

Private mCells As Collection, mBefore As Collection, mAfter As Collection
Private mReady As Boolean

Public Sub NxFunctionWrapOpen(ByVal target As Range, ByVal errorMode As Boolean)
    Dim view As New FNxFunctionWrap
    mReady = False
    view.BindTarget target, errorMode
    view.Show vbModal
End Sub

Public Function NxWrappedFormula(ByVal original As String, ByVal functionName As String, ByVal argument As String) As String
    Dim expression As String, inner As String, i As Long, depth As Long, brackets As Long, braces As Long
    Dim quote As String, ch As String, splitAt As Long, validOuter As Boolean
    If functionName <> "ROUND" And functionName <> "ROUNDUP" And functionName <> "ROUNDDOWN" And functionName <> "IFERROR" Then NxRaiseContractError "지원하지 않는 함수입니다."
    expression = original
    If Left$(expression, 1) = "=" Then expression = Mid$(expression, 2)
    ' Update only the exact same outer function. Other wrappers remain nested.
    If UCase$(Left$(expression, Len(functionName) + 1)) = functionName & "(" And Right$(expression, 1) = ")" Then
        inner = Mid$(expression, Len(functionName) + 2, Len(expression) - Len(functionName) - 2)
        validOuter = True
        For i = 1 To Len(inner)
            ch = Mid$(inner, i, 1)
            If Len(quote) > 0 Then
                If ch = quote Then
                    If Mid$(inner, i + 1, 1) = quote Then i = i + 1 Else quote = ""
                End If
            Else
                Select Case ch
                    Case """", "'": quote = ch
                    Case "[": brackets = brackets + 1
                    Case "]": brackets = brackets - 1
                    Case "{": braces = braces + 1
                    Case "}": braces = braces - 1
                    Case "(": If brackets = 0 Then depth = depth + 1
                    Case ")": If brackets = 0 Then depth = depth - 1
                    Case ","
                        If depth = 0 And brackets = 0 And braces = 0 Then
                            If splitAt > 0 Then validOuter = False
                            splitAt = i
                        End If
                End Select
                If depth < 0 Or brackets < 0 Or braces < 0 Then validOuter = False
            End If
        Next i
        If validOuter And splitAt > 0 And depth = 0 And brackets = 0 And braces = 0 And Len(quote) = 0 Then expression = Left$(inner, splitAt - 1)
    End If
    NxWrappedFormula = "=" & functionName & "(" & expression & "," & argument & ")"
    If Len(NxWrappedFormula) > 8192 Then NxRaiseContractError "변경할 수식이 Excel 길이 제한을 넘습니다."
End Function

Public Function NxFunctionArgument(ByVal errorMode As Boolean, ByVal optionIndex As Long, ByVal text As String) As String
    Dim number As Double
    If errorMode And optionIndex = 0 Then NxFunctionArgument = """""": Exit Function
    If errorMode And optionIndex = 2 Then
        NxFunctionArgument = """" & Replace(text, """", """""") & """"
        Exit Function
    End If
    If Not IsNumeric(text) Then NxRaiseContractError "숫자를 입력하세요."
    number = CDbl(text)
    If Not errorMode Then
        If number <> Fix(number) Or number < -308 Or number > 308 Then NxRaiseContractError "자릿수는 -308~308 사이 정수로 입력하세요."
    End If
    NxFunctionArgument = Trim$(Str$(number))
End Function

Private Function ReadFormula(ByVal cell As Range) As String
    On Error GoTo Legacy
    ReadFormula = CStr(CallByName(cell, "Formula2", VbGet))
    Exit Function
Legacy:
    Err.Clear
    ReadFormula = CStr(cell.Formula)
End Function

Private Sub WriteFormula(ByVal cell As Range, ByVal formula As String)
    Dim ignored As Variant, supports As Boolean
    On Error Resume Next
    ignored = CallByName(cell, "Formula2", VbGet)
    supports = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If supports Then CallByName cell, "Formula2", VbLet, formula Else cell.Formula = formula
End Sub

Private Function IsSpill(ByVal cell As Range) As Boolean
    On Error Resume Next
    IsSpill = CBool(CallByName(cell, "HasSpill", VbGet))
    Err.Clear
End Function

Private Function Candidates(ByVal target As Range) As Range
    Dim constants As Range, formulas As Range, combined As Range
    On Error Resume Next
    Set constants = target.SpecialCells(xlCellTypeConstants, xlNumbers)
    Set formulas = target.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If constants Is Nothing Then
        Set combined = formulas
    ElseIf formulas Is Nothing Then
        Set combined = constants
    Else
        Set combined = Application.Union(constants, formulas)
    End If
    If Not combined Is Nothing Then Set Candidates = Application.Intersect(target, combined)
End Function

Private Function Eligible(ByVal cell As Range, ByVal includeHidden As Boolean) As Boolean
    If Not includeHidden Then
        If cell.EntireRow.Hidden Or cell.EntireColumn.Hidden Then Exit Function
    End If
    If cell.HasArray Or IsSpill(cell) Then Exit Function
    If Not cell.ListObject Is Nothing Then Exit Function
    Eligible = True
End Function

Public Function NxFunctionPreview(ByVal target As Range, ByVal functionName As String, ByVal argument As String, ByVal includeHidden As Boolean) As String
    Dim selected As Range, cell As Range, sample As String, count As Long, skipped As Long
    Set selected = Candidates(target)
    If selected Is Nothing Then NxFunctionPreview = "적용할 숫자나 수식이 없습니다.": Exit Function
    If selected.CountLarge > 10000 Then NxRaiseContractError "숫자·수식 셀을 10,000개 이내로 선택하세요."
    For Each cell In selected.Cells
        If Eligible(cell, includeHidden) Then
            count = count + 1
            If Len(sample) = 0 Then
                If cell.HasFormula Then sample = cell.Address(False, False) & ": " & ReadFormula(cell) & vbCrLf & NxWrappedFormula(ReadFormula(cell), functionName, argument) Else sample = cell.Address(False, False) & ": " & CStr(cell.Value2) & vbCrLf & NxWrappedFormula(Trim$(Str$(CDbl(cell.Value2))), functionName, argument)
            End If
        Else
            skipped = skipped + 1
        End If
    Next cell
    NxFunctionPreview = "적용 " & count & "개 / 숨김·배열·스필·표 제외 " & skipped & "개" & vbCrLf & sample
End Function

Public Sub NxFunctionApply(ByVal target As Range, ByVal functionName As String, ByVal argument As String, ByVal includeHidden As Boolean)
    Dim selected As Range, cell As Range, expression As String, i As Long, applied As Long
    Dim guard As CNxStateGuard, failure As Long, detail As String, state As Variant
    mReady = False
    If target.Worksheet.ProtectContents Or target.Worksheet.Parent.ReadOnly Then NxRaiseContractError "편집 가능한 시트에서 실행하세요."
    Set selected = Candidates(target)
    If selected Is Nothing Then NxRaiseContractError "적용할 숫자나 수식이 없습니다."
    If selected.CountLarge > 10000 Then NxRaiseContractError "숫자·수식 셀을 10,000개 이내로 선택하세요."
    Set mCells = New Collection: Set mBefore = New Collection: Set mAfter = New Collection
    For Each cell In selected.Cells
        If Eligible(cell, includeHidden) Then
            If cell.HasFormula Then
                expression = ReadFormula(cell): state = Array(True, expression)
            Else
                expression = Trim$(Str$(CDbl(cell.Value2))): state = Array(False, cell.Value2)
            End If
            mCells.Add cell: mBefore.Add state: mAfter.Add NxWrappedFormula(expression, functionName, argument)
        End If
    Next cell
    If mCells.Count = 0 Then NxRaiseContractError "적용 가능한 셀이 없습니다. 숨김·배열·스필·표 셀은 제외됩니다."
    Set guard = New CNxStateGuard
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    For i = 1 To mCells.Count
        applied = i
        WriteFormula mCells(i), CStr(mAfter(i))
    Next i
    guard.Restore
    mReady = True
    NxFunctionUndoArm
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description
    On Error Resume Next
    For i = applied To 1 Step -1
        Set cell = mCells(i): state = mBefore(i)
        If state(0) Then WriteFormula cell, CStr(state(1)) Else cell.Value2 = state(1)
    Next i
    If Err.Number <> 0 Then detail = detail & " / 복원: " & Err.Description
    guard.Restore
    On Error GoTo 0
    Err.Raise failure, "NxFunctionApply", detail
End Sub

Public Sub NxFunctionUndoArm()
    If mReady Then Application.OnUndo "내엑셀 함수 감싸기 취소", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!NxFunctionUndoLast"
End Sub

Public Sub NxFunctionUndoLast()
    Dim i As Long, cell As Range, state As Variant, guard As CNxStateGuard, detail As String
    If Not mReady Then Exit Sub
    On Error GoTo Failed
    For i = 1 To mCells.Count
        Set cell = mCells(i)
        If cell.Worksheet.ProtectContents Or cell.Worksheet.Parent.ReadOnly Then NxRaiseContractError "편집 가능한 시트에서 복원하세요."
        If ReadFormula(cell) <> CStr(mAfter(i)) Then NxRaiseContractError "적용 후 변경된 셀이 있어 복원하지 않았습니다."
    Next i
    Set guard = New CNxStateGuard
    Application.EnableEvents = False: Application.ScreenUpdating = False
    For i = mCells.Count To 1 Step -1
        Set cell = mCells(i): state = mBefore(i)
        If state(0) Then WriteFormula cell, CStr(state(1)) Else cell.Value2 = state(1)
    Next i
    mReady = False
    Set mCells = Nothing: Set mBefore = Nothing: Set mAfter = Nothing
    guard.Restore
    Exit Sub
Failed:
    detail = Err.Description
    If Not guard Is Nothing Then guard.Restore
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation, "내엑셀 - 함수 복원"
End Sub
