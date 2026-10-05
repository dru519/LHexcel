Attribute VB_Name = "NxFormulaTools"
Option Explicit

Private mCells As Collection, mBefore As Collection, mAfter As Collection
Private mBackup As CNxFormulaNoteBackup
Private mAppended As Collection
Private mReady As Boolean
Private mReport As Worksheet

Public Sub NxFormulaToolsOpen(ByVal target As Range, ByVal references As Boolean)
    Dim view As New FNxFormulaTools
    If references Then
        Set mReport = Nothing
    Else
        mReady = False
        If Not mBackup Is Nothing Then mBackup.Dispose
        Set mBackup = Nothing
    End If
    view.BindTarget target, references
    view.Show vbModal
End Sub

Public Function NxFormulaCells(ByVal target As Range) As Range
    Dim used As Range, found As Range
    Set used = Intersect(target, target.Worksheet.UsedRange)
    If used Is Nothing Then Exit Function
    On Error Resume Next
    Set found = used.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    If Not found Is Nothing Then Set NxFormulaCells = Intersect(found, target)
End Function

Public Sub NxFormulaNotesApply(ByVal target As Range, ByVal showNotes As Boolean, Optional ByVal replaceNotes As Boolean = False)
    Dim formulas As Range, cell As Range, old As String, value As String, state As Variant
    Dim i As Long, applied As Long, guard As CNxStateGuard, failure As Long, detail As String
    If Not mBackup Is Nothing Then mBackup.Dispose
    Set mBackup = Nothing
    mReady = False
    If target.Worksheet.ProtectContents Or target.Worksheet.Parent.ReadOnly Then NxRaiseContractError "편집 가능한 시트에서 실행하세요."
    Set formulas = NxFormulaCells(target)
    If formulas Is Nothing Then NxRaiseContractError "선택 범위에 수식이 없습니다."
    If formulas.CountLarge > 10000 Then NxRaiseContractError "수식 셀을 10,000개 이내로 선택하세요."
    Set mCells = New Collection: Set mBefore = New Collection: Set mAfter = New Collection
    Set mAppended = New Collection
    For Each cell In formulas.Cells
        old = ""
        state = Array(False, "", False)
        If Not cell.Comment Is Nothing Then
            old = cell.Comment.Text
            state = Array(True, old, cell.Comment.Visible)
        End If
        value = CStr(cell.Formula)
        If Not replaceNotes And Len(old) > 0 Then
            If Right$(old, Len(value)) = value Then value = old Else value = old & vbCrLf & value
        End If
        If Len(value) > 32767 Then NxRaiseContractError "메모의 최대 글자 수를 초과하는 셀이 있습니다."
        mCells.Add cell: mBefore.Add state: mAfter.Add value
        mAppended.Add (Not replaceNotes And Len(old) > 0 And Len(value) > Len(old))
    Next cell
    Set guard = New CNxStateGuard
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set mBackup = New CNxFormulaNoteBackup
    For i = 1 To mCells.Count
        state = mBefore(i)
        If CBool(state(0)) And CStr(mAfter(i)) <> CStr(state(1)) Then mBackup.Capture mCells(i), i
    Next i
    mBackup.Hide
    For i = 1 To mCells.Count
        applied = i
        Set cell = mCells(i)
        state = mBefore(i)
        If cell.Comment Is Nothing Then cell.AddComment
        If CBool(mAppended(i)) Then
            cell.Comment.Text Text:=Mid$(CStr(mAfter(i)), Len(CStr(state(1))) + 1), Start:=Len(CStr(state(1))) + 1, Overwrite:=False
        ElseIf CStr(mAfter(i)) <> CStr(state(1)) Or Not CBool(state(0)) Then
            cell.Comment.Text Text:=CStr(mAfter(i))
        End If
        cell.Comment.Visible = showNotes
    Next i
    guard.Restore
    mReady = True
    NxFormulaNotesUndoArm
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description
    On Error Resume Next
    For i = applied To 1 Step -1
        RestoreNote i
    Next i
    If Err.Number <> 0 Then detail = detail & " / 메모 복원: " & Err.Description
    If Not mBackup Is Nothing Then mBackup.Dispose
    Set mBackup = Nothing
    guard.Restore
    On Error GoTo 0
    Err.Raise failure, "NxFormulaNotesApply", detail
End Sub

Private Sub RestoreNote(ByVal index As Long)
    Dim cell As Range, state As Variant
    Set cell = mCells(index): state = mBefore(index)
    If state(0) Then
        If CStr(mAfter(index)) <> CStr(state(1)) Then mBackup.Restore cell, index
        cell.Comment.Visible = CBool(state(2))
    ElseIf Not cell.Comment Is Nothing Then
        cell.Comment.Delete
    End If
End Sub

Public Sub NxFormulaNotesUndoArm()
    If mReady Then Application.OnUndo "셀수식 메모 기록 취소", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!NxFormulaNotesUndoLast"
End Sub

Public Sub NxFormulaNotesUndoLast()
    Dim i As Long, cell As Range, guard As CNxStateGuard, detail As String
    If Not mReady Then Exit Sub
    On Error GoTo Failed
    For i = 1 To mCells.Count
        Set cell = mCells(i)
        If cell.Worksheet.ProtectContents Or cell.Worksheet.Parent.ReadOnly Then NxRaiseContractError "편집 가능한 시트에서 복원하세요."
        If cell.Comment Is Nothing Then NxRaiseContractError "적용 후 메모가 변경되어 복원을 중단했습니다."
        If cell.Comment.Text <> CStr(mAfter(i)) Then NxRaiseContractError "적용 후 메모가 변경되어 복원을 중단했습니다."
    Next i
    Set guard = New CNxStateGuard
    Application.EnableEvents = False: Application.ScreenUpdating = False
    For i = mCells.Count To 1 Step -1
        RestoreNote i
    Next i
    If Not mBackup Is Nothing Then mBackup.Dispose
    Set mBackup = Nothing
    mReady = False
    guard.Restore
    Exit Sub
Failed:
    detail = Err.Description
    If Not guard Is Nothing Then guard.Restore
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation, "내엑셀 - 메모 복원"
End Sub

Public Function NxFormulaReferenceReport(ByVal target As Range, ByVal newWorkbook As Boolean) As Worksheet
    Dim formulas As Range, cell As Range, refs As Range, area As Range, rows As New Collection
    Dim book As Workbook, report As Worksheet, source As Worksheet, previous As Object
    Dim row As Variant, values() As Variant, i As Long, j As Long, status As String, expression As String
    Dim guard As CNxStateGuard, failure As Long, detail As String, table As Range
    Set mReport = Nothing
    Set formulas = NxFormulaCells(target)
    If formulas Is Nothing Then NxRaiseContractError "선택 범위에 수식이 없습니다."
    If formulas.CountLarge > 10000 Then NxRaiseContractError "수식 셀을 10,000개 이내로 선택하세요."
    Set source = target.Worksheet
    If Not newWorkbook And (source.Parent.ProtectStructure Or source.Parent.ReadOnly) Then NxRaiseContractError "새 시트를 추가할 수 없습니다. 새 통합문서를 선택하세요."
    Set previous = ActiveSheet
    Set guard = New CNxStateGuard
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    source.Activate
    For Each cell In formulas.Cells
        Set refs = Nothing
        expression = CStr(cell.Formula)
        On Error Resume Next
        Set refs = cell.DirectPrecedents
        Err.Clear
        On Error GoTo Failed
        status = NxFormulaReferenceStatus(expression, Not refs Is Nothing)
        If refs Is Nothing Then
            rows.Add Array(cell.Address(External:=True), expression, "", "", status)
        Else
            For Each area In refs.Areas
                If area.CountLarge = 1 Then
                    If IsError(area.Value2) Then detail = area.Text Else detail = CStr(area.Value2)
                Else
                    detail = Format$(area.CountLarge, "#,##0") & "개 셀"
                End If
                rows.Add Array(cell.Address(External:=True), expression, area.Address(External:=True), detail, status)
                If rows.Count > 20000 Then NxRaiseContractError "참조표가 20,000행을 초과합니다. 선택 범위를 나누어 실행하세요."
            Next area
        End If
    Next cell
    ' Finish source inspection before creating/activating the output.
    If newWorkbook Then
        Set book = Workbooks.Add(xlWBATWorksheet)
        Set report = book.Worksheets(1)
    Else
        Set book = source.Parent
        Set report = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
    End If
    report.Name = NxDataNextNormalizationSheetName(book, "수식 참조표")
    ReDim values(1 To rows.Count + 1, 1 To 5)
    row = Array("수식 셀", "수식", "참조 위치", "참조 값", "확인 상태")
    For j = 0 To 4: values(1, j + 1) = row(j): Next j
    For i = 1 To rows.Count
        row = rows(i)
        For j = 0 To 4: values(i + 1, j + 1) = row(j): Next j
    Next i
    Set table = report.Range("A1").Resize(rows.Count + 1, 5)
    table.NumberFormat = "@"
    table.Value2 = values
    For i = 1 To rows.Count
        row = rows(i)
        NxFormulaSourceLink report.Cells(i + 1, 1), source, CStr(row(0))
        If Len(CStr(row(2))) > 0 Then NxFormulaSourceLink report.Cells(i + 1, 3), source, CStr(row(2))
    Next i
    NxFormulaReportStyle table
    guard.Restore
    Set mReport = report
    NxFormulaReferenceReveal
    Set NxFormulaReferenceReport = report
    Exit Function
Failed:
    failure = Err.Number: detail = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not report Is Nothing Then
        If newWorkbook Then book.Close False Else report.Delete
    End If
    If Not previous Is Nothing Then previous.Activate
    guard.Restore
    On Error GoTo 0
    Err.Raise failure, "NxFormulaReferenceReport", detail
End Function

Private Sub NxFormulaSourceLink(ByVal anchor As Range, ByVal source As Worksheet, ByVal externalAddress As String)
    Dim filePath As String, subAddress As String
    subAddress = "'" & Replace$(source.Name, "'", "''") & "'!" & Mid$(externalAddress, InStrRev(externalAddress, "!") + 1)
    If Not anchor.Worksheet.Parent Is source.Parent Then
        If Len(source.Parent.Path) > 0 Then
            filePath = source.Parent.FullName
        Else
            ' An unsaved workbook has no persistent file target yet.
            subAddress = externalAddress
        End If
    End If
    anchor.Worksheet.Hyperlinks.Add Anchor:=anchor, Address:=filePath, SubAddress:=subAddress, TextToDisplay:=externalAddress, ScreenTip:=NxFormulaLinkGuide()
End Sub

Public Function NxFormulaLinkGuide() As String
    NxFormulaLinkGuide = "클릭하면 원본으로 이동합니다. 이동되지 않으면 원본 파일을 열고 파일·시트 이름을 확인하세요. 이름이 바뀌었으면 참조표를 다시 만드세요."
End Function

Public Function NxFormulaReferenceStatus(ByVal expression As String, ByVal found As Boolean) As String
    Dim clean As String, i As Long, ch As String, quoted As Boolean, kinds As String, compact As String
    ' Ignore text literals: punctuation inside a displayed string is not a reference.
    i = 1
    Do While i <= Len(expression)
        ch = Mid$(expression, i, 1)
        If ch = """" Then
            If quoted And Mid$(expression, i + 1, 1) = """" Then
                i = i + 1
            Else
                quoted = Not quoted
            End If
        ElseIf Not quoted Then
            clean = clean & ch
        End If
        i = i + 1
    Loop
    compact = UCase$(Replace(clean, " ", ""))
    If InStr(compact, "INDIRECT(") > 0 Or InStr(compact, "OFFSET(") > 0 Then kinds = "동적 참조"
    If InStr(clean, "!") > 0 Then
        If Len(kinds) > 0 Then kinds = kinds & " · "
        If InStr(clean, "[") > 0 Then kinds = kinds & "외부 파일 참조" Else kinds = kinds & "다른 시트 참조"
    ElseIf InStr(clean, "[") > 0 Then
        If Len(kinds) > 0 Then kinds = kinds & " · "
        kinds = kinds & "표 구조 참조"
    End If
    If Len(kinds) > 0 Then
        NxFormulaReferenceStatus = kinds & " 별도 확인"
        If found Then NxFormulaReferenceStatus = "확인된 범위만 표시 · " & NxFormulaReferenceStatus
    ElseIf found Then
        NxFormulaReferenceStatus = "같은 시트의 직접 참조 확인"
    Else
        NxFormulaReferenceStatus = "직접 참조를 확인하지 못함 · 상수 수식 또는 이름 정의 확인"
    End If
End Function

Public Sub NxFormulaReferenceReveal()
    If mReport Is Nothing Then Exit Sub
    mReport.Parent.Activate
    mReport.Activate
    mReport.Range("A1").Select
End Sub

Private Sub NxFormulaReportStyle(ByVal table As Range)
    Dim edge As Variant
    NxApplyBodyToken table, True
    NxApplyHeaderToken table.Rows(1)
    table.Font.Name = "맑은 고딕": table.Font.Size = 9
    table.Rows(1).Interior.Color = NxDrawTableHeaderFill(NX_ROLE_STYLE_MODE_MONO)
    For Each edge In Array(xlEdgeTop, xlEdgeBottom, xlInsideHorizontal, xlInsideVertical)
        With table.Borders(CLng(edge))
            .LineStyle = xlContinuous: .Color = RGB(110, 110, 110): .Weight = xlThin
        End With
    Next edge
    table.Borders(xlEdgeLeft).LineStyle = xlNone: table.Borders(xlEdgeRight).LineStyle = xlNone
    table.Borders(xlEdgeBottom).Weight = xlThick
    table.Rows(1).Borders(xlEdgeBottom).Weight = xlThick
    table.Columns(1).ColumnWidth = 30: table.Columns(2).ColumnWidth = 48
    table.Columns(3).ColumnWidth = 30: table.Columns(4).ColumnWidth = 22: table.Columns(5).ColumnWidth = 44
    table.WrapText = True: table.Rows.AutoFit
End Sub
