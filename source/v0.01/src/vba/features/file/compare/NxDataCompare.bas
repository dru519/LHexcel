Attribute VB_Name = "NxDataCompare"
Option Explicit

Private Const MAX_CELLS As Long = 1000000
Private Const BLOCK_CELLS As Long = 20000
Private mReport As Workbook
Private mBusy As Boolean, mCancel As Boolean
Private mView As Object
Private mReadSeconds As Double, mCompareSeconds As Double, mWriteSeconds As Double, mPaintSeconds As Double
Private mDifferences As Long
Private mPaintAddresses As String

Public Sub NxDataCompareBind(ByVal view As Object)
    If mBusy Then NxRaiseContractError "이미 데이터를 비교하고 있습니다."
    Set mView = view
End Sub

Public Sub NxDataCompareCancel()
    If mBusy Then mCancel = True
End Sub

Public Sub NxCompareRequireDestination(ByVal destination As Workbook)
    Dim book As Workbook, found As Boolean
    If destination Is Nothing Then NxRaiseContractError "결과를 추가할 통합문서가 없습니다. 새 창을 선택하세요."
    For Each book In Application.Workbooks
        If book Is destination Then found = True: Exit For
    Next book
    If Not found Then NxRaiseContractError "결과를 추가할 통합문서가 닫혔습니다. 새 창을 선택하세요."
    If destination.IsAddin Or destination.ReadOnly Or destination.ProtectStructure Then NxRaiseContractError "현재 파일에 시트를 추가할 수 없습니다. 새 창을 선택하세요."
End Sub

Public Sub NxCompareAppendResultSheets(ByVal report As Workbook, ByVal destination As Workbook)
    Dim before As Long, i As Long, sheet As Worksheet, first As Worksheet
    Dim guard As CNxStateGuard, failure As Long, detail As String
    Dim names As Object, used As Object, key As Variant, link As Hyperlink, links As New Collection
    Dim entry As Variant, oldName As String, newName As String, subAddress As String
    NxCompareRequireDestination destination
    If report Is destination Then NxRaiseContractError "결과 통합문서를 분리해서 생성하지 못했습니다."
    before = destination.Sheets.Count
    Set guard = New CNxStateGuard
    On Error GoTo Failed
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set names = CreateObject("Scripting.Dictionary"): names.CompareMode = vbTextCompare
    Set used = CreateObject("Scripting.Dictionary"): used.CompareMode = vbTextCompare
    For Each sheet In destination.Worksheets: used(sheet.Name) = True: Next sheet
    For Each sheet In report.Worksheets: used(sheet.Name) = True: Next sheet
    For Each sheet In report.Worksheets
        oldName = sheet.Name
        newName = NxSafeSheetName(oldName, used)
        names.Add oldName, newName
        For Each link In sheet.Hyperlinks
            If Not link.Range.HasFormula Then
                If Len(link.Address) = 0 Then links.Add Array(link, link.SubAddress)
            End If
        Next link
    Next sheet
    For Each sheet In report.Worksheets
        oldName = sheet.Name
        sheet.Name = CStr(names(oldName))
        If oldName = "비교 결과" Then RebindDifferenceLinks sheet, names
        If oldName = "비교요약" Or Left$(oldName, 3) = "셀차이" Then
            If VarType(sheet.Range("Q1").Value2) = vbString Then
                If names.Exists(CStr(sheet.Range("Q1").Value2)) Then sheet.Range("Q1").Value2 = names(CStr(sheet.Range("Q1").Value2))
            End If
        End If
        If oldName = "비교요약" Then
            For i = 6 To sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
                If names.Exists(CStr(sheet.Cells(i, 12).Value2)) Then sheet.Cells(i, 12).Value2 = names(CStr(sheet.Cells(i, 12).Value2))
                If names.Exists(CStr(sheet.Cells(i, 13).Value2)) Then sheet.Cells(i, 13).Value2 = names(CStr(sheet.Cells(i, 13).Value2))
            Next i
        End If
    Next sheet
    For Each entry In links
        Set link = entry(0): subAddress = CStr(entry(1))
        For Each key In names.Keys
            oldName = "'" & Replace$(CStr(key), "'", "''") & "'!"
            If Left$(subAddress, Len(oldName)) = oldName Then
                link.SubAddress = "'" & Replace$(CStr(names(key)), "'", "''") & "'!" & Mid$(subAddress, Len(oldName) + 1)
                Exit For
            End If
        Next key
    Next entry
    report.Worksheets.Copy After:=destination.Sheets(before)
    Set first = destination.Sheets(before + 1)
    report.Close SaveChanges:=False
    guard.Restore
    destination.Activate: first.Activate
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    For i = destination.Sheets.Count To before + 1 Step -1
        destination.Sheets(i).Delete
    Next i
    guard.Restore
    On Error GoTo 0
    Err.Raise failure, "NxCompareAppendResultSheets", detail
End Sub

Private Sub RebindDifferenceLinks(ByVal sheet As Worksheet, ByVal names As Object)
    Dim lastRow As Long, formulas As Variant, row As Long, col As Long, key As Variant
    lastRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
    If lastRow < 6 Then Exit Sub
    formulas = sheet.Range("D6:E" & CStr(lastRow)).Formula
    For row = 1 To UBound(formulas, 1)
        For col = 1 To 2
            For Each key In names.Keys
                formulas(row, col) = Replace$(CStr(formulas(row, col)), _
                    "#'" & Replace$(CStr(key), "'", "''") & "'!", _
                    "#'" & Replace$(CStr(names(key)), "'", "''") & "'!")
            Next key
        Next col
    Next row
    sheet.Range("D6:E" & CStr(lastRow)).Formula = formulas
End Sub

Public Function NxDataCompareMetrics() As String
    NxDataCompareMetrics = CStr(mReadSeconds) & "|" & CStr(mCompareSeconds) & "|" & CStr(mWriteSeconds) & "|" & CStr(mPaintSeconds) & "|" & CStr(mDifferences)
End Function

Private Function Elapsed(ByVal started As Double) As Double
    Elapsed = Timer - started
    If Elapsed < 0 Then Elapsed = Elapsed + 86400#
End Function

Private Sub Progress(ByVal stage As String, ByVal done As Long, ByVal total As Long)
    Application.StatusBar = stage & " · " & done & " / " & total & " · Esc로 취소"
    If Not mView Is Nothing Then mView.UpdateCompareProgress stage, done, total
    DoEvents
    If mCancel Then Err.Raise 18, "NxDataCompare", "데이터 비교를 취소했습니다."
End Sub

Public Function NxDataCompareUsedRange(ByVal sheet As Worksheet) As Range
    Dim lastRow As Range, lastColumn As Range
    Set lastRow = sheet.Cells.Find(What:="*", After:=sheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    Set lastColumn = sheet.Cells.Find(What:="*", After:=sheet.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If lastRow Is Nothing Then
        Set NxDataCompareUsedRange = sheet.Range("A1")
    Else
        Set NxDataCompareUsedRange = sheet.Range(sheet.Cells(1, 1), sheet.Cells(lastRow.Row, lastColumn.Column))
    End If
End Function

Private Sub ValidateRange(ByVal target As Range)
    If target Is Nothing Then NxRaiseContractError "비교할 두 범위를 선택하세요."
    If target.Areas.Count <> 1 Then NxRaiseContractError "연속된 하나의 범위를 선택하세요."
    If target.CountLarge > MAX_CELLS Then NxRaiseContractError "비교 범위는 각각 1,000,000셀 이내로 선택하세요."
    If target.Parent.Parent Is ThisWorkbook Then NxRaiseContractError "추가 기능 내부 시트는 비교할 수 없습니다."
End Sub

Private Function ReadBlock(ByVal source As Range, ByVal offset As Long, ByVal count As Long, ByVal formulas As Boolean) As Variant
    Dim target As Range, values As Variant, singleton(1 To 1, 1 To 1) As Variant
    If offset > source.Rows.Count Then Exit Function
    If count > source.Rows.Count - offset + 1 Then count = source.Rows.Count - offset + 1
    Set target = source.Cells(offset, 1).Resize(count, source.Columns.Count)
    If formulas Then values = target.Formula Else values = target.Value2
    If target.CountLarge = 1 Then
        singleton(1, 1) = values: ReadBlock = singleton
    Else
        ReadBlock = values
    End If
End Function

Private Function SameValue(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If EmptyValue(leftValue) And EmptyValue(rightValue) Then SameValue = True: Exit Function
    If VarType(leftValue) <> VarType(rightValue) Then Exit Function
    If IsError(leftValue) Then
        SameValue = (CStr(leftValue) = CStr(rightValue))
    ElseIf IsEmpty(leftValue) Then
        SameValue = True
    ElseIf VarType(leftValue) = vbString Then
        SameValue = (StrComp(CStr(leftValue), CStr(rightValue), vbBinaryCompare) = 0)
    Else
        SameValue = (leftValue = rightValue)
    End If
End Function

Private Function EmptyValue(ByVal value As Variant) As Boolean
    If IsEmpty(value) Then
        EmptyValue = True
    ElseIf VarType(value) = vbString Then
        EmptyValue = (Len(value) = 0)
    End If
End Function

Private Sub WriteBlock(ByVal output As Worksheet, ByVal offset As Long, ByVal values As Variant)
    Dim target As Range
    If Not IsArray(values) Then Exit Sub
    Set target = output.Cells(offset + 4, 1).Resize(UBound(values, 1), UBound(values, 2))
    ' Text format prevents formula-looking source strings becoming active formulas.
    target.NumberFormat = "@"
    target.Value2 = values
End Sub

Private Sub Paint(ByVal leftSheet As Worksheet, ByVal rightSheet As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstColumn As Long, ByVal lastColumn As Long)
    Dim started As Double, address As String
    If firstColumn = 0 Then Exit Sub
    started = Timer
    address = ColumnName(firstColumn) & CStr(firstRow + 4) & ":" & ColumnName(lastColumn) & CStr(lastRow + 4)
    If Len(mPaintAddresses) + Len(address) + 1 > 220 Then FlushPaint leftSheet, rightSheet
    If Len(mPaintAddresses) > 0 Then mPaintAddresses = mPaintAddresses & ","
    mPaintAddresses = mPaintAddresses & address
    mPaintSeconds = mPaintSeconds + Elapsed(started)
End Sub

Private Function ColumnName(ByVal number As Long) As String
    Do While number > 0
        number = number - 1
        ColumnName = Chr$(65 + number Mod 26) & ColumnName
        number = number \ 26
    Loop
End Function

Private Sub FlushPaint(ByVal leftSheet As Worksheet, ByVal rightSheet As Worksheet)
    If Len(mPaintAddresses) = 0 Then Exit Sub
    leftSheet.Range(mPaintAddresses).Interior.Color = RGB(255, 199, 206)
    rightSheet.Range(mPaintAddresses).Interior.Color = RGB(255, 199, 206)
    mPaintAddresses = vbNullString
End Sub

Public Function NxDataCompareCreate(ByVal featureId As String, ByVal leftRange As Range, ByVal rightRange As Range, Optional ByVal compareFormulas As Boolean = False) As CNxResult
    Dim rows As Long, columns As Long, blockRows As Long, offset As Long, count As Long, row As Long, col As Long
    Dim lv As Variant, rv As Variant, lf As Variant, rf As Variant, different As Boolean, mask() As Boolean
    Dim leftSheet As Worksheet, rightSheet As Worksheet, output As Workbook, sourceBook As Workbook
    Dim differences As Worksheet, detailRow As Long
    Dim runStart As Long, pendingStart As Long, pendingEnd As Long, pendingTop As Long, pendingBottom As Long
    Dim started As Double, priorStatus As Variant, priorCancel As XlEnableCancelKey
    Dim errorNumber As Long, detail As String
    Dim leftRows As Long, leftColumns As Long, rightRows As Long, rightColumns As Long
    Dim leftPresent As Boolean, rightPresent As Boolean
    Dim priorScreen As Boolean, priorEvents As Boolean
    If mBusy Then NxRaiseContractError "이미 데이터를 비교하고 있습니다."
    ValidateRange leftRange: ValidateRange rightRange
    leftRows = leftRange.Rows.Count: leftColumns = leftRange.Columns.Count
    rightRows = rightRange.Rows.Count: rightColumns = rightRange.Columns.Count
    rows = Application.Max(leftRows, rightRows)
    columns = Application.Max(leftColumns, rightColumns)
    If CDbl(rows) * columns > MAX_CELLS Then NxRaiseContractError "두 범위를 맞춘 결과가 1,000,000셀을 넘습니다."
    If rows + 4 > leftRange.Worksheet.Rows.Count Then NxRaiseContractError "결과 안내 행을 위한 공간이 부족합니다."
    If Not mReport Is Nothing Then NxRaiseContractError "이전 비교 결과 처리가 끝나지 않았습니다."
    priorStatus = Application.StatusBar: priorCancel = Application.EnableCancelKey
    priorScreen = Application.ScreenUpdating: priorEvents = Application.EnableEvents
    Set sourceBook = leftRange.Worksheet.Parent
    mBusy = True: mCancel = False: mDifferences = 0
    mReadSeconds = 0: mCompareSeconds = 0: mWriteSeconds = 0: mPaintSeconds = 0
    On Error GoTo Failed
    Application.ScreenUpdating = False: Application.EnableEvents = False
    mPaintAddresses = vbNullString
    Application.EnableCancelKey = xlErrorHandler
    Progress "비교 준비", 0, rows
    Set output = Workbooks.Add(xlWBATWorksheet)
    Set leftSheet = output.Worksheets(1): leftSheet.Name = "기준 데이터"
    Set rightSheet = output.Worksheets.Add(After:=leftSheet): rightSheet.Name = "비교 데이터"
    leftSheet.Range("A1").Value2 = "기준 데이터": rightSheet.Range("A1").Value2 = "비교 데이터"
    leftSheet.Range("A2").NumberFormat = "@": rightSheet.Range("A2").NumberFormat = "@"
    leftSheet.Range("A2").Value2 = leftRange.Address(External:=True)
    rightSheet.Range("A2").Value2 = rightRange.Address(External:=True)
    If featureId = "NX-FILE-SHEET-COMPARE" Then
        Set differences = output.Worksheets.Add(After:=rightSheet): differences.Name = "비교 결과"
        differences.Range("A1").Value2 = "시트 비교 결과"
        differences.Range("A2:B2").NumberFormat = "@"
        differences.Range("A2").Value2 = leftRange.Address(External:=True)
        differences.Range("B2").Value2 = rightRange.Address(External:=True)
        differences.Range("A5:E5").Value2 = Array("셀 위치", IIf(compareFormulas, "기준 값·수식", "기준값"), IIf(compareFormulas, "비교 값·수식", "비교값"), "기준 이동", "비교 이동")
        detailRow = 6
    End If
    NxDrawStyleGeneratedTable output, leftSheet.Cells(5, 1).Resize(rows, columns), 0
    NxDrawStyleGeneratedTable output, rightSheet.Cells(5, 1).Resize(rows, columns), 0
    blockRows = BLOCK_CELLS \ columns: If blockRows < 1 Then blockRows = 1
    For offset = 1 To rows Step blockRows
        count = blockRows: If offset + count - 1 > rows Then count = rows - offset + 1
        started = Timer
        lv = ReadBlock(leftRange, offset, count, False): rv = ReadBlock(rightRange, offset, count, False)
        If compareFormulas Then
            lf = ReadBlock(leftRange, offset, count, True): rf = ReadBlock(rightRange, offset, count, True)
        End If
        mReadSeconds = mReadSeconds + Elapsed(started)
        Progress "데이터 읽기", offset - 1, rows
        started = Timer
        ReDim mask(1 To count, 1 To columns)
        For row = 1 To count
            For col = 1 To columns
                different = False
                leftPresent = (offset + row - 1 <= leftRows And col <= leftColumns)
                rightPresent = (offset + row - 1 <= rightRows And col <= rightColumns)
                If leftPresent Xor rightPresent Then
                    If leftPresent Then
                        different = Not EmptyValue(lv(row, col))
                        If compareFormulas And Not different Then different = Not EmptyValue(lf(row, col))
                    Else
                        different = Not EmptyValue(rv(row, col))
                        If compareFormulas And Not different Then different = Not EmptyValue(rf(row, col))
                    End If
                ElseIf leftPresent And rightPresent Then
                    different = Not SameValue(lv(row, col), rv(row, col))
                    If compareFormulas And Not different Then different = Not SameValue(lf(row, col), rf(row, col))
                End If
                mask(row, col) = different
                If different Then mDifferences = mDifferences + 1
            Next col
        Next row
        mCompareSeconds = mCompareSeconds + Elapsed(started)
        started = Timer
        If Not differences Is Nothing Then WriteDifferences differences, detailRow, mask, offset, count, columns, lv, rv, lf, rf, compareFormulas, leftSheet, rightSheet
        WriteBlock leftSheet, offset, lv: WriteBlock rightSheet, offset, rv
        mWriteSeconds = mWriteSeconds + Elapsed(started)
        Progress "차이 표시", offset - 1, rows
        For row = 1 To count
            runStart = 0
            For col = 1 To columns + 1
                different = False
                If col <= columns Then different = mask(row, col)
                If different Then
                    If runStart = 0 Then runStart = col
                ElseIf runStart > 0 Then
                    If pendingStart = runStart And pendingEnd = col - 1 And pendingBottom = offset + row - 2 Then
                        pendingBottom = offset + row - 1
                    Else
                        Paint leftSheet, rightSheet, pendingTop, pendingBottom, pendingStart, pendingEnd
                        pendingStart = runStart: pendingEnd = col - 1
                        pendingTop = offset + row - 1: pendingBottom = pendingTop
                    End If
                    runStart = 0
                End If
            Next col
            If row Mod 100 = 0 Then Progress "차이 표시", offset + row - 1, rows
        Next row
    Next offset
    Paint leftSheet, rightSheet, pendingTop, pendingBottom, pendingStart, pendingEnd
    started = Timer: FlushPaint leftSheet, rightSheet: mPaintSeconds = mPaintSeconds + Elapsed(started)
    leftSheet.Range("A3").Value2 = "차이 " & Format$(mDifferences, "#,##0") & "셀 · 위치 기준 · " & IIf(compareFormulas, "값·수식 비교", "값 비교")
    rightSheet.Range("A3").Value2 = leftSheet.Range("A3").Value2
    leftSheet.Range("A4").Value2 = "원본을 보존한 값 사본입니다. 저장은 직접 선택하세요."
    rightSheet.Range("A4").Value2 = leftSheet.Range("A4").Value2
    If Not differences Is Nothing Then
        NxDrawStyleGeneratedTable output, differences.Range("A5").Resize(detailRow - 5, 5), 1
        differences.Range("A3").Value2 = "차이 " & Format$(mDifferences, "#,##0") & "셀"
        If mDifferences = 0 Then differences.Range("A6").Value2 = "차이가 없습니다."
        differences.Columns("A:A").ColumnWidth = 14
        differences.Columns("B:C").ColumnWidth = 30
        differences.Columns("D:E").ColumnWidth = 16
        If detailRow > 6 Then
            With differences.Range("D6:E" & CStr(detailRow - 1)).Font
                .Color = RGB(0, 102, 204): .Underline = xlUnderlineStyleSingle
            End With
        End If
        leftSheet.Hyperlinks.Add Anchor:=leftSheet.Range("A4"), Address:=vbNullString, SubAddress:="'비교 결과'!A5", TextToDisplay:="차이 목록으로 이동"
        rightSheet.Hyperlinks.Add Anchor:=rightSheet.Range("A4"), Address:=vbNullString, SubAddress:="'비교 결과'!A5", TextToDisplay:="차이 목록으로 이동"
    End If
    Set mReport = output: Set output = Nothing
    Set NxDataCompareCreate = NxCreateResult(featureId, NxSuccess, "complete", sourceBook.FullName, False, vbNullString, "data_compare_complete")
Done:
    Application.StatusBar = priorStatus
    Application.EnableCancelKey = priorCancel
    Application.ScreenUpdating = priorScreen: Application.EnableEvents = priorEvents
    mBusy = False: Set mView = Nothing
    Exit Function
Failed:
    errorNumber = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not output Is Nothing Then output.Close SaveChanges:=False
    Application.StatusBar = priorStatus: Application.EnableCancelKey = priorCancel
    Application.ScreenUpdating = priorScreen: Application.EnableEvents = priorEvents
    mBusy = False: Set mView = Nothing
    On Error GoTo 0
    Err.Raise errorNumber, "NxDataCompare", detail
End Function

Private Sub WriteDifferences(ByVal sheet As Worksheet, ByRef nextRow As Long, ByRef mask() As Boolean, _
    ByVal offset As Long, ByVal count As Long, ByVal columns As Long, ByRef lv As Variant, ByRef rv As Variant, _
    ByRef lf As Variant, ByRef rf As Variant, ByVal compareFormulas As Boolean, ByVal leftSheet As Worksheet, ByVal rightSheet As Worksheet)
    Dim row As Long, col As Long, total As Long, index As Long, address As String, snapshotAddress As String
    Dim values() As Variant, links() As Variant, target As Range
    For row = 1 To count
        For col = 1 To columns
            If mask(row, col) Then total = total + 1
        Next col
    Next row
    If total = 0 Then Exit Sub
    ReDim values(1 To total, 1 To 3): ReDim links(1 To total, 1 To 2)
    For row = 1 To count
        For col = 1 To columns
            If mask(row, col) Then
                index = index + 1
                address = ColumnName(col) & CStr(offset + row - 1)
                snapshotAddress = ColumnName(col) & CStr(offset + row + 3)
                values(index, 1) = address
                If compareFormulas Then
                    values(index, 2) = DifferenceValue(lf, row, col): values(index, 3) = DifferenceValue(rf, row, col)
                Else
                    values(index, 2) = DifferenceValue(lv, row, col): values(index, 3) = DifferenceValue(rv, row, col)
                End If
                links(index, 1) = DifferenceLink(leftSheet.Name, snapshotAddress, address)
                links(index, 2) = DifferenceLink(rightSheet.Name, snapshotAddress, address)
            End If
        Next col
    Next row
    Set target = sheet.Cells(nextRow, 1).Resize(total, 3)
    target.NumberFormat = "@": target.Value2 = values
    Set target = sheet.Cells(nextRow, 4).Resize(total, 2)
    target.NumberFormat = "General": target.Formula = links
    nextRow = nextRow + total
End Sub

Private Function DifferenceValue(ByRef values As Variant, ByVal row As Long, ByVal col As Long) As Variant
    If Not IsArray(values) Then Exit Function
    If row > UBound(values, 1) Or col > UBound(values, 2) Then Exit Function
    DifferenceValue = values(row, col)
End Function

Private Function DifferenceLink(ByVal sheetName As String, ByVal address As String, ByVal label As String) As String
    DifferenceLink = "=HYPERLINK(""#'" & Replace$(sheetName, "'", "''") & "'!" & address & """,""" & label & """)"
End Function

Public Sub NxDataCompareFinish(ByVal success As Boolean)
    If mReport Is Nothing Then Exit Sub
    If success Then
        mReport.Activate
        mReport.Worksheets(1).Activate
        mReport.Worksheets(1).Range("A5").Select
        If mReport.Worksheets.Count = 3 Then
            If mReport.Worksheets(3).Name = "비교 결과" Then mReport.Worksheets(3).Activate
        End If
    Else
        mReport.Close SaveChanges:=False
    End If
    Set mReport = Nothing
End Sub
