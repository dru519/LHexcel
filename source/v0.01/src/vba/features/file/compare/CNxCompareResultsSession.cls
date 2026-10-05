VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxCompareResultsSession"
Option Explicit

Private mReport As Workbook
Private mSheets As Collection
Private mRows As Variant, mLocations As Variant, mExpected As Variant
Private mBasePath As String, mComparePath As String
Private mCount As Long, mReleased As Boolean, mNavigating As Boolean

Public Sub Configure(ByVal report As Workbook)
    Dim summary As Worksheet, sheet As Worksheet, block As Range, values As Variant
    Dim count As Long, missing As Long, lastRow As Long, summaryLast As Long, row As Long
    Dim index As Long, sheetIndex As Long, column As Long, chars As Double, reason As String
    If Not mReport Is Nothing Or mReleased Then NxRaiseContractError "결과 탐색 세션은 다시 사용할 수 없습니다."
    If report Is Nothing Then NxRaiseContractError "비교 보고서를 찾을 수 없습니다."
    If report.IsAddin Then NxRaiseContractError "내엑셀 비교 결과가 있는 통합문서를 선택하세요."
    Set summary = ResolveSummary(report)
    If Literal(summary.Range("J1")) <> "NX_COMPARE_NAV_V1" Then NxRaiseContractError "내엑셀 비교 보고서가 아닙니다."
    If Literal(summary.Range("A1")) <> "내엑셀 통합문서 비교" Then NxRaiseContractError "비교 보고서 제목을 확인하세요."
    mBasePath = Literal(summary.Range("B2")): mComparePath = Literal(summary.Range("B3"))
    If Len(NxWorkbookCompareNavigationPathIssue(mBasePath)) > 0 Or Len(NxWorkbookCompareNavigationPathIssue(mComparePath)) > 0 Then NxRaiseContractError "비교 원본 경로가 올바르지 않습니다."
    If StrComp(mBasePath, mComparePath, vbTextCompare) = 0 Then NxRaiseContractError "비교 원본 경로가 같습니다."
    summaryLast = summary.Cells(summary.Rows.Count, 1).End(xlUp).Row
    If summaryLast < 6 Or summaryLast > 32000 Then NxRaiseContractError "비교 요약 범위를 확인하세요."
    values = LiteralBlock(summary.Range("A6:F" & CStr(summaryLast)))
    For row = 1 To UBound(values, 1)
        reason = CStr(values(row, 2))
        Select Case reason
            Case "기준 문서에 없음", "비교 문서에 없음": missing = missing + 1
            Case "일치", "차이 있음"
            Case Else: NxRaiseContractError "비교 요약 상태가 올바르지 않습니다."
        End Select
    Next row
    count = missing
    Set mSheets = New Collection: mSheets.Add summary
    For Each sheet In report.Worksheets
        If IsReportDetail(sheet, summary) Then
            lastRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
            If lastRow > 32000 Then NxRaiseContractError "비교 상세 범위를 확인하세요."
            If lastRow >= 2 Then count = count + lastRow - 1
            mSheets.Add sheet
        End If
    Next sheet
    If count > 200000 Then NxRaiseContractError "결과 탐색은 200,000건까지 지원합니다. 보고서에서 확인하세요."
    If Literal(summary.Range("B4")) <> CStr(count) Then NxRaiseContractError "요약과 상세 차이 수가 달라 보고서를 다시 생성해야 합니다."
    Set mReport = report: mCount = count
    If count = 0 Then Exit Sub
    ReDim mRows(1 To count, 1 To 9)
    ReDim mExpected(1 To count, 1 To 7)
    ReDim mLocations(1 To count, 1 To 2)
    For row = 1 To UBound(values, 1)
        reason = CStr(values(row, 2))
        If reason = "기준 문서에 없음" Or reason = "비교 문서에 없음" Then
            index = index + 1: mLocations(index, 1) = 1: mLocations(index, 2) = row + 5
            For column = 1 To 9: mRows(index, column) = vbNullString: Next column
            For column = 1 To 6: mExpected(index, column) = CStr(values(row, column)): Next column
            mExpected(index, 7) = vbNullString
            mRows(index, 1) = CStr(values(row, 1)): mRows(index, 2) = "A1": mRows(index, 3) = reason
            mRows(index, 8) = IIf(reason = "기준 문서에 없음", "0", "1")
            mRows(index, 9) = IIf(reason = "비교 문서에 없음", "0", "1")
            CountCharacters index, chars
        End If
    Next row
    For sheetIndex = 2 To mSheets.Count
        Set sheet = mSheets.Item(sheetIndex)
        lastRow = sheet.Cells(sheet.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then
            values = LiteralBlock(sheet.Range("A2:G" & CStr(lastRow)))
            For row = 1 To UBound(values, 1)
                index = index + 1: mLocations(index, 1) = sheetIndex: mLocations(index, 2) = row + 1
                For column = 1 To 7
                    mRows(index, column) = CStr(values(row, column)): mExpected(index, column) = mRows(index, column)
                Next column
                If Len(NxWorkbookCompareNavigationCellIssue(CStr(mRows(index, 2)))) > 0 Then NxRaiseContractError "비교 셀 주소가 올바르지 않습니다."
                mRows(index, 8) = "1": mRows(index, 9) = "1"
                CountCharacters index, chars
            Next row
        End If
    Next sheetIndex
End Sub

Private Function ResolveSummary(ByVal report As Workbook) As Worksheet
    Dim sheet As Worksheet, candidate As Worksheet, active As Object, name As String, count As Long
    Set active = report.ActiveSheet
    If TypeOf active Is Worksheet Then
        If Not active.Range("Q1").HasFormula And Not IsError(active.Range("Q1").Value2) Then
            name = CStr(active.Range("Q1").Value2)
            For Each sheet In report.Worksheets
                If sheet.Name = name Then
                    If Not sheet.Range("J1").HasFormula And Not IsError(sheet.Range("J1").Value2) Then
                        If CStr(sheet.Range("J1").Value2) = "NX_COMPARE_NAV_V1" Then Set ResolveSummary = sheet: Exit Function
                    End If
                End If
            Next sheet
        End If
    End If
    For Each sheet In report.Worksheets
        If Not sheet.Range("J1").HasFormula And Not IsError(sheet.Range("J1").Value2) Then
            If CStr(sheet.Range("J1").Value2) = "NX_COMPARE_NAV_V1" Then count = count + 1: Set candidate = sheet
        End If
    Next sheet
    If count <> 1 Then NxRaiseContractError "탐색할 비교 결과의 요약 또는 차이 시트를 선택하세요."
    Set ResolveSummary = candidate
End Function

Private Function IsReportDetail(ByVal sheet As Worksheet, ByVal summary As Worksheet) As Boolean
    If sheet Is summary Then Exit Function
    If sheet.Range("Q1").HasFormula Or IsError(sheet.Range("Q1").Value2) Then Exit Function
    If Len(CStr(summary.Range("Q1").Value2)) > 0 Then
        IsReportDetail = (CStr(sheet.Range("Q1").Value2) = summary.Name And Left$(sheet.Name, 3) = "셀차이")
    Else
        IsReportDetail = IsDetailName(sheet.Name)
    End If
End Function

Private Sub CountCharacters(ByVal index As Long, ByRef total As Double)
    Dim column As Long
    For column = 1 To 9
        If Len(mRows(index, column)) > 131072 Then NxRaiseContractError "비교 결과 텍스트가 너무 깁니다."
        total = total + Len(mRows(index, column))
    Next column
    If total > 16777216# Then NxRaiseContractError "결과 탐색 텍스트가 너무 큽니다. 보고서에서 확인하세요."
End Sub

Public Property Get Rows() As Variant
    Rows = mRows
End Property

Public Function IsAlive() As Boolean
    Dim candidate As Workbook
    If mReleased Or mReport Is Nothing Then Exit Function
    On Error GoTo Closed
    For Each candidate In Application.Workbooks
        If candidate Is mReport Then IsAlive = True: Exit Function
    Next candidate
Closed:
End Function

Public Function Navigate(ByVal index As Long, ByVal side As Long) As String
    Dim sheet As Worksheet, summary As Worksheet, candidate As Worksheet
    Dim values As Variant, column As Long, row As Long, linkColumn As Long, issue As String, link As Hyperlink
    Dim found As Boolean
    Navigate = "보고서가 닫히거나 변경됐습니다. 결과 탐색을 다시 여세요."
    If mNavigating Or Not IsAlive() Then Exit Function
    If index < 1 Or index > mCount Or (side <> 0 And side <> 1) Then Exit Function
    If CStr(mRows(index, 8 + side)) <> "1" Then Navigate = "해당 원본에는 이 시트가 없습니다.": Exit Function
    mNavigating = True
    On Error GoTo Failed
    Set sheet = mSheets.Item(CLng(mLocations(index, 1))): Set summary = mSheets.Item(1)
    For Each candidate In mReport.Worksheets
        If candidate Is sheet Then found = True
    Next candidate
    If Not found Then GoTo Finished
    found = False
    For Each candidate In mReport.Worksheets
        If candidate Is summary Then found = True
    Next candidate
    If Not found Then GoTo Finished
    If Literal(summary.Range("B2")) <> mBasePath Or Literal(summary.Range("B3")) <> mComparePath Then GoTo Finished
    row = CLng(mLocations(index, 2))
    If sheet Is summary Then
        values = LiteralBlock(sheet.Cells(row, 1).Resize(1, 6)): linkColumn = 5 + side
    Else
        values = LiteralBlock(sheet.Cells(row, 1).Resize(1, 7)): linkColumn = 8 + side
    End If
    For column = 1 To UBound(values, 2)
        If CStr(values(1, column)) <> CStr(mExpected(index, column)) Then GoTo Finished
    Next column
    If Not sheet Is summary Then
        If sheet.Cells(row, 4 + side).Hyperlinks.Count = 1 Then
            If NxWorkbookCompareNavigateSnapshot(sheet, row, side) Then Navigate = "비교 사본의 " & CStr(mRows(index, 2)) & "로 이동했습니다."
            GoTo Finished
        End If
    End If
    If sheet.Cells(row, linkColumn).Hyperlinks.Count <> 1 Then GoTo Finished
    Set link = sheet.Cells(row, linkColumn).Hyperlinks(1)
    If NxWorkbookCompareNavigate(sheet, link, issue) Then
        Navigate = IIf(side = 0, "기준", "비교") & " 원본 " & CStr(mRows(index, 1)) & "!" & CStr(mRows(index, 2)) & "로 이동했습니다."
    ElseIf Len(issue) > 0 Then
        Navigate = issue
    End If
Finished:
    mNavigating = False
    Exit Function
Failed:
    Navigate = "보고서 또는 원본을 확인할 수 없습니다. 작업을 마친 뒤 다시 열어 주세요."
    Resume Finished
End Function

Public Sub ReleaseSession()
    mReleased = True
    Set mReport = Nothing: Set mSheets = Nothing
    mRows = Empty: mExpected = Empty: mLocations = Empty
End Sub

Private Function Literal(ByVal cell As Range) As String
    If cell.HasFormula Or IsError(cell.Value2) Then NxRaiseContractError "비교 보고서 필드는 일반 값이어야 합니다."
    Literal = CStr(cell.Value2)
End Function

Private Function LiteralBlock(ByVal block As Range) As Variant
    Dim hasFormula As Variant, values As Variant, row As Long, column As Long
    hasFormula = block.HasFormula
    If IsNull(hasFormula) Then NxRaiseContractError "비교 보고서 필드에는 수식을 사용할 수 없습니다."
    If CBool(hasFormula) Then NxRaiseContractError "비교 보고서 필드에는 수식을 사용할 수 없습니다."
    values = block.Value2
    For row = 1 To UBound(values, 1)
        For column = 1 To UBound(values, 2)
            If IsError(values(row, column)) Then NxRaiseContractError "비교 보고서의 오류값을 확인하세요."
        Next column
    Next row
    LiteralBlock = values
End Function

Private Function IsDetailName(ByVal name As String) As Boolean
    Dim suffix As String
    If name = "셀차이" Then IsDetailName = True: Exit Function
    If Left$(name, 4) <> "셀차이_" Then Exit Function
    suffix = Mid$(name, 5)
    If Len(suffix) = 0 Or Len(suffix) > 5 Then Exit Function
    If suffix Like "*[!0-9]*" Then Exit Function
    IsDetailName = CLng(suffix) >= 2 And CLng(suffix) <= 32767 And suffix = CStr(CLng(suffix))
End Function
