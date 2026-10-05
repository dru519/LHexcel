Attribute VB_Name = "NxPrivacyScan"
Option Explicit

#If Mac Then
#ElseIf VBA7 Then
Private Declare PtrSafe Function NxPrivacyEscState Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Private Declare PtrSafe Function NxPrivacyForeground Lib "user32" Alias "GetForegroundWindow" () As LongPtr
#Else
Private Declare Function NxPrivacyEscState Lib "user32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer
Private Declare Function NxPrivacyForeground Lib "user32" Alias "GetForegroundWindow" () As Long
#End If
Private mKeyboardCancelActive As Boolean
Private mCancelWasForeground As Boolean

Public Sub NxPrivacyBeginCancellation()
    mCancelWasForeground = False
#If Mac Then
#Else
    Call NxPrivacyEscState(27)
#End If
    mKeyboardCancelActive = True
End Sub

Public Sub NxPrivacyEndCancellation()
    mKeyboardCancelActive = False
    mCancelWasForeground = False
End Sub

Public Sub NxPrivacyYield()
    NxPrivacyCheckEscape
    DoEvents
    NxPrivacyCheckEscape
End Sub

Private Sub NxPrivacyCheckEscape()
    Dim keyState As Integer, isForeground As Boolean, cancel As Boolean
    If Not mKeyboardCancelActive Then Exit Sub
#If Mac Then
#Else
    isForeground = (NxPrivacyForeground() = Application.hWnd)
    keyState = NxPrivacyEscState(27)
    If isForeground Then
        cancel = (keyState < 0)
        ' Best-effort short-tap fallback, alongside Excel's native error handler
        ' and current down state. Other processes may consume the historical bit.
        If mCancelWasForeground And (keyState And 1) <> 0 Then cancel = True
    End If
    mCancelWasForeground = isForeground
    If cancel Then Err.Raise 18, "NxPrivacyScan", "개인정보 검사를 중단했습니다."
#End If
End Sub

Public Function NxPrivacyScanSelection(ByVal source As Range) As Collection
    Dim findings As New Collection
    Dim cell As Range
    Dim kind As String
    Dim masked As Variant
    Dim record As Collection
    If source Is Nothing Or source.Areas.Count <> 1 Then NxRaiseContractError "개인정보를 점검할 연속 범위를 선택하세요."
    If source.CountLarge > 49999 Then NxRaiseContractError "개인정보 점검은 한 번에 49,999셀까지 지원합니다."
    For Each cell In source.Cells
        kind = vbNullString
        masked = NxPrivacyMaskedText(cell.Value2, kind)
        If Len(kind) > 0 Then
            Set record = New Collection
            record.Add kind
            record.Add source.Worksheet.Name
            record.Add cell.Address(False, False, xlA1, False)
            record.Add masked
            findings.Add record
        End If
    Next cell
    Set NxPrivacyScanSelection = findings
End Function

Public Function NxPrivacyScanWorkbook(ByVal book As Workbook, ByRef checkedCells As Double, ByRef incomplete As Collection) As Collection
    Dim findings As New Collection, sheet As Worksheet, issue As String
    If book Is Nothing Then NxRaiseContractError "점검할 통합문서가 없습니다."
    If book Is ThisWorkbook Then NxRaiseContractError "내엑셀 추가기능 파일은 점검 대상이 아닙니다."
    checkedCells = 0
    Set incomplete = New Collection
    ' Do not filter Visible: Hidden and VeryHidden sheets are deliberately included.
    For Each sheet In book.Worksheets
        issue = NxPrivacyScanWorksheet(sheet, findings, checkedCells)
        If Len(issue) > 0 Then incomplete.Add Array(sheet.Name, issue)
    Next sheet
    Set NxPrivacyScanWorkbook = findings
End Function

Private Function NxPrivacyScanWorksheet(ByVal sheet As Worksheet, ByVal findings As Collection, ByRef checkedCells As Double) As String
    Dim used As Range, expected As Double, visibility As String, issue As String
    Dim checkedBefore As Double, visitedBlocks As Long, headers As Object
    On Error GoTo Failed
    Set used = sheet.UsedRange
    expected = Application.WorksheetFunction.CountA(used)
    If expected = 0 Then Exit Function
    If checkedCells + expected > 200000 Then
        NxPrivacyScanWorksheet = "전체 검사 한도 200,000셀을 초과해 이 시트는 검사하지 못했습니다."
        Exit Function
    End If
    checkedBefore = checkedCells
    Select Case sheet.Visible
        Case xlSheetVisible: visibility = "표시"
        Case xlSheetHidden: visibility = "숨김"
        Case xlSheetVeryHidden: visibility = "매우 숨김"
    End Select
    Set headers = CreateObject("Scripting.Dictionary")
    issue = NxPrivacyScanBlock(used, sheet.Name, visibility, findings, checkedCells, visitedBlocks, headers)
    If Len(issue) > 0 Then NxPrivacyScanWorksheet = issue: Exit Function
    If checkedCells - checkedBefore <> expected Then NxRaiseContractError "검사 셀 수가 일치하지 않습니다."
    Exit Function
Failed:
    If Err.Number = 18 Then Err.Raise 18, "NxPrivacyScan", "개인정보 검사를 중단했습니다."
    NxPrivacyScanWorksheet = "셀 값을 끝까지 읽지 못했습니다. 오류 " & CStr(Err.Number)
    Err.Clear
End Function

Private Function NxPrivacyScanBlock(ByVal block As Range, ByVal sheetName As String, ByVal visibility As String, _
    ByVal findings As Collection, ByRef checkedCells As Double, ByRef visitedBlocks As Long, ByVal headers As Object) As String
    Dim rowCount As Long, columnCount As Long, half As Long, rowIndex As Long, columnIndex As Long
    Dim values As Variant, value As Variant, kind As String, issue As String, cell As Range
    Dim header As String, columnKey As String, contextKind As String, date1904 As Boolean, firstColumn As Long
    visitedBlocks = visitedBlocks + 1
    NxPrivacyYield
    If visitedBlocks > 8192 Then
        NxPrivacyScanBlock = "데이터가 넓게 분산되어 영역 확인 한도에 도달했습니다. 이 시트의 검사는 미완료입니다."
        Exit Function
    End If
    If Application.WorksheetFunction.CountA(block) = 0 Then Exit Function
    rowCount = block.Rows.Count: columnCount = block.Columns.Count
    ' SpecialCells can dirty Saved and rejects protected sheets. CountA/Value2
    ' preserve source state. Split large sparse extents before materializing values.
    If block.CountLarge > 4096 Then
        If rowCount >= columnCount Then
            half = rowCount \ 2
            issue = NxPrivacyScanBlock(block.Resize(half, columnCount), sheetName, visibility, findings, checkedCells, visitedBlocks, headers)
            If Len(issue) = 0 Then issue = NxPrivacyScanBlock(block.Cells(half + 1, 1).Resize(rowCount - half, columnCount), sheetName, visibility, findings, checkedCells, visitedBlocks, headers)
        Else
            half = columnCount \ 2
            issue = NxPrivacyScanBlock(block.Resize(rowCount, half), sheetName, visibility, findings, checkedCells, visitedBlocks, headers)
            If Len(issue) = 0 Then issue = NxPrivacyScanBlock(block.Cells(1, half + 1).Resize(rowCount, columnCount - half), sheetName, visibility, findings, checkedCells, visitedBlocks, headers)
        End If
        NxPrivacyScanBlock = issue
        Exit Function
    End If
    values = block.Value2
    date1904 = block.Worksheet.Parent.Date1904
    firstColumn = block.Column
    For rowIndex = 1 To rowCount
        For columnIndex = 1 To columnCount
            If block.CountLarge = 1 Then value = values Else value = values(rowIndex, columnIndex)
            ' Formula results containing "" are strings, not Empty, so remain counted.
            If Not IsEmpty(value) Then
                checkedCells = checkedCells + 1
                If checkedCells Mod 256 = 0 Then NxPrivacyYield
                columnKey = CStr(firstColumn + columnIndex - 1)
                header = NxPrivacyHeaderKind(value)
                kind = vbNullString
                If Len(header) > 0 Then
                    headers(columnKey) = header
                Else
                    contextKind = vbNullString
                    If headers.Exists(columnKey) Then contextKind = headers(columnKey)
                    kind = NxPrivacyAuditKind(value, contextKind, date1904)
                    If kind = "이름 후보" Then
                        If Not NxPrivacyHasNearbyIdentity(block.Cells(rowIndex, columnIndex)) Then kind = vbNullString
                    End If
                End If
                If Len(kind) > 0 Then
                    If findings.Count >= 25000 Then
                        NxPrivacyScanBlock = "의심 항목 표시 한도 25,000건에 도달해 이 시트의 나머지 셀은 검사하지 못했습니다."
                        Exit Function
                    End If
                    Set cell = block.Cells(rowIndex, columnIndex)
                    findings.Add Array(kind, sheetName, cell.Row, cell.Address(False, False), visibility, CStr(value))
                End If
            End If
        Next columnIndex
    Next rowIndex
End Function

Private Function NxPrivacyHasNearbyIdentity(ByVal cell As Range) As Boolean
    Dim nearby As Range, values As Variant, r As Long, c As Long, kind As String
    Set nearby = cell.Worksheet.Range(cell.Worksheet.Cells(Application.Max(1, cell.Row - 2), Application.Max(1, cell.Column - 2)), _
        cell.Worksheet.Cells(Application.Min(cell.Worksheet.Rows.Count, cell.Row + 2), Application.Min(cell.Worksheet.Columns.Count, cell.Column + 2)))
    values = nearby.Value2
    For r = 1 To UBound(values, 1)
        For c = 1 To UBound(values, 2)
            kind = NxPrivacyAuditKind(values(r, c), vbNullString)
            If kind = "연락처" Or kind = "주민등록번호" Then NxPrivacyHasNearbyIdentity = True: Exit Function
        Next c
    Next r
End Function

Private Function NxPrivacyHeaderKind(ByVal value As Variant) As String
    Dim text As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    text = LCase$(Replace$(Replace$(Trim$(CStr(value)), " ", ""), vbLf, ""))
    Select Case text
        Case "이름", "성명", "고객명", "담당자명", "직원명", "성명(한글)", "name": NxPrivacyHeaderKind = "이름"
        Case "생년월일", "생일", "출생일", "생년월일(8자리)", "birthdate", "dateofbirth": NxPrivacyHeaderKind = "생년월일"
        Case "주민등록번호", "주민번호": NxPrivacyHeaderKind = "주민등록번호"
        Case "연락처", "전화번호", "휴대폰", "휴대폰번호", "휴대전화", "휴대전화번호", "핸드폰", "전화", "tel", "phone": NxPrivacyHeaderKind = "연락처"
    End Select
End Function

' Audit-only policy. Do not narrow the separate masking or AI safety detector.
Public Function NxPrivacyAuditKind(ByVal value As Variant, ByVal header As String, Optional ByVal date1904 As Boolean = False) As String
    Dim text As String, digits As String, parsed As Date, dateInput As Variant
    Dim index As Long, code As Long, prefix As String, validPhone As Boolean
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    text = Trim$(CStr(value))
    If Len(text) = 0 Or Len(NxPrivacyHeaderKind(value)) > 0 Then Exit Function
    If header = "이름" Then
        If Len(text) < 2 Or Len(text) > 5 Then Exit Function
        For index = 1 To Len(text)
            code = AscW(Mid$(text, index, 1))
            If code < 0 Then code = code + 65536
            If code < 44032 Or code > 55203 Then Exit Function
        Next index
        NxPrivacyAuditKind = "이름": Exit Function
    End If
    If header = "생년월일" Then
        dateInput = value
        If IsNumeric(value) Then
            If CDbl(value) > 0 And CDbl(value) < 100000 And date1904 Then dateInput = CDbl(value) + 1462
        End If
        If NxDateTryParseValue(dateInput, Date, parsed) Then
            If parsed <= Date And Year(parsed) >= 1900 Then NxPrivacyAuditKind = "생년월일"
        End If
        Exit Function
    End If
    If Len(header) = 0 And Len(text) >= 2 And Len(text) <= 4 Then
        If InStr("김이박최정강조윤장임한오서신권황안송전홍유고문양손배백허남심노하곽성차주우구민진나지엄채원천방공현함변염여추도소석선설마길연위표명기반왕금옥육인맹제모탁국어은편용예경봉사부", Left$(text, 1)) > 0 Then
            For index = 1 To Len(text)
                code = AscW(Mid$(text, index, 1))
                If code < 0 Then code = code + 65536
                If code < 44032 Or code > 55203 Then Exit For
            Next index
            If index > Len(text) Then NxPrivacyAuditKind = "이름 후보": Exit Function
        End If
    End If
    ' Retain masked identifiers as findings; validate the visible birth date.
    digits = Replace$(text, "-", "")
    If Len(digits) = 13 And InStr(digits, "*") > 0 Then
        If Right$(digits, 7) = "*******" Or (Mid$(digits, 7, 1) >= "1" And Mid$(digits, 7, 1) <= "4" And Right$(digits, 6) = "******") Then
            If NxDateTryParseValue(text, Date, parsed) Then
                If parsed <= Date And Year(parsed) >= 1900 Then NxPrivacyAuditKind = "주민등록번호"
            End If
        End If
        Exit Function
    End If
    digits = Replace$(Replace$(Replace$(Replace$(text, "-", ""), " ", ""), "(", ""), ")", "")
    If Len(digits) = 0 Or digits Like "*[!0-9]*" Then Exit Function
    If Len(digits) = 13 And Mid$(digits, 7, 1) >= "1" And Mid$(digits, 7, 1) <= "4" Then
        If NxDateTryParseValue(digits, Date, parsed) Then
            If parsed <= Date And Year(parsed) >= 1900 Then NxPrivacyAuditKind = "주민등록번호"
        End If
        Exit Function
    End If
    ' A missing leading zero is restored only in an explicitly labelled phone column.
    If header = "연락처" And Left$(digits, 1) <> "0" And (Len(digits) = 9 Or Len(digits) = 10) Then digits = "0" & digits
    prefix = Left$(digits, 3)
    If Left$(digits, 2) = "02" Then
        validPhone = (Len(digits) = 9 Or Len(digits) = 10)
    Else
        Select Case prefix
            Case "010", "070", "050": validPhone = (Len(digits) = 11)
            Case "011", "016", "017", "018", "019", "031", "032", "033", "041", "042", "043", "044", "051", "052", "053", "054", "055", "061", "062", "063", "064"
                validPhone = (Len(digits) = 10 Or Len(digits) = 11)
        End Select
    End If
    If validPhone Then NxPrivacyAuditKind = "연락처"
End Function

Public Function NxPrivacyScanIdentity(ByVal findings As Collection, ByVal incomplete As Collection, ByVal checkedCells As Double) As String
    Dim record As Variant, value As Variant, result As String
    result = CStr(checkedCells) & "|"
    For Each record In findings
        For Each value In record
            result = result & CStr(Len(CStr(value))) & ":" & CStr(value)
        Next value
    Next record
    For Each record In incomplete
        result = result & "!" & CStr(Len(CStr(record(0)))) & ":" & CStr(record(0)) & CStr(record(1))
    Next record
    NxPrivacyScanIdentity = result
End Function

Public Function NxPrivacyCreateWorkbookReport(ByVal sourceBook As Workbook, ByVal findings As Collection, ByVal incomplete As Collection, ByVal checkedCells As Double) As Worksheet
    Dim reportBook As Workbook, report As Worksheet, record As Variant
    Dim rowIndex As Long, errorNumber As Long, errorDescription As String
    Dim sourceSheet As Worksheet, sourceCell As Range, linkAddress As String, linkTarget As String, link As Hyperlink
    On Error GoTo Failed
    If (findings.Count + incomplete.Count + 6) * 6 > 200000 Then NxRaiseContractError "개인정보 점검 결과가 표시 한도를 초과했습니다."
    Set reportBook = Workbooks.Add(xlWBATWorksheet)
    Set report = reportBook.Worksheets(1)
    report.Name = "개인정보"
    report.Cells.NumberFormat = "@"
    report.Range("A1").Value2 = "개인정보"
    report.Range("B1").Value2 = sourceBook.Name
    report.Range("C1").Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    report.Range("A2").Value2 = "검사 범위"
    report.Range("B2").Value2 = "이름·생년월일·주민등록번호·연락처 (숨김 시트 포함)"
    report.Range("A3").Value2 = "검사 결과"
    If incomplete.Count > 0 Then
        report.Range("B3").Value2 = "검사 미완료: 아래 제외·실패 내역을 확인하세요. 개인정보 없음으로 판단할 수 없습니다."
    ElseIf findings.Count = 0 Then
        report.Range("B3").Value2 = "검사한 범위에서 개인정보 의심 항목이 발견되지 않았습니다."
    Else
        report.Range("B3").Value2 = "개인정보 의심 항목 " & CStr(findings.Count) & "건. 실제 개인정보 여부를 확인하세요."
    End If
    report.Range("A4").Value2 = "검사 셀 수"
    report.Range("B4").Value2 = checkedCells
    report.Range("C4").Value2 = "외부 전송 없음 · 원본 변경 없음"
    report.Range("A5").Value2 = "원본 위치를 클릭해 확인하세요. 저장하지 않은 원본은 열려 있어야 합니다. 숨김 시트는 직접 표시한 뒤 이동하세요."
    report.Range("A6").Value2 = "유형": report.Range("B6").Value2 = "발견 값"
    report.Range("C6").Value2 = "원본 위치"
    rowIndex = 7
    For Each record In findings
        If rowIndex Mod 64 = 0 Then NxPrivacyYield
        report.Cells(rowIndex, 1).Value2 = record(0)
        report.Cells(rowIndex, 2).Value2 = record(5)
        report.Cells(rowIndex, 3).Value2 = CStr(record(1)) & "!" & CStr(record(3))
        Set sourceSheet = sourceBook.Worksheets(CStr(record(1)))
        Set sourceCell = sourceSheet.Range(CStr(record(3)))
        If sourceSheet.Visible = xlSheetVisible Then
            linkAddress = vbNullString
            If Len(sourceBook.Path) > 0 Then
                linkAddress = sourceBook.FullName
                linkTarget = "'" & Replace$(sourceSheet.Name, "'", "''") & "'!" & sourceCell.Address
            Else
                linkTarget = sourceCell.Address(True, True, xlA1, True)
            End If
            report.Hyperlinks.Add Anchor:=report.Cells(rowIndex, 3), Address:=linkAddress, _
                SubAddress:=linkTarget, TextToDisplay:=CStr(record(1)) & "!" & CStr(record(3)), _
                ScreenTip:="원본 셀로 이동합니다. 원본 이름이나 저장 위치가 바뀌면 다시 검사하세요."
            If report.Cells(rowIndex, 3).Hyperlinks.Count <> 1 Then NxRaiseContractError "원본 위치 링크를 생성하지 못했습니다."
        Else
            report.Cells(rowIndex, 3).Value2 = CStr(record(1)) & "!" & CStr(record(3)) & " [숨김 · 시트 표시 후 재검사]"
        End If
        rowIndex = rowIndex + 1
    Next record
    For Each record In incomplete
        report.Cells(rowIndex, 1).Value2 = "검사 미완료"
        report.Cells(rowIndex, 3).Value2 = record(0)
        report.Cells(rowIndex, 2).Value2 = record(1)
        rowIndex = rowIndex + 1
    Next record
    report.Range("A6:C" & CStr(Application.Max(7, rowIndex - 1))).AutoFilter
    NxDrawStyleGeneratedTable reportBook, report.Range("A6:C" & CStr(Application.Max(7, rowIndex - 1)))
    ' Generated table styling must not make navigation look like ordinary text.
    For Each link In report.Hyperlinks
        link.Range.Font.Color = RGB(0, 102, 204)
        link.Range.Font.Underline = xlUnderlineStyleSingle
    Next link
    report.Rows(1).Font.Bold = True: report.Rows(6).Font.Bold = True
    report.Columns("A").ColumnWidth = 20: report.Columns("B").ColumnWidth = 36
    report.Columns("C").ColumnWidth = 48
    report.Range("A7:C" & CStr(Application.Max(7, rowIndex - 1))).WrapText = True
    report.Range("A7:C" & CStr(Application.Max(7, rowIndex - 1))).Rows.AutoFit
    report.Range("A1").Select
    Set NxPrivacyCreateWorkbookReport = report
    Exit Function
Failed:
    errorNumber = Err.Number: errorDescription = Err.Description
    On Error Resume Next
    If Not reportBook Is Nothing Then reportBook.Close SaveChanges:=False
    On Error GoTo 0
    If errorNumber = 0 Then errorNumber = NX_CONTRACT_ERROR
    Err.Raise errorNumber, "NxPrivacyScan", errorDescription
End Function

Public Function NxPrivacyCreateScanReport(ByVal source As Range) As Worksheet
    Dim findings As Collection, incomplete As Collection, checkedCells As Double
    If source Is Nothing Then NxRaiseContractError "점검할 통합문서가 없습니다."
    Set findings = NxPrivacyScanWorkbook(source.Worksheet.Parent, checkedCells, incomplete)
    Set NxPrivacyCreateScanReport = NxPrivacyCreateWorkbookReport(source.Worksheet.Parent, findings, incomplete, checkedCells)
End Function
