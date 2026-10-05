Attribute VB_Name = "NxWorkbookCompare"
Option Explicit

Private Const NX_COMPARE_MAX_CELLS As Long = 200000
Private Const NX_COMPARE_DETAIL_LAST_ROW As Long = 32000 ' Two links per row; Excel limit is 65,530 per worksheet.
Private Const NX_COMPARE_NAV_MARKER As String = "NX_COMPARE_NAV_V1"
Private mCompareNavigating As Boolean
Private mCompareRunning As Boolean, mCompareCancel As Boolean
Private mCompareProgress As Object
Private mOpenReport As Workbook
Private mSnapshotLeft As Worksheet, mSnapshotRight As Worksheet
#If VBA7 Then
    Private Declare PtrSafe Sub CompareSleep Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
#Else
    Private Declare Sub CompareSleep Lib "kernel32" Alias "Sleep" (ByVal milliseconds As Long)
#End If

Public Sub NxWorkbookCompareBindProgress(ByVal view As Object)
    If mCompareRunning Then NxRaiseContractError "이미 통합문서를 비교하고 있습니다."
    Set mCompareProgress = view
End Sub

Public Sub NxWorkbookCompareCancel()
    If mCompareRunning Then mCompareCancel = True
End Sub

Private Sub CompareProgress(ByVal stage As String, ByVal done As Long, ByVal total As Long)
    Application.StatusBar = stage & " · " & CStr(done) & " / " & CStr(total) & " · Esc로 취소"
    If Not mCompareProgress Is Nothing Then mCompareProgress.UpdateCompareProgress stage, done, total
    DoEvents
    If mCompareCancel Then Err.Raise 18, "LHexcel.File.Compare", "통합문서 비교를 취소했습니다."
End Sub

Public Function NxWorkbookCompareReport(ByVal basePath As String, ByVal comparePath As String, _
    ByVal outputPath As String, ByVal compareFormats As Boolean, Optional ByVal baseSheetName As String = "", Optional ByVal compareSheetName As String = "") As CNxResult
    Dim baseBook As Workbook, compareBook As Workbook, reportBook As Workbook
    Dim baseOwned As Boolean, compareOwned As Boolean
    Dim temporaryPath As String, errorNumber As Long, errorDescription As String
    Dim priorStatus As Variant, priorCancel As XlEnableCancelKey
    If mCompareRunning Then NxRaiseContractError "이미 통합문서를 비교하고 있습니다."
    priorStatus = Application.StatusBar: priorCancel = Application.EnableCancelKey
    mCompareRunning = True: mCompareCancel = False
    On Error GoTo Failed
    Application.EnableCancelKey = xlErrorHandler
    CompareProgress "비교 파일 확인", 0, 1
    NxWorkbookCompareValidatePaths basePath, comparePath, outputPath
    Set baseBook = NxSafeWorkbookOpen(basePath, baseOwned)
    Set compareBook = NxSafeWorkbookOpen(comparePath, compareOwned)
    ' All bounds are established before Workbooks.Add so failed preflight leaves no report.
    NxWorkbookComparePreflight baseBook, compareBook, baseSheetName, compareSheetName
    Set reportBook = NxWorkbookCompareBuildReport(baseBook, compareBook, basePath, comparePath, compareFormats, baseSheetName, compareSheetName)
    If Len(outputPath) = 0 Then
        Set mOpenReport = reportBook
        Set reportBook = Nothing
        NxSafeWorkbookClose compareBook, compareOwned
        NxSafeWorkbookClose baseBook, baseOwned
        CompareRestoreState priorStatus, priorCancel
        mCompareRunning = False: Set mCompareProgress = Nothing
        Set NxWorkbookCompareReport = NxCreateResult(NX_FEATURE_FILE_WORKBOOK_COMPARE, NxSuccess, _
            "complete", mOpenReport.Name, False, vbNullString, "workbook_compare_complete")
        Exit Function
    End If
    CompareProgress "보고서 저장", 1, 1
    temporaryPath = NxWorkbookCompareTemporaryPath(outputPath)
    reportBook.SaveAs Filename:=temporaryPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    reportBook.Close SaveChanges:=False
    Set reportBook = Nothing
    NxWorkbookCompareAtomicPublish temporaryPath, outputPath
    NxSafeWorkbookClose compareBook, compareOwned
    NxSafeWorkbookClose baseBook, baseOwned
    Set NxWorkbookCompareReport = NxCreateResult(NX_FEATURE_FILE_WORKBOOK_COMPARE, NxSuccess, _
        "complete", outputPath, False, vbNullString, "workbook_compare_complete")
    CompareRestoreState priorStatus, priorCancel
    mCompareRunning = False: Set mCompareProgress = Nothing
    Exit Function
Failed:
    errorNumber = Err.Number: errorDescription = Err.Description: Err.Clear
    On Error Resume Next
    If Not reportBook Is Nothing Then reportBook.Close SaveChanges:=False
    If Not mOpenReport Is Nothing Then mOpenReport.Close SaveChanges:=False
    Set mOpenReport = Nothing
    If compareOwned Then
        If Not compareBook Is Nothing Then NxSafeWorkbookClose compareBook, compareOwned
    End If
    If baseOwned Then
        If Not baseBook Is Nothing Then NxSafeWorkbookClose baseBook, baseOwned
    End If
    If Len(temporaryPath) > 0 Then If Len(Dir$(temporaryPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then Kill temporaryPath
    CompareRestoreState priorStatus, priorCancel
    mCompareRunning = False: Set mCompareProgress = Nothing
    On Error GoTo 0
    Err.Raise errorNumber, "LHexcel.File.Compare", errorDescription
End Function

Public Sub NxWorkbookCompareFinishReport(ByVal success As Boolean)
    If mOpenReport Is Nothing Then Exit Sub
    If success Then
        mOpenReport.Activate
        mOpenReport.Worksheets("비교요약").Activate
        mOpenReport.Worksheets("비교요약").Range("A1").Select
    Else
        mOpenReport.Close SaveChanges:=False
    End If
    Set mOpenReport = Nothing
End Sub

Private Sub CompareRestoreState(ByVal priorStatus As Variant, ByVal priorCancel As XlEnableCancelKey)
    ' Excel can coerce a Boolean Variant into visible "FALSE" text on assignment.
    If IsEmpty(priorStatus) Or VarType(priorStatus) = vbBoolean Then
        Application.StatusBar = False
    Else
        Application.StatusBar = priorStatus
    End If
    Application.EnableCancelKey = priorCancel
End Sub

Private Sub NxWorkbookCompareValidatePaths(ByVal basePath As String, ByVal comparePath As String, ByVal outputPath As String)
    If Len(Trim$(basePath)) = 0 Or Len(Dir$(basePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then NxRaiseContractError "기준 통합문서를 찾을 수 없습니다."
    If Len(Trim$(comparePath)) = 0 Or Len(Dir$(comparePath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then NxRaiseContractError "비교 통합문서를 찾을 수 없습니다."
    If StrComp(basePath, comparePath, vbTextCompare) = 0 Then NxRaiseContractError "서로 다른 통합문서를 선택하세요."
    If Len(outputPath) = 0 Then Exit Sub
    If LCase$(Right$(Trim$(outputPath), 5)) <> ".xlsx" Then NxRaiseContractError "비교 보고서는 .xlsx로 저장해야 합니다."
    NxWorkbookCompareRequireVacantOutput outputPath
End Sub

Private Sub NxWorkbookCompareRequireVacantOutput(ByVal outputPath As String)
    Dim folderPath As String, separator As Long
    separator = InStrRev(outputPath, Application.PathSeparator)
    If separator = 0 Then NxRaiseContractError "비교 보고서 저장 폴더를 확인하세요."
    folderPath = Left$(outputPath, separator - 1)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then NxRaiseContractError "비교 보고서 저장 폴더를 찾을 수 없습니다."
    If Len(Dir$(outputPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then NxRaiseContractError "같은 이름의 비교 보고서가 이미 있습니다. 다른 이름을 선택하세요."
End Sub

Private Sub NxWorkbookComparePreflight(ByVal baseBook As Workbook, ByVal compareBook As Workbook, ByVal baseSheetName As String, ByVal compareSheetName As String)
    Dim names As Object, sheetName As Variant, baseSheet As Worksheet, compareSheet As Worksheet
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, lastColumn As Long, sheetCells As Double, totalCells As Double
    Set names = NxWorkbookComparePairs(baseBook, compareBook, baseSheetName, compareSheetName)
    For Each sheetName In names.Keys
        Set baseSheet = NxWorkbookCompareWorksheet(baseBook, CStr(sheetName))
        Set compareSheet = NxWorkbookCompareWorksheet(compareBook, CStr(names(sheetName)))
        If Not baseSheet Is Nothing And Not compareSheet Is Nothing Then
            NxWorkbookCompareUnionBounds baseSheet, compareSheet, firstRow, lastRow, firstColumn, lastColumn, sheetCells
            If sheetCells > NX_COMPARE_MAX_CELLS Then NxRaiseContractError "시트별 비교 범위는 200,000셀을 넘을 수 없습니다: " & CStr(sheetName)
            totalCells = totalCells + sheetCells
            If totalCells > NX_COMPARE_MAX_CELLS Then NxRaiseContractError "전체 비교 범위는 200,000셀을 넘을 수 없습니다."
        Else
            If Not baseSheet Is Nothing Then totalCells = totalCells + NxDataCompareUsedRange(baseSheet).CountLarge
            If Not compareSheet Is Nothing Then totalCells = totalCells + NxDataCompareUsedRange(compareSheet).CountLarge
            If totalCells > NX_COMPARE_MAX_CELLS Then NxRaiseContractError "전체 비교 범위는 200,000셀을 넘을 수 없습니다."
        End If
    Next sheetName
End Sub

Private Function NxWorkbookCompareBuildReport(ByVal baseBook As Workbook, ByVal compareBook As Workbook, _
    ByVal baseSourcePath As String, ByVal compareSourcePath As String, ByVal compareFormats As Boolean, ByVal baseSheetName As String, ByVal compareSheetName As String) As Workbook
    Dim reportBook As Workbook, summarySheet As Worksheet, detailSheet As Worksheet, resultSheet As Worksheet
    Dim names As Object, sheetName As Variant, baseSheet As Worksheet, compareSheet As Worksheet
    Dim summaryRow As Long, detailRow As Long, differenceCount As Long, totalDifferenceCount As Long
    Dim errorNumber As Long, errorDescription As String, firstDetail As Range
    On Error GoTo Failed
    Set reportBook = Workbooks.Add(xlWBATWorksheet)
    NxRouteAvailabilityStartEvents
    Set summarySheet = reportBook.Worksheets(1): summarySheet.Name = "비교요약"
    Set detailSheet = reportBook.Worksheets.Add(After:=summarySheet): detailSheet.Name = "셀차이"
    NxWorkbookCompareInitializeSummary summarySheet, baseSourcePath, compareSourcePath, compareFormats
    NxWorkbookCompareInitializeDetails detailSheet
    Set names = NxWorkbookComparePairs(baseBook, compareBook, baseSheetName, compareSheetName)
    summaryRow = 6: detailRow = 2
    For Each sheetName In names.Keys
        If detailRow > NX_COMPARE_DETAIL_LAST_ROW Then NxWorkbookCompareNextDetailSheet detailSheet, detailRow
        Set firstDetail = detailSheet.Cells(detailRow, 1)
        Set baseSheet = NxWorkbookCompareWorksheet(baseBook, CStr(sheetName))
        Set compareSheet = NxWorkbookCompareWorksheet(compareBook, CStr(names(sheetName)))
        Set mSnapshotLeft = NxCompareSnapshot(reportBook, summarySheet, baseSheet, "기준")
        Set mSnapshotRight = NxCompareSnapshot(reportBook, summarySheet, compareSheet, "비교")
        summarySheet.Cells(summaryRow, 12).Value2 = mSnapshotLeft.Name
        summarySheet.Cells(summaryRow, 13).Value2 = mSnapshotRight.Name
        summarySheet.Cells(summaryRow, 11).Value2 = CStr(names(sheetName))
        summarySheet.Columns("K:M").Hidden = True
        If baseSheet Is Nothing Then
            NxWorkbookCompareWriteSummary summarySheet, summaryRow, CStr(sheetName), "기준 문서에 없음", 1: differenceCount = 1
            NxDataCompareUsedRange(mSnapshotRight).Interior.Color = RGB(255, 199, 206)
        ElseIf compareSheet Is Nothing Then
            NxWorkbookCompareWriteSummary summarySheet, summaryRow, CStr(sheetName), "비교 문서에 없음", 1: differenceCount = 1
            NxDataCompareUsedRange(mSnapshotLeft).Interior.Color = RGB(255, 199, 206)
        Else
            differenceCount = NxWorkbookCompareSheets(baseSheet, compareSheet, detailSheet, detailRow, compareFormats, baseSourcePath, compareSourcePath)
            NxWorkbookCompareWriteSummary summarySheet, summaryRow, CStr(sheetName), IIf(differenceCount = 0, "일치", "차이 있음"), differenceCount
        End If
        If Not baseSheet Is Nothing And Not compareSheet Is Nothing And differenceCount > 0 Then
            summarySheet.Hyperlinks.Add Anchor:=summarySheet.Cells(summaryRow, 4), Address:=vbNullString, _
                SubAddress:=NxWorkbookCompareSelfAnchor(firstDetail), TextToDisplay:="차이 상세 보기"
        Else
            summarySheet.Cells(summaryRow, 4).Value2 = IIf(differenceCount = 0, "동일", "시트 추가/삭제")
        End If
        If Not baseSheet Is Nothing Then NxWorkbookCompareWriteCellLink summarySheet.Cells(summaryRow, 5), baseSourcePath, CStr(sheetName), "A1", "기준 원본으로 이동"
        If Not compareSheet Is Nothing Then NxWorkbookCompareWriteCellLink summarySheet.Cells(summaryRow, 6), compareSourcePath, CStr(names(sheetName)), "A1", "비교 원본으로 이동"
        totalDifferenceCount = totalDifferenceCount + differenceCount: summaryRow = summaryRow + 1
    Next sheetName
    summarySheet.Range("B4").Value2 = totalDifferenceCount
    summarySheet.Range("A5:F" & CStr(summaryRow - 1)).AutoFilter
    NxDrawStyleGeneratedTable reportBook, summarySheet.Range("A5:F" & CStr(summaryRow - 1))
    summarySheet.Range("D6:F" & CStr(summaryRow - 1)).Font.Color = RGB(0, 102, 204)
    summarySheet.Columns("A:G").AutoFit
    summarySheet.Columns("B:B").ColumnWidth = 26
    summarySheet.Columns("G:G").ColumnWidth = 62
    summarySheet.Range("G2:G3").WrapText = True
    summarySheet.Rows("2:3").AutoFit
    For Each resultSheet In reportBook.Worksheets
        If resultSheet Is summarySheet Or NxWorkbookCompareIsDetailName(resultSheet.Name) Then
            resultSheet.Range("Q1").Value2 = summarySheet.Name
            resultSheet.Columns("Q:S").Hidden = True
        End If
        If NxWorkbookCompareIsDetailName(resultSheet.Name) Then
            resultSheet.Range("A1:O" & CStr(Application.Max(2, resultSheet.UsedRange.Rows.Count))).AutoFilter
            NxDrawStyleGeneratedTable reportBook, resultSheet.Range("A1:O" & CStr(Application.Max(2, resultSheet.UsedRange.Rows.Count)))
            resultSheet.Range("H2:I" & CStr(Application.Max(2, resultSheet.UsedRange.Rows.Count))).Font.Color = RGB(0, 102, 204)
            resultSheet.Columns("A:O").AutoFit
            resultSheet.Columns("D:G").ColumnWidth = 30
            resultSheet.Columns("D:G").WrapText = True
            resultSheet.Columns("F:O").Hidden = True
            resultSheet.Range("D2:E" & CStr(Application.Max(2, resultSheet.UsedRange.Rows.Count))).Font.Color = RGB(0, 102, 204)
        End If
    Next resultSheet
    Set mSnapshotLeft = Nothing: Set mSnapshotRight = Nothing
    Set NxWorkbookCompareBuildReport = reportBook
    Exit Function
Failed:
    errorNumber = Err.Number: errorDescription = Err.Description
    On Error Resume Next
    If Not reportBook Is Nothing Then reportBook.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise errorNumber, "LHexcel.File.Compare.Report", errorDescription
End Function

Private Function NxCompareSnapshot(ByVal book As Workbook, ByVal summary As Worksheet, ByVal source As Worksheet, ByVal prefix As String) As Worksheet
    Dim result As Worksheet, data As Range, target As Range
    Set result = book.Worksheets.Add(Before:=summary)
    result.Name = NxDataNextNormalizationSheetName(book, prefix & " 데이터")
    result.Tab.Color = IIf(prefix = "기준", RGB(91, 155, 213), RGB(237, 125, 49))
    If source Is Nothing Then
        result.Range("A1").Value2 = "해당 시트가 없습니다."
    Else
        Set data = NxDataCompareUsedRange(source)
        If data.CountLarge > NX_COMPARE_MAX_CELLS Then NxRaiseContractError "사본 범위가 200,000셀을 넘습니다: " & source.Name
        Set target = result.Range(data.Address)
        target.NumberFormat = "@"
        target.Value2 = data.Value2
        NxDrawStyleGeneratedTable book, target
    End If
    Set NxCompareSnapshot = result
End Function

Private Function NxWorkbookComparePairs(ByVal baseBook As Workbook, ByVal compareBook As Workbook, ByVal baseName As String, ByVal compareName As String) As Object
    Dim names As Object, key As Variant, sheet As Worksheet
    If Len(baseName) = 0 And Len(compareName) = 0 Then
        Set names = NxWorkbookCompareSheetNames(baseBook, compareBook)
        For Each key In names.Keys: names(key) = CStr(key): Next key
    Else
        Set sheet = NxWorkbookCompareWorksheet(baseBook, baseName)
        If sheet Is Nothing Then NxRaiseContractError "기준 시트를 찾을 수 없습니다: " & baseName
        Set sheet = NxWorkbookCompareWorksheet(compareBook, compareName)
        If sheet Is Nothing Then NxRaiseContractError "비교 시트를 찾을 수 없습니다: " & compareName
        Set names = CreateObject("Scripting.Dictionary")
        names.CompareMode = vbTextCompare
        names.Add baseName, compareName
    End If
    Set NxWorkbookComparePairs = names
End Function

Private Function NxWorkbookCompareSheetNames(ByVal baseBook As Workbook, ByVal compareBook As Workbook) As Object
    Dim names As Object, sheet As Worksheet
    Set names = CreateObject("Scripting.Dictionary"): names.CompareMode = vbTextCompare
    For Each sheet In baseBook.Worksheets
        If Not names.Exists(sheet.Name) Then names.Add sheet.Name, True
    Next sheet
    For Each sheet In compareBook.Worksheets
        If Not names.Exists(sheet.Name) Then names.Add sheet.Name, True
    Next sheet
    Set NxWorkbookCompareSheetNames = names
End Function

Private Function NxWorkbookCompareWorksheet(ByVal book As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set NxWorkbookCompareWorksheet = book.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function NxWorkbookCompareSheets(ByVal baseSheet As Worksheet, ByVal compareSheet As Worksheet, ByRef detailSheet As Worksheet, ByRef detailRow As Long, ByVal compareFormats As Boolean, ByVal baseSourcePath As String, ByVal compareSourcePath As String) As Long
    Dim firstRow As Long, lastRow As Long, firstColumn As Long, lastColumn As Long, sheetCells As Double
    Dim rowIndex As Long, columnIndex As Long, baseCell As Range, compareCell As Range, reason As String
    Dim reasons As Variant, native As Boolean, ordinal As Long, baseRange As Range, compareRange As Range
    Dim leftRows As Variant, rightRows As Variant
    Dim batch(1 To 1024, 1 To 15) As Variant, buffered As Long, batchStart As Long
    Dim leftValues As Variant, rightValues As Variant, leftFormulas As Variant, rightFormulas As Variant
    NxWorkbookCompareUnionBounds baseSheet, compareSheet, firstRow, lastRow, firstColumn, lastColumn, sheetCells
    If sheetCells = 0 Then Exit Function
    Set baseRange = baseSheet.Range(baseSheet.Cells(firstRow, firstColumn), baseSheet.Cells(lastRow, lastColumn))
    Set compareRange = compareSheet.Range(compareSheet.Cells(firstRow, firstColumn), compareSheet.Cells(lastRow, lastColumn))
    leftValues = CompareGridValues(baseRange, False): rightValues = CompareGridValues(compareRange, False)
    leftFormulas = CompareGridValues(baseRange, True): rightFormulas = CompareGridValues(compareRange, True)
    batchStart = detailRow
    native = CompareNativeReasons(baseRange, compareRange, reasons, compareFormats, leftValues, rightValues, leftFormulas, rightFormulas)
    If Not native Then
        leftRows = CompareCapture(baseRange, compareFormats, leftValues, leftFormulas)
        rightRows = CompareCapture(compareRange, compareFormats, rightValues, rightFormulas)
    End If
    For rowIndex = firstRow To lastRow
        For columnIndex = firstColumn To lastColumn
            If native Then
                reason = CStr(reasons(ordinal))
            Else
                reason = NxWorkbookCompareCellDifference(leftRows, rightRows, ordinal)
            End If
            If Len(reason) > 0 Then
                Set baseCell = baseSheet.Cells(rowIndex, columnIndex): Set compareCell = compareSheet.Cells(rowIndex, columnIndex)
                If detailRow > NX_COMPARE_DETAIL_LAST_ROW Then
                    CompareFlushDetails detailSheet, batchStart, batch, buffered, baseSourcePath, compareSourcePath
                    NxWorkbookCompareNextDetailSheet detailSheet, detailRow
                    batchStart = detailRow
                End If
                buffered = buffered + 1
                NxWorkbookCompareWriteDetail batch, buffered, baseSheet.Name, baseCell.Address(False, False), reason, _
                    leftValues(rowIndex - firstRow + 1, columnIndex - firstColumn + 1), rightValues(rowIndex - firstRow + 1, columnIndex - firstColumn + 1), _
                    leftFormulas(rowIndex - firstRow + 1, columnIndex - firstColumn + 1), rightFormulas(rowIndex - firstRow + 1, columnIndex - firstColumn + 1), baseSourcePath, compareSourcePath, _
                    CStr(baseCell.NumberFormat), CStr(compareCell.NumberFormat), NxWorkbookCompareMergeSignature(baseCell), _
                    NxWorkbookCompareMergeSignature(compareCell), baseCell.Text, compareCell.Text
                detailRow = detailRow + 1: NxWorkbookCompareSheets = NxWorkbookCompareSheets + 1
                If buffered = 1024 Then
                    CompareFlushDetails detailSheet, batchStart, batch, buffered, baseSourcePath, compareSourcePath
                    batchStart = detailRow
                End If
            End If
            ordinal = ordinal + 1
            If ordinal Mod 256 = 0 Then CompareProgress IIf(native, "보고서 작성", "셀 비교"), ordinal, CLng(sheetCells)
        Next columnIndex
    Next rowIndex
    CompareFlushDetails detailSheet, batchStart, batch, buffered, baseSourcePath, compareSourcePath
End Function

Private Function CompareGridValues(ByVal target As Range, ByVal formulas As Boolean) As Variant
    Dim values As Variant, scalarGrid(1 To 1, 1 To 1) As Variant
    If formulas Then values = target.Formula Else values = target.Value2
    If target.CountLarge = 1 Then
        scalarGrid(1, 1) = values
        CompareGridValues = scalarGrid
    Else
        CompareGridValues = values
    End If
End Function

Private Sub CompareFlushDetails(ByVal sheet As Worksheet, ByVal firstRow As Long, ByRef batch As Variant, _
    ByRef count As Long, ByVal basePath As String, ByVal comparePath As String)
    Dim values() As Variant, r As Long, c As Long, target As Range, paint As String, address As String
    If count = 0 Then Exit Sub
    ReDim values(1 To count, 1 To 15)
    For r = 1 To count
        For c = 1 To 15: values(r, c) = batch(r, c): Next c
    Next r
    Set target = sheet.Cells(firstRow, 1).Resize(count, 15)
    target.NumberFormat = "@"
    target.Value2 = values
    For r = 1 To count
        sheet.Hyperlinks.Add Anchor:=sheet.Cells(firstRow + r - 1, 4), Address:=vbNullString, SubAddress:=NxWorkbookCompareSelfAnchor(mSnapshotLeft.Range(CStr(values(r, 2)))), TextToDisplay:=CStr(values(r, 4))
        sheet.Hyperlinks.Add Anchor:=sheet.Cells(firstRow + r - 1, 5), Address:=vbNullString, SubAddress:=NxWorkbookCompareSelfAnchor(mSnapshotRight.Range(CStr(values(r, 2)))), TextToDisplay:=CStr(values(r, 5))
        address = CStr(values(r, 2))
        NxWorkbookCompareWriteCellLink sheet.Cells(firstRow + r - 1, 8), basePath, CStr(values(r, 1)), address, "기준 원본으로 이동"
        NxWorkbookCompareWriteCellLink sheet.Cells(firstRow + r - 1, 9), comparePath, CStr(values(r, 1)), address, "비교 원본으로 이동"
        If Len(paint) + Len(address) > 220 Then
            mSnapshotLeft.Range(paint).Interior.Color = RGB(255, 199, 206)
            mSnapshotRight.Range(paint).Interior.Color = RGB(255, 199, 206)
            paint = vbNullString
        End If
        If Len(paint) > 0 Then paint = paint & ","
        paint = paint & address
        If r Mod 128 = 0 Then CompareProgress "원본 이동 링크 작성", r, count
    Next r
    If Len(paint) > 0 Then
        mSnapshotLeft.Range(paint).Interior.Color = RGB(255, 199, 206)
        mSnapshotRight.Range(paint).Interior.Color = RGB(255, 199, 206)
    End If
    count = 0
End Sub

Public Function NxWorkbookCompareNavigateSnapshot(ByVal detail As Worksheet, ByVal row As Long, ByVal side As Long) As Boolean
    Dim summary As Worksheet, target As Worksheet, link As Hyperlink
    Dim i As Long, name As String, address As String
    On Error GoTo Invalid
    If side < 0 Or side > 1 Or row < 2 Then Exit Function
    Set summary = detail.Parent.Worksheets(CStr(detail.Range("Q1").Value2))
    If NxWorkbookCompareReadLiteral(summary.Range("J1")) <> NX_COMPARE_NAV_MARKER Then Exit Function
    For i = 6 To summary.Cells(summary.Rows.Count, 1).End(xlUp).Row
        If NxWorkbookCompareReadLiteral(summary.Cells(i, 1)) = NxWorkbookCompareReadLiteral(detail.Cells(row, 1)) Then
            name = NxWorkbookCompareReadLiteral(summary.Cells(i, 12 + side)): Exit For
        End If
    Next i
    If Len(name) = 0 Then Exit Function
    address = NxWorkbookCompareReadLiteral(detail.Cells(row, 2))
    If Len(NxWorkbookCompareNavigationCellIssue(address)) > 0 Then Exit Function
    Set target = detail.Parent.Worksheets(name)
    Set link = detail.Cells(row, 4 + side).Hyperlinks(1)
    If Len(link.Address) <> 0 Or link.SubAddress <> NxWorkbookCompareSelfAnchor(target.Range(address)) Then Exit Function
    Application.Goto target.Range(address), True
    NxWorkbookCompareNavigateSnapshot = True
Invalid:
End Function

Private Function CompareNativeReasons(ByVal leftRange As Range, ByVal rightRange As Range, ByRef reasons As Variant, ByVal compareFormats As Boolean, _
    ByRef leftValues As Variant, ByRef rightValues As Variant, ByRef leftFormulas As Variant, ByRef rightFormulas As Variant) As Boolean
    Dim service As Object, leftRows As Variant, rightRows As Variant
    Dim job As String, state As String, started As Double, elapsed As Double, detail As String, failure As Long
    Set service = NxHostCreateWorkbookCompare()
    If service Is Nothing Then Exit Function
    leftRows = CompareCapture(leftRange, compareFormats, leftValues, leftFormulas)
    rightRows = CompareCapture(rightRange, compareFormats, rightValues, rightFormulas)
    On Error GoTo Failed
    job = CStr(service.Start(leftRows, rightRows))
    started = Timer
    Do
        state = CStr(service.GetStatus(job))
        If state <> "running" Then Exit Do
        CompareProgress "DLL 값·수식 비교", CLng(service.GetProgress(job)), 100
        CompareSleep 10
        elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400#
        If elapsed > 60 Then NxRaiseContractError "비교 시간이 초과되었습니다."
    Loop
    If state <> "succeeded" Then NxRaiseContractError "DLL 비교가 취소되었거나 실패했습니다."
    reasons = service.GetResult(job)
    If UBound(reasons) - LBound(reasons) + 1 <> leftRange.CountLarge Then NxRaiseContractError "DLL 비교 결과 개수가 올바르지 않습니다."
    CompareNativeReasons = True
    Exit Function
Failed:
    failure = Err.Number: detail = Err.Description
    On Error Resume Next
    If Len(job) > 0 Then service.Cancel job
    On Error GoTo 0
    ' Never repeat an accepted/cancelled job through a second engine.
    Err.Raise failure, "LHexcel.File.Compare", detail
End Function

Private Function CompareCapture(ByVal target As Range, ByVal compareFormats As Boolean, ByRef values As Variant, ByRef formulas As Variant) As Variant
    Dim hasFormula As Variant
    Dim flags() As Boolean, rows() As Variant
    Dim rowHeights() As Variant, columnWidths() As Variant
    Dim rowCount As Long, columnCount As Long, row As Long, column As Long, index As Long
    Dim cell As Range, mergeState As Variant, hasMerges As Boolean
    rowCount = target.Rows.Count: columnCount = target.Columns.Count
    ReDim flags(1 To rowCount, 1 To columnCount)
    ReDim rows(0 To rowCount * columnCount - 1, 0 To 4)
    hasFormula = target.HasFormula
    mergeState = target.MergeCells
    If IsNull(mergeState) Then hasMerges = True Else hasMerges = CBool(mergeState)
    If compareFormats Then
        ReDim rowHeights(1 To rowCount): ReDim columnWidths(1 To columnCount)
        For row = 1 To rowCount: rowHeights(row) = target.Rows(row).EntireRow.RowHeight: Next row
        For column = 1 To columnCount: columnWidths(column) = target.Columns(column).EntireColumn.ColumnWidth: Next column
    End If
    If IsNull(hasFormula) Then
        ' SpecialCells(formulas) dirties saved formula workbooks on supported Excel.
        ' Inspect only formula-shaped strings; HasFormula distinguishes literal =text.
        For row = 1 To rowCount
            For column = 1 To columnCount
                If VarType(formulas(row, column)) = vbString Then
                    If Left$(formulas(row, column), 1) = "=" Then flags(row, column) = CBool(target.Cells(row, column).HasFormula)
                End If
            Next column
        Next row
    ElseIf CBool(hasFormula) Then
        For row = 1 To rowCount
            For column = 1 To columnCount: flags(row, column) = True: Next column
        Next row
    End If
    For row = 1 To rowCount
        For column = 1 To columnCount
            rows(index, 0) = IIf(flags(row, column), "1", "0")
            rows(index, 1) = vbNullString
            If flags(row, column) Then rows(index, 1) = NxWorkbookCompareCellText(formulas(row, column))
            rows(index, 2) = NxWorkbookCompareCellText(values(row, column))
            rows(index, 3) = "-": rows(index, 4) = vbNullString
            ' Pure blank cells are not data, even when they carry formatting/merges.
            If rows(index, 0) = "0" And (rows(index, 2) = "#EMPTY" Or rows(index, 2) = "8:") Then
                rows(index, 2) = "#EMPTY"
            ElseIf hasMerges Or compareFormats Then
                Set cell = target.Cells(row, column)
                If hasMerges Then rows(index, 3) = NxWorkbookCompareMergeSignature(cell)
                If compareFormats Then rows(index, 4) = NxWorkbookCompareFormatSignature(cell, rowHeights(row), columnWidths(column))
            End If
            index = index + 1
            If index Mod 2048 = 0 Then CompareProgress "셀 데이터 읽기", index, rowCount * columnCount
        Next column
    Next row
    CompareCapture = rows
End Function

Private Sub NxWorkbookCompareNextDetailSheet(ByRef detailSheet As Worksheet, ByRef detailRow As Long)
    Dim book As Workbook
    Set book = detailSheet.Parent
    Set detailSheet = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
    detailSheet.Name = "셀차이_" & CStr(book.Worksheets.Count - 1)
    NxWorkbookCompareInitializeDetails detailSheet
    detailRow = 2
End Sub

Private Sub NxWorkbookCompareUnionBounds(ByVal baseSheet As Worksheet, ByVal compareSheet As Worksheet, ByRef firstRow As Long, ByRef lastRow As Long, ByRef firstColumn As Long, ByRef lastColumn As Long, ByRef cellCount As Double)
    Dim baseRange As Range, compareRange As Range
    Set baseRange = NxDataCompareUsedRange(baseSheet): Set compareRange = NxDataCompareUsedRange(compareSheet)
    firstRow = Application.Min(baseRange.Row, compareRange.Row)
    lastRow = Application.Max(baseRange.Row + baseRange.Rows.Count - 1, compareRange.Row + compareRange.Rows.Count - 1)
    firstColumn = Application.Min(baseRange.Column, compareRange.Column)
    lastColumn = Application.Max(baseRange.Column + baseRange.Columns.Count - 1, compareRange.Column + compareRange.Columns.Count - 1)
    cellCount = CDbl(lastRow - firstRow + 1) * CDbl(lastColumn - firstColumn + 1)
End Sub

Private Function NxWorkbookCompareCellDifference(ByRef leftRows As Variant, ByRef rightRows As Variant, ByVal ordinal As Long) As String
    If leftRows(ordinal, 0) <> rightRows(ordinal, 0) Then
        NxWorkbookCompareCellDifference = NxWorkbookCompareAppendReason(NxWorkbookCompareCellDifference, "수식/상수")
    ElseIf leftRows(ordinal, 0) = "1" Then
        If leftRows(ordinal, 1) <> rightRows(ordinal, 1) Then NxWorkbookCompareCellDifference = NxWorkbookCompareAppendReason(NxWorkbookCompareCellDifference, "수식")
    End If
    If leftRows(ordinal, 2) <> rightRows(ordinal, 2) Then
        NxWorkbookCompareCellDifference = NxWorkbookCompareAppendReason(NxWorkbookCompareCellDifference, "값")
    End If
    If leftRows(ordinal, 3) <> rightRows(ordinal, 3) Then NxWorkbookCompareCellDifference = NxWorkbookCompareAppendReason(NxWorkbookCompareCellDifference, "병합")
    If leftRows(ordinal, 4) <> rightRows(ordinal, 4) Then NxWorkbookCompareCellDifference = NxWorkbookCompareAppendReason(NxWorkbookCompareCellDifference, "서식")
End Function

Private Function NxWorkbookCompareCellText(ByVal value As Variant) As String
    If IsError(value) Then
        NxWorkbookCompareCellText = NxWorkbookCompareErrorText(value)
    ElseIf IsEmpty(value) Then
        NxWorkbookCompareCellText = "#EMPTY"
    Else
        NxWorkbookCompareCellText = CStr(VarType(value)) & ":" & CStr(value)
    End If
End Function

Private Function NxWorkbookCompareErrorText(ByVal value As Variant) As String
    ' Explicit conversion preserves the error number without WorksheetFunction dispatch.
    NxWorkbookCompareErrorText = "#ERROR:" & CStr(value)
End Function

Private Function NxWorkbookCompareMergeSignature(ByVal cell As Range) As String
    If cell.MergeCells Then NxWorkbookCompareMergeSignature = cell.MergeArea.Address(False, False) Else NxWorkbookCompareMergeSignature = "-"
End Function

Private Function NxWorkbookCompareFormatSignature(ByVal cell As Range, ByVal rowHeight As Variant, ByVal columnWidth As Variant) As String
    NxWorkbookCompareFormatSignature = CStr(cell.NumberFormat) & "|" & CStr(cell.Font.Name) & "|" & CStr(cell.Font.Size) & "|" & CStr(cell.Font.Bold) & "|" & CStr(cell.Font.Italic) & "|" & CStr(cell.Font.Color) & "|" & CStr(cell.Interior.Color) & "|" & CStr(cell.HorizontalAlignment) & "|" & CStr(cell.VerticalAlignment) & "|" & NxWorkbookCompareBorderSignature(cell, xlEdgeLeft) & "|" & NxWorkbookCompareBorderSignature(cell, xlEdgeTop) & "|" & NxWorkbookCompareBorderSignature(cell, xlEdgeRight) & "|" & NxWorkbookCompareBorderSignature(cell, xlEdgeBottom) & "|" & CStr(rowHeight) & "|" & CStr(columnWidth)
End Function

Private Function NxWorkbookCompareBorderSignature(ByVal cell As Range, ByVal edge As XlBordersIndex) As String
    With cell.Borders(edge)
        NxWorkbookCompareBorderSignature = CStr(.LineStyle) & ":" & CStr(.Weight) & ":" & CStr(.Color)
    End With
End Function

Private Function NxWorkbookCompareAppendReason(ByVal currentReason As String, ByVal nextReason As String) As String
    If Len(currentReason) = 0 Then NxWorkbookCompareAppendReason = nextReason Else NxWorkbookCompareAppendReason = currentReason & ", " & nextReason
End Function

Private Sub NxWorkbookCompareInitializeSummary(ByVal sheet As Worksheet, ByVal baseSourcePath As String, _
    ByVal compareSourcePath As String, ByVal compareFormats As Boolean)
    sheet.Range("A1").Value2 = "내엑셀 통합문서 비교"
    sheet.Range("A2").Value2 = "기준 문서": NxWorkbookCompareWriteLiteral sheet.Range("B2"), baseSourcePath
    sheet.Range("A3").Value2 = "비교 문서": NxWorkbookCompareWriteLiteral sheet.Range("B3"), compareSourcePath
    sheet.Range("A4").Value2 = "전체 차이": sheet.Range("B4").Value2 = 0
    sheet.Range("C4").Value2 = "서식 비교": sheet.Range("D4").Value2 = IIf(compareFormats, "포함", "제외")
    NxWorkbookCompareWriteLiteral sheet.Range("F2"), "비교 제외"
    NxWorkbookCompareWriteLiteral sheet.Range("G2"), "차트, 도형, 조건부서식, 정의된 이름"
    sheet.Range("A5").Value2 = "시트": sheet.Range("B5").Value2 = "상태": sheet.Range("C5").Value2 = "차이 수": sheet.Range("D5").Value2 = "확인"
    sheet.Range("E5").Value2 = "기준 원본": sheet.Range("F5").Value2 = "비교 원본"
    sheet.Range("F3").Value2 = "원본 이동 안내": sheet.Range("G3").Value2 = "내엑셀 추가 기능이 열린 상태에서 원본으로 이동합니다. 파일 위치와 시트 이름을 유지하세요. 숨김 대상은 원본에서 표시한 뒤 이동하세요."
    NxWorkbookCompareWriteLiteral sheet.Range("J1"), NX_COMPARE_NAV_MARKER
    sheet.Columns("J").Hidden = True
    sheet.Range("A1:D1").Font.Bold = True: sheet.Range("A5:F5").Font.Bold = True
End Sub

Private Sub NxWorkbookCompareInitializeDetails(ByVal sheet As Worksheet)
    sheet.Range("A1").Value2 = "시트": sheet.Range("B1").Value2 = "셀": sheet.Range("C1").Value2 = "차이"
    sheet.Range("D1").Value2 = "기준 값": sheet.Range("E1").Value2 = "비교 값": sheet.Range("F1").Value2 = "기준 수식": sheet.Range("G1").Value2 = "비교 수식"
    sheet.Range("H1").Value2 = "기준 원본": sheet.Range("I1").Value2 = "비교 원본"
    sheet.Range("J1").Value2 = "기준 표시 형식": sheet.Range("K1").Value2 = "비교 표시 형식"
    sheet.Range("L1").Value2 = "기준 병합 범위": sheet.Range("M1").Value2 = "비교 병합 범위"
    sheet.Range("N1").Value2 = "기준 화면 표시": sheet.Range("O1").Value2 = "비교 화면 표시"
    sheet.Hyperlinks.Add Anchor:=sheet.Range("P1"), Address:=vbNullString, SubAddress:="'비교요약'!A1", TextToDisplay:="요약으로"
    sheet.Range("A1:P1").Font.Bold = True
End Sub

Private Sub NxWorkbookCompareWriteCellLink(ByVal target As Range, ByVal sourcePath As String, ByVal sheetName As String, ByVal cellAddress As String, ByVal labelText As String)
    ' SheetFollowHyperlink cannot cancel Excel's default navigation. The link therefore
    ' points only to itself; the add-in resolves literal report fields after validation.
    target.Hyperlinks.Add Anchor:=target, Address:=vbNullString, _
        SubAddress:=NxWorkbookCompareSelfAnchor(target), TextToDisplay:=labelText, ScreenTip:="원본 파일의 " & sheetName & "!" & cellAddress & "(으)로 이동: " & sourcePath
End Sub

Private Function NxWorkbookCompareSelfAnchor(ByVal target As Range) As String
    NxWorkbookCompareSelfAnchor = "'" & Replace$(target.Worksheet.Name, "'", "''") & "'!" & target.Address(False, False, xlA1)
End Function

Public Sub NxWorkbookCompareFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    Dim issue As String
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    If NxWorkbookCompareNavigate(Sh, Target, issue) Then Exit Sub
    If Len(issue) > 0 Then MsgBox issue, vbExclamation, "통합문서 비교"
End Sub

Public Function NxWorkbookCompareResolveLink(ByVal reportSheet As Worksheet, ByVal target As Hyperlink, _
    ByRef sourcePath As String, ByRef sheetName As String, ByRef cellAddress As String, ByRef issue As String) As Boolean
    Dim summarySheet As Worksheet, linkCell As Range, sourceRow As Long, rowIndex As Long, summaryLastRow As Long
    Dim otherPath As String, labelText As String, stateText As String, foundSheet As Boolean
    sourcePath = vbNullString: sheetName = vbNullString: cellAddress = vbNullString: issue = vbNullString
    On Error GoTo InvalidReport
    If reportSheet Is Nothing Then Exit Function
    Set summarySheet = NxWorkbookCompareWorksheet(reportSheet.Parent, NxWorkbookCompareReadLiteral(reportSheet.Range("Q1")))
    If summarySheet Is Nothing Then Set summarySheet = NxWorkbookCompareWorksheet(reportSheet.Parent, "비교요약")
    If summarySheet Is Nothing Then Exit Function
    If NxWorkbookCompareReadLiteral(summarySheet.Range("J1")) <> NX_COMPARE_NAV_MARKER Then Exit Function
    ' Ordinary summary/detail navigation is not an original-file link.
    If target Is Nothing Then Exit Function
    Set linkCell = target.Range
    If linkCell.CountLarge <> 1 Then Exit Function
    If Len(target.Address) <> 0 Then Exit Function
    If StrComp(target.SubAddress, NxWorkbookCompareSelfAnchor(linkCell), vbBinaryCompare) <> 0 Then Exit Function
    issue = "원본 이동 정보를 확인할 수 없습니다. 통합문서 비교 보고서를 다시 만들어 주세요."
    If reportSheet.Parent.IsAddin Then Exit Function
    If NxWorkbookCompareReadLiteral(summarySheet.Range("A1")) <> "내엑셀 통합문서 비교" Then Exit Function
    If NxWorkbookCompareReadLiteral(summarySheet.Range("A2")) <> "기준 문서" Then Exit Function
    If NxWorkbookCompareReadLiteral(summarySheet.Range("A3")) <> "비교 문서" Then Exit Function
    If Not NxWorkbookCompareHeadersMatch(summarySheet.Range("A5:F5"), "시트|상태|차이 수|확인|기준 원본|비교 원본") Then Exit Function
    summaryLastRow = summarySheet.Cells(summarySheet.Rows.Count, 1).End(xlUp).Row
    If summaryLastRow < 6 Or summaryLastRow > NX_COMPARE_DETAIL_LAST_ROW Then Exit Function
    If target Is Nothing Then Exit Function
    Set linkCell = target.Range
    If linkCell.CountLarge <> 1 Then Exit Function
    If Not linkCell.Worksheet Is reportSheet Then Exit Function
    If Len(target.Address) <> 0 Then Exit Function
    If StrComp(target.SubAddress, NxWorkbookCompareSelfAnchor(linkCell), vbBinaryCompare) <> 0 Then Exit Function
    If reportSheet Is summarySheet Then
        If linkCell.Row < 6 Or linkCell.Row > summaryLastRow Then Exit Function
        Select Case linkCell.Column
            Case 5: sourceRow = 2
            Case 6: sourceRow = 3
            Case Else: Exit Function
        End Select
        sheetName = NxWorkbookCompareReadLiteral(reportSheet.Cells(linkCell.Row, 1))
        cellAddress = "A1"
        stateText = NxWorkbookCompareReadLiteral(reportSheet.Cells(linkCell.Row, 2))
    Else
        If Not NxWorkbookCompareIsDetailName(reportSheet.Name) Then Exit Function
        If Not NxWorkbookCompareHeadersMatch(reportSheet.Range("A1:I1"), "시트|셀|차이|기준 값|비교 값|기준 수식|비교 수식|기준 원본|비교 원본") Then Exit Function
        If linkCell.Row < 2 Or linkCell.Row > NX_COMPARE_DETAIL_LAST_ROW Then Exit Function
        Select Case linkCell.Column
            Case 4: sourceRow = 2
            Case 5: sourceRow = 3
            Case 8: sourceRow = 2
            Case 9: sourceRow = 3
            Case Else: Exit Function
        End Select
        sheetName = NxWorkbookCompareReadLiteral(reportSheet.Cells(linkCell.Row, 1))
        cellAddress = NxWorkbookCompareReadLiteral(reportSheet.Cells(linkCell.Row, 2))
        If Len(NxWorkbookCompareReadLiteral(reportSheet.Cells(linkCell.Row, 3))) = 0 Then Exit Function
        For rowIndex = 6 To summaryLastRow
            If StrComp(NxWorkbookCompareReadLiteral(summarySheet.Cells(rowIndex, 1)), sheetName, vbBinaryCompare) = 0 Then
                stateText = NxWorkbookCompareReadLiteral(summarySheet.Cells(rowIndex, 2))
                foundSheet = True: Exit For
            End If
        Next rowIndex
        If Not foundSheet Then Exit Function
        If stateText <> "차이 있음" Then Exit Function
    End If
    Select Case stateText
        Case "일치", "차이 있음"
        Case "기준 문서에 없음": If sourceRow = 2 Then Exit Function
        Case "비교 문서에 없음": If sourceRow = 3 Then Exit Function
        Case Else: Exit Function
    End Select
    labelText = IIf(sourceRow = 2, "기준 원본으로 이동", "비교 원본으로 이동")
    If reportSheet Is summarySheet Or linkCell.Column >= 8 Then
        If NxWorkbookCompareReadLiteral(linkCell) <> labelText Then Exit Function
    Else
        If linkCell.HasFormula Then Exit Function
    End If
    If Len(sheetName) = 0 Or Len(sheetName) > 31 Then Exit Function
    If Len(NxWorkbookCompareNavigationCellIssue(cellAddress)) > 0 Then Exit Function
    sourcePath = NxWorkbookCompareReadLiteral(summarySheet.Cells(sourceRow, 2))
    otherPath = NxWorkbookCompareReadLiteral(summarySheet.Cells(5 - sourceRow, 2))
    issue = NxWorkbookCompareNavigationPathIssue(sourcePath)
    If Len(issue) > 0 Then Exit Function
    issue = NxWorkbookCompareNavigationPathIssue(otherPath)
    If Len(issue) > 0 Then Exit Function
    If StrComp(sourcePath, otherPath, vbTextCompare) = 0 Then GoTo InvalidReport
    If sourceRow = 3 And reportSheet Is summarySheet Then
        If Len(NxWorkbookCompareReadLiteral(summarySheet.Cells(linkCell.Row, 11))) > 0 Then sheetName = NxWorkbookCompareReadLiteral(summarySheet.Cells(linkCell.Row, 11))
    End If
    NxWorkbookCompareResolveLink = True
    Exit Function
InvalidReport:
    issue = "원본 이동 정보를 확인할 수 없습니다. 통합문서 비교 보고서를 다시 만들어 주세요."
End Function

Public Function NxWorkbookCompareNavigate(ByVal reportSheet As Worksheet, ByVal target As Hyperlink, ByRef issue As String) As Boolean
    Dim sourcePath As String, sheetName As String, cellAddress As String
    Dim sourceBook As Workbook, candidate As Workbook, sourceSheet As Worksheet, sourceCell As Range
    Dim sourceWindow As Window, candidateWindow As Window
    Dim sourceHandle As Integer, openedByUs As Boolean, stateCaptured As Boolean
    Dim priorSecurity As MsoAutomationSecurity, priorEvents As Boolean, restoreIssue As String
    issue = vbNullString
    If mCompareNavigating Then Exit Function
    If Not NxWorkbookCompareResolveLink(reportSheet, target, sourcePath, sheetName, cellAddress, issue) Then Exit Function
    mCompareNavigating = True
    On Error GoTo Failed
    priorSecurity = Application.AutomationSecurity: priorEvents = Application.EnableEvents: stateCaptured = True
    Application.EnableEvents = False
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, sourcePath, vbTextCompare) = 0 Then Set sourceBook = candidate: Exit For
    Next candidate
    If sourceBook Is Nothing Then
        If Not CreateObject("Scripting.FileSystemObject").FileExists(sourcePath) Then
            issue = "원본 파일을 찾을 수 없습니다. 파일이 이동되거나 이름이 변경되었는지 확인하세요: " & sourcePath
            GoTo CleanUp
        End If
        ' Inspection checks a snapshot. Deny writes/replacement of the original until
        ' Excel has opened that exact file, so it cannot change between check and open.
        sourceHandle = FreeFile
        Open sourcePath For Binary Access Read Lock Write As #sourceHandle
        issue = NxSafeWorkbookInputIssue(sourcePath)
        If Len(issue) > 0 Then GoTo CleanUp
        Application.AutomationSecurity = msoAutomationSecurityForceDisable
        Set sourceBook = Application.Workbooks.Open(Filename:=sourcePath, UpdateLinks:=0, ReadOnly:=True, _
            Password:=vbNullString, WriteResPassword:=vbNullString, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        openedByUs = True
        If Not sourceBook.ReadOnly Then issue = "원본을 읽기 전용으로 열지 못해 이동을 중단했습니다.": GoTo CleanUp
        If StrComp(sourceBook.FullName, sourcePath, vbTextCompare) <> 0 Then issue = "선택한 원본과 열린 파일 경로가 달라 이동을 중단했습니다.": GoTo CleanUp
    End If
    Set sourceSheet = NxWorkbookCompareWorksheet(sourceBook, sheetName)
    If sourceSheet Is Nothing Then
        issue = "원본 시트를 찾을 수 없습니다. 시트가 삭제되거나 이름이 변경되었는지 확인하세요: " & sheetName
        GoTo CleanUp
    End If
    If sourceSheet.Visible <> xlSheetVisible Then issue = "원본 시트가 숨김 상태입니다. 원본에서 시트를 표시한 뒤 다시 이동하세요: " & sheetName: GoTo CleanUp
    Set sourceCell = sourceSheet.Range(cellAddress)
    If sourceCell.EntireRow.Hidden Or sourceCell.EntireColumn.Hidden Then issue = "원본 셀이 숨김 행 또는 열에 있습니다. 원본에서 표시한 뒤 다시 이동하세요: " & cellAddress: GoTo CleanUp
    For Each candidateWindow In sourceBook.Windows
        If candidateWindow.Visible Then Set sourceWindow = candidateWindow: Exit For
    Next candidateWindow
    If sourceWindow Is Nothing Then issue = "원본 통합문서 창이 숨김 상태입니다. 원본 창을 표시한 뒤 다시 이동하세요.": GoTo CleanUp
    sourceWindow.Activate
    sourceSheet.Activate
    Application.Goto Reference:=sourceCell, Scroll:=True
    NxWorkbookCompareNavigate = True
CleanUp:
    On Error Resume Next
    If Not NxWorkbookCompareNavigate Then
        If openedByUs Then sourceBook.Close SaveChanges:=False
    End If
    If sourceHandle <> 0 Then Close #sourceHandle
    If stateCaptured Then
        Err.Clear
        Application.AutomationSecurity = priorSecurity
        If Err.Number <> 0 Then restoreIssue = "매크로 보안 설정"
        Err.Clear
        Application.EnableEvents = priorEvents
        If Err.Number <> 0 Then restoreIssue = restoreIssue & " 이벤트 설정"
    End If
    On Error GoTo 0
    mCompareNavigating = False
    If Len(restoreIssue) > 0 Then
        issue = issue & " Excel 상태 복원에 실패했습니다. 설정을 확인하세요: " & restoreIssue
        NxWorkbookCompareNavigate = False
    End If
    Exit Function
Failed:
    issue = "원본을 안전하게 열거나 해당 셀로 이동할 수 없습니다. 파일 권한, 암호, 보호 상태를 확인하세요. " & Err.Description
    Resume CleanUp
End Function

Private Function NxWorkbookCompareReadLiteral(ByVal cell As Range) As String
    If cell.HasFormula Or IsError(cell.Value2) Then Err.Raise 5, , "비교 보고서 원본 이동 필드는 일반 텍스트여야 합니다."
    NxWorkbookCompareReadLiteral = CStr(cell.Value2)
End Function

Private Function NxWorkbookCompareHeadersMatch(ByVal headers As Range, ByVal expected As String) As Boolean
    Dim labels As Variant, index As Long
    labels = Split(expected, "|")
    If headers.CountLarge <> UBound(labels) + 1 Then Exit Function
    For index = 0 To UBound(labels)
        If NxWorkbookCompareReadLiteral(headers.Cells(1, index + 1)) <> CStr(labels(index)) Then Exit Function
    Next index
    NxWorkbookCompareHeadersMatch = True
End Function

Private Function NxWorkbookCompareIsDetailName(ByVal sheetName As String) As Boolean
    Dim suffix As String
    If sheetName = "셀차이" Then NxWorkbookCompareIsDetailName = True: Exit Function
    If Left$(sheetName, 4) <> "셀차이_" Then Exit Function
    suffix = Mid$(sheetName, 5)
    If Len(suffix) = 0 Or Len(suffix) > 5 Then Exit Function
    If suffix Like "*[!0-9]*" Then Exit Function
    If CLng(suffix) < 2 Or CLng(suffix) > 32767 Then Exit Function
    NxWorkbookCompareIsDetailName = (suffix = CStr(CLng(suffix)))
End Function

Public Function NxWorkbookCompareNavigationCellIssue(ByVal address As String) As String
    Dim index As Long, letterCount As Long, columnNumber As Long, rowText As String, character As String
    NxWorkbookCompareNavigationCellIssue = "원본 셀 주소가 올바르지 않습니다. 비교 보고서를 다시 만들어 주세요."
    If Len(address) < 2 Or Len(address) > 10 Then Exit Function
    For index = 1 To Len(address)
        character = Mid$(address, index, 1)
        If character < "A" Or character > "Z" Then Exit For
        letterCount = letterCount + 1
        If letterCount > 3 Then Exit Function
        columnNumber = columnNumber * 26 + AscW(character) - AscW("A") + 1
    Next index
    If letterCount = 0 Or columnNumber > 16384 Then Exit Function
    rowText = Mid$(address, letterCount + 1)
    If Len(rowText) = 0 Or Len(rowText) > 7 Then Exit Function
    If rowText Like "*[!0-9]*" Then Exit Function
    If Left$(rowText, 1) = "0" Or CLng(rowText) > 1048576 Then Exit Function
    NxWorkbookCompareNavigationCellIssue = vbNullString
End Function

Public Function NxWorkbookCompareNavigationPathIssue(ByVal sourcePath As String) As String
    Dim parts As Variant, segment As Variant, index As Long, stem As String, rest As String
    NxWorkbookCompareNavigationPathIssue = "원본 경로는 절대 드라이브 또는 UNC 공유 경로의 .xlsx/.xlsm 파일이어야 합니다. URL, 상대 경로, 장치 경로는 사용할 수 없습니다."
    If Len(sourcePath) < 8 Or Len(sourcePath) > 259 Or sourcePath <> Trim$(sourcePath) Then Exit Function
    If Left$(sourcePath, 2) = "\\" Then
        If Left$(sourcePath, 4) = "\\?\" Or Left$(sourcePath, 4) = "\\.\" Then Exit Function
        parts = Split(Mid$(sourcePath, 3), "\")
        If UBound(parts) < 2 Then Exit Function
        ' WebDAV redirector syntax is not an SMB source path.
        If CStr(parts(0)) Like "*[!A-Za-z0-9_.-]*" Then Exit Function
        If StrComp(CStr(parts(1)), "DavWWWRoot", vbTextCompare) = 0 Then Exit Function
    Else
        If Not UCase$(Left$(sourcePath, 1)) Like "[A-Z]" Then Exit Function
        If Mid$(sourcePath, 2, 2) <> ":\" Then Exit Function
        parts = Split(Mid$(sourcePath, 4), "\")
    End If
    For Each segment In parts
        rest = CStr(segment)
        If Len(rest) = 0 Or rest = "." Or rest = ".." Then Exit Function
        If Right$(rest, 1) = "." Or Right$(rest, 1) = " " Then Exit Function
        For index = 1 To Len(rest)
            If AscW(Mid$(rest, index, 1)) >= 0 And AscW(Mid$(rest, index, 1)) < 32 Then Exit Function
            If InStr(1, ":/<>""|?*", Mid$(rest, index, 1), vbBinaryCompare) > 0 Then Exit Function
        Next index
        stem = UCase$(Split(rest, ".")(0))
        Select Case stem
            Case "CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$": Exit Function
        End Select
        If stem Like "COM[1-9]" Or stem Like "LPT[1-9]" Then Exit Function
    Next segment
    Select Case LCase$(Mid$(sourcePath, InStrRev(sourcePath, ".") + 1))
        Case "xlsx", "xlsm": NxWorkbookCompareNavigationPathIssue = vbNullString
    End Select
End Function

Private Sub NxWorkbookCompareWriteSummary(ByVal sheet As Worksheet, ByVal targetRow As Long, ByVal sheetName As String, ByVal statusText As String, ByVal differenceCount As Long)
    NxWorkbookCompareWriteLiteral sheet.Cells(targetRow, 1), sheetName
    NxWorkbookCompareWriteLiteral sheet.Cells(targetRow, 2), statusText
    sheet.Cells(targetRow, 3).Value2 = differenceCount
    NxWorkbookCompareWriteLiteral sheet.Cells(targetRow, 4), IIf(differenceCount = 0, "-", "셀차이 시트 확인")
End Sub

Private Sub NxWorkbookCompareWriteDetail(ByRef batch As Variant, ByVal targetRow As Long, ByVal sheetName As String, ByVal cellAddress As String, ByVal reason As String, ByVal baseValue As Variant, ByVal compareValue As Variant, ByVal baseFormula As Variant, ByVal compareFormula As Variant, ByVal baseSourcePath As String, ByVal compareSourcePath As String, _
    ByVal baseFormat As String, ByVal compareFormat As String, ByVal baseMerge As String, ByVal compareMerge As String, ByVal baseDisplay As String, ByVal compareDisplay As String)
    Dim rowValues(1 To 1, 1 To 15) As Variant, column As Long
    rowValues(1, 1) = sheetName
    rowValues(1, 2) = cellAddress
    rowValues(1, 3) = reason
    rowValues(1, 4) = NxWorkbookCompareDisplayValue(baseValue)
    rowValues(1, 5) = NxWorkbookCompareDisplayValue(compareValue)
    rowValues(1, 6) = NxWorkbookCompareDisplayValue(baseFormula)
    rowValues(1, 7) = NxWorkbookCompareDisplayValue(compareFormula)
    rowValues(1, 10) = baseFormat: rowValues(1, 11) = compareFormat
    rowValues(1, 12) = baseMerge: rowValues(1, 13) = compareMerge
    rowValues(1, 14) = baseDisplay: rowValues(1, 15) = compareDisplay
    For column = 1 To 15: batch(targetRow, column) = rowValues(1, column): Next column
End Sub

Private Function NxWorkbookCompareDisplayValue(ByVal value As Variant) As String
    If IsEmpty(value) Then
        NxWorkbookCompareDisplayValue = "(빈칸)"
    ElseIf IsError(value) Then
        NxWorkbookCompareDisplayValue = CStr(value)
    ElseIf VarType(value) = vbString Then
        If Len(value) = 0 Then NxWorkbookCompareDisplayValue = "(빈칸)" Else NxWorkbookCompareDisplayValue = value
    Else
        NxWorkbookCompareDisplayValue = CStr(value)
    End If
End Function

Private Sub NxWorkbookCompareWriteLiteral(ByVal target As Range, ByVal valueText As String)
    target.NumberFormat = "@": target.Value2 = valueText
End Sub

Private Function NxWorkbookCompareTemporaryPath(ByVal outputPath As String) As String
    NxWorkbookCompareTemporaryPath = Left$(outputPath, Len(outputPath) - 5) & ".nx-" & Replace$(NxCreateRunUuid(), "-", vbNullString) & ".tmp.xlsx"
End Function

Public Sub NxWorkbookCompareAtomicPublish(ByVal temporaryPath As String, ByVal outputPath As String)
    NxWorkbookCompareRequireVacantOutput outputPath
    If Len(Dir$(temporaryPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then NxRaiseContractError "비교 보고서 임시 파일을 찾을 수 없습니다."
    Name temporaryPath As outputPath
End Sub
