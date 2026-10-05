Attribute VB_Name = "NxCopySave"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal fileNamePointer As LongPtr, ByVal desiredAccess As Long, _
    ByVal shareMode As Long, ByVal securityAttributes As LongPtr, ByVal creationDisposition As Long, _
    ByVal flagsAndAttributes As Long, ByVal templateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal objectHandle As LongPtr) As Long
Private Declare PtrSafe Function MoveFileExW Lib "kernel32" (ByVal existingFileNamePointer As LongPtr, _
    ByVal newFileNamePointer As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function CreateFileW Lib "kernel32" (ByVal fileNamePointer As Long, ByVal desiredAccess As Long, _
    ByVal shareMode As Long, ByVal securityAttributes As Long, ByVal creationDisposition As Long, _
    ByVal flagsAndAttributes As Long, ByVal templateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal objectHandle As Long) As Long
Private Declare Function MoveFileExW Lib "kernel32" (ByVal existingFileNamePointer As Long, _
    ByVal newFileNamePointer As Long, ByVal flags As Long) As Long
#End If

Private Const NX_GENERIC_READ As Long = &H80000000
Private Const NX_GENERIC_WRITE As Long = &H40000000
Private Const NX_OPEN_EXISTING As Long = 3
Private Const NX_FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const NX_INVALID_HANDLE_VALUE As Long = -1
Private Const NX_MOVEFILE_WRITE_THROUGH As Long = &H8

' Copy/save commands always operate on a new workbook, preserve the original, and restore Excel state.
Public Function NxRunSheetCopySave(ByVal sourceSheet As Worksheet, ByVal outputPath As String) As CNxResult
    Dim oldBook As Workbook, oldSheet As Worksheet, oldSelection As Object
    Dim copyBook As Workbook, copySheet As Worksheet, tempPath As String, failure As String
    On Error GoTo Failed
    If sourceSheet Is Nothing Then NxRaiseContractError "Sheet copy requires a worksheet"
    NxFileRequireNewOutput outputPath, ".xlsx": Set oldBook = ActiveWorkbook: Set oldSheet = ActiveSheet: Set oldSelection = Application.Selection
    tempPath = NxFileTemporaryOutputPath(outputPath, ".xlsx")
    ' The worksheet object copy method can deadlock inside a modal plan-form event.
    ' Rebuild the owned workbook from worksheet content so the original remains
    ' untouched and the modal execution stack never copies or opens a workbook.
    Set copyBook = Workbooks.Add(xlWBATWorksheet)
    Set copySheet = copyBook.Worksheets.Item(1)
    NxCopyWorksheetContent sourceSheet, copySheet
    copyBook.SaveAs Filename:=tempPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    copyBook.Close SaveChanges:=False
    Set copyBook = Nothing
    NxFileCommitTemporaryOutput tempPath, outputPath
    Set NxRunSheetCopySave = NxCreateResult("NX-FILE-SHEET-COPY-SAVE", NxSuccess, "complete", outputPath, False, vbNullString, "result_success")
    GoTo Restore
Failed:
    failure = Err.Description
    On Error Resume Next
    If Not copyBook Is Nothing Then copyBook.Close SaveChanges:=False
    NxFileDeleteCreatedFile tempPath, True
    Set NxRunSheetCopySave = NxCreateResult("NX-FILE-SHEET-COPY-SAVE", NxEnvironmentError, "copy_save", outputPath, False, failure, "result_environment_error")
Restore:
    On Error Resume Next
    If Not oldBook Is Nothing Then oldBook.Activate: If Not oldSheet Is Nothing Then oldSheet.Activate: If Not oldSelection Is Nothing Then oldSelection.Select
    On Error GoTo 0
End Function

Private Sub NxCopyWorksheetContent(ByVal sourceSheet As Worksheet, ByVal destinationSheet As Worksheet)
    Dim sourceUsed As Range, destinationUsed As Range
    Dim sourceShape As Object, pastedShape As Object
    Dim shapeIndex As Long, shapeCount As Long
    Dim stage As String, failure As String
    On Error GoTo Failed
    If sourceSheet Is Nothing Or destinationSheet Is Nothing Then NxRaiseContractError "Sheet copy content endpoints are required"
    stage = "sheet-name"
    destinationSheet.Name = sourceSheet.Name
    stage = "used-range"
    Set sourceUsed = sourceSheet.UsedRange
    Set destinationUsed = destinationSheet.Range(sourceUsed.Address)
    sourceUsed.Copy Destination:=destinationUsed
    stage = "row-dimensions"
    NxCopyRangeDimensions sourceUsed, destinationUsed, True
    stage = "column-dimensions"
    NxCopyRangeDimensions sourceUsed, destinationUsed, False
    stage = "shapes"
    shapeCount = sourceSheet.Shapes.Count
    For shapeIndex = 1 To shapeCount
        Set sourceShape = sourceSheet.Shapes.Item(shapeIndex)
        If sourceShape.Type = 3 Then
            NxCopyChartShape sourceShape, destinationSheet
        ElseIf sourceShape.Type = 4 Then
            ' Cell notes are already preserved by the range copy. Copying their
            ' backing comment shape a second time raises an Excel object error.
        Else
            sourceShape.Copy
            destinationSheet.Paste
            Set pastedShape = destinationSheet.Shapes.Item(destinationSheet.Shapes.Count)
            pastedShape.Left = sourceShape.Left
            pastedShape.Top = sourceShape.Top
            pastedShape.Width = sourceShape.Width
            pastedShape.Height = sourceShape.Height
            pastedShape.Name = sourceShape.Name
            Set pastedShape = Nothing
        End If
        Set sourceShape = Nothing
    Next shapeIndex
    stage = "page-setup"
    NxCopyPageSetup sourceSheet.PageSetup, destinationSheet.PageSetup
    stage = "tab-color"
    destinationSheet.Tab.ColorIndex = sourceSheet.Tab.ColorIndex
    Application.CutCopyMode = False
    Exit Sub
Failed:
    failure = Err.Description
    NxRaiseContractError "Sheet content copy failed at " & stage & ": " & failure
End Sub

Private Sub NxCopyChartShape(ByVal sourceShape As Object, ByVal destinationSheet As Worksheet)
    Dim sourceChart As Object, destinationChartObject As Object, destinationChart As Object
    Dim sourceSeries As Object, destinationSeries As Object
    Set sourceChart = sourceShape.Chart
    Set destinationChartObject = destinationSheet.ChartObjects.Add(sourceShape.Left, sourceShape.Top, sourceShape.Width, sourceShape.Height)
    destinationChartObject.Name = sourceShape.Name
    Set destinationChart = destinationChartObject.Chart
    destinationChart.ChartType = sourceChart.ChartType
    For Each sourceSeries In sourceChart.SeriesCollection
        Set destinationSeries = destinationChart.SeriesCollection.NewSeries
        destinationSeries.Formula = NxLocalizeSeriesFormula(sourceSeries.Formula, sourceShape.Parent.Parent.Name)
        Set destinationSeries = Nothing
    Next sourceSeries
    destinationChart.HasTitle = sourceChart.HasTitle
    If sourceChart.HasTitle Then destinationChart.ChartTitle.Text = sourceChart.ChartTitle.Text
    destinationChart.HasLegend = sourceChart.HasLegend
    If sourceChart.HasLegend Then destinationChart.Legend.Position = sourceChart.Legend.Position
End Sub

Private Function NxLocalizeSeriesFormula(ByVal sourceFormula As String, ByVal workbookName As String) As String
    NxLocalizeSeriesFormula = Replace(sourceFormula, "[" & workbookName & "]", vbNullString, 1, -1, vbTextCompare)
End Function

Private Sub NxCopyPageSetup(ByVal sourceSetup As PageSetup, ByVal destinationSetup As PageSetup)
    Dim stage As String, failure As String
    On Error GoTo Failed
    If sourceSetup Is Nothing Or destinationSetup Is Nothing Then NxRaiseContractError "Sheet copy page setup endpoints are required"
    stage = "orientation"
    destinationSetup.Orientation = sourceSetup.Orientation
    stage = "paper-size"
    destinationSetup.PaperSize = sourceSetup.PaperSize
    stage = "zoom"
    destinationSetup.Zoom = sourceSetup.Zoom
    stage = "fit-wide"
    destinationSetup.FitToPagesWide = sourceSetup.FitToPagesWide
    stage = "fit-tall"
    destinationSetup.FitToPagesTall = sourceSetup.FitToPagesTall
    stage = "margins"
    destinationSetup.LeftMargin = sourceSetup.LeftMargin
    destinationSetup.RightMargin = sourceSetup.RightMargin
    destinationSetup.TopMargin = sourceSetup.TopMargin
    destinationSetup.BottomMargin = sourceSetup.BottomMargin
    destinationSetup.HeaderMargin = sourceSetup.HeaderMargin
    destinationSetup.FooterMargin = sourceSetup.FooterMargin
    stage = "centering"
    destinationSetup.CenterHorizontally = sourceSetup.CenterHorizontally
    destinationSetup.CenterVertically = sourceSetup.CenterVertically
    stage = "print-ranges"
    destinationSetup.PrintArea = sourceSetup.PrintArea
    destinationSetup.PrintTitleRows = sourceSetup.PrintTitleRows
    destinationSetup.PrintTitleColumns = sourceSetup.PrintTitleColumns
    stage = "print-options"
    destinationSetup.PrintGridlines = sourceSetup.PrintGridlines
    destinationSetup.PrintHeadings = sourceSetup.PrintHeadings
    stage = "headers"
    destinationSetup.LeftHeader = sourceSetup.LeftHeader
    destinationSetup.CenterHeader = sourceSetup.CenterHeader
    destinationSetup.RightHeader = sourceSetup.RightHeader
    destinationSetup.LeftFooter = sourceSetup.LeftFooter
    destinationSetup.CenterFooter = sourceSetup.CenterFooter
    destinationSetup.RightFooter = sourceSetup.RightFooter
    Exit Sub
Failed:
    failure = Err.Description
    NxRaiseContractError "Sheet page setup copy failed at " & stage & ": " & failure
End Sub

Public Function NxRunRangeCopySave(ByVal sourceRange As Range, ByVal outputPath As String) As CNxResult
    Dim oldBook As Workbook, oldSheet As Worksheet, oldSelection As Object, outputBook As Workbook, tempPath As String, failure As String
    On Error GoTo Failed
    If sourceRange Is Nothing Then NxRaiseContractError "Range copy requires a range"
    NxFileRequireNewOutput outputPath, ".xlsx": Set oldBook = ActiveWorkbook: Set oldSheet = ActiveSheet: Set oldSelection = Application.Selection
    tempPath = NxFileTemporaryOutputPath(outputPath, ".xlsx"): Set outputBook = Workbooks.Add(xlWBATWorksheet)
    sourceRange.Copy Destination:=outputBook.Worksheets(1).Range("A1")
    NxCopyRangeDimensions sourceRange, outputBook.Worksheets(1).Range("A1"), True
    NxCopyRangeDimensions sourceRange, outputBook.Worksheets(1).Range("A1"), False
    outputBook.SaveAs Filename:=tempPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    outputBook.Close SaveChanges:=False
    Set outputBook = Nothing
    NxFileCommitTemporaryOutput tempPath, outputPath
    Set NxRunRangeCopySave = NxCreateResult("NX-FILE-RANGE-COPY-SAVE", NxSuccess, "complete", outputPath, False, vbNullString, "result_success"): GoTo Restore
Failed:
    failure = Err.Description
    On Error Resume Next
    If Not outputBook Is Nothing Then outputBook.Close SaveChanges:=False
    NxFileDeleteCreatedFile tempPath, True
    Set NxRunRangeCopySave = NxCreateResult("NX-FILE-RANGE-COPY-SAVE", NxEnvironmentError, "copy_save", outputPath, False, failure, "result_environment_error")
Restore:
    On Error Resume Next: If Not oldBook Is Nothing Then oldBook.Activate: If Not oldSheet Is Nothing Then oldSheet.Activate: If Not oldSelection Is Nothing Then oldSelection.Select
    On Error GoTo 0
End Function

Private Sub NxCopyRangeDimensions(ByVal source As Range, ByVal destination As Range, ByVal rows As Boolean)
    Dim count As Long, index As Long, runStart As Long, size As Double, previousSize As Double
    Dim hidden As Boolean, previousHidden As Boolean, part As Range
    Dim uniformSize As Variant, uniformHidden As Variant
    If rows Then count = source.Rows.Count Else count = source.Columns.Count
    If rows Then
        uniformSize = source.EntireRow.RowHeight: uniformHidden = source.EntireRow.Hidden
    Else
        uniformSize = source.EntireColumn.ColumnWidth: uniformHidden = source.EntireColumn.Hidden
    End If
    If Not IsNull(uniformSize) And Not IsNull(uniformHidden) Then
        NxApplyDimensionRun destination, rows, 1, count, CDbl(uniformSize), CBool(uniformHidden)
        Exit Sub
    End If
    runStart = 1
    For index = 1 To count
        If rows Then
            Set part = source.Rows(index).EntireRow
            size = part.RowHeight
        Else
            Set part = source.Columns(index).EntireColumn
            size = part.ColumnWidth
        End If
        hidden = part.Hidden
        If index > 1 Then
            If size <> previousSize Or hidden <> previousHidden Then
                NxApplyDimensionRun destination, rows, runStart, index - runStart, previousSize, previousHidden
                runStart = index
            End If
        End If
        previousSize = size: previousHidden = hidden
    Next index
    NxApplyDimensionRun destination, rows, runStart, count - runStart + 1, previousSize, previousHidden
End Sub

Private Sub NxApplyDimensionRun(ByVal destination As Range, ByVal rows As Boolean, ByVal first As Long, ByVal count As Long, ByVal size As Double, ByVal hidden As Boolean)
    Dim target As Range
    If rows Then
        Set target = destination.Cells(first, 1).Resize(count, 1).EntireRow
        target.RowHeight = size
    Else
        Set target = destination.Cells(1, first).Resize(1, count).EntireColumn
        target.ColumnWidth = size
    End If
    target.Hidden = hidden
End Sub

Private Sub NxFileRequireNewOutput(ByVal outputPath As String, ByVal requiredExtension As String)
    Dim parentFolder As String
    If Len(Trim$(outputPath)) = 0 Then NxRaiseContractError "Copy-save output path is required"
    If Len(requiredExtension) < 2 Or Left$(requiredExtension, 1) <> "." Then NxRaiseContractError "Copy-save extension contract is invalid"
    If InStr(outputPath, vbNullChar) > 0 Or InStr(outputPath, vbCr) > 0 Or InStr(outputPath, vbLf) > 0 Then NxRaiseContractError "Copy-save output path contains control characters"
    If InStr(outputPath, "*") > 0 Or InStr(outputPath, "?") > 0 Then NxRaiseContractError "Copy-save output path contains wildcard characters"
    If LCase$(Right$(outputPath, Len(requiredExtension))) <> LCase$(requiredExtension) Then NxRaiseContractError "Copy-save output must use " & requiredExtension
    If Len(outputPath) > 218 Then NxRaiseContractError "Copy-save output path exceeds the Excel path limit"
    If NxFilePathExists(outputPath) Then NxRaiseContractError "Copy-save output already exists; automatic overwrite is forbidden"
    parentFolder = NxFileParentFolder(outputPath)
    If Not NxFileDirectoryExists(parentFolder) Then NxRaiseContractError "Copy-save output folder does not exist"
End Sub

Private Function NxFileTemporaryOutputPath(ByVal outputPath As String, ByVal requiredExtension As String) As String
    Dim candidate As String
    NxFileRequireNewOutput outputPath, requiredExtension
    candidate = Left$(outputPath, Len(outputPath) - Len(requiredExtension)) & ".nx-copy-save.tmp" & requiredExtension
    If NxFilePathExists(candidate) Then NxRaiseContractError "Copy-save temporary output already exists; automatic overwrite is forbidden"
    NxFileTemporaryOutputPath = candidate
End Function

Private Sub NxFileCommitTemporaryOutput(ByVal temporaryPath As String, ByVal finalPath As String)
    Dim startedAt As Double
    Dim lockAvailable As Boolean
    Dim moveError As Long
    If Not NxFilePathExists(temporaryPath) Then NxRaiseContractError "Copy-save temporary output is missing"
    If NxFilePathExists(finalPath) Then NxRaiseContractError "Copy-save output already exists; automatic overwrite is forbidden"
    startedAt = Timer
    Do
        lockAvailable = NxFileTryExclusiveHandle(temporaryPath)
        If lockAvailable Then Exit Do
        If NxFileElapsedSeconds(startedAt) >= 10# Then _
            NxRaiseContractError "Copy-save temporary output remained locked"
        DoEvents
    Loop
    If MoveFileExW(StrPtr(temporaryPath), StrPtr(finalPath), NX_MOVEFILE_WRITE_THROUGH) = 0 Then
        moveError = Err.LastDllError
        NxRaiseContractError "Copy-save output commit failed (Win32 " & CStr(moveError) & ")"
    End If
    If Not NxFilePathExists(finalPath) Or NxFilePathExists(temporaryPath) Then _
        NxRaiseContractError "Copy-save output commit was not materialized"
End Sub

Private Function NxFileTryExclusiveHandle(ByVal filePath As String) As Boolean
#If VBA7 Then
    Dim fileHandle As LongPtr
#Else
    Dim fileHandle As Long
#End If
    fileHandle = CreateFileW(StrPtr(filePath), NX_GENERIC_READ Or NX_GENERIC_WRITE, 0, 0, _
        NX_OPEN_EXISTING, NX_FILE_ATTRIBUTE_NORMAL, 0)
    If fileHandle = NX_INVALID_HANDLE_VALUE Then Exit Function
    CloseHandle fileHandle
    NxFileTryExclusiveHandle = True
End Function

Private Function NxFileElapsedSeconds(ByVal startedAt As Double) As Double
    Dim currentValue As Double
    currentValue = Timer
    If currentValue >= startedAt Then
        NxFileElapsedSeconds = currentValue - startedAt
    Else
        NxFileElapsedSeconds = 86400# - startedAt + currentValue
    End If
End Function

Private Sub NxFileDeleteCreatedFile(ByVal filePath As String, Optional ByVal suppressErrors As Boolean = False)
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    Dim normalizedPath As String, ownedTemporary As Boolean
    On Error GoTo Failed
    If Len(filePath) = 0 Or Not NxFilePathExists(filePath) Then Exit Sub
    normalizedPath = LCase$(filePath)
    ownedTemporary = InStr(1, normalizedPath, ".nx-copy-save.tmp.xlsx", vbBinaryCompare) > 0
    If Not ownedTemporary Then NxRaiseContractError "Refusing to delete an unowned copy-save path"
    If (GetAttr(filePath) And vbDirectory) = vbDirectory Then NxRaiseContractError "Refusing to delete a copy-save directory"
    Kill filePath
    If NxFilePathExists(filePath) Then NxRaiseContractError "Copy-save temporary output could not be removed"
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    If suppressErrors Then
        Debug.Print "NX-COPY-SAVE:CLEANUP_FAILED:" & failureDescription
        Exit Sub
    End If
    Err.Raise failureNumber, failureSource, failureDescription
End Sub

Private Function NxFileParentFolder(ByVal filePath As String) As String
    Dim separatorPosition As Long
    If Not ((Len(filePath) >= 3 And Mid$(filePath, 2, 1) = ":" And Mid$(filePath, 3, 1) = "\") Or Left$(filePath, 2) = "\\") Then NxRaiseContractError "Copy-save output path must be absolute"
    separatorPosition = InStrRev(filePath, "\")
    If separatorPosition = 0 Then NxRaiseContractError "Copy-save output path must be absolute"
    If separatorPosition = 3 And Mid$(filePath, 2, 1) = ":" Then
        NxFileParentFolder = Left$(filePath, separatorPosition)
    Else
        NxFileParentFolder = Left$(filePath, separatorPosition - 1)
    End If
    If Len(NxFileParentFolder) = 0 Then NxRaiseContractError "Copy-save output folder is invalid"
End Function

Private Function NxFilePathExists(ByVal filePath As String) As Boolean
    Dim attributes As Long
    On Error GoTo Missing
    attributes = GetAttr(filePath)
    NxFilePathExists = True
    Exit Function
Missing:
    NxFilePathExists = False
End Function

Private Function NxFileDirectoryExists(ByVal folderPath As String) As Boolean
    Dim attributes As Long
    On Error GoTo Missing
    attributes = GetAttr(folderPath)
    NxFileDirectoryExists = ((attributes And vbDirectory) = vbDirectory)
    Exit Function
Missing:
    NxFileDirectoryExists = False
End Function
