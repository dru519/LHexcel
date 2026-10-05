Attribute VB_Name = "NxPngExport"
Option Explicit

Public Function NxExportRangePng(ByVal source As Range, ByVal outputPath As String, Optional ByVal includeHidden As Boolean = False) As CNxResult
    Dim tempBook As Workbook, tempSheet As Worksheet, chart As ChartObject
    Dim guard As CNxStateGuard, image As Shape
    Dim tempPath As String, fileSystem As Object
    Dim stage As String, detail As String
    On Error GoTo Failed
    stage = "입력 확인"
    If source Is Nothing Or source.Areas.Count <> 1 Then NxRaiseContractError "하나의 연속된 범위만 선택할 수 있습니다."
    If source.Cells.CountLarge > 10000 Then NxRaiseContractError "범위 PNG는 10,000개 셀을 초과할 수 없습니다."
    NxPngValidateOutput outputPath, vbNullString
    If source.Width <= 0 Or source.Height <= 0 Then NxRaiseContractError "보이는 셀 범위를 선택하세요."
    Set guard = New CNxStateGuard
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    tempPath = fileSystem.BuildPath(fileSystem.GetParentFolderName(outputPath), "nx-image-" & NxCreateRunUuid() & "." & fileSystem.GetExtensionName(outputPath))
    Application.EnableEvents = False
    Application.ScreenUpdating = True
    stage = "임시 통합 문서 만들기"
    Set tempBook = Workbooks.Add(xlWBATWorksheet)
    Set tempSheet = tempBook.Worksheets(1)
    stage = "범위 그림 복사"
    NxPngCopyRangePicture source
    DoEvents
    stage = "임시 차트 만들기"
    tempBook.Activate
    Set chart = tempSheet.ChartObjects.Add(0, 0, source.Width, source.Height)
    chart.Activate
    stage = "범위 그림 붙여넣기"
    chart.Chart.Paste
    DoEvents
    If chart.Chart.Shapes.Count <> 1 Then NxRaiseContractError "선택 범위 그림을 준비하지 못했습니다. 다시 실행하세요."
    Set image = chart.Chart.Shapes(1)
    image.Left = 0: image.Top = 0
    image.Width = source.Width: image.Height = source.Height
    chart.Chart.Refresh
    DoEvents
    stage = "PNG 내보내기"
    If Not chart.Chart.Export(Filename:=tempPath, FilterName:=NxImageExportFilter(outputPath)) Then NxRaiseContractError "선택 범위 그림 저장에 실패했습니다."
    If Not fileSystem.FileExists(tempPath) Then NxRaiseContractError "선택 범위 그림 파일을 만들지 못했습니다."
    If fileSystem.GetFile(tempPath).Size = 0 Then NxRaiseContractError "선택 범위 그림 파일이 비어 있습니다."
    NxPngValidateOutput outputPath, vbNullString
    Name tempPath As outputPath
    tempPath = vbNullString
    stage = "임시 통합 문서 닫기"
    tempBook.Close SaveChanges:=False
    Set tempBook = Nothing
    Application.CutCopyMode = False
    guard.Restore
    If Len(Dir$(outputPath)) = 0 Then NxRaiseContractError "범위 PNG 결과가 생성되지 않았습니다. (was not materialized)"
    Set NxExportRangePng = NxCreateResult("NX-FILE-RANGE-PNG", NxSuccess, "complete", outputPath, False, vbNullString, "result_success")
    Exit Function
Failed:
    detail = Err.Description
    On Error Resume Next
    If Not tempBook Is Nothing Then tempBook.Close SaveChanges:=False
    If Len(tempPath) > 0 Then
        If Not fileSystem Is Nothing Then
            If fileSystem.FileExists(tempPath) Then fileSystem.DeleteFile tempPath
        End If
    End If
    Application.CutCopyMode = False
    If Not guard Is Nothing Then guard.Restore
    On Error GoTo 0
    Set NxExportRangePng = NxCreateResult("NX-FILE-RANGE-PNG", NxEnvironmentError, "png_export", outputPath, False, _
        "stage=" & stage & " :: " & detail, "result_environment_error")
End Function

Private Sub NxPngCopyRangePicture(ByVal source As Range)
    Dim attempt As Long, failureNumber As Long, failureDescription As String
    ' Retry only this non-mutating copy stage; never repeat the file commit.
    For attempt = 1 To 3
        source.Worksheet.Parent.Activate
        source.Worksheet.Activate
        On Error Resume Next
        Err.Clear
        source.CopyPicture Appearance:=xlScreen, Format:=xlPicture
        failureNumber = Err.Number
        failureDescription = Err.Description
        Err.Clear
        On Error GoTo 0
        If failureNumber = 0 Then Exit Sub
        DoEvents
    Next attempt
    Err.Raise failureNumber, "NxPngExport.CopyPicture", failureDescription
End Sub

Public Function NxExportChartPng(ByVal chartObject As Object, ByVal outputPath As String, Optional ByVal replaceToken As String = vbNullString) As CNxResult
    Dim stage As String, detail As String
    On Error GoTo Failed
    stage = "차트 PNG 내보내기"
    NxPngExportChart chartObject, outputPath, replaceToken
    If Len(Dir$(outputPath)) = 0 Then NxRaiseContractError "차트 PNG 결과가 생성되지 않았습니다. (was not materialized)"
    Set NxExportChartPng = NxCreateResult("NX-FILE-CHART-PNG", NxSuccess, "complete", outputPath, False, vbNullString, "result_success")
    Exit Function
Failed:
    detail = Err.Description
    Set NxExportChartPng = NxCreateResult("NX-FILE-CHART-PNG", NxEnvironmentError, "png_export", outputPath, False, _
        "stage=" & stage & " :: " & detail, "result_environment_error")
End Function
