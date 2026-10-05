Attribute VB_Name = "NxFileController"
Option Explicit

Public Sub NxRibbonOpenFile(ByVal control As Object)
    NxFileOpenManager
End Sub

Public Sub NxFileOpenManager(Optional ByVal featureId As String = vbNullString)
    If Len(featureId) = 0 Then NxRaiseContractError "파일관리 목록에서 실행할 기능을 선택하세요."
    NxRouteFeature featureId
End Sub

Public Sub NxFileOpenFeatureDialog(ByVal featureId As String)
    Dim consolidateForm As FNxFileConsolidate
    Dim outputForm As FNxFileOutput
    Dim compareForm As FNxWorkbookCompare
    Dim dataCompareForm As FNxDataCompare
    Dim renameForm As FNxBatchRename
    Dim pdfForm As FNxPdf
    Dim requestedReport As String

    Select Case featureId
        Case "NX-FILE-CONSOLIDATE"
            Set consolidateForm = New FNxFileConsolidate
            consolidateForm.BindFeature
            consolidateForm.Show vbModal
        Case "NX-FILE-RANGE-PNG"
            NxFileSaveRangeFromRibbon
        Case "NX-FILE-CHART-PNG"
            Set outputForm = New FNxFileOutput
            outputForm.BindFeature featureId
            outputForm.Show vbModal
        Case "NX-FILE-FILE-COMPARE"
            Set compareForm = New FNxWorkbookCompare
            compareForm.BindFeature
            compareForm.Show vbModal
            Unload compareForm
            Set compareForm = Nothing
        Case NX_FEATURE_FILE_WORKBOOK_COMPARE, "NX-FILE-SHEET-COMPARE"
            Set dataCompareForm = New FNxDataCompare
            dataCompareForm.BindFeature featureId
            dataCompareForm.Show vbModal
            Unload dataCompareForm
            Set dataCompareForm = Nothing
        Case NX_FEATURE_FILE_BATCH_RENAME, NX_FEATURE_FILE_SHEET_BATCH_RENAME
            Set renameForm = New FNxBatchRename
            renameForm.BindFeature featureId
            renameForm.Show vbModal
        Case NX_FEATURE_FILE_PDF_CURRENT_SHEET, NX_FEATURE_FILE_PDF_EACH_SHEET, _
             NX_FEATURE_FILE_PDF_SELECTED_COMBINED, NX_FEATURE_FILE_PDF_ALL_COMBINED, _
             NX_FEATURE_FILE_PDF_SETTINGS
            Set pdfForm = New FNxPdf
            pdfForm.BindFeature featureId
            pdfForm.Show vbModal
            Unload pdfForm
            Set pdfForm = Nothing
        Case Else
            NxRaiseContractError "지원하지 않는 파일 기능입니다."
    End Select
End Sub

Public Sub NxFileSaveRangeFromRibbon()
    Dim result As CNxResult, target As Range, selectedPath As Variant
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "그림으로 저장할 범위를 먼저 선택하세요."
    Set target = Application.Selection
    selectedPath = Application.GetSaveAsFilename(InitialFileName:=NxFileAutomaticRangeImagePath(target), _
        FileFilter:="PNG 파일 (*.png), *.png, JPG 파일 (*.jpg), *.jpg", _
        Title:="그림 저장 폴더·이름·파일 형식 선택")
    If VarType(selectedPath) = vbBoolean Then Exit Sub
    Set result = NxFileRun(NX_FEATURE_FILE_RANGE_PNG, CStr(selectedPath), selectedTarget:=target)
    If result Is Nothing Then NxRaiseContractError "그림 저장 결과를 확인할 수 없습니다."
    If result.Outcome <> NxSuccess And result.Outcome <> NxCancelled Then NxRaiseContractError result.Recovery
End Sub

Public Function NxFileSaveRangeAutomatic(ByVal source As Range) As CNxResult
    Dim current As Range, outputPath As String
    If source Is Nothing Then NxRaiseContractError "그림으로 저장할 범위를 먼저 선택하세요."
    If source.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위를 선택하세요."
    If source.CountLarge > 10000 Then NxRaiseContractError "그림 저장은 한 번에 10,000셀까지 가능합니다."
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "그림으로 저장할 범위를 다시 선택하세요."
    Set current = Application.Selection
    If Not current.Worksheet Is source.Worksheet Then NxRaiseContractError "선택한 시트가 변경되었습니다."
    If current.Address <> source.Address Then NxRaiseContractError "선택 범위가 변경되었습니다."
    outputPath = NxFileAutomaticRangeImagePath(source)
    Set NxFileSaveRangeAutomatic = NxFileRun(NX_FEATURE_FILE_RANGE_PNG, outputPath, selectedTarget:=source, silentRangeImage:=True)
End Function

Public Function NxFileAutomaticRangeImagePath(ByVal source As Range) As String
    Dim folder As String, stem As String, candidate As String, bookName As String
    Dim invalid As Variant, item As Variant, number As Long, fileSystem As Object
    If source Is Nothing Then NxRaiseContractError "그림 저장 범위가 필요합니다."
    folder = source.Worksheet.Parent.Path
    If Len(folder) = 0 Then folder = Application.DefaultFilePath
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If Not fileSystem.FolderExists(folder) Then NxRaiseContractError "그림을 저장할 문서 폴더를 찾을 수 없습니다. 문서를 로컬 폴더에 저장하고 다시 실행하세요."
    bookName = fileSystem.GetBaseName(source.Worksheet.Parent.Name)
    stem = bookName & "_" & source.Worksheet.Name & "_" & Replace$(source.Address(False, False), ":", "-")
    invalid = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    For Each item In invalid
        stem = Replace$(stem, CStr(item), "_")
    Next item
    stem = Left$(stem, 100) & "_" & Format$(Now, "yyyymmdd_hhnnss")
    For number = 0 To 9999
        candidate = fileSystem.BuildPath(folder, stem & IIf(number = 0, vbNullString, "_" & CStr(number)) & ".png")
        If Not fileSystem.FileExists(candidate) And Not fileSystem.FolderExists(candidate) Then
            NxFileAutomaticRangeImagePath = candidate
            Exit Function
        End If
    Next number
    NxRaiseContractError "그림 저장 이름을 만들지 못했습니다. 저장 폴더를 확인하세요."
End Function

Public Sub NxFileRunCopySaveFromRibbon(ByVal featureId As String)
    Dim chosenPath As Variant
    Dim defaultName As String
    Dim result As CNxResult
    Dim target As Object, picker As Object, completion As FNxCopySaveResult

    If ActiveWorkbook Is Nothing Then NxRaiseContractError "열린 통합문서가 없습니다."
    Select Case featureId
        Case NX_FEATURE_FILE_SHEET_COPY_SAVE
            If Not TypeOf ActiveSheet Is Worksheet Then NxRaiseContractError "현재 워크시트를 찾을 수 없습니다."
            Set target = ActiveSheet
            defaultName = ActiveWorkbook.Path & Application.PathSeparator & ActiveSheet.Name & "_사본.xlsx"
        Case NX_FEATURE_FILE_RANGE_COPY_SAVE
            If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "저장할 범위를 먼저 선택하세요."
            Set target = Application.Selection
            defaultName = ActiveWorkbook.Path & Application.PathSeparator & ActiveSheet.Name & "_선택범위.xlsx"
        Case Else
            NxRaiseContractError "지원하지 않는 사본저장 기능입니다."
    End Select

    If featureId = NX_FEATURE_FILE_RANGE_COPY_SAVE Then
        Set picker = Application.FileDialog(4)
        picker.Title = "선택범위 사본을 저장할 폴더"
        picker.AllowMultiSelect = False
        If picker.Show <> -1 Then Exit Sub
        chosenPath = NxFileVacantCopyPath(CStr(picker.SelectedItems(1)), ActiveSheet.Name)
    Else
        If Len(ActiveWorkbook.Path) = 0 Then defaultName = featureId & ".xlsx"
    chosenPath = Application.GetSaveAsFilename(InitialFileName:=defaultName, _
        FileFilter:="Excel 통합문서 (*.xlsx), *.xlsx", Title:="내엑셀 - 사본 저장")
    If VarType(chosenPath) = vbBoolean Then
        If chosenPath = False Then Exit Sub
    End If
    End If
    Set result = NxFileRun(featureId, CStr(chosenPath), Empty, False, target)
    If result Is Nothing Then NxRaiseContractError "저장 결과를 확인할 수 없습니다."
    If result.Outcome = NxCancelled Then Exit Sub
    If result.Outcome <> NxSuccess Then NxRaiseContractError result.Recovery
    If featureId = NX_FEATURE_FILE_RANGE_COPY_SAVE Then
        Set completion = New FNxCopySaveResult
        completion.BindPath CStr(chosenPath)
        completion.Show vbModal
    End If
End Sub

Public Function NxFileVacantCopyPath(ByVal folder As String, ByVal sheetName As String) As String
    Dim fs As Object, stem As String, candidate As String, n As Long, token As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Not fs.FolderExists(folder) Then NxRaiseContractError "저장 폴더를 찾을 수 없습니다."
    stem = sheetName & "_선택범위"
    For Each token In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
        stem = Replace$(stem, CStr(token), "_")
    Next token
    For n = 0 To 9999
        candidate = fs.BuildPath(folder, stem & IIf(n = 0, vbNullString, "_" & CStr(n)) & ".xlsx")
        If Not fs.FileExists(candidate) And Not fs.FolderExists(candidate) Then NxFileVacantCopyPath = candidate: Exit Function
    Next n
    NxRaiseContractError "사용할 파일 이름을 만들지 못했습니다."
End Function

Public Function NxFileRun(ByVal featureId As String, ByVal outputPath As String, _
    Optional ByVal inputPaths As Variant, _
    Optional ByVal allowPartial As Boolean = False, Optional ByVal selectedTarget As Object, _
    Optional ByVal secondApproval As Boolean = False, Optional ByVal workflowOptions As Object, _
    Optional ByVal silentRangeImage As Boolean = False) As CNxResult
    Dim commandObject As New CNxFileFeatureCommand
    Dim command As INxFeatureCommand
    Dim definition As CNxFeatureDefinition
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim result As CNxResult
    Dim sourceWorkbook As Workbook
    Dim sourceWasSaved As Boolean
    Dim approval As CNxApproval
    Dim failure As Long, detail As String
    On Error GoTo Failed
    If silentRangeImage Then
        If featureId <> NX_FEATURE_FILE_RANGE_PNG Then NxRaiseContractError "자동 그림 저장 대상이 올바르지 않습니다."
        NxPngValidateOutput outputPath, vbNullString
    End If
    Set sourceWorkbook = ActiveWorkbook
    If sourceWorkbook Is Nothing Then NxRaiseContractError "File execution requires an active source workbook"
    sourceWasSaved = sourceWorkbook.Saved
    commandObject.Configure featureId, outputPath, inputPaths, allowPartial, selectedTarget, secondApproval, workflowOptions
    Set command = commandObject
    Set definition = NxFileFeatureDefinition(featureId)
    Set ticket = router.Prepare(definition, command)
    If ticket.Decision.ResolvedGrade <> NxExecutionPlanned Then NxRaiseContractError "File execution grade is invalid"
    If silentRangeImage Then
        ' The ribbon action approves this new, collision-free local image only.
        Set approval = router.TakePlannedApproval(ticket, True)
        Set result = router.RunApproved(ticket, approval)
    Else
        Set result = NxRunPlannedFile(ticket, command, definition)
    End If
    If sourceWasSaved And Not result.SourceChanged And NxFileOutputPreservesSourceSavedState(featureId) Then
        sourceWorkbook.Saved = True
    End If
    If (featureId = NX_FEATURE_FILE_WORKBOOK_COMPARE Or featureId = "NX-FILE-FILE-COMPARE") And Len(outputPath) = 0 Then NxWorkbookCompareFinishReport (result.Outcome = NxSuccess)
    If featureId = NX_FEATURE_FILE_WORKBOOK_COMPARE Or featureId = "NX-FILE-SHEET-COMPARE" Then NxDataCompareFinish (result.Outcome = NxSuccess)
    Set NxFileRun = result
    Exit Function
Failed:
    failure = Err.Number: detail = Err.Description
    On Error Resume Next
    If (featureId = NX_FEATURE_FILE_WORKBOOK_COMPARE Or featureId = "NX-FILE-FILE-COMPARE") And Len(outputPath) = 0 Then NxWorkbookCompareFinishReport False
    If featureId = NX_FEATURE_FILE_WORKBOOK_COMPARE Or featureId = "NX-FILE-SHEET-COMPARE" Then NxDataCompareFinish False
    On Error GoTo 0
    Err.Raise failure, "NxFileRun", detail
End Function

Public Function NxRunPlannedFile(ByVal ticket As CNxExecutionTicket, ByVal command As INxFeatureCommand, _
    ByVal definition As CNxFeatureDefinition) As CNxResult
    Dim router As New CNxExecutionRouter
    Set NxRunPlannedFile = router.ExecuteUserAction(ticket)
End Function

Private Function NxFileOutputPreservesSourceSavedState(ByVal featureId As String) As Boolean
    Select Case featureId
        Case NX_FEATURE_FILE_SHEET_COPY_SAVE, NX_FEATURE_FILE_RANGE_COPY_SAVE, _
             NX_FEATURE_FILE_RANGE_PNG, NX_FEATURE_FILE_CHART_PNG, _
             NX_FEATURE_FILE_PDF_CURRENT_SHEET, NX_FEATURE_FILE_PDF_EACH_SHEET, _
             NX_FEATURE_FILE_PDF_SELECTED_COMBINED, NX_FEATURE_FILE_PDF_ALL_COMBINED, _
             NX_FEATURE_FILE_PDF_SETTINGS
            NxFileOutputPreservesSourceSavedState = True
    End Select
End Function

Private Function NxFileResultSummary(ByVal featureId As String) As String
    Select Case featureId
        Case NX_FEATURE_FILE_FOLDER_CREATE: NxFileResultSummary = "계획된 폴더 생성"
        Case NX_FEATURE_FILE_BATCH_RENAME, NX_FEATURE_FILE_SHEET_BATCH_RENAME: NxFileResultSummary = "검토한 이름 변경"
        Case NX_FEATURE_FILE_PDF_SETTINGS: NxFileResultSummary = "PDF 기본 설정 저장"
        Case Else: NxFileResultSummary = "새 파일 출력"
    End Select
End Function

Private Function NxFileTargetSummary(ByVal featureId As String) As String
    If featureId = "NX-FILE-FOLDER-CREATE" Then
        NxFileTargetSummary = "승인된 기준 폴더"
    Else
        NxFileTargetSummary = "선택 대상"
    End If
End Function

Private Function NxFileRecoverySummary(ByVal featureId As String) As String
    Select Case featureId
        Case NX_FEATURE_FILE_FOLDER_CREATE
            NxFileRecoverySummary = "원본 통합문서는 유지되며 계획에서 승인한 폴더만 생성합니다."
        Case NX_FEATURE_FILE_BATCH_RENAME
            NxFileRecoverySummary = "2단계 임시 이름과 저널을 사용해 실패 시 원래 파일명으로 복구합니다."
        Case NX_FEATURE_FILE_SHEET_BATCH_RENAME
            NxFileRecoverySummary = "2단계 임시 이름을 사용해 실패 시 원래 시트명으로 복구합니다."
        Case NX_FEATURE_FILE_PDF_SETTINGS
            NxFileRecoverySummary = "임시 파일과 백업을 사용해 설정 파일을 원자적으로 교체합니다."
        Case Else
            NxFileRecoverySummary = "원본을 유지하고 출력 파일만 변경합니다."
    End Select
End Function

Public Function NxFileActionSummary(ByVal featureId As String) As String
    Select Case featureId
        Case "NX-FILE-MANNER-SAVE": NxFileActionSummary = "각 표시 시트를 A1·100%로 맞추고 현재 파일을 저장합니다."
        Case "NX-FILE-CONSOLIDATE": NxFileActionSummary = "선택 방식에 따라 표시 시트를 모으거나 한 시트에 값으로 이어붙여 새 통합문서를 만듭니다."
        Case "NX-FILE-RANGE-PNG": NxFileActionSummary = "선택 범위를 원본 변경 없이 PNG 또는 JPG로 저장합니다."
        Case "NX-FILE-CHART-PNG": NxFileActionSummary = "선택 차트를 원본 변경 없이 PNG 또는 JPG로 저장합니다."
        Case "NX-FILE-FOLDER-CREATE": NxFileActionSummary = NxFolderCreateSummary()
        Case "NX-FILE-SHEET-COPY-SAVE": NxFileActionSummary = "Copy a worksheet to a new workbook without changing the original."
        Case "NX-FILE-RANGE-COPY-SAVE": NxFileActionSummary = "Copy a range to a new workbook without changing the original."
        Case NX_FEATURE_FILE_WORKBOOK_COMPARE: NxFileActionSummary = "두 통합문서를 읽기 전용으로 비교해 새 xlsx 보고서를 만듭니다."
        Case NX_FEATURE_FILE_BATCH_RENAME: NxFileActionSummary = "검토한 파일 이름을 2단계 저널 방식으로 일괄 변경합니다."
        Case NX_FEATURE_FILE_SHEET_BATCH_RENAME: NxFileActionSummary = "현재 통합문서의 시트 이름을 2단계 방식으로 일괄 변경합니다."
        Case NX_FEATURE_FILE_PDF_CURRENT_SHEET: NxFileActionSummary = "현재 시트를 새 PDF 파일로 저장합니다."
        Case NX_FEATURE_FILE_PDF_EACH_SHEET: NxFileActionSummary = "보이는 각 시트를 개별 PDF 파일로 저장합니다."
        Case NX_FEATURE_FILE_PDF_SELECTED_COMBINED: NxFileActionSummary = "선택한 시트를 한 PDF 파일로 저장합니다."
        Case NX_FEATURE_FILE_PDF_ALL_COMBINED: NxFileActionSummary = "통합문서 전체를 한 PDF 파일로 저장합니다."
        Case NX_FEATURE_FILE_PDF_SETTINGS: NxFileActionSummary = "PDF 기본 저장 폴더·이름·품질을 로컬 설정에 저장합니다."
        Case Else: NxRaiseContractError "Unknown file feature"
    End Select
End Function
