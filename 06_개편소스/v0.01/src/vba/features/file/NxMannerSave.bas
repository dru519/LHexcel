Attribute VB_Name = "NxMannerSave"
Option Explicit

Public Sub NxMannerSaveFromRibbon()
    Dim result As CNxResult
    Set result = NxRunMannerSave(Application.ActiveWorkbook)
    If result.Outcome = NxSuccess Then Exit Sub
    If result.Outcome = NxCancelled Then Exit Sub
    NxRaiseContractError result.Recovery
End Sub

Public Function NxRunMannerSave(ByVal sourceBook As Workbook) As CNxResult
    Dim view As Window, sheet As Worksheet, firstSheet As Worksheet, originalSheet As Object
    Dim snapshots As New Collection, snapshot As Variant, selectedSheets As Variant
    Dim observer As New CNxMannerSaveEvents
    Dim previousEvents As Boolean, previousUpdating As Boolean, suspended As Boolean
    Dim stateCaptured As Boolean, saved As Boolean, cancelled As Boolean, saveDialogAccepted As Boolean
    Dim failureDetail As String, outputPath As String, index As Long, phase As String

    previousEvents = Application.EnableEvents
    previousUpdating = Application.ScreenUpdating
    outputPath = "active-workbook"
    phase = "사전 검사"
    On Error GoTo Failed
    If sourceBook Is Nothing Then NxRaiseContractError "저장할 통합문서를 여세요."
    outputPath = sourceBook.FullName
    If sourceBook Is ThisWorkbook Then NxRaiseContractError "내엑셀 추가 기능 자체는 매너 저장할 수 없습니다."
    If sourceBook.IsAddin Then NxRaiseContractError "추가 기능 파일은 매너 저장할 수 없습니다."
    If Not (sourceBook Is Application.ActiveWorkbook) Then NxRaiseContractError "현재 통합문서에서 실행하세요."
    If sourceBook.ReadOnly Then NxRaiseContractError "읽기 전용 파일은 매너 저장할 수 없습니다."
    If Not previousEvents Then NxRaiseContractError "Excel 이벤트가 꺼져 있습니다. 저장 이벤트를 복원한 뒤 실행하세요."
    Set view = Application.ActiveWindow
    If view Is Nothing Then NxRaiseContractError "보이는 통합문서 창이 없습니다."
    For Each sheet In sourceBook.Worksheets
        If sheet.Visible = xlSheetVisible Then
            If firstSheet Is Nothing Then Set firstSheet = sheet
        End If
    Next sheet
    If firstSheet Is Nothing Then NxRaiseContractError "보이는 워크시트가 없습니다."

    Set originalSheet = view.ActiveSheet
    ReDim selectedSheets(0 To view.SelectedSheets.Count - 1)
    For index = 1 To view.SelectedSheets.Count
        selectedSheets(index - 1) = view.SelectedSheets(index).Name
    Next index
    stateCaptured = True
    phase = "포커스셀 일시 정지"
    NxFocusSuspendForOperation
    suspended = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    For Each sheet In sourceBook.Worksheets
        If sheet.Visible = xlSheetVisible Then
            phase = sheet.Name & " 시트 선택"
            sheet.Select
            phase = sheet.Name & " 선택 범위 보존"
            ReDim snapshot(0 To 5)
            Set snapshot(0) = sheet
            Set snapshot(1) = Application.Selection
            Set snapshot(2) = Application.ActiveCell
            phase = sheet.Name & " 화면 위치 보존"
            snapshot(3) = view.Zoom
            snapshot(4) = view.ScrollRow
            snapshot(5) = view.ScrollColumn
            snapshots.Add snapshot
            phase = sheet.Name & " 확대율 변경"
            view.Zoom = 100
            phase = sheet.Name & " A1 이동"
            Application.Goto sheet.Range("A1"), True
            phase = sheet.Name & " 스크롤 위치 변경"
            view.ScrollRow = 1
            view.ScrollColumn = 1
        End If
    Next sheet
    phase = "첫 시트 선택"
    firstSheet.Select
    Application.Goto firstSheet.Range("A1"), True
    view.Zoom = 100
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousUpdating

    ' Do not let Excel serialize the internal conditional-format focus rules.
    ' Keep normal WorkbookBeforeSave/AfterSave handlers enabled as well.
    phase = "저장 전 포커스셀 정리"
    If Not NxFocusControllerBeforeSave(sourceBook) Then NxRaiseContractError "포커스셀 저장 정리에 실패했습니다."
    observer.Start sourceBook
    phase = "파일 저장"
    If Len(sourceBook.Path) = 0 Then
        saveDialogAccepted = Application.Dialogs(xlDialogSaveAs).Show
        If Not saveDialogAccepted Then cancelled = True
    Else
        sourceBook.Save
    End If
    saved = observer.SaveSucceeded
    If Not saved Then
        cancelled = True
        GoTo RestoreFailure
    End If
    outputPath = sourceBook.FullName
    GoTo Finished
Failed:
    failureDetail = phase & ": " & Err.Description
RestoreFailure:
    On Error Resume Next
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    If stateCaptured Then NxRestoreMannerViews view, snapshots, originalSheet, selectedSheets
    If Err.Number <> 0 Then failureDetail = failureDetail & " 화면 복원 확인 필요: " & Err.Description
    On Error GoTo 0
Finished:
    On Error Resume Next
    observer.StopListening
    Application.EnableEvents = previousEvents
    Application.ScreenUpdating = previousUpdating
    NxFocusControllerAfterSave sourceBook, saved
    If suspended Then NxFocusResumeAfterOperation
    On Error GoTo 0
    If saved Then
        Set NxRunMannerSave = NxCreateResult(NX_FEATURE_FILE_MANNER_SAVE, NxSuccess, "complete", outputPath, True, vbNullString, "result_success")
    ElseIf cancelled And Len(failureDetail) = 0 Then
        Set NxRunMannerSave = NxCreateResult(NX_FEATURE_FILE_MANNER_SAVE, NxCancelled, "save_cancelled", outputPath, False, vbNullString, "result_cancelled")
    Else
        Set NxRunMannerSave = NxCreateResult(NX_FEATURE_FILE_MANNER_SAVE, NxEnvironmentError, "manner_save", outputPath, False, failureDetail, "result_environment_error")
    End If
End Function

Private Sub NxRestoreMannerViews(ByVal view As Window, ByVal snapshots As Collection, _
    ByVal originalSheet As Object, ByVal selectedSheets As Variant)
    Dim snapshot As Variant, sheet As Worksheet, target As Object, cell As Range
    view.Activate
    For Each snapshot In snapshots
        Set sheet = snapshot(0)
        sheet.Select
        Set target = snapshot(1)
        target.Select
        Set cell = snapshot(2)
        cell.Activate
        view.Zoom = snapshot(3)
        view.ScrollRow = snapshot(4)
        view.ScrollColumn = snapshot(5)
    Next snapshot
    originalSheet.Select
    If UBound(selectedSheets) > LBound(selectedSheets) Then originalSheet.Parent.Sheets(selectedSheets).Select
    originalSheet.Activate
End Sub
