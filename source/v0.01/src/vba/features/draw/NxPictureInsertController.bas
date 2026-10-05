Attribute VB_Name = "NxPictureInsertController"
Option Explicit

Public Function NxPictureInsertNativePreview(ByVal request As CNxPictureInsertRequest, ByVal preview As CNxPictureInsertPreview) As Boolean
    Dim service As Object, paths As Collection, addresses As Collection, area As Range
    Dim rows() As Variant, index As Long
    Set service = NxHostCreatePicturePreview()
    If service Is Nothing Then NxPictureInsertNativePreview = True: Exit Function
    If Not preview.Revalidate(request) Then NxRaiseContractError "그림 미리보기 대상이 변경되었습니다."
    Set paths = request.Files: Set addresses = preview.TargetAddresses
    ReDim rows(0 To preview.InsertCount - 1, 0 To 3)
    For index = 1 To preview.InsertCount
        Set area = request.Target.Worksheet.Range(CStr(addresses(index)))
        rows(index - 1, 0) = CStr(paths(index)): rows(index - 1, 1) = CStr(addresses(index))
        rows(index - 1, 2) = CDbl(area.Width): rows(index - 1, 3) = CDbl(area.Height)
    Next index
    NxPictureInsertNativePreview = CBool(service.ShowPreview(rows, request.SizeMode, CStr(Application.Hwnd)))
    If NxPictureInsertNativePreview Then
        If Not preview.Revalidate(request) Then NxRaiseContractError "미리보기 중 그림 또는 대상 셀이 변경되었습니다. 다시 확인하세요."
    End If
End Function

Public Function NxPictureInsertRun(ByVal target As Range, ByVal filePaths As Collection, Optional ByVal sizeMode As String = "FIT") As CNxResult
    Dim request As CNxPictureInsertRequest, preview As CNxPictureInsertPreview
    Set request = NxPictureInsertCreateRequest(target, filePaths, sizeMode)
    Set preview = NxPictureInsertBuildPreview(request)
    Set NxPictureInsertRun = NxPictureInsertRunSealed(request, preview)
End Function

Public Function NxPictureInsertCreateRequest(ByVal target As Range, ByVal filePaths As Collection, Optional ByVal sizeMode As String = "FIT") As CNxPictureInsertRequest
    Dim request As New CNxPictureInsertRequest
    request.Configure target, filePaths, sizeMode
    Set NxPictureInsertCreateRequest = request
End Function

Public Function NxPictureInsertRunSealed(ByVal request As CNxPictureInsertRequest, ByVal preview As CNxPictureInsertPreview) As CNxResult
    Dim journal As New CNxPictureInsertJournal
    Dim commandObject As New CNxPictureInsertCommand, command As INxFeatureCommand
    If request Is Nothing Or preview Is Nothing Then NxRaiseContractError "봉인된 그림 삽입 요청과 미리보기가 필요합니다."
    If Not preview.IsSealed Then NxRaiseContractError "봉인되지 않은 그림 삽입 미리보기입니다."
    commandObject.Configure request, preview, journal: Set command = commandObject
    Set NxPictureInsertRunSealed = NxRunDrawingFeatureCommand(command, NxDrawingTargetSummary(request.Target))
End Function

Public Function NxPictureInsertBuildPreviewFor(ByVal target As Range, ByVal filePaths As Collection, Optional ByVal sizeMode As String = "FIT") As CNxPictureInsertPreview
    Dim request As CNxPictureInsertRequest
    Set request = NxPictureInsertCreateRequest(target, filePaths, sizeMode)
    Set NxPictureInsertBuildPreviewFor = NxPictureInsertBuildPreview(request)
End Function

Public Function NxPictureInsertPickFiles() As Collection
    Dim picker As FileDialog, answer As New Collection
    Dim originalWorkbook As Workbook, originalSheet As Object, originalSelection As Object
    Dim originalScreenUpdating As Boolean, originalEnableEvents As Boolean, originalDisplayAlerts As Boolean
    Dim originalCalculation As XlCalculation
    Dim selectedPath As Variant
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    Dim restored As Boolean, stateCaptured As Boolean
    On Error GoTo Failed
    Set originalWorkbook = Application.ActiveWorkbook
    Set originalSheet = Application.ActiveSheet
    Set originalSelection = Application.Selection
    originalScreenUpdating = Application.ScreenUpdating
    originalEnableEvents = Application.EnableEvents
    originalDisplayAlerts = Application.DisplayAlerts
    originalCalculation = Application.Calculation
    stateCaptured = True
    Set picker = Application.FileDialog(3)
    With picker
        .AllowMultiSelect = True
        .Title = "삽입할 그림 파일 선택"
        .Filters.Clear
        .Filters.Add "지원 이미지", "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff"
        If .Show <> -1 Then GoTo Cancelled
        If .SelectedItems.Count < 1 Then GoTo Cancelled
        If .SelectedItems.Count > 500 Then GoTo InvalidSelection
        For Each selectedPath In .SelectedItems
            If Len(Trim$(CStr(selectedPath))) = 0 Then GoTo InvalidSelection
            answer.Add CStr(selectedPath)
        Next selectedPath
    End With
    If stateCaptured Then restored = RestorePictureInsertPickerState(originalWorkbook, originalSheet, originalSelection, originalScreenUpdating, originalEnableEvents, originalDisplayAlerts, originalCalculation)
    If Not restored Then NxRaiseContractError "그림 파일 선택기 상태 복원에 실패했습니다."
    Set NxPictureInsertPickFiles = answer
    Exit Function
Cancelled:
    If stateCaptured Then restored = RestorePictureInsertPickerState(originalWorkbook, originalSheet, originalSelection, originalScreenUpdating, originalEnableEvents, originalDisplayAlerts, originalCalculation)
    If Not restored Then NxRaiseContractError "그림 파일 선택 취소 후 상태 복원에 실패했습니다."
    Set NxPictureInsertPickFiles = Nothing
    Exit Function
InvalidSelection:
    If stateCaptured Then restored = RestorePictureInsertPickerState(originalWorkbook, originalSheet, originalSelection, originalScreenUpdating, originalEnableEvents, originalDisplayAlerts, originalCalculation)
    If Not restored Then NxRaiseContractError "그림 파일 선택 초과 후 상태 복원에 실패했습니다."
    NxRaiseContractError "그림 파일은 1개 이상 500개 이하로 선택해야 합니다."
    Exit Function
Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    If stateCaptured Then
        restored = RestorePictureInsertPickerState(originalWorkbook, originalSheet, originalSelection, originalScreenUpdating, originalEnableEvents, originalDisplayAlerts, originalCalculation)
    Else
        Debug.Print "NX-DRAW-INSERT-PICTURE:PICKER_STATE_CAPTURE_FAILED"
    End If
    Err.Raise failureNumber, failureSource, failureDescription
End Function

Private Function RestorePictureInsertPickerState(ByVal originalWorkbook As Workbook, ByVal originalSheet As Object, ByVal originalSelection As Object, _
    ByVal originalScreenUpdating As Boolean, ByVal originalEnableEvents As Boolean, ByVal originalDisplayAlerts As Boolean, _
    ByVal originalCalculation As XlCalculation) As Boolean
    Dim restoreFailed As Boolean
    ' FileDialog normally leaves the selection unchanged. Do not activate/select
    ' again or reassign Calculation when the captured state already matches.
    If PictureInsertPickerStateMatches(originalWorkbook, originalSheet, originalSelection, originalScreenUpdating, originalEnableEvents, originalDisplayAlerts, originalCalculation) Then
        RestorePictureInsertPickerState = True
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    If Not originalWorkbook Is Nothing Then
        If Not Application.ActiveWorkbook Is originalWorkbook Then originalWorkbook.Activate
    End If
    If Err.Number <> 0 Then restoreFailed = True: Err.Clear
    If Not originalSheet Is Nothing Then
        If Not Application.ActiveSheet Is originalSheet Then originalSheet.Activate
    End If
    If Err.Number <> 0 Then restoreFailed = True: Err.Clear
    If Not originalSelection Is Nothing Then
        If Not PictureInsertPickerSelectionMatches(originalSelection) Then originalSelection.Select
    End If
    If Err.Number <> 0 Then restoreFailed = True: Err.Clear
    If Application.ScreenUpdating <> originalScreenUpdating Then Application.ScreenUpdating = originalScreenUpdating
    If Err.Number <> 0 Then restoreFailed = True: Err.Clear
    If Application.EnableEvents <> originalEnableEvents Then Application.EnableEvents = originalEnableEvents
    If Err.Number <> 0 Then restoreFailed = True: Err.Clear
    If Application.DisplayAlerts <> originalDisplayAlerts Then Application.DisplayAlerts = originalDisplayAlerts
    If Err.Number <> 0 Then restoreFailed = True: Err.Clear
    If Application.Calculation <> originalCalculation Then Application.Calculation = originalCalculation
    If Err.Number <> 0 Then restoreFailed = True: Err.Clear
    On Error GoTo 0
    If restoreFailed Or Not PictureInsertPickerStateMatches(originalWorkbook, originalSheet, originalSelection, originalScreenUpdating, originalEnableEvents, originalDisplayAlerts, originalCalculation) Then
        Debug.Print "NX-DRAW-INSERT-PICTURE:PICKER_RESTORE_FAILED"
        Exit Function
    End If
    RestorePictureInsertPickerState = True
End Function

Private Function PictureInsertPickerStateMatches(ByVal originalWorkbook As Workbook, ByVal originalSheet As Object, ByVal originalSelection As Object, _
    ByVal originalScreenUpdating As Boolean, ByVal originalEnableEvents As Boolean, ByVal originalDisplayAlerts As Boolean, _
    ByVal originalCalculation As XlCalculation) As Boolean
    On Error GoTo Failed
    If Not originalWorkbook Is Nothing Then
        If Application.ActiveWorkbook Is Nothing Then Exit Function
        If Not Application.ActiveWorkbook Is originalWorkbook Then Exit Function
    End If
    If Not originalSheet Is Nothing Then
        If Application.ActiveSheet Is Nothing Then Exit Function
        If Not Application.ActiveSheet Is originalSheet Then Exit Function
    End If
    If Not PictureInsertPickerSelectionMatches(originalSelection) Then Exit Function
    If Application.ScreenUpdating <> originalScreenUpdating Then Exit Function
    If Application.EnableEvents <> originalEnableEvents Then Exit Function
    If Application.DisplayAlerts <> originalDisplayAlerts Then Exit Function
    If Application.Calculation <> originalCalculation Then Exit Function
    PictureInsertPickerStateMatches = True
    Exit Function
Failed:
    PictureInsertPickerStateMatches = False
End Function

Private Function PictureInsertPickerSelectionMatches(ByVal originalSelection As Object) As Boolean
    Dim current As Object, selectedShapes As Object, currentShapes As Object, index As Long
    On Error GoTo Different
    Set current = Application.Selection
    If originalSelection Is Nothing Then
        PictureInsertPickerSelectionMatches = (current Is Nothing)
        Exit Function
    End If
    If current Is Nothing Then Exit Function
    If TypeName(current) <> TypeName(originalSelection) Then Exit Function
    ' Excel returns fresh Range/Picture wrappers on each Selection read.
    ' Compare the underlying selection, not COM wrapper identity.
    Select Case TypeName(originalSelection)
        Case "Range"
            If Not current.Worksheet Is originalSelection.Worksheet Then Exit Function
            PictureInsertPickerSelectionMatches = (current.Address(External:=True) = originalSelection.Address(External:=True))
        Case "Picture", "Shape"
            If Not current.Parent Is originalSelection.Parent Then Exit Function
            PictureInsertPickerSelectionMatches = (current.Name = originalSelection.Name)
        Case "ShapeRange", "DrawingObjects"
            ' Multi-selection is a fresh DrawingObjects wrapper in native Excel.
            If TypeName(originalSelection) = "DrawingObjects" Then
                Set selectedShapes = originalSelection.ShapeRange
                Set currentShapes = current.ShapeRange
            Else
                Set selectedShapes = originalSelection
                Set currentShapes = current
            End If
            If currentShapes.Count <> selectedShapes.Count Then Exit Function
            For index = 1 To currentShapes.Count
                If Not currentShapes.Item(index).Parent Is selectedShapes.Item(index).Parent Then Exit Function
                If currentShapes.Item(index).ID <> selectedShapes.Item(index).ID Then Exit Function
            Next index
            PictureInsertPickerSelectionMatches = True
        Case Else
            PictureInsertPickerSelectionMatches = (current Is originalSelection)
    End Select
    Exit Function
Different:
    PictureInsertPickerSelectionMatches = False
End Function
