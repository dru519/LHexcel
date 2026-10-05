Attribute VB_Name = "NxDrawController"
Option Explicit

Public Sub NxDrawApplyTitleTable(ByVal target As Range, Optional ByVal preserveAlignment As Boolean = True, Optional ByVal includeHeader As Boolean = True, Optional ByVal autoFitColumns As Boolean = False)
    NxDrawApplyTableOptions target, NX_FEATURE_DRAW_TITLE_TABLE, NX_ROLE_STYLE_MODE_MONO, "제목", _
        IIf(includeHeader, 1, 0), NxHeaderFontSize(), xlCenter, NxBodyFontSize(), _
        IIf(preserveAlignment, xlGeneral, xlCenter), False, autoFitColumns
End Sub

Public Sub NxDrawApplyBusinessTable(ByVal target As Range, Optional ByVal includeTotalRow As Boolean = False, Optional ByVal preserveAlignment As Boolean = True, Optional ByVal includeHeader As Boolean = True, Optional ByVal autoFitColumns As Boolean = False)
    NxDrawApplyTableOptions target, NX_FEATURE_DRAW_BUSINESS_TABLE, NX_ROLE_STYLE_MODE_MONO, "열린 표", _
        IIf(includeHeader, 1, 0), 9#, xlCenter, 9#, _
        IIf(preserveAlignment, xlGeneral, xlCenter), includeTotalRow, autoFitColumns
End Sub

Public Sub NxDrawApplyTableOptions(ByVal target As Range, ByVal featureId As String, ByVal displayMode As String, _
    ByVal tableStyle As String, ByVal headerRowCount As Long, ByVal headerFontSize As Double, _
    ByVal headerAlignment As Long, ByVal bodyFontSize As Double, ByVal bodyAlignment As Long, _
    Optional ByVal includeTotalRow As Boolean = False, Optional ByVal autoFitColumns As Boolean = False)

    Dim snapshot As New CNxFormatSnapshot
    Dim header As Range
    Dim body As Range
    Dim total As Range
    Dim errorNumber As Long
    Dim errorDescription As String
    EnsureDrawingTarget target
    If headerRowCount < 0 Or headerRowCount > 5 Or headerRowCount > target.Rows.Count Then _
        NxRaiseContractError "머리글 행 수가 선택 범위와 맞지 않습니다."
    snapshot.Capture target, Not NxRunnerOwnsTableRollback(featureId, target)
    On Error GoTo Failed
    If featureId = NX_FEATURE_DRAW_TITLE_TABLE Then
        NxDrawApplyTitleOptions target, displayMode, tableStyle, headerFontSize, headerAlignment
    ElseIf featureId = NX_FEATURE_DRAW_BUSINESS_TABLE Then
        If headerRowCount > 0 Then
            Set header = target.Rows(1).Resize(headerRowCount, target.Columns.Count)
            NxApplyHeaderToken header
            header.Interior.Color = NxDrawTableHeaderFill(displayMode)
            header.Font.Size = headerFontSize
            header.HorizontalAlignment = headerAlignment
            If UCase$(displayMode) = NX_ROLE_STYLE_MODE_MONO Then
                header.Font.Color = RGB(32, 32, 32)
            End If
        End If
        If target.Rows.Count > headerRowCount Then
            Set body = target.Offset(headerRowCount, 0).Resize(target.Rows.Count - headerRowCount, target.Columns.Count)
            NxApplyBodyToken body, (bodyAlignment = xlGeneral)
            body.Font.Size = bodyFontSize
            If bodyAlignment <> xlGeneral Then body.HorizontalAlignment = bodyAlignment
            If UCase$(displayMode) = NX_ROLE_STYLE_MODE_MONO Then body.Font.Color = RGB(32, 32, 32)
        End If
        NxDrawApplyTableBorders target, header, tableStyle, displayMode
        If includeTotalRow And Not body Is Nothing Then
            Set total = body.Rows(body.Rows.Count)
            NxApplyTotalToken total
            total.Font.Size = bodyFontSize
            total.Font.Bold = True
            If bodyAlignment <> xlGeneral Then total.HorizontalAlignment = bodyAlignment
            If UCase$(displayMode) = NX_ROLE_STYLE_MODE_MONO Then
                total.Font.Color = RGB(32, 32, 32)
                total.Interior.Color = RGB(242, 242, 242)
            End If
            NxDrawSetBorder total.Borders(xlEdgeTop), NxRoleStyleTokenColor("EMPHASIS", displayMode), xlMedium
        End If
        ' A one-row body may also be the total row: the header boundary wins.
        If Not header Is Nothing Then NxDrawSetBorder header.Borders(xlEdgeBottom), NxRoleStyleTokenColor("LINE", displayMode), xlThick
    Else
        NxRaiseContractError "표 옵션은 제목표 또는 실무표에만 적용할 수 있습니다."
    End If
    If autoFitColumns Then target.Columns.AutoFit
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    snapshot.Restore
    On Error GoTo 0
    If errorNumber = 0 Then errorNumber = NX_CONTRACT_ERROR
    Err.Raise errorNumber, "NxDrawController.NxDrawApplyTableOptions", errorDescription
End Sub

Public Function NxDrawTableHeaderFill(ByVal displayMode As String) As Long
    If UCase$(displayMode) = NX_ROLE_STYLE_MODE_MONO Then
        NxDrawTableHeaderFill = RGB(230, 230, 230)
    Else
        NxDrawTableHeaderFill = RGB(216, 226, 240)
    End If
End Function

' For a caller-owned NEW output only. Its caller closes that output on failure;
' do not snapshot or repaint source/copy workbooks just to format a generated report.
Public Sub NxDrawStyleGeneratedTable(ByVal outputBook As Workbook, ByVal target As Range, Optional ByVal headerRows As Long = 1)
    Dim header As Range
    If outputBook Is Nothing Or target Is Nothing Then NxRaiseContractError "생성된 결과 표가 필요합니다."
    If Not target.Worksheet.Parent Is outputBook Then NxRaiseContractError "결과 표의 통합문서가 다릅니다."
    If Len(outputBook.Path) > 0 Or outputBook.IsAddin Or target.Areas.Count <> 1 Then NxRaiseContractError "새 결과 통합문서에만 결과 표 스타일을 적용합니다."
    If headerRows < 0 Or headerRows > target.Rows.Count Then NxRaiseContractError "결과 표 제목행을 확인하세요."
    NxApplyBodyToken target, True
    target.Font.Size = 9#: target.Font.Color = RGB(32, 32, 32)
    If headerRows > 0 Then
        Set header = target.Rows(1).Resize(headerRows, target.Columns.Count)
        NxApplyHeaderToken header
        header.Font.Size = 9#: header.Font.Color = RGB(32, 32, 32)
        header.Interior.Color = NxDrawTableHeaderFill(NX_ROLE_STYLE_MODE_MONO)
    End If
    NxDrawApplyTableBorders target, header, "열린 표", NX_ROLE_STYLE_MODE_MONO
    If Not header Is Nothing Then NxDrawSetBorder header.Borders(xlEdgeBottom), NxRoleStyleTokenColor("LINE", NX_ROLE_STYLE_MODE_MONO), xlThick
End Sub

Private Sub NxDrawApplyTitleOptions(ByVal target As Range, ByVal displayMode As String, ByVal tableStyle As String, _
    ByVal fontSize As Double, ByVal alignment As Long)

    NxApplyHeaderToken target
    target.Font.Size = fontSize
    target.HorizontalAlignment = alignment
    target.Font.Bold = (tableStyle = "제목")
    If UCase$(displayMode) = NX_ROLE_STYLE_MODE_MONO Then
        target.Font.Color = RGB(32, 32, 32)
        target.Interior.Pattern = xlNone
    End If
    NxDrawClearInner target
    target.Borders(xlEdgeBottom).LineStyle = xlContinuous
    target.Borders(xlEdgeBottom).Color = NxRoleStyleTokenColor("EMPHASIS", displayMode)
    target.Borders(xlEdgeBottom).Weight = IIf(tableStyle = "제목", xlMedium, xlThin)
End Sub

Private Sub NxDrawApplyTableBorders(ByVal target As Range, ByVal header As Range, ByVal tableStyle As String, ByVal displayMode As String)
    Dim borderColor As Long
    borderColor = NxRoleStyleTokenColor("LINE", displayMode)
    If tableStyle <> "닫힌 표" And tableStyle <> "열린 표" Then _
        NxRaiseContractError "표 테두리 스타일이 올바르지 않습니다."
    NxDrawSetBorder target.Borders(xlEdgeLeft), borderColor, xlThin
    NxDrawSetBorder target.Borders(xlEdgeTop), borderColor, xlThin
    NxDrawSetBorder target.Borders(xlEdgeRight), borderColor, xlThin
    NxDrawSetBorder target.Borders(xlEdgeBottom), borderColor, xlThick
    If target.Rows.Count > 1 Then NxDrawSetBorder target.Borders(xlInsideHorizontal), borderColor, xlThin
    If target.Columns.Count > 1 Then NxDrawSetBorder target.Borders(xlInsideVertical), borderColor, xlThin
    If tableStyle = "열린 표" Then
        target.Borders(xlEdgeLeft).LineStyle = xlNone
        target.Borders(xlEdgeRight).LineStyle = xlNone
    End If
End Sub

Private Sub NxDrawSetBorder(ByVal border As Border, ByVal colorValue As Long, ByVal weightValue As XlBorderWeight)
    border.LineStyle = xlContinuous
    border.Color = colorValue
    border.Weight = weightValue
End Sub

Public Function NxDrawRun(ByVal featureId As String, ByVal target As Range, Optional ByVal totalRow As Boolean = False, _
    Optional ByVal preserveAlignment As Boolean = True, Optional ByVal includeHeader As Boolean = True, _
    Optional ByVal autoFitColumns As Boolean = False, Optional ByVal displayMode As String = "MONO", _
    Optional ByVal tableStyle As String = "열린 표", Optional ByVal headerRowCount As Long = 1, _
    Optional ByVal headerFontSize As Double = 9#, Optional ByVal headerAlignment As Long = xlCenter, _
    Optional ByVal bodyFontSize As Double = 9#, Optional ByVal bodyAlignment As Long = xlGeneral) As CNxResult

    Dim request As New CNxDrawingRequest, commandObject As New CNxDrawingFeatureCommand, command As INxFeatureCommand
    If Not includeHeader Then headerRowCount = 0
    request.ConfigureRange featureId, target, totalRow, preserveAlignment, includeHeader, autoFitColumns, _
        displayMode, tableStyle, headerRowCount, headerFontSize, headerAlignment, bodyFontSize, bodyAlignment
    Set command = commandObject
    commandObject.Configure request
    Set NxDrawRun = NxRunDrawingFeatureCommand(command, NxDrawingTargetSummary(request.Target))
End Function

Public Function NxDrawFitPictureRun(ByVal pictureShape As Shape, ByVal target As Range, Optional ByVal pictureMargin As Double = 4#, Optional ByVal moveAndSize As Boolean = True) As CNxResult
    Dim request As New CNxDrawingRequest, commandObject As New CNxDrawingFeatureCommand, command As INxFeatureCommand
    request.ConfigurePicture pictureShape, target, pictureMargin, moveAndSize: Set command = commandObject: commandObject.Configure request
    Set NxDrawFitPictureRun = NxRunDrawingFeatureCommand(command, NxDrawingTargetSummary(request.Target))
End Function

Public Sub NxDrawValidateSelection(ByVal target As Range)
    EnsureDrawingTarget target
End Sub

Public Sub NxDrawValidatePictureSelection(ByVal pictureShape As Shape, ByVal target As Range, Optional ByVal pictureMargin As Double = 4#)
    If pictureShape Is Nothing Then NxRaiseContractError "그림 하나를 선택해야 합니다."
    If pictureMargin < 0 Then NxRaiseContractError "그림 여백은 0 이상이어야 합니다."
    EnsureDrawingTarget target
    If pictureShape.Width <= 0 Or pictureShape.Height <= 0 Then NxRaiseContractError "그림 크기가 0입니다."
End Sub

Public Function NxDrawingTargetSummary(ByVal target As Range) As String
    If target Is Nothing Then NxRaiseContractError "표시할 작업 대상 범위가 필요합니다."
    NxDrawingTargetSummary = target.Worksheet.Parent.Name & " / " & target.Worksheet.Name & _
        "!" & target.Address(True, True, xlA1, False)
End Function

Public Function NxRunDrawingFeatureCommand(ByVal command As INxFeatureCommand, _
    Optional ByVal targetSummary As String = vbNullString) As CNxResult
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Set ticket = router.Prepare(NxDrawingFeatureDefinition(command.FeatureId), command)
    Select Case ticket.Decision.ResolvedGrade
        Case NxExecutionFast
            Set NxRunDrawingFeatureCommand = router.ExecuteFast(ticket)
        Case NxExecutionGuarded
            Set NxRunDrawingFeatureCommand = router.ExecuteUserAction(ticket)
        Case NxExecutionPlanned
            Set NxRunDrawingFeatureCommand = NxShowDrawingExecutionPlan(ticket, targetSummary)
        Case Else
            NxRaiseContractError "Drawing execution grade is invalid"
    End Select
End Function

Private Function NxShowDrawingExecutionPlan(ByVal ticket As CNxExecutionTicket, ByVal targetSummary As String) As CNxResult
    Dim router As New CNxExecutionRouter
    Set NxShowDrawingExecutionPlan = router.ExecuteUserAction(ticket)
End Function

Private Sub EnsureDrawingTarget(ByVal target As Range)
    Dim table As ListObject
    If target Is Nothing Or target.Areas.Count <> 1 Then NxRaiseContractError "표/그리기 대상은 연속된 한 범위여야 합니다."
    If target.Worksheet.ProtectContents Or target.Worksheet.ProtectDrawingObjects Then NxRaiseContractError "보호된 시트에는 적용할 수 없습니다."
    If target.CountLarge > 100000 Then NxRaiseContractError "표/그리기 셀 한도를 초과했습니다."
    For Each table In target.Worksheet.ListObjects
        If Not Intersect(table.Range, target) Is Nothing Then NxRaiseContractError "Excel 표 객체와 겹치는 범위는 명시적 충돌 확인 후 적용해야 합니다."
    Next table
    If target.FormatConditions.Count > 0 Then NxRaiseContractError "조건부서식과 충돌할 수 있어 명시적 확인 후 적용해야 합니다."
End Sub
