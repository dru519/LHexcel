Attribute VB_Name = "NxDataUi"
Option Explicit

Public Sub NxRibbonOpenData(ByVal control As Object)
    NxDataOpenManager
End Sub

Public Sub NxDataOpenManager(Optional ByVal featureId As String = vbNullString)
    If Len(featureId) = 0 Then NxRaiseContractError "데이터 목록에서 실행할 기능을 선택하세요."
    NxRouteFeature featureId
End Sub

Public Sub NxDataOpenFeatureDialog(ByVal featureId As String)
    Dim analyzeForm As New FNxDataAnalyze
    Dim source As Range
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "데이터 범위를 먼저 선택하세요."
    Set source = Application.Selection
    If source.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 선택하세요."
    Select Case featureId
        Case "NX-DATA-UNIQUE-COUNT", "NX-DATA-DUPLICATE-LIST"
            analyzeForm.BindSelection source, featureId
            analyzeForm.Show vbModal
        Case Else
            NxRaiseContractError "지원하지 않는 데이터 기능입니다."
    End Select
End Sub

Public Function NxDataPreviewText(ByVal featureId As String, ByVal source As Range, ByVal hasTitleRow As Boolean, _
    ByVal keyColumn As Long, ByVal trimText As Boolean, ByVal caseSensitive As Boolean, _
    ByVal includeHidden As Boolean) As String

    Dim request As New CNxDataRequest
    Dim groups As Collection
    Dim group As CNxDataGroup
    Dim duplicateGroups As Long
    Dim duplicateRows As Long
    Dim sample As String
    source.Select
    request.Configure featureId, source, hasTitleRow, keyColumn, trimText, caseSensitive, includeHidden
    Set groups = NxDataGroups(request)
    For Each group In groups
        If group.Count > 1 Then duplicateGroups = duplicateGroups + 1: duplicateRows = duplicateRows + group.Count
        If Len(sample) < 400 And groups.Count <= 10 Then sample = sample & vbCrLf & "- " & group.Key & " · " & CStr(group.Count) & "건"
    Next group
    NxDataPreviewText = "선택 셀 " & Format$(request.ActualCellCount, "#,##0") & "개" & vbCrLf & _
        "고유 그룹 " & Format$(groups.Count, "#,##0") & "개" & vbCrLf & _
        "중복 그룹 " & Format$(duplicateGroups, "#,##0") & "개 · 중복 행 " & Format$(duplicateRows, "#,##0") & "개" & sample
End Function

Public Function NxDataRun(ByVal featureId As String, ByVal source As Range, ByVal keyColumn As Long, _
    ByVal trimText As Boolean, ByVal caseSensitive As Boolean, ByVal includeHidden As Boolean, _
    Optional ByVal hasTitleRow As Boolean = True) As CNxResult

    Dim request As New CNxDataRequest
    Dim commandObject As New CNxDataFeatureCommand
    Dim command As INxFeatureCommand
    source.Select
    request.Configure featureId, source, hasTitleRow, keyColumn, trimText, caseSensitive, includeHidden
    commandObject.Configure request
    Set command = commandObject
    Set NxDataRun = NxDataRunCommand(command, NxDataActionSummary(featureId))
End Function

Public Function NxDataRunNormalization(ByVal source As Range, ByVal preset As String, ByVal dateFormat As String, _
    Optional ByVal outputMode As String = "새 시트 만들기") As CNxResult
    Dim request As New CNxDataNormalizeRequest
    Dim commandObject As New CNxDataNormalizeCommand
    Dim command As INxFeatureCommand
    Dim result As CNxResult, failure As Long, detail As String
    On Error GoTo Failed
    Set source = NxDataContentRange(source)
    source.Worksheet.Activate: source.Select
    request.Configure source, preset, dateFormat, outputMode
    commandObject.Configure request
    Set command = commandObject
    If request.OutputMode = "원본 변경" Then
        Set result = NxDataRunCommand(command, _
            "정규화된 값과 표시 형식을 선택 범위에 적용합니다.")
    Else
        Set result = NxDataRunCommand(command, _
            "정규화된 값을 새 시트에 생성합니다.")
    End If
    commandObject.CompleteOutput result
    Set NxDataRunNormalization = result
    Exit Function
Failed:
    failure = Err.Number: detail = Err.Description: Err.Clear
    On Error Resume Next
    commandObject.DiscardOutput
    If Err.Number <> 0 Then detail = detail & " | 정규화 결과 정리 실패: " & Err.Description
    On Error GoTo 0
    Err.Raise failure, "NxDataRunNormalization", detail
End Function


Public Function NxDataFormContractHas(ByVal controlName As String) As Boolean
    Select Case controlName
        Case "cboFeature", "txtRangeSummary", "cboHeader", "cboKeyColumn", _
            "cboTrim", "cboCase", "cboHidden", "txtPreview", "cmdPreview", "cmdExecute", "cmdCancel"
            NxDataFormContractHas = True
    End Select
End Function

Public Function NxDataUsesGeneratedFormSources() As Boolean
    NxDataUsesGeneratedFormSources = True
End Function

Public Function NxDataRunCommand(ByVal command As INxFeatureCommand, ByVal actionSummary As String) As CNxResult
    Dim definition As CNxFeatureDefinition
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim approval As CNxApproval
    Set definition = NxDataFeatureDefinition(command.FeatureId)
    Set ticket = router.Prepare(definition, command)
    Select Case ticket.Decision.ResolvedGrade
        Case NxExecutionFast
            Set NxDataRunCommand = router.ExecuteFast(ticket)
        Case NxExecutionGuarded
            Set NxDataRunCommand = router.ExecuteUserAction(ticket)
        Case NxExecutionPlanned
            Set NxDataRunCommand = NxDataRunPlannedFrame(command, ticket, router, actionSummary)
        Case Else
            NxRaiseContractError "Data execution grade is invalid"
    End Select
End Function

Public Function NxDataRunPlannedFrame(ByVal command As INxFeatureCommand, ByVal ticket As CNxExecutionTicket, _
    ByVal router As CNxExecutionRouter, ByVal actionSummary As String) As CNxResult
    Set NxDataRunPlannedFrame = router.ExecuteUserAction(ticket)
End Function

' Compatibility entry point for existing native tests; production uses the same ticket consumption.
Public Function NxDataRunPlannedFrameForTest(ByVal command As INxFeatureCommand, ByVal ticket As CNxExecutionTicket, _
    ByVal router As CNxExecutionRouter, ByVal actionSummary As String) As CNxResult
    Dim approval As CNxApproval
    Set approval = router.TakePlannedApproval(ticket, True)
    Set NxDataRunPlannedFrameForTest = router.RunApproved(ticket, approval)
End Function

Private Function NxDataActionSummary(ByVal featureId As String) As String
    Select Case featureId
        Case "NX-DATA-UNIQUE-COUNT": NxDataActionSummary = "고유값과 발생 건수를 읽기 전용으로 계산합니다."
        Case "NX-DATA-DUPLICATE-LIST": NxDataActionSummary = "중복 결과를 새 시트에 값으로 생성합니다."
        Case Else: NxRaiseContractError "Data action is outside the closed contract"
    End Select
End Function
