Attribute VB_Name = "NxTemplateUi"
Option Explicit

Public Sub NxRibbonOpenTemplate(ByVal control As Object)
    NxRibbonExecuteTag "nx1|entry|NX-ENTRY-TEMPLATE"
End Sub

Public Sub NxTemplateOpenManager(Optional ByVal featureId As String = vbNullString)
    If Len(featureId) = 0 Then NxRaiseContractError "템플릿 목록에서 실행할 기능을 선택하세요."
    NxRouteFeature featureId
End Sub

Public Sub NxTemplateOpenFeatureDialog(ByVal featureId As String)
    Dim managerForm As New FNxTemplateManager
    Dim actionForm As New FNxTemplateAction, result As CNxResult, sourceSheet As Worksheet
    Select Case featureId
        Case "NX-TPL-LIST", "NX-TPL-LOAD", "NX-TPL-RENAME", "NX-TPL-DELETE"
            managerForm.BindContext ActiveWorkbook, featureId
            managerForm.Show vbModal
        Case "NX-TPL-REGISTER-SHEET"
            If TypeName(ActiveSheet) <> "Worksheet" Then NxRaiseContractError "등록할 현재시트를 선택하세요."
            Set sourceSheet = ActiveSheet
            Set result = NxTemplateShowRegistration(ActiveWorkbook, sourceSheet, "sheet_used_range")
        Case Else
            NxRaiseContractError "지원하지 않는 템플릿 기능입니다."
    End Select
End Sub

Public Function NxTemplateShowRegistration(ByVal targetBook As Workbook, ByVal sourceSheet As Worksheet, _
    ByVal sourceKind As String, Optional ByVal filePath As String = vbNullString) As CNxResult
    Dim registerForm As New FNxTemplateRegister, failureNumber As Long, failureText As String
    On Error GoTo Failed
    registerForm.BindSource targetBook, sourceSheet, sourceKind, filePath
    registerForm.Show vbModal
    Set NxTemplateShowRegistration = registerForm.OperationResult
    Unload registerForm
    Exit Function
Failed:
    failureNumber = Err.Number
    failureText = Err.Description
    On Error Resume Next
    Unload registerForm
    On Error GoTo 0
    Err.Raise failureNumber, "NxTemplateShowRegistration", failureText
End Function

Public Function NxTemplateRunRequestFromManager(ByVal request As CNxTemplateRequest, ByVal targetBook As Workbook) As CNxResult
    Dim commandObject As New CNxTemplateFeatureCommand, command As INxFeatureCommand
    NxTemplateValidateUserBook targetBook
    targetBook.Activate
    commandObject.ConfigureRegistration request
    Set command = commandObject
    Set NxTemplateRunRequestFromManager = RunTemplateCommand(command, _
        "템플릿 '" & request.DisplayName & "'을 등록합니다. " & CStr(request.SourceRanges.Count) & "개 시트 / " & _
        CStr(request.ActualCellCount) & "개 사용 셀. 원본 통합문서는 저장하거나 변경하지 않습니다.")
End Function

Public Function NxTemplateRunManagerAction(ByVal action As String, ByVal targetBook As Workbook, _
    ByVal templateId As String, ByVal riskConfirmed As Boolean, _
    Optional ByVal displayName As String = vbNullString, Optional ByVal description As String = vbNullString, _
    Optional ByVal category As Variant, Optional ByVal quarantineId As String = vbNullString, Optional ByVal moveDirection As Long = 0) As CNxResult

    Dim commandObject As New CNxTemplateFeatureCommand, command As INxFeatureCommand
    Dim record As CNxTemplateRecord, records As Collection, summary As String, cells As Long
    NxTemplateValidateUserBook targetBook
    targetBook.Activate
    Select Case action
        Case "use", "copy"
            Set record = TemplateRecordById(templateId)
            cells = CLng(CDbl(record.RowCount) * CDbl(record.ColumnCount))
            If action = "copy" Then
                commandObject.ConfigureCopy templateId, targetBook, riskConfirmed, cells, Len(record.RiskFlags) > 0
                summary = "템플릿을 새 미저장 통합문서로 엽니다. 저장소 원본은 편집하지 않습니다."
            Else
                commandObject.ConfigureLoad templateId, targetBook, riskConfirmed, cells, Len(record.RiskFlags) > 0
                summary = "템플릿의 모든 시트를 '" & targetBook.Name & "'에 새 시트로 추가합니다. 기존 시트는 변경하지 않습니다."
            End If
        Case "catalog"
            Set records = NxTemplateList()
            commandObject.ConfigureCatalog targetBook, records.Count
            summary = "템플릿 관리표를 '" & targetBook.Name & "'의 새 시트에 생성합니다. 기존 관리표는 보존합니다."
        Case "rename"
            commandObject.ConfigureRename templateId, displayName, description, category
            summary = TemplateActionSummary("NX-TPL-RENAME")
        Case "delete"
            commandObject.ConfigureDelete templateId
            summary = TemplateActionSummary("NX-TPL-DELETE")
        Case "restore"
            commandObject.ConfigureRestore templateId, quarantineId
            summary = "삭제한 템플릿을 무결성 검사 후 보관함으로 복원합니다."
        Case "move"
            commandObject.ConfigureMove templateId, moveDirection
            summary = "템플릿 보관함의 표시 순서를 변경합니다."
        Case Else
            NxRaiseContractError "지원하지 않는 템플릿 관리 작업입니다."
    End Select
    Set command = commandObject
    Set NxTemplateRunManagerAction = RunTemplateCommand(command, summary)
End Function

Public Function NxTemplateRecordsForForm() As Collection
    Set NxTemplateRecordsForForm = NxTemplateListMetadata(NxTemplateStoreRoot(), False)
End Function

Public Function NxTemplatePreviewText(ByVal source As Range, ByVal retained As Collection, ByVal includeHidden As Boolean, ByVal exclusionAccepted As Boolean) As String
    Dim analysis As CNxTemplateAnalysis
    Set analysis = NxTemplateAnalyze(source, retained, includeHidden, exclusionAccepted)
    NxTemplatePreviewText = NxTemplateAnalysisPreviewText(analysis)
End Function

Public Function NxTemplateAnalysisPreviewText(ByVal analysis As CNxTemplateAnalysis) As String
    Dim riskText As String
    If analysis Is Nothing Then NxRaiseContractError "Template analysis is required"
    If analysis.HasBlockedResource Then
        NxTemplateAnalysisPreviewText = "등록 차단: " & analysis.BlockReason
        Exit Function
    End If
    riskText = JoinText(analysis.RiskFlags, ", ")
    If Len(riskText) = 0 Then riskText = "없음"
    NxTemplateAnalysisPreviewText = "범위 " & analysis.SourceAddress & vbCrLf & _
        CStr(analysis.RowCount) & "행 × " & CStr(analysis.ColumnCount) & "열" & vbCrLf & _
        "보존 상수 " & CStr(analysis.StoredConstantCount) & "개 · 제외 상수 " & CStr(analysis.ExcludedConstantCount) & "개" & vbCrLf & _
        "수식 " & CStr(analysis.FormulaCount) & "개 · 위험 표시 " & riskText
    If analysis.HasExcludedDisplayResource Then NxTemplateAnalysisPreviewText = NxTemplateAnalysisPreviewText & vbCrLf & "제외 리소스: " & analysis.ExcludedResourceSummary
End Function

Public Function NxTemplateRunRegistrationFromForm(ByVal featureId As String, ByVal source As Range, ByVal retained As Collection, _
    ByVal displayName As String, ByVal description As String, ByVal includeHidden As Boolean, _
    ByVal exclusionAccepted As Boolean, ByVal riskConfirmed As Boolean) As CNxResult

    Dim analysis As CNxTemplateAnalysis, request As New CNxTemplateRequest
    Dim commandObject As New CNxTemplateFeatureCommand, command As INxFeatureCommand
    Set analysis = NxTemplateAnalyze(source, retained, includeHidden, exclusionAccepted)
    If analysis.HasBlockedResource Then NxRaiseContractError analysis.BlockReason
    request.ConfigureRegistration featureId, SourceKindFor(featureId), source, retained, displayName, description, includeHidden, exclusionAccepted, riskConfirmed
    request.BindAnalysis analysis
    commandObject.ConfigureRegistration request
    Set command = commandObject
    Set NxTemplateRunRegistrationFromForm = RunTemplateCommand(command, "선택 범위를 매크로 없는 개인 템플릿으로 등록합니다.")
End Function

Public Function NxTemplateRunFromForm(ByVal featureId As String, ByVal templateId As String, ByVal riskConfirmed As Boolean, _
    Optional ByVal displayName As String = vbNullString, Optional ByVal description As String = vbNullString) As CNxResult

    Dim commandObject As New CNxTemplateFeatureCommand, command As INxFeatureCommand
    Dim record As CNxTemplateRecord, records As Collection
    Select Case featureId
        Case "NX-TPL-LOAD"
            Set record = TemplateRecordById(templateId)
            commandObject.ConfigureLoad templateId, ActiveWorkbook, riskConfirmed, CLng(CDbl(record.RowCount) * CDbl(record.ColumnCount)), (Len(record.RiskFlags) > 0)
        Case "NX-TPL-RENAME"
            commandObject.ConfigureRename templateId, displayName, description
        Case "NX-TPL-DELETE"
            commandObject.ConfigureDelete templateId
        Case Else
            NxRaiseContractError "Form template action is outside the closed contract"
    End Select
    Set command = commandObject
    Set NxTemplateRunFromForm = RunTemplateCommand(command, TemplateActionSummary(featureId))
End Function

Public Function NxTemplateFormContractHas(ByVal controlName As String) As Boolean
    Select Case controlName
        Case "cboSourceKind", "txtTargetSummary", "cmdSelectRetained", "txtRetainedSummary", "txtName", "txtDescription", _
            "cboHiddenPolicy", "txtPreview", "cboRiskConsent", "cmdPreview", "cmdRegister", "cboTemplates", _
            "txtTemplateDetails", "cmdLoadNewSheet", "cmdRename", "cmdDelete", "cmdCancel"
            NxTemplateFormContractHas = True
    End Select
End Function

Public Function NxTemplateRibbonCallbackTarget() As String
    NxTemplateRibbonCallbackTarget = "NxTemplateOpenManager"
End Function

Public Function NxTemplateUsesGeneratedFormSources() As Boolean
    NxTemplateUsesGeneratedFormSources = True
End Function

Private Function RunTemplateCommand(ByVal command As INxFeatureCommand, ByVal actionSummary As String) As CNxResult
    Dim definition As CNxFeatureDefinition, router As New CNxExecutionRouter, ticket As CNxExecutionTicket
    Set definition = NxTemplateFeatureDefinition(command.FeatureId)
    Set ticket = router.Prepare(definition, command)
    Select Case ticket.Decision.ResolvedGrade
        Case NxExecutionFast
            Set RunTemplateCommand = router.ExecuteFast(ticket)
        Case NxExecutionGuarded
            Set RunTemplateCommand = router.ExecuteUserAction(ticket)
        Case NxExecutionPlanned
            Set RunTemplateCommand = RunPlannedTemplate(ticket, command, definition, router, actionSummary)
        Case Else
            NxRaiseContractError "Template execution grade is invalid"
    End Select
End Function

Private Function RunPlannedTemplate(ByVal ticket As CNxExecutionTicket, ByVal command As INxFeatureCommand, _
    ByVal definition As CNxFeatureDefinition, ByVal router As CNxExecutionRouter, ByVal actionSummary As String) As CNxResult
    Set RunPlannedTemplate = router.ExecuteUserAction(ticket)
End Function

Private Function TemplateRecordById(ByVal templateId As String) As CNxTemplateRecord
    Dim record As CNxTemplateRecord
    Set record = NxTemplateReadRecordFolder(NxTemplateTemplateFolderPath(templateId), False, False)
    If record.Status <> "healthy" Then NxRaiseContractError "손상된 템플릿은 불러올 수 없습니다."
    Set TemplateRecordById = record
End Function

Private Function SourceKindFor(ByVal featureId As String) As String
    If featureId = "NX-TPL-REGISTER-SHEET" Then
        SourceKindFor = "sheet_used_range"
    Else
        NxRaiseContractError "Template registration feature is invalid"
    End If
End Function

Private Function TemplateActionSummary(ByVal featureId As String) As String
    Select Case featureId
        Case "NX-TPL-LOAD": TemplateActionSummary = "검증된 템플릿을 새 시트 A1부터 불러옵니다."
        Case "NX-TPL-RENAME": TemplateActionSummary = "템플릿 표시 이름과 설명을 변경합니다."
        Case "NX-TPL-DELETE": TemplateActionSummary = "템플릿을 복구 가능한 격리 보관소로 이동합니다."
    End Select
End Function

Private Function JoinText(ByVal values As Collection, ByVal delimiter As String) As String
    Dim item As Variant
    For Each item In values
        If Len(JoinText) > 0 Then JoinText = JoinText & delimiter
        JoinText = JoinText & CStr(item)
    Next item
End Function

Private Function CreateTemplateFrameTextCatalog() As CNxFrameTextCatalog
    Dim catalog As New CNxFrameTextCatalog
    catalog.Configure "성공", "취소", "입력 오류", "환경 오류", "부분 실패", _
        "없음", "완료하지 못한 작업을 확인하세요.", "통합문서가 변경됨", _
        "통합문서가 변경되지 않음", "복구 필요 없음", "결과를 확인하세요."
    Set CreateTemplateFrameTextCatalog = catalog
End Function
