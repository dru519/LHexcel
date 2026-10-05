Attribute VB_Name = "NxRouteAvailability"
Option Explicit

Public Const NX_ROUTE_AVAILABLE As String = "available"
Public Const NX_ROUTE_NEEDS_INPUT As String = "needs_input"
Public Const NX_ROUTE_UNAVAILABLE As String = "unavailable"

Private mAvailabilityEvents As CNxRouteAvailabilityEvents

Public Sub NxRouteAvailabilityStartEvents()
    If Not mAvailabilityEvents Is Nothing Then Exit Sub
    Set mAvailabilityEvents = New CNxRouteAvailabilityEvents
    mAvailabilityEvents.Configure Application
End Sub

Public Sub NxRouteAvailabilityStopEvents()
    If mAvailabilityEvents Is Nothing Then Exit Sub
    mAvailabilityEvents.Disconnect
    Set mAvailabilityEvents = Nothing
End Sub

Public Sub NxRouteAvailabilityNotifyChanged()
    NxRibbonInvalidateAvailability
    NxNavigatorRefreshAvailability
End Sub

Public Function NxRouteKey(ByVal routeKind As String, ByVal identifier As String) As String
    If routeKind <> "feature" And routeKind <> "command" Then NxRaiseContractError "Route kind is invalid"
    If Not NxRibbonIdentifierIsSafe(identifier) Then NxRaiseContractError "Route identifier is invalid"
    NxRouteKey = routeKind & ":" & identifier
End Function

Public Function NxRouteKeyFromRibbonTag(ByVal rawTag As String) As String
    Dim parts As Variant
    parts = Split(rawTag, "|", -1, vbBinaryCompare)
    If UBound(parts) <> 2 Or CStr(parts(0)) <> "nx1" Then NxRaiseContractError "Ribbon route tag is invalid"
    NxRouteKeyFromRibbonTag = NxRouteKey(CStr(parts(1)), CStr(parts(2)))
End Function

Public Function NxRouteAvailabilityState(ByVal routeKey As String) As String
    Dim reason As String
    NxRouteAvailabilityState = NxEvaluateRouteAvailability(routeKey, reason)
End Function

Public Function NxRouteAvailabilityReason(ByVal routeKey As String) As String
    Dim reason As String
    Call NxEvaluateRouteAvailability(routeKey, reason)
    NxRouteAvailabilityReason = reason
End Function

Public Function NxRouteAvailabilitySummary(ByVal routeKey As String) As String
    Dim reason As String
    Dim state As String
    state = NxEvaluateRouteAvailability(routeKey, reason)
    reason = Replace$(Replace$(Replace$(reason, "|", "/"), vbCr, " "), vbLf, " ")
    NxRouteAvailabilitySummary = state & "|" & reason
End Function

Public Function NxRouteAvailabilityRibbonValue(ByVal routeKey As String) As String
    If NxRouteAvailabilityState(routeKey) = NX_ROUTE_UNAVAILABLE Then
        NxRouteAvailabilityRibbonValue = "false"
    Else
        NxRouteAvailabilityRibbonValue = "true"
    End If
End Function

Public Function NxRouteAvailabilityMenuTip(ByVal routeKey As String) As String
    Dim reason As String
    reason = NxRouteAvailabilityReason(routeKey)
    If Len(reason) = 0 Then
        NxRouteAvailabilityMenuTip = NxGeneratedNavigationRouteField(routeKey, "description_ko")
    Else
        NxRouteAvailabilityMenuTip = reason
    End If
End Function

Public Sub NxRouteAvailabilityEnsureExecutable(ByVal routeKey As String)
    Dim reason As String
    If NxEvaluateRouteAvailability(routeKey, reason) = NX_ROUTE_UNAVAILABLE Then
        If Len(reason) = 0 Then reason = "현재 Excel 상태에서는 이 기능을 실행할 수 없습니다."
        NxRaiseContractError reason
    End If
End Sub

Private Function NxEvaluateRouteAvailability(ByVal routeKey As String, ByRef reason As String) As String
    Dim inputContext As String
    Dim launchSurface As String
    Dim mutationScope As String
    Dim selectionMode As String
    Dim ribbonMetadata As String, fields As Variant

    On Error GoTo Unavailable
    If Not NxDistributionCanExecute() Then
        reason = NxDistributionMessage()
        NxEvaluateRouteAvailability = NX_ROUTE_UNAVAILABLE
        Exit Function
    End If
    routeKey = NxCanonicalRouteKey(routeKey)
    ribbonMetadata = NxRibbonOnlyCommandMetadata(routeKey)
    reason = vbNullString
    If NxGeneratedNavigationRouteExists(routeKey) Then
        inputContext = NxGeneratedNavigationRouteField(routeKey, "input_context")
        selectionMode = NxGeneratedNavigationRouteField(routeKey, "selection_mode")
        launchSurface = NxGeneratedNavigationRouteField(routeKey, "launch_surface")
        mutationScope = NxGeneratedNavigationRouteField(routeKey, "mutation_scope")
    ElseIf Len(ribbonMetadata) > 0 Then
        fields = Split(ribbonMetadata, "|")
        inputContext = CStr(fields(0)): selectionMode = CStr(fields(1))
        launchSurface = CStr(fields(2)): mutationScope = CStr(fields(3))
    ElseIf Not NxTryUnexposedAiRouteMetadata(routeKey, inputContext, selectionMode, launchSurface, mutationScope) Then
        reason = "현재 버전에서 사용할 수 없는 항목입니다."
        NxEvaluateRouteAvailability = NX_ROUTE_UNAVAILABLE
        Exit Function
    End If
    If Not Application.Ready Then
        reason = "Excel 작업이 끝난 뒤 다시 실행하세요."
        NxEvaluateRouteAvailability = NX_ROUTE_UNAVAILABLE
        Exit Function
    End If

    Select Case inputContext
        Case "active_window"
            If Application.ActiveWindow Is Nothing Then
                reason = "활성 Excel 창이 필요합니다."
                NxEvaluateRouteAvailability = NX_ROUTE_UNAVAILABLE
                Exit Function
            End If
        Case "active_workbook"
            If Application.ActiveWorkbook Is Nothing Then
                reason = "활성 통합문서를 먼저 여세요."
                NxEvaluateRouteAvailability = NX_ROUTE_UNAVAILABLE
                Exit Function
            End If
        Case "active_worksheet"
            If Application.ActiveWorkbook Is Nothing Or Not NxAvailabilityHasWorksheet() Then
                reason = "활성 워크시트가 필요합니다."
                NxEvaluateRouteAvailability = NX_ROUTE_UNAVAILABLE
                Exit Function
            End If
        Case "range_selection", "active_data_region"
            If Not NxAvailabilityHasRangeSelection() Then
                NxEvaluateRouteAvailability = NxMissingInputState(launchSurface, "셀 범위를 선택하세요.", reason)
                Exit Function
            End If
        Case "shape_selection"
            If Not NxAvailabilityHasShapeSelection() Then
                NxEvaluateRouteAvailability = NxMissingInputState(launchSurface, "차트나 그림을 선택하세요.", reason)
                Exit Function
            End If
        Case "range_or_shape"
            If Not (NxAvailabilityHasRangeSelection() Or NxAvailabilityHasShapeSelection()) Then
                NxEvaluateRouteAvailability = NxMissingInputState(launchSurface, "셀 범위나 그림을 선택하세요.", reason)
                Exit Function
            End If
        Case "selection_or_clipboard"
            If Not NxAvailabilityHasRangeSelection() Then
                reason = "셀 범위를 선택하거나 클립보드 내용을 준비하세요."
                NxEvaluateRouteAvailability = NX_ROUTE_NEEDS_INPUT
                Exit Function
            End If
        Case "selection_optional", "excel_ready"
        Case Else
            reason = "지원되지 않는 입력 조건입니다."
            NxEvaluateRouteAvailability = NX_ROUTE_UNAVAILABLE
            Exit Function
    End Select

    If selectionMode = "required" And Not NxAvailabilityHasUsableSelection(inputContext) Then
        NxEvaluateRouteAvailability = NxMissingInputState(launchSurface, "실행할 대상을 선택하세요.", reason)
        Exit Function
    End If

    If mutationScope = "document" Then
        If NxAvailabilitySheetProtected() Then
            reason = "보호된 시트에서는 이 작업을 실행할 수 없습니다."
            NxEvaluateRouteAvailability = NX_ROUTE_UNAVAILABLE
            Exit Function
        End If
        If NxAvailabilityWorkbookReadOnly() Then
            reason = "읽기 전용 통합문서에서는 이 작업을 실행할 수 없습니다."
            NxEvaluateRouteAvailability = NX_ROUTE_UNAVAILABLE
            Exit Function
        End If
    End If

    NxEvaluateRouteAvailability = NX_ROUTE_AVAILABLE
    Exit Function

Unavailable:
    Err.Clear
    reason = "현재 Excel 상태를 확인할 수 없습니다."
    NxEvaluateRouteAvailability = NX_ROUTE_UNAVAILABLE
End Function

Private Function NxTryUnexposedAiRouteMetadata( _
    ByVal routeKey As String, _
    ByRef inputContext As String, _
    ByRef selectionMode As String, _
    ByRef launchSurface As String, _
    ByRef mutationScope As String) As Boolean

    Dim featureId As String
    Const FEATURE_ROUTE_PREFIX As String = "feature:"

    If Left$(routeKey, Len(FEATURE_ROUTE_PREFIX)) <> FEATURE_ROUTE_PREFIX Then Exit Function
    featureId = Mid$(routeKey, Len(FEATURE_ROUTE_PREFIX) + 1)
    Select Case featureId
        Case NX_FEATURE_AI_SUMMARY, NX_FEATURE_AI_CLEAN, NX_FEATURE_AI_FORMULA, _
             NX_FEATURE_AI_WRITE, NX_FEATURE_AI_IMAGE
            inputContext = "selection_optional"
            selectionMode = "optional"
            launchSurface = "workbench"
            mutationScope = "none"
            NxTryUnexposedAiRouteMetadata = True
    End Select
End Function

Private Function NxMissingInputState(ByVal launchSurface As String, ByVal prompt As String, ByRef reason As String) As String
    reason = prompt
    Select Case launchSurface
        Case "dialog", "workbench", "hybrid"
            NxMissingInputState = NX_ROUTE_NEEDS_INPUT
        Case Else
            NxMissingInputState = NX_ROUTE_UNAVAILABLE
    End Select
End Function

Private Function NxAvailabilityHasWorksheet() As Boolean
    On Error GoTo Missing
    NxAvailabilityHasWorksheet = (TypeName(Application.ActiveSheet) = "Worksheet")
Missing:
End Function

Private Function NxAvailabilityHasRangeSelection() As Boolean
    On Error GoTo Missing
    NxAvailabilityHasRangeSelection = (TypeName(Application.Selection) = "Range")
Missing:
End Function

Private Function NxAvailabilityHasShapeSelection() As Boolean
    Dim selectionType As String
    On Error GoTo Missing
    selectionType = TypeName(Application.Selection)
    NxAvailabilityHasShapeSelection = (Len(selectionType) > 0 And selectionType <> "Range" And selectionType <> "Nothing")
Missing:
End Function

Private Function NxAvailabilityHasUsableSelection(ByVal inputContext As String) As Boolean
    Select Case inputContext
        Case "shape_selection": NxAvailabilityHasUsableSelection = NxAvailabilityHasShapeSelection()
        Case "range_or_shape": NxAvailabilityHasUsableSelection = NxAvailabilityHasRangeSelection() Or NxAvailabilityHasShapeSelection()
        Case Else: NxAvailabilityHasUsableSelection = NxAvailabilityHasRangeSelection()
    End Select
End Function

Private Function NxAvailabilitySheetProtected() As Boolean
    On Error GoTo Protected
    If TypeName(Application.ActiveSheet) <> "Worksheet" Then Exit Function
    NxAvailabilitySheetProtected = CBool(Application.ActiveSheet.ProtectContents)
    Exit Function
Protected:
    NxAvailabilitySheetProtected = True
End Function

Private Function NxAvailabilityWorkbookReadOnly() As Boolean
    On Error GoTo ReadOnly
    If Application.ActiveWorkbook Is Nothing Then Exit Function
    NxAvailabilityWorkbookReadOnly = CBool(Application.ActiveWorkbook.ReadOnly)
    Exit Function
ReadOnly:
    NxAvailabilityWorkbookReadOnly = True
End Function
