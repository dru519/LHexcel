VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDataFeatureCommand"
Option Explicit
Implements INxFeatureCommand

Private mRequest As CNxDataRequest
Private mGroups As Collection
Private mFeatureId As String
Private mExpectedOutputCells As Long
Private mActualCellCount As Long
Private mConfigured As Boolean
Private mPlanBuilt As Boolean
Private mAnalysisValues As Variant
Private mAnalysisOutput As Worksheet

Public Property Get AnalysisOutput() As Worksheet
    Set AnalysisOutput = mAnalysisOutput
End Property

Public Sub DiscardAnalysisOutput()
    Dim book As Workbook, sourceBook As Workbook, priorAlerts As Boolean, priorEvents As Boolean
    Dim failureNumber As Long, detail As String
    If mAnalysisOutput Is Nothing Then Exit Sub
    Set sourceBook = mRequest.SourceRange.Worksheet.Parent
    Set book = mAnalysisOutput.Parent
    If mAnalysisOutput Is mRequest.SourceRange.Worksheet Then NxRaiseContractError "원본 시트는 분석 결과 정리 대상이 아닙니다."
    priorAlerts = Application.DisplayAlerts: priorEvents = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False: Application.DisplayAlerts = False
    If mRequest.OutputMode = "새창으로 보기" Then
        If book Is sourceBook Or book Is ThisWorkbook Then NxRaiseContractError "원본 통합문서는 분석 결과 정리 대상이 아닙니다."
        book.Close SaveChanges:=False
    Else
        mAnalysisOutput.Delete
    End If
    Set mAnalysisOutput = Nothing
    Application.DisplayAlerts = priorAlerts: Application.EnableEvents = priorEvents
    Exit Sub
Failed:
    failureNumber = Err.Number: detail = Err.Description: Err.Clear
    On Error Resume Next
    Application.DisplayAlerts = priorAlerts: Application.EnableEvents = priorEvents
    On Error GoTo 0
    Err.Raise failureNumber, "CNxDataFeatureCommand.DiscardAnalysisOutput", "이번 작업에서 생성한 결과를 정리하지 못했습니다: " & detail
End Sub

Public Sub Configure(ByVal request As CNxDataRequest)
    If mConfigured Then NxRaiseContractError "Data command can be configured only once"
    If request Is Nothing Then NxRaiseContractError "Data command requires a request"
    Set mRequest = request
    mFeatureId = request.FeatureId
    mActualCellCount = request.ActualCellCount
    mConfigured = True
End Sub

Private Property Get INxFeatureCommand_FeatureId() As String
    EnsureConfigured
    INxFeatureCommand_FeatureId = mFeatureId
End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan
    EnsureConfigured
    If mPlanBuilt Then NxRaiseContractError "Data command plan can be built only once"
    Set plan = NxCreatePlan(mFeatureId, context.Fingerprint, context.SourceStateDigest)
    Select Case mFeatureId
        Case "NX-DATA-UNIQUE-COUNT", "NX-DATA-DUPLICATE-LIST"
            mRequest.RequireContext context
            plan.DeclareObservedCells mRequest.ActualCellCount
            If mRequest.IsAnalysis Then
                mAnalysisValues = NxDataAnalysisValues(mRequest)
                mExpectedOutputCells = NxDataAnalysisOutputCells(mFeatureId, mAnalysisValues)
                plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxWorksheetCreate, context.WorkbookIdentity, _
                    "remove only the command-created analysis workbook or sheet", mExpectedOutputCells)
            Else
                Set mGroups = NxDataGroups(mRequest)
            End If
            If Not mRequest.IsAnalysis And mFeatureId = "NX-DATA-DUPLICATE-LIST" Then
                mExpectedOutputCells = NxDataDuplicateOutputCellCount(mGroups)
                plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxWorksheetCreate, context.WorkbookIdentity, _
                    "delete only the command-created duplicate result sheet", mExpectedOutputCells)
            End If
        Case Else
            NxRaiseContractError "Data command feature is outside the closed contract"
    End Select
    plan.Seal
    mPlanBuilt = True
    Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    Dim serviceResult As CNxResult
    EnsureConfigured
    If Not mPlanBuilt Then NxRaiseContractError "Data command requires its sealed plan"
    If context Is Nothing Then NxRaiseContractError "Data command requires approved runtime facts"
    If approvedPlan Is Nothing Then NxRaiseContractError "Data command requires approved runtime facts"
    If approvedPlan.FeatureId <> mFeatureId Then NxRaiseContractError "Data approved plan identity mismatch"

    If mRequest.IsAnalysis Then
        mRequest.RequireContext context
        Set serviceResult = NxDataCreateAnalysisResult(mRequest, mAnalysisValues, mAnalysisOutput)
        Set INxFeatureCommand_Execute = NxDataRebindResult(serviceResult, context.WorkbookIdentity)
        Exit Function
    End If

    Select Case mFeatureId
        Case "NX-DATA-UNIQUE-COUNT"
            mRequest.RequireContext context
            Set serviceResult = NxCreateResult(mFeatureId, NxSuccess, "complete", context.WorkbookIdentity, False, vbNullString, "result_success")
        Case "NX-DATA-DUPLICATE-LIST"
            mRequest.RequireContext context
            Set serviceResult = NxDataCreateDuplicateResult(mRequest, mGroups)
        Case Else
            NxRaiseContractError "Data command feature is outside the closed contract"
    End Select
    Set INxFeatureCommand_Execute = NxDataRebindResult(serviceResult, context.WorkbookIdentity)
End Function

Private Function NxDataRebindResult(ByVal source As CNxResult, ByVal target As String) As CNxResult
    Set NxDataRebindResult = NxCreateResult(source.FeatureId, source.Outcome, source.Stage, target, source.SourceChanged, source.Recovery, source.MessageKey)
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Data command is not configured"
End Sub
