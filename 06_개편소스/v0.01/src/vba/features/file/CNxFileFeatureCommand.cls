VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFileFeatureCommand"
Option Explicit
Implements INxFeatureCommand

Private mFeatureId As String
Private mOutputPath As String
Private mInputPaths As Variant
Private mAllowPartial As Boolean
Private mSecondApproval As Boolean
Private mSelectedTarget As Object
Private mWorkflowOptions As Object
Private mConfigured As Boolean
Private mPlanBuilt As Boolean
Private mConsolidationMode As String
Private mSkipHeaders As Boolean
Private mCreateContents As Boolean
Private mRenameUndo As Boolean
Private mUndoJournal As Object
Private mUndoBook As Workbook

Public Sub Configure(ByVal featureId As String, ByVal outputPath As String, Optional ByVal inputPaths As Variant, Optional ByVal allowPartial As Boolean = False, Optional ByVal selectedTarget As Object, Optional ByVal secondApproval As Boolean = False, Optional ByVal workflowOptions As Object)
    If mConfigured Then NxRaiseContractError "File command configuration can only occur once"
    Select Case featureId
        Case "NX-FILE-CONSOLIDATE", "NX-FILE-RANGE-PNG", "NX-FILE-CHART-PNG", _
             "NX-FILE-SHEET-COPY-SAVE", "NX-FILE-RANGE-COPY-SAVE", _
             "NX-FILE-WORKBOOK-COMPARE", "NX-FILE-SHEET-COMPARE", "NX-FILE-FILE-COMPARE", "NX-FILE-BATCH-RENAME", "NX-FILE-SHEET-BATCH-RENAME", _
             "NX-FILE-PDF-CURRENT-SHEET", "NX-FILE-PDF-EACH-SHEET", _
             "NX-FILE-PDF-SELECTED-COMBINED", "NX-FILE-PDF-ALL-COMBINED", "NX-FILE-PDF-SETTINGS"
        Case Else: NxRaiseContractError "File command feature is outside the closed contract"
    End Select
    If Len(Trim$(outputPath)) = 0 And featureId <> NX_FEATURE_FILE_WORKBOOK_COMPARE And featureId <> "NX-FILE-SHEET-COMPARE" And featureId <> "NX-FILE-FILE-COMPARE" Then NxRaiseContractError "File command output path is required"
    mFeatureId = featureId
    mOutputPath = outputPath
    mAllowPartial = allowPartial
    mSecondApproval = secondApproval
    If IsArray(inputPaths) Then mInputPaths = inputPaths
    If Not selectedTarget Is Nothing Then Set mSelectedTarget = selectedTarget
    If Not workflowOptions Is Nothing Then Set mWorkflowOptions = workflowOptions
    If featureId = NX_FEATURE_FILE_BATCH_RENAME Or featureId = NX_FEATURE_FILE_SHEET_BATCH_RENAME Then
        If Not mWorkflowOptions Is Nothing Then
            If mWorkflowOptions.Exists("rename_undo") Then mRenameUndo = WorkflowBoolean("rename_undo")
        End If
        If mRenameUndo Then
            If featureId = NX_FEATURE_FILE_SHEET_BATCH_RENAME Then
                If selectedTarget Is Nothing Then NxRaiseContractError "복구 대상 통합문서를 확인하세요."
                If TypeName(selectedTarget) <> "Workbook" Then NxRaiseContractError "복구 대상 통합문서가 필요합니다."
                Set mUndoBook = selectedTarget
            End If
            Set mUndoJournal = NxBatchRenameUndoJournal(featureId = NX_FEATURE_FILE_BATCH_RENAME, mUndoBook)
            RequireCurrentUndo
        End If
    End If
    mConsolidationMode = "SHEETS"
    If featureId = NX_FEATURE_FILE_CONSOLIDATE Then
        If Not workflowOptions Is Nothing Then
            mConsolidationMode = WorkflowText("consolidation_mode")
            mSkipHeaders = WorkflowBoolean("skip_headers")
            If workflowOptions.Exists("create_contents") Then mCreateContents = WorkflowBoolean("create_contents")
            If mConsolidationMode <> "SHEETS" And mConsolidationMode <> "ROWS" And mConsolidationMode <> "BOTH" Then NxRaiseContractError "파일 통합 방식이 올바르지 않습니다."
        End If
    End If
    mConfigured = True
End Sub

Private Property Get INxFeatureCommand_FeatureId() As String
    EnsureConfigured
    INxFeatureCommand_FeatureId = mFeatureId
End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan
    EnsureConfigured
    If mPlanBuilt Then NxRaiseContractError "File command plan can be built only once"
    If context Is Nothing Then NxRaiseContractError "File command requires context"
    If mRenameUndo Then RequireCurrentUndo
    Set plan = NxCreatePlan(mFeatureId, context.Fingerprint, context.SourceStateDigest)
    plan.DeclareObservedCells context.ActualCellCount
    ' The command owns the binary/package write; the dispatcher must not pre-create
    ' an empty marker because the feature produces a validated workbook or PNG.
    If (mFeatureId = NX_FEATURE_FILE_WORKBOOK_COMPARE Or mFeatureId = "NX-FILE-SHEET-COMPARE" Or mFeatureId = "NX-FILE-FILE-COMPARE") And Len(mOutputPath) = 0 Then
        plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxWorksheetCreate, context.WorkbookIdentity, _
            "close only the unsaved comparison report", 1000000)
    Else
        plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxBinaryPackageWrite, mOutputPath, NxFileCommandCompensator())
    End If
    plan.Seal
    mPlanBuilt = True
    Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    Dim decisions As Object, preview As Collection
    Dim changedCount As Long, compareResult As CNxResult
    EnsureConfigured
    If Not mPlanBuilt Or context Is Nothing Or approvedPlan Is Nothing Then NxRaiseContractError "File command requires approved runtime facts"
    If approvedPlan.FeatureId <> mFeatureId Then NxRaiseContractError "File approved plan identity mismatch"
    If mRenameUndo Then
        RequireCurrentUndo
        changedCount = NxBatchRenameUndoLast(mFeatureId = NX_FEATURE_FILE_BATCH_RENAME, mUndoBook)
        Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, NxSuccess, "complete", mOutputPath, _
            mFeatureId = NX_FEATURE_FILE_SHEET_BATCH_RENAME, vbNullString, "이름 복구 완료")
        Exit Function
    End If
    Select Case mFeatureId
        Case "NX-FILE-CONSOLIDATE"
            If Not IsArray(mInputPaths) Then NxRaiseContractError "Consolidation input files are required"
            Set INxFeatureCommand_Execute = NxRunConsolidation(mInputPaths, mOutputPath, mAllowPartial, mSecondApproval, mConsolidationMode, mSkipHeaders, mCreateContents)
        Case "NX-FILE-RANGE-PNG"
            If mSelectedTarget Is Nothing Then NxRaiseContractError "저장할 셀 범위를 선택하세요."
            If TypeName(mSelectedTarget) <> "Range" Then NxRaiseContractError "저장할 셀 범위를 확인하세요."
            Set INxFeatureCommand_Execute = NxExportRangePng(mSelectedTarget, mOutputPath)
        Case "NX-FILE-CHART-PNG"
            If mSelectedTarget Is Nothing Then NxRaiseContractError "Chart PNG requires a selected chart object"
            Set INxFeatureCommand_Execute = NxExportChartPng(mSelectedTarget, mOutputPath)
        Case "NX-FILE-SHEET-COPY-SAVE"
            If mSelectedTarget Is Nothing Or TypeName(mSelectedTarget) <> "Worksheet" Then NxRaiseContractError "Sheet copy requires a worksheet target"
            Set INxFeatureCommand_Execute = NxRunSheetCopySave(mSelectedTarget, mOutputPath)
        Case "NX-FILE-RANGE-COPY-SAVE"
            If mSelectedTarget Is Nothing Or TypeName(mSelectedTarget) <> "Range" Then NxRaiseContractError "Range copy requires a range target"
            Set INxFeatureCommand_Execute = NxRunRangeCopySave(mSelectedTarget, mOutputPath)
        Case "NX-FILE-WORKBOOK-COMPARE", "NX-FILE-SHEET-COMPARE", "NX-FILE-FILE-COMPARE"
            If Not mWorkflowOptions Is Nothing Then
                If mWorkflowOptions.Exists("data_compare") Then
                    Set compareResult = NxDataCompareCreate(mFeatureId, mWorkflowOptions.Item("left_range"), mWorkflowOptions.Item("right_range"), WorkflowBoolean("compare_formulas"))
                    Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, compareResult.Outcome, compareResult.Stage, context.WorkbookIdentity, False, compareResult.Recovery, compareResult.MessageKey)
                    Exit Function
                End If
            End If
            Set compareResult = NxWorkbookCompareReport(WorkflowText("base_path"), _
                WorkflowText("compare_path"), mOutputPath, WorkflowBoolean("compare_formats"), OptionalWorkflowText("base_sheet"), OptionalWorkflowText("compare_sheet"))
            If Len(mOutputPath) = 0 Then
                Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, compareResult.Outcome, compareResult.Stage, _
                    context.WorkbookIdentity, False, compareResult.Recovery, compareResult.MessageKey)
            Else
                Set INxFeatureCommand_Execute = compareResult
            End If
        Case "NX-FILE-BATCH-RENAME"
            Set preview = WorkflowCollection("preview")
            changedCount = NxBatchRenameCommit(preview)
            Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, NxSuccess, "complete", mOutputPath, _
                False, vbNullString, "batch_rename_complete")
        Case "NX-FILE-SHEET-BATCH-RENAME"
            Set preview = WorkflowCollection("preview")
            changedCount = NxSheetBatchRenameCommit(ActiveWorkbook, preview)
            Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, NxSuccess, "complete", mOutputPath, _
                True, vbNullString, "sheet_batch_rename_complete")
        Case "NX-FILE-PDF-CURRENT-SHEET", "NX-FILE-PDF-EACH-SHEET", _
             "NX-FILE-PDF-SELECTED-COMBINED", "NX-FILE-PDF-ALL-COMBINED", "NX-FILE-PDF-SETTINGS"
            Set INxFeatureCommand_Execute = NxPdfExecute(mFeatureId, WorkflowText("output_folder"), _
                WorkflowText("base_name"), WorkflowText("quality"))
    End Select
End Function

Private Function NxFileCommandCompensator() As String
    If mRenameUndo Then
        NxFileCommandCompensator = "retain exact unresolved journal entries if restoring original names fails"
        Exit Function
    End If
    Select Case mFeatureId
        Case "NX-FILE-BATCH-RENAME": NxFileCommandCompensator = "rollback every journaled temporary or target name"
        Case "NX-FILE-SHEET-BATCH-RENAME": NxFileCommandCompensator = "restore every journaled worksheet name"
        Case "NX-FILE-PDF-SETTINGS": NxFileCommandCompensator = "restore prior settings file or remove new settings file"
        Case Else: NxFileCommandCompensator = "delete only the command-created output"
    End Select
End Function

Private Sub RequireCurrentUndo()
    Dim currentJournal As Object, targetPath As String
    Set currentJournal = NxBatchRenameUndoJournal(mFeatureId = NX_FEATURE_FILE_BATCH_RENAME, mUndoBook)
    If Not currentJournal Is mUndoJournal Then NxRaiseContractError "직전 이름 변경이 달라졌습니다. 복구를 다시 요청하세요."
    targetPath = NxBatchRenameUndoTarget(mFeatureId = NX_FEATURE_FILE_BATCH_RENAME, mUndoBook)
    If StrComp(targetPath, mOutputPath, vbTextCompare) <> 0 Then NxRaiseContractError "복구 대상 경로가 변경되었습니다."
End Sub

Private Function WorkflowText(ByVal key As String) As String
    EnsureWorkflowKey key
    WorkflowText = CStr(mWorkflowOptions.Item(key))
End Function

Private Function OptionalWorkflowText(ByVal key As String) As String
    If mWorkflowOptions Is Nothing Then Exit Function
    If mWorkflowOptions.Exists(key) Then OptionalWorkflowText = CStr(mWorkflowOptions.Item(key))
End Function

Private Function WorkflowBoolean(ByVal key As String) As Boolean
    EnsureWorkflowKey key
    WorkflowBoolean = CBool(mWorkflowOptions.Item(key))
End Function

Private Function WorkflowCollection(ByVal key As String) As Collection
    EnsureWorkflowKey key
    If TypeName(mWorkflowOptions.Item(key)) <> "Collection" Then NxRaiseContractError "File workflow collection is invalid"
    Set WorkflowCollection = mWorkflowOptions.Item(key)
End Function

Private Sub EnsureWorkflowKey(ByVal key As String)
    If mWorkflowOptions Is Nothing Then NxRaiseContractError "File workflow options are required"
    If Not mWorkflowOptions.Exists(key) Then NxRaiseContractError "File workflow option is missing: " & key
End Sub

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "File command is not configured"
End Sub
