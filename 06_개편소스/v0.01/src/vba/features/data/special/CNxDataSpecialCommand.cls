VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDataSpecialCommand"
Option Explicit
Implements INxFeatureCommand

Private mFeatureId As String
Private mSource As Range
Private mScanWorkbook As Workbook
Private mScanFindings As Collection
Private mScanIncomplete As Collection
Private mScanCheckedCells As Double
Private mScanIdentity As String
Private mScanReport As Worksheet
Private mScanReportBook As Workbook
Private mOutput As Range
Private mSourceValues As Variant
Private mSourceFormulas As Variant
Private mSourceRows As Long
Private mSourceColumns As Long
Private mContextFingerprint As String
Private mReferenceDate As Date
Private mConfigured As Boolean
Private mPlanBuilt As Boolean
Private mOptions As CNxDataSpecialOptions
Private mResultSheet As Worksheet
Private mResultBook As Workbook
Private mDestination As New CNxDataOutput

Public Sub ConfigureOptions(ByVal featureId As String, ByVal source As Range, ByVal output As Range, ByVal options As CNxDataSpecialOptions)
    If options Is Nothing Then NxRaiseContractError "변환 옵션이 필요합니다."
    If Not options.IsConfigured Then NxRaiseContractError "변환 옵션을 확인하세요."
    ConfigureTransform featureId, source, output, options.ReferenceDate
    Set mOptions = options
    If CreatesTransformResult Then mDestination.Configure source, options.OutputMode
End Sub

Public Property Get ResultSheet() As Worksheet
    Set ResultSheet = mResultSheet
End Property

Public Sub DiscardTransformResult()
    Dim alerts As Boolean, events As Boolean, failure As Long, detail As String
    If CreatesTransformResult Then
        mDestination.Rollback
        Set mResultSheet = Nothing: Set mResultBook = Nothing
        Exit Sub
    End If
    alerts = Application.DisplayAlerts: events = Application.EnableEvents
    On Error GoTo Failed
    Application.DisplayAlerts = False: Application.EnableEvents = False
    If Not mResultBook Is Nothing Then
        If mResultBook Is mSource.Worksheet.Parent Then NxRaiseContractError "원본 통합문서는 정리할 수 없습니다."
        mResultBook.Close SaveChanges:=False
    ElseIf Not mResultSheet Is Nothing Then
        If mResultSheet Is mSource.Worksheet Then NxRaiseContractError "원본 시트는 정리할 수 없습니다."
        mResultSheet.Delete
    End If
    Set mResultBook = Nothing: Set mResultSheet = Nothing
    Application.DisplayAlerts = alerts: Application.EnableEvents = events
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description: Err.Clear
    Application.DisplayAlerts = alerts: Application.EnableEvents = events
    Err.Raise failure, "NxDataSpecial.DiscardTransformResult", detail
End Sub

Private Function CreatesTransformResult() As Boolean
    If mOptions Is Nothing Then Exit Function
    CreatesTransformResult = (mOptions.OutputMode = "오른쪽 새 열 삽입" Or mOptions.OutputMode = "새 시트 만들기" Or mOptions.OutputMode = "새 통합문서")
End Function

Private Function InsertsColumns() As Boolean
    If mOptions Is Nothing Then Exit Function
    InsertsColumns = (mOptions.OutputMode = "오른쪽 새 열 삽입")
End Function

Public Sub ConfigureTransform(ByVal featureId As String, ByVal source As Range, ByVal output As Range, ByVal referenceDate As Date)
    EnsureMutable
    Select Case featureId
        Case NX_FEATURE_DATA_AGE, NX_FEATURE_DATA_KOREAN_MONEY, NX_FEATURE_DATA_PRIVACY_MASK
        Case Else: NxRaiseContractError "지원하지 않는 데이터 변환 기능입니다."
    End Select
    If source Is Nothing Or output Is Nothing Then NxRaiseContractError "원본과 결과 범위가 필요합니다."
    If source.Areas.Count <> 1 Or output.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 사용할 수 있습니다."
    If source.Rows.Count <> output.Rows.Count Or source.Columns.Count <> output.Columns.Count Then NxRaiseContractError "원본과 결과 범위 크기가 다릅니다."
    Set mSource = source
    Set mOutput = output
    mSourceRows = source.Rows.Count
    mSourceColumns = source.Columns.Count
    mSourceValues = source.Value2
    mSourceFormulas = source.Formula
    mFeatureId = featureId
    mReferenceDate = referenceDate
    mContextFingerprint = NxContextFactory.CaptureCurrentFingerprint()
    mConfigured = True
End Sub

Public Property Get ScanReport() As Worksheet
    Set ScanReport = mScanReport
End Property

Public Sub DiscardScanReport()
    Dim book As Workbook, priorEvents As Boolean, errorNumber As Long, detail As String
    If mScanReportBook Is Nothing Then Exit Sub
    priorEvents = Application.EnableEvents
    On Error GoTo Failed
    ' Match the created workbook object, never the active workbook or a reused name.
    For Each book In Application.Workbooks
        If book Is mScanReportBook Then
            If book Is mScanWorkbook Then NxRaiseContractError "원본 통합문서는 보고서 정리 대상이 아닙니다."
            If book Is ThisWorkbook Then NxRaiseContractError "추가기능 파일은 보고서 정리 대상이 아닙니다."
            Application.EnableEvents = False
            book.Close SaveChanges:=False
            Exit For
        End If
    Next book
    Set mScanReport = Nothing
    Set mScanReportBook = Nothing
    Application.EnableEvents = priorEvents
    Exit Sub
Failed:
    errorNumber = Err.Number: detail = Err.Description: Err.Clear
    On Error Resume Next
    Application.EnableEvents = priorEvents
    On Error GoTo 0
    Err.Raise errorNumber, "NxPrivacyScan.DiscardScanReport", detail
End Sub

Public Sub ConfigureScan(ByVal source As Range)
    If source Is Nothing Then NxRaiseContractError "점검할 통합문서가 없습니다."
    ConfigureWorkbookScan source.Worksheet.Parent
End Sub

Public Sub ConfigureWorkbookScan(ByVal book As Workbook)
    EnsureMutable
    Set mScanWorkbook = book
    Set mScanFindings = NxPrivacyScanWorkbook(book, mScanCheckedCells, mScanIncomplete)
    mScanIdentity = NxPrivacyScanIdentity(mScanFindings, mScanIncomplete, mScanCheckedCells)
    mFeatureId = NX_FEATURE_DATA_PRIVACY_SCAN
    mContextFingerprint = NxContextFactory.CaptureCurrentFingerprint()
    mConfigured = True
End Sub

Private Property Get INxFeatureCommand_FeatureId() As String
    EnsureConfigured
    INxFeatureCommand_FeatureId = mFeatureId
End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan
    Dim declaredCells As Long
    Dim mutationMask As Long
    EnsureConfigured
    If mPlanBuilt Then NxRaiseContractError "데이터 특수 기능 계획은 한 번만 만들 수 있습니다."
    RequireContext context
    Set plan = NxCreatePlan(mFeatureId, context.Fingerprint, context.SourceStateDigest)
    If mFeatureId = NX_FEATURE_DATA_PRIVACY_SCAN Then
        plan.DeclareObservedCells mScanCheckedCells
        declaredCells = (mScanFindings.Count + mScanIncomplete.Count + 6) * 6
        plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxWorksheetCreate, "new-workbook:privacy-report", _
            "close only the command-created privacy report workbook without saving", declaredCells)
    Else
        plan.DeclareObservedCells CLng(mSource.CountLarge)
        If InsertsColumns Then
            plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxColumnInsert, context.WorkbookIdentity, _
                "delete only columns inserted by this conversion", CLng(mSource.CountLarge))
        ElseIf CreatesTransformResult Then
            plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxWorksheetCreate, context.WorkbookIdentity, _
                "delete only the command-created transform sheet or close its created workbook", CLng(mSource.CountLarge))
        Else
            mutationMask = NxCellValueMask
            If Not mOptions Is Nothing Then mutationMask = mutationMask Or NxCellNumberFormatMask
            plan.AddSideEffect NxCreateFeatureCellEffect(mFeatureId, context.SelectionTargetIdentity, mutationMask, CLng(mOutput.CountLarge))
        End If
    End If
    plan.Seal
    mPlanBuilt = True
    Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    Dim resultValues As Variant
    Dim report As Worksheet
    Dim findings As Collection, incomplete As Collection, checkedCells As Double
    Dim errorNumber As Long, detail As String, sourceChanged As Boolean
    On Error GoTo Failed
    EnsureConfigured
    If Not mPlanBuilt Or context Is Nothing Or approvedPlan Is Nothing Then NxRaiseContractError "승인된 데이터 특수 기능 계획이 필요합니다."
    If approvedPlan.FeatureId <> mFeatureId Then NxRaiseContractError "데이터 특수 기능 계획 ID가 다릅니다."
    RequireContext context
    If mFeatureId = NX_FEATURE_DATA_PRIVACY_SCAN Then
        Set findings = NxPrivacyScanWorkbook(mScanWorkbook, checkedCells, incomplete)
        If NxPrivacyScanIdentity(findings, incomplete, checkedCells) <> mScanIdentity Then NxRaiseContractError "점검 이후 의심 항목이나 검사 범위가 변경되었습니다. 다시 실행하세요."
        Set report = NxPrivacyCreateWorkbookReport(mScanWorkbook, findings, incomplete, checkedCells)
        Set mScanReportBook = report.Parent
        Set mScanReport = report
        ' Worksheet-create results retain the approved source identity. ScanReport owns the output location.
        If incomplete.Count > 0 Then
            Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, NxPartialFailure, "scan_incomplete", context.WorkbookIdentity, False, _
                "검사 미완료 항목이 있습니다. 결과 통합문서를 확인하세요.", "result_partial_failure")
        Else
            Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, NxSuccess, "complete", context.WorkbookIdentity, False, vbNullString, "result_success")
        End If
    Else
        resultValues = BuildTransformValues()
        If CreatesTransformResult Then
            Set mOutput = mDestination.CreateOutput(mFeatureId, approvedPlan)
            Set mResultSheet = mDestination.ResultSheet
            If mOptions.OutputMode = "새 통합문서" Then Set mResultBook = mResultSheet.Parent
            mOutput.NumberFormat = "@"
            mOutput.Value2 = resultValues
            mOutput.NumberFormat = TransformNumberFormat()
        Else
            If mOptions Is Nothing Then
                NxApplyApprovedCellValues approvedPlan, context, 0, 0, mSourceRows, mSourceColumns, resultValues
            Else
                ApplyTransformVector approvedPlan, context, resultValues
            End If
        End If
        sourceChanged = True
        If Not mResultBook Is Nothing Then sourceChanged = False
        Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, NxSuccess, "complete", context.WorkbookIdentity, sourceChanged, vbNullString, "result_success")
    End If
    Exit Function
Failed:
    errorNumber = Err.Number: detail = Err.Description: Err.Clear
    On Error Resume Next
    DiscardScanReport
    If Err.Number <> 0 Then detail = detail & " | 보고서 정리 실패: " & Err.Description
    Err.Clear
    DiscardTransformResult
    If Err.Number <> 0 Then detail = detail & " | 변환 결과 정리 실패: " & Err.Description
    On Error GoTo 0
    If errorNumber = 18 And mFeatureId = NX_FEATURE_DATA_PRIVACY_SCAN Then
        Set INxFeatureCommand_Execute = NxCreateResult(mFeatureId, NxCancelled, "cancelled", context.WorkbookIdentity, False, "개인정보 검사를 중단했습니다.", "result_cancelled")
        Exit Function
    End If
    Err.Raise errorNumber, "NxDataSpecialCommand", detail
End Function

Private Sub ApplyTransformVector(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, ByRef matrix As Variant)
    ' Text format prevents header/formula-like strings from becoming formulas.
    NxApplyApprovedCellNumberFormats approvedPlan, context, 0, 0, mSourceRows, mSourceColumns, "@"
    NxApplyApprovedCellValues approvedPlan, context, 0, 0, mSourceRows, mSourceColumns, matrix
    NxApplyApprovedCellNumberFormats approvedPlan, context, 0, 0, mSourceRows, mSourceColumns, TransformNumberFormat()
End Sub

Private Function TransformNumberFormat() As String
    TransformNumberFormat = "@"
    If mFeatureId = NX_FEATURE_DATA_AGE Then
        If mOptions.AgeFormat = "숫자" Then TransformNumberFormat = "General"
    End If
End Function

Private Function BuildTransformValues() As Variant
    Dim values As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim sourceValue As Variant
    ReDim values(1 To mSourceRows, 1 To mSourceColumns)
    For rowIndex = 1 To mSourceRows
        For columnIndex = 1 To mSourceColumns
            sourceValue = SourceValueAt(rowIndex, columnIndex)
            If Not mOptions Is Nothing Then
                If mOptions.HasHeader And ((mSourceRows = 1 And columnIndex = 1) Or (mSourceRows > 1 And rowIndex = 1)) Then
                    If IsError(sourceValue) Then sourceValue = vbNullString
                    values(rowIndex, columnIndex) = sourceValue
                Else
                    values(rowIndex, columnIndex) = mOptions.TransformValue(mFeatureId, sourceValue)
                End If
            Else
              Select Case mFeatureId
                Case NX_FEATURE_DATA_AGE: values(rowIndex, columnIndex) = NxAgeValue(sourceValue, mReferenceDate)
                Case NX_FEATURE_DATA_KOREAN_MONEY: values(rowIndex, columnIndex) = NxKoreanMoneyText(sourceValue)
                Case NX_FEATURE_DATA_PRIVACY_MASK: values(rowIndex, columnIndex) = NxPrivacyMaskedText(sourceValue)
                Case Else: NxRaiseContractError "지원하지 않는 데이터 변환 기능입니다."
              End Select
            End If
        Next columnIndex
    Next rowIndex
    BuildTransformValues = values
End Function

Private Function SourceValueAt(ByVal rowIndex As Long, ByVal columnIndex As Long) As Variant
    If mSourceRows = 1 And mSourceColumns = 1 Then
        SourceValueAt = mSourceValues
    Else
        SourceValueAt = mSourceValues(rowIndex, columnIndex)
    End If
End Function

Private Sub RequireContext(ByVal context As CNxExecutionContext)
    If context Is Nothing Then NxRaiseContractError "실행 문맥이 없습니다."
    If context.Fingerprint <> mContextFingerprint Then NxRaiseContractError "선택 범위가 변경되었습니다."
    If mFeatureId = NX_FEATURE_DATA_PRIVACY_SCAN Then
        If Not Application.ActiveWorkbook Is mScanWorkbook Then NxRaiseContractError "점검할 통합문서가 변경되었습니다."
    Else
        If Not NxDataSpecialSourceMatches(mSource, mSourceValues, mSourceRows, mSourceColumns) Then NxRaiseContractError "원본 데이터가 실행 전에 변경되었습니다."
        If Not NxDataSpecialSourceMatches(mSource, mSourceFormulas, mSourceRows, mSourceColumns, True) Then NxRaiseContractError "원본 수식이 실행 전에 변경되었습니다."
        If context.ActualCellCount <> CDbl(mOutput.CountLarge) Then NxRaiseContractError "결과 범위 크기가 변경되었습니다."
    End If
End Sub

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "데이터 특수 기능 명령이 구성되지 않았습니다."
End Sub

Private Sub EnsureMutable()
    If mConfigured Then NxRaiseContractError "데이터 특수 기능 명령은 한 번만 구성할 수 있습니다."
End Sub
