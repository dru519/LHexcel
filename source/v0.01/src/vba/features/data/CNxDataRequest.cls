VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDataRequest"
Option Explicit

Private mFeatureId As String
Private mSourceRange As Range
Private mHasTitleRow As Boolean
Private mKeyColumn As Long
Private mTrimText As Boolean
Private mCaseSensitive As Boolean
Private mIncludeHidden As Boolean
Private mActualCellCount As Long
Private mSelectionFingerprint As String
Private mConfigured As Boolean
Private mAnalysis As Boolean
Private mRowDirection As Boolean
Private mWholeRecord As Boolean
Private mSkipBlank As Boolean
Private mIncludeRatio As Boolean
Private mSortMode As String
Private mCompareColumn As Long
Private mDenseCompare As Boolean
Private mOutputMode As String
Private mValues As Variant
Private mFormulas As Variant
Private mRows As Long
Private mColumns As Long
Private mHiddenAxes As String

Public Sub ConfigureAnalysis(ByVal featureId As String, ByVal source As Range, ByVal hasHeader As Boolean, _
    ByVal rowDirection As Boolean, ByVal wholeRecord As Boolean, ByVal skipBlank As Boolean, _
    ByVal cleanSpaces As Boolean, ByVal includeRatio As Boolean, ByVal sortMode As String, _
    ByVal leftIndex As Long, ByVal rightIndex As Long, ByVal ignoreCase As Boolean, _
    ByVal denseCompare As Boolean, ByVal includeHidden As Boolean, ByVal outputMode As String)
    Dim policy As New CNxLimitsPolicy, axisCount As Long, itemCount As Long
    If mConfigured Then NxRaiseContractError "Data request can be configured only once"
    If featureId <> "NX-DATA-UNIQUE-COUNT" And featureId <> "NX-DATA-DUPLICATE-LIST" Then NxRaiseContractError "지원하지 않는 분석 기능입니다."
    If source Is Nothing Then NxRaiseContractError "분석할 범위가 필요합니다."
    If source.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 선택하세요."
    If source.Worksheet.Parent Is ThisWorkbook Then NxRaiseContractError "일반 통합문서에서 실행하세요."
    If CDbl(source.CountLarge) > CDbl(policy.MaximumCellsFor(featureId)) Then NxRaiseContractError "분석 범위가 안전 한도를 초과합니다."
    If outputMode <> "새창으로 보기" And outputMode <> "새 시트 만들기" Then NxRaiseContractError "분석 출력은 새창 또는 새 시트만 지원합니다. 행과 열 삽입은 지원하지 않습니다."
    If sortMode <> "값 오름차순" And sortMode <> "값 내림차순" And sortMode <> "원본 순서" Then NxRaiseContractError "집계 정렬 방식을 확인하세요."
    mRows = source.Rows.Count: mColumns = source.Columns.Count
    If rowDirection Then axisCount = mRows: itemCount = mColumns Else axisCount = mColumns: itemCount = mRows
    If hasHeader And itemCount < 2 Then NxRaiseContractError "제목 다음에 분석할 값이 없습니다."
    If featureId = "NX-DATA-DUPLICATE-LIST" Then
        If leftIndex < 1 Or leftIndex > axisCount Or rightIndex < 1 Or rightIndex > axisCount Then NxRaiseContractError "기준과 대조 열 또는 행을 확인하세요."
        If leftIndex = rightIndex Then NxRaiseContractError "기준과 대조 열 또는 행은 서로 달라야 합니다."
    End If
    RequireCurrentSelection source
    Set mSourceRange = source
    mFeatureId = featureId: mHasTitleRow = hasHeader: mKeyColumn = leftIndex
    mRowDirection = rowDirection: mWholeRecord = wholeRecord: mSkipBlank = skipBlank
    mTrimText = cleanSpaces: mIncludeRatio = includeRatio: mSortMode = sortMode
    mCompareColumn = rightIndex: mCaseSensitive = Not ignoreCase: mDenseCompare = denseCompare
    mIncludeHidden = includeHidden: mOutputMode = outputMode
    mActualCellCount = CLng(CDbl(source.CountLarge))
    mValues = NxDataAnalysisRangeMatrix(source, False)
    mFormulas = NxDataAnalysisRangeMatrix(source, True)
    mHiddenAxes = NxDataAnalysisHiddenAxes(source, rowDirection)
    mSelectionFingerprint = NxContextFactory.CaptureCurrentFingerprint()
    mAnalysis = True: mConfigured = True
End Sub

Public Sub RequireFreshAnalysis()
    EnsureConfigured
    If Not mAnalysis Then NxRaiseContractError "전용 분석 요청이 필요합니다."
    RequireCurrentSelection mSourceRange
    If NxContextFactory.CaptureCurrentFingerprint() <> mSelectionFingerprint Then NxRaiseContractError "분석 대상이 변경되었습니다. 미리보기를 다시 실행하세요."
    If mSourceRange.Rows.Count <> mRows Or mSourceRange.Columns.Count <> mColumns Then NxRaiseContractError "분석 범위 크기가 변경되었습니다."
    If Not NxDataAnalysisMatrixMatches(mValues, NxDataAnalysisRangeMatrix(mSourceRange, False)) Then NxRaiseContractError "미리보기 이후 원본 값이 변경되었습니다."
    If Not NxDataAnalysisMatrixMatches(mFormulas, NxDataAnalysisRangeMatrix(mSourceRange, True)) Then NxRaiseContractError "미리보기 이후 원본 수식이 변경되었습니다."
    If NxDataAnalysisHiddenAxes(mSourceRange, mRowDirection) <> mHiddenAxes Then NxRaiseContractError "미리보기 이후 숨김 또는 필터 상태가 변경되었습니다."
End Sub

Public Sub Configure(ByVal featureId As String, ByVal source As Range, ByVal hasTitleRow As Boolean, _
    ByVal keyColumn As Long, ByVal trimText As Boolean, ByVal caseSensitive As Boolean, _
    ByVal includeHidden As Boolean)

    Dim policy As New CNxLimitsPolicy
    If mConfigured Then NxRaiseContractError "Data request can be configured only once"
    If featureId <> "NX-DATA-UNIQUE-COUNT" And featureId <> "NX-DATA-DUPLICATE-LIST" Then NxRaiseContractError "Data request feature is outside the current closed contract"
    If source Is Nothing Then NxRaiseContractError "Data request requires a source range"
    If source.Areas.Count <> 1 Then NxRaiseContractError "Data request requires one contiguous source range"
    If keyColumn < 1 Or keyColumn > source.Columns.Count Then NxRaiseContractError "Data key column is outside the source range"
    If hasTitleRow And source.Rows.Count < 2 Then NxRaiseContractError "Data source has no row below its title"
    If CDbl(source.CountLarge) > CDbl(policy.MaximumCellsFor(featureId)) Then NxRaiseContractError "Data source exceeds the exact feature limit"
    RequireCurrentSelection source

    mFeatureId = featureId
    Set mSourceRange = source
    mHasTitleRow = hasTitleRow
    mKeyColumn = keyColumn
    mTrimText = trimText
    mCaseSensitive = caseSensitive
    mIncludeHidden = includeHidden
    mActualCellCount = CLng(CDbl(source.CountLarge))
    mSelectionFingerprint = NxContextFactory.CaptureCurrentFingerprint()
    mConfigured = True
End Sub

Public Sub RequireContext(ByVal context As CNxExecutionContext)
    EnsureConfigured
    If context Is Nothing Then NxRaiseContractError "Data request requires an execution context"
    If context.Fingerprint <> mSelectionFingerprint Then NxRaiseContractError "Data source no longer matches the captured selection"
    If context.ActualCellCount <> CDbl(mActualCellCount) Then NxRaiseContractError "Data source cell count changed"
    If mAnalysis Then RequireFreshAnalysis
End Sub

Public Property Get FeatureId() As String: EnsureConfigured: FeatureId = mFeatureId: End Property
Public Property Get SourceRange() As Range: EnsureConfigured: Set SourceRange = mSourceRange: End Property
Public Property Get HasTitleRow() As Boolean: EnsureConfigured: HasTitleRow = mHasTitleRow: End Property
Public Property Get KeyColumn() As Long: EnsureConfigured: KeyColumn = mKeyColumn: End Property
Public Property Get TrimText() As Boolean: EnsureConfigured: TrimText = mTrimText: End Property
Public Property Get CaseSensitive() As Boolean: EnsureConfigured: CaseSensitive = mCaseSensitive: End Property
Public Property Get IncludeHidden() As Boolean: EnsureConfigured: IncludeHidden = mIncludeHidden: End Property
Public Property Get ActualCellCount() As Long: EnsureConfigured: ActualCellCount = mActualCellCount: End Property
Public Property Get IsAnalysis() As Boolean: EnsureConfigured: IsAnalysis = mAnalysis: End Property
Public Property Get RowDirection() As Boolean: EnsureConfigured: RowDirection = mRowDirection: End Property
Public Property Get WholeRecord() As Boolean: EnsureConfigured: WholeRecord = mWholeRecord: End Property
Public Property Get SkipBlank() As Boolean: EnsureConfigured: SkipBlank = mSkipBlank: End Property
Public Property Get IncludeRatio() As Boolean: EnsureConfigured: IncludeRatio = mIncludeRatio: End Property
Public Property Get SortMode() As String: EnsureConfigured: SortMode = mSortMode: End Property
Public Property Get CompareColumn() As Long: EnsureConfigured: CompareColumn = mCompareColumn: End Property
Public Property Get DenseCompare() As Boolean: EnsureConfigured: DenseCompare = mDenseCompare: End Property
Public Property Get OutputMode() As String: EnsureConfigured: OutputMode = mOutputMode: End Property
Public Property Get SnapshotValues() As Variant: EnsureConfigured: SnapshotValues = mValues: End Property
Public Property Get HiddenAxes() As String: EnsureConfigured: HiddenAxes = mHiddenAxes: End Property

Private Sub RequireCurrentSelection(ByVal source As Range)
    Dim selected As Range
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "Data source must be the current range selection"
    Set selected = Application.Selection
    If Not (selected.Parent Is source.Parent) Then NxRaiseContractError "Data source must be the current range selection"
    If selected.Address(True, True, xlA1, False) <> source.Address(True, True, xlA1, False) Then NxRaiseContractError "Data source must be the current range selection"
End Sub

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Data request is not configured"
End Sub
