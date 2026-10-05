VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxTemplateRequest"
Option Explicit

Private mFeatureId As String
Private mSourceKind As String
Private mSourceRange As Range
Private mRetainedAreas As Collection
Private mDisplayName As String
Private mDescription As String
Private mCategory As String

Private mIncludeHidden As Boolean
Private mExclusionAccepted As Boolean
Private mRiskConfirmed As Boolean
Private mHasRiskyFormula As Boolean
Private mConfigured As Boolean
Private mSourceRanges As Collection
Private mRetainAll As Boolean
Private mActualCellCount As Long

Public Property Get Category() As String
    Category = mCategory
End Property

Public Property Let Category(ByVal value As String)
    If Len(value) > 80 Then NxRaiseContractError "분류는 80자 이내로 입력하세요."
    mCategory = Trim$(value)
End Property

Public Sub ConfigureRegistration(ByVal featureId As String, ByVal sourceKind As String, _
    ByVal source As Range, ByVal retained As Collection, ByVal displayName As String, _
    ByVal description As String, ByVal includeHidden As Boolean, _
    ByVal exclusionAccepted As Boolean, ByVal riskConfirmed As Boolean)

    Dim policy As New CNxLimitsPolicy
    If mConfigured Then NxRaiseContractError "Template request can be configured only once"
    If featureId <> "NX-TPL-REGISTER-SHEET" Then NxRaiseContractError "Template registration feature is outside the closed contract"
    If sourceKind <> "sheet_used_range" Then NxRaiseContractError "Template source kind is outside the closed contract"
    If source Is Nothing Or retained Is Nothing Then NxRaiseContractError "Template registration requires source and retained areas"
    If CDbl(source.CountLarge) > CDbl(policy.MaximumCellsFor(featureId)) Then NxRaiseContractError "Template source exceeds the exact feature limit"
    If Len(Trim$(displayName)) = 0 Then NxRaiseContractError "Template display name is required"
    mFeatureId = featureId
    mSourceKind = sourceKind
    Set mSourceRange = source
    Set mSourceRanges = New Collection
    mSourceRanges.Add source
    mActualCellCount = CLng(CDbl(source.CountLarge))
    Set mRetainedAreas = SnapshotRanges(retained)
    mDisplayName = Trim$(displayName)
    mDescription = Trim$(description)
    mIncludeHidden = includeHidden
    mExclusionAccepted = exclusionAccepted
    mRiskConfirmed = riskConfirmed
    mConfigured = True
End Sub

Public Sub ConfigureWorkbookRegistration(ByVal sourceBook As Workbook, ByVal sourceKind As String, _
    ByVal displayName As String, ByVal description As String, ByVal retainAll As Boolean, _
    ByVal exclusionAccepted As Boolean, ByVal riskConfirmed As Boolean)

    Dim sheet As Worksheet, source As Range, retained As Collection, analysis As CNxTemplateAnalysis
    Dim policy As New CNxLimitsPolicy, totalCells As Double
    If mConfigured Then NxRaiseContractError "Template request can be configured only once"
    If sourceBook Is Nothing Then NxRaiseContractError "등록할 통합문서가 없습니다."
    If sourceBook.IsAddin Or sourceBook Is ThisWorkbook Then NxRaiseContractError "일반 통합문서만 등록할 수 있습니다."
    If sourceBook.Sheets.Count <> sourceBook.Worksheets.Count Then NxRaiseContractError "차트 시트나 매크로 시트가 있는 통합문서는 등록할 수 없습니다."
    If sourceKind <> "workbook" And sourceKind <> "file" Then NxRaiseContractError "Template workbook source kind is invalid"
    If Len(Trim$(displayName)) = 0 Then NxRaiseContractError "템플릿 이름을 입력하세요."
    Set mSourceRanges = New Collection
    mRetainAll = retainAll
    For Each sheet In sourceBook.Worksheets
        If sheet.Visible = xlSheetVisible Then
            Set source = sheet.UsedRange
            totalCells = totalCells + CDbl(source.CountLarge)
            If totalCells > CDbl(policy.MaximumCellsFor("NX-TPL-REGISTER-SHEET")) Then NxRaiseContractError "통합문서 전체 사용 셀이 템플릿 한도를 초과합니다."
            Set retained = New Collection
            If retainAll Then retained.Add source
            Set analysis = NxTemplateAnalyze(source, retained, True, exclusionAccepted)
            If analysis.RiskFlags.Count > 0 Then mHasRiskyFormula = True
            mSourceRanges.Add source
        End If
    Next sheet
    If mSourceRanges.Count = 0 Then NxRaiseContractError "등록할 표시 워크시트가 없습니다."
    If mHasRiskyFormula And Not riskConfirmed Then NxRaiseContractError "위험 수식 확인에 동의해야 합니다."
    mFeatureId = "NX-TPL-REGISTER-SHEET"
    mSourceKind = sourceKind
    Set mSourceRange = mSourceRanges.Item(1)
    Set mRetainedAreas = New Collection
    If retainAll Then mRetainedAreas.Add mSourceRange
    mDisplayName = Trim$(displayName)
    mDescription = Trim$(description)
    mIncludeHidden = True
    mExclusionAccepted = exclusionAccepted
    mRiskConfirmed = riskConfirmed
    mActualCellCount = CLng(totalCells)
    mConfigured = True
End Sub

Public Property Get SourceRanges() As Collection
    EnsureConfigured
    Set SourceRanges = SnapshotRanges(mSourceRanges)
End Property

Public Function RetainedForSource(ByVal source As Range) As Collection
    Dim retained As New Collection
    EnsureConfigured
    If mSourceKind = "sheet_used_range" Then
        Set retained = SnapshotRanges(mRetainedAreas)
    ElseIf mRetainAll Then
        retained.Add source
    End If
    Set RetainedForSource = retained
End Function

Public Sub BindAnalysis(ByVal analysis As CNxTemplateAnalysis)
    EnsureConfigured
    If analysis Is Nothing Or Not analysis.IsSealed Then NxRaiseContractError "Template request requires a sealed analysis"
    mHasRiskyFormula = (analysis.RiskFlags.Count > 0)
    If mHasRiskyFormula And Not mRiskConfirmed Then NxRaiseContractError "Risky template formulas require explicit confirmation"
End Sub

Public Property Get FeatureId() As String: EnsureConfigured: FeatureId = mFeatureId: End Property
Public Property Get SourceKind() As String: EnsureConfigured: SourceKind = mSourceKind: End Property
Public Property Get SourceRange() As Range: EnsureConfigured: Set SourceRange = mSourceRange: End Property
Public Property Get DisplayName() As String: EnsureConfigured: DisplayName = mDisplayName: End Property
Public Property Get Description() As String: EnsureConfigured: Description = mDescription: End Property
Public Property Get IncludeHidden() As Boolean: EnsureConfigured: IncludeHidden = mIncludeHidden: End Property
Public Property Get ExclusionAccepted() As Boolean: EnsureConfigured: ExclusionAccepted = mExclusionAccepted: End Property
Public Property Get RiskConfirmed() As Boolean: EnsureConfigured: RiskConfirmed = mRiskConfirmed: End Property
Public Property Get HasRiskyFormula() As Boolean: EnsureConfigured: HasRiskyFormula = mHasRiskyFormula: End Property
Public Property Get ActualCellCount() As Long: EnsureConfigured: ActualCellCount = mActualCellCount: End Property

Public Property Get RetainedAreas() As Collection
    EnsureConfigured
    Set RetainedAreas = SnapshotRanges(mRetainedAreas)
End Property

Private Function SnapshotRanges(ByVal source As Collection) As Collection
    Dim result As New Collection, item As Variant
    For Each item In source
        result.Add item
    Next item
    Set SnapshotRanges = result
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Template request is not configured"
End Sub
