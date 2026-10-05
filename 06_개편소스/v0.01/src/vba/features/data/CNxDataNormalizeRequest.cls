VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDataNormalizeRequest"
Option Explicit

Private Const NX_NORMALIZE_TABLE_MAX_CELLS As Long = 50000

Private mSourceRange As Range
Private mPreset As String
Private mDateFormat As String
Private mOutputMode As String
Private mActualCellCount As Long
Private mSelectionFingerprint As String
Private mConfigured As Boolean

Public Sub Configure(ByVal source As Range, ByVal preset As String, ByVal dateFormat As String, _
    Optional ByVal outputMode As String = "새 시트 만들기")
    Dim policy As New CNxLimitsPolicy
    If mConfigured Then NxRaiseContractError "Data normalization request can be configured only once"
    If source Is Nothing Then NxRaiseContractError "Data normalization requires a source range"
    If source.Areas.Count <> 1 Then NxRaiseContractError "Data normalization requires one contiguous source range"
    If CDbl(source.CountLarge) > CDbl(policy.MaximumCellsFor("NX-DATA-NORMALIZE")) Then
        NxRaiseContractError "Data normalization source exceeds the exact feature limit"
    End If
    RequireCurrentSelection source
    mPreset = NxDataNormalizePreset(preset)
    mDateFormat = NxDataNormalizeDateFormat(dateFormat)
    mOutputMode = NxDataNormalizeOutputMode(outputMode)
    If mPreset = "표 구조 정리" And CDbl(source.CountLarge) > CDbl(NX_NORMALIZE_TABLE_MAX_CELLS) Then
        NxRaiseContractError "표 구조 정리는 최대 50,000개 셀까지 실행할 수 있습니다."
    End If

    Set mSourceRange = source
    mActualCellCount = CLng(CDbl(source.CountLarge))
    mSelectionFingerprint = NxContextFactory.CaptureCurrentFingerprint()
    mConfigured = True
End Sub
Public Sub RequireContext(ByVal context As CNxExecutionContext)
    EnsureConfigured
    If context Is Nothing Then NxRaiseContractError "Data normalization requires an execution context"
    If context.Fingerprint <> mSelectionFingerprint Then
        NxRaiseContractError "Data normalization source no longer matches the captured selection"
    End If
    If context.ActualCellCount <> CDbl(mActualCellCount) Then
        NxRaiseContractError "Data normalization source cell count changed"
    End If
End Sub

Public Property Get FeatureId() As String: FeatureId = "NX-DATA-NORMALIZE": End Property
Public Property Get SourceRange() As Range: EnsureConfigured: Set SourceRange = mSourceRange: End Property
Public Property Get Preset() As String: EnsureConfigured: Preset = mPreset: End Property
Public Property Get DateFormat() As String: EnsureConfigured: DateFormat = mDateFormat: End Property
Public Property Get OutputMode() As String: EnsureConfigured: OutputMode = mOutputMode: End Property
Public Property Get ActualCellCount() As Long: EnsureConfigured: ActualCellCount = mActualCellCount: End Property

Private Sub RequireCurrentSelection(ByVal source As Range)
    Dim selected As Range
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "Data normalization source must be the current range selection"
    Set selected = Application.Selection
    If Not (selected.Parent Is source.Parent) Then NxRaiseContractError "Data normalization source must be the current range selection"
    If selected.Address(True, True, xlA1, False) <> source.Address(True, True, xlA1, False) Then
        NxRaiseContractError "Data normalization source must be the current range selection"
    End If
End Sub

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Data normalization request is not configured"
End Sub
