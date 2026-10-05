VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxTemplateAnalysis"
Option Explicit

Private mSourceAddress As String
Private mRowCount As Long
Private mColumnCount As Long
Private mCellCount As Double
Private mStoredConstantCount As Long
Private mExcludedConstantCount As Long
Private mFormulaCount As Long
Private mIncludeHidden As Boolean
Private mRiskFlags As Collection
Private mRiskDetails As Collection
Private mHasBlockedResource As Boolean
Private mBlockReason As String
Private mHasExcludedDisplayResource As Boolean
Private mExcludedResourceSummary As String
Private mConfigured As Boolean
Private mIsSealed As Boolean

Private Sub Class_Initialize()
    Set mRiskFlags = New Collection
    Set mRiskDetails = New Collection
End Sub

Friend Sub Configure(ByVal source As Range, ByVal includeHidden As Boolean)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Template analysis can be configured only once"
    If source Is Nothing Then NxRaiseContractError "Template analysis requires a source range"
    mSourceAddress = source.Address(True, True, xlA1, False)
    mRowCount = source.Rows.Count
    mColumnCount = source.Columns.Count
    mCellCount = CDbl(source.CountLarge)
    mIncludeHidden = includeHidden
    mConfigured = True
End Sub

Friend Sub CountStoredConstant()
    EnsureConfiguredMutable
    mStoredConstantCount = mStoredConstantCount + 1
End Sub

Friend Sub CountExcludedConstant()
    EnsureConfiguredMutable
    mExcludedConstantCount = mExcludedConstantCount + 1
End Sub

Friend Sub CountFormula()
    EnsureConfiguredMutable
    mFormulaCount = mFormulaCount + 1
End Sub

Friend Sub AddRisk(ByVal flag As String, ByVal cellAddress As String, ByVal functionName As String)
    EnsureConfiguredMutable
    If Len(flag) = 0 Or Len(cellAddress) = 0 Or Len(functionName) = 0 Then NxRaiseContractError "Template risk detail is incomplete"
    If Not HasRiskFlagInternal(flag) Then mRiskFlags.Add flag
    mRiskDetails.Add cellAddress & "|" & flag & "|" & functionName
End Sub

Friend Sub MarkBlocked(ByVal reason As String)
    EnsureConfiguredMutable
    If Len(Trim$(reason)) = 0 Then NxRaiseContractError "Template block reason is required"
    If Not mHasBlockedResource Then mBlockReason = reason
    mHasBlockedResource = True
End Sub

Friend Sub MarkExcludedDisplayResource(ByVal summary As String)
    EnsureConfiguredMutable
    If Len(Trim$(summary)) = 0 Then NxRaiseContractError "Template exclusion summary is required"
    mHasExcludedDisplayResource = True
    If Len(mExcludedResourceSummary) = 0 Then
        mExcludedResourceSummary = summary
    ElseIf InStr(1, mExcludedResourceSummary, summary, vbBinaryCompare) = 0 Then
        mExcludedResourceSummary = mExcludedResourceSummary & ", " & summary
    End If
End Sub

Friend Sub Seal()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Template analysis is not configured"
    mIsSealed = True
End Sub

Public Function HasRiskFlag(ByVal flag As String) As Boolean
    EnsureReadable
    HasRiskFlag = HasRiskFlagInternal(flag)
End Function

Public Property Get SourceAddress() As String: EnsureReadable: SourceAddress = mSourceAddress: End Property
Public Property Get RowCount() As Long: EnsureReadable: RowCount = mRowCount: End Property
Public Property Get ColumnCount() As Long: EnsureReadable: ColumnCount = mColumnCount: End Property
Public Property Get CellCount() As Double: EnsureReadable: CellCount = mCellCount: End Property
Public Property Get StoredConstantCount() As Long: EnsureReadable: StoredConstantCount = mStoredConstantCount: End Property
Public Property Get ExcludedConstantCount() As Long: EnsureReadable: ExcludedConstantCount = mExcludedConstantCount: End Property
Public Property Get FormulaCount() As Long: EnsureReadable: FormulaCount = mFormulaCount: End Property
Public Property Get IncludeHidden() As Boolean: EnsureReadable: IncludeHidden = mIncludeHidden: End Property
Public Property Get HasBlockedResource() As Boolean: EnsureReadable: HasBlockedResource = mHasBlockedResource: End Property
Public Property Get BlockReason() As String: EnsureReadable: BlockReason = mBlockReason: End Property
Public Property Get HasExcludedDisplayResource() As Boolean: EnsureReadable: HasExcludedDisplayResource = mHasExcludedDisplayResource: End Property
Public Property Get ExcludedResourceSummary() As String: EnsureReadable: ExcludedResourceSummary = mExcludedResourceSummary: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property

Public Property Get RiskFlags() As Collection
    Dim snapshot As New Collection, item As Variant
    EnsureReadable
    For Each item In mRiskFlags
        snapshot.Add CStr(item)
    Next item
    Set RiskFlags = snapshot
End Property

Public Property Get RiskDetails() As Collection
    Dim snapshot As New Collection, item As Variant
    EnsureReadable
    For Each item In mRiskDetails
        snapshot.Add CStr(item)
    Next item
    Set RiskDetails = snapshot
End Property

Private Function HasRiskFlagInternal(ByVal flag As String) As Boolean
    Dim item As Variant
    For Each item In mRiskFlags
        If CStr(item) = flag Then HasRiskFlagInternal = True: Exit Function
    Next item
End Function

Private Sub EnsureConfiguredMutable()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Template analysis is not configured"
End Sub

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed template analysis cannot be changed"
End Sub

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Template analysis must be sealed before use"
End Sub
