VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxTemplateRecord"
Option Explicit

Private mTemplateId As String
Private mDisplayName As String
Private mDescription As String
Private mPackageSha256 As String
Private mLogicalSha256 As String
Private mSourceKind As String
Private mRowCount As Long
Private mColumnCount As Long
Private mRiskFlags As String
Private mStatus As String
Private mCreatedUtc As String
Private mModifiedUtc As String
Private mCategory As String
Private mLastUsedUtc As String
Private mConfigured As Boolean
Private mIsSealed As Boolean

Public Sub Configure(ByVal templateId As String, ByVal displayName As String, ByVal description As String, _
    ByVal packageSha256 As String, ByVal logicalSha256 As String, ByVal sourceKind As String, _
    ByVal rowCount As Long, ByVal columnCount As Long, ByVal riskFlags As String, _
    ByVal status As String, ByVal createdUtc As String, ByVal modifiedUtc As String, _
    Optional ByVal category As String = vbNullString, Optional ByVal lastUsedUtc As String = vbNullString)

    If mConfigured Or mIsSealed Then NxRaiseContractError "Template record can be configured only once"
    If Not NxTemplateIdIsValid(templateId) Then NxRaiseContractError "Template record ID is invalid"
    If Len(Trim$(displayName)) = 0 Then NxRaiseContractError "Template record display name is required"
    If Not IsSha256(packageSha256) Or Not IsSha256(logicalSha256) Then NxRaiseContractError "Template record hash is invalid"
    If sourceKind <> "range" And sourceKind <> "sheet_used_range" And sourceKind <> "workbook" And sourceKind <> "file" Then NxRaiseContractError "Template record source kind is invalid"
    If rowCount <= 0 Or columnCount <= 0 Then NxRaiseContractError "Template record dimensions are invalid"
    If CDbl(rowCount) * CDbl(columnCount) > 100000# Then NxRaiseContractError "Template record dimensions exceed the feature limit"
    If status <> "healthy" And status <> "damaged" Then NxRaiseContractError "Template record status is invalid"
    If Len(createdUtc) = 0 Or Len(modifiedUtc) = 0 Then NxRaiseContractError "Template record timestamps are required"
    If Len(category) > 80 Or Len(displayName) > 200 Or Len(description) > 10000 Then NxRaiseContractError "Template metadata text is too long"
    mTemplateId = templateId
    mDisplayName = displayName
    mDescription = description
    mPackageSha256 = LCase$(packageSha256)
    mLogicalSha256 = LCase$(logicalSha256)
    mSourceKind = sourceKind
    mRowCount = rowCount
    mColumnCount = columnCount
    mRiskFlags = riskFlags
    mStatus = status
    mCreatedUtc = createdUtc
    mModifiedUtc = modifiedUtc
    mCategory = Trim$(category)
    mLastUsedUtc = lastUsedUtc
    mConfigured = True
End Sub

Public Sub Seal()
    If Not mConfigured Or mIsSealed Then NxRaiseContractError "Template record cannot be sealed"
    mIsSealed = True
End Sub

Public Property Get TemplateId() As String: EnsureReadable: TemplateId = mTemplateId: End Property
Public Property Get DisplayName() As String: EnsureReadable: DisplayName = mDisplayName: End Property
Public Property Get Description() As String: EnsureReadable: Description = mDescription: End Property
Public Property Get PackageSha256() As String: EnsureReadable: PackageSha256 = mPackageSha256: End Property
Public Property Get LogicalSha256() As String: EnsureReadable: LogicalSha256 = mLogicalSha256: End Property
Public Property Get SourceKind() As String: EnsureReadable: SourceKind = mSourceKind: End Property
Public Property Get RowCount() As Long: EnsureReadable: RowCount = mRowCount: End Property
Public Property Get ColumnCount() As Long: EnsureReadable: ColumnCount = mColumnCount: End Property
Public Property Get RiskFlags() As String: EnsureReadable: RiskFlags = mRiskFlags: End Property
Public Property Get Status() As String: EnsureReadable: Status = mStatus: End Property
Public Property Get CreatedUtc() As String: EnsureReadable: CreatedUtc = mCreatedUtc: End Property
Public Property Get ModifiedUtc() As String: EnsureReadable: ModifiedUtc = mModifiedUtc: End Property
Public Property Get Category() As String: EnsureReadable: Category = mCategory: End Property
Public Property Get LastUsedUtc() As String: EnsureReadable: LastUsedUtc = mLastUsedUtc: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property

Private Function IsSha256(ByVal value As String) As Boolean
    Dim index As Long, character As String
    If Len(value) <> 64 Then Exit Function
    value = LCase$(value)
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If InStr(1, "0123456789abcdef", character, vbBinaryCompare) = 0 Then Exit Function
    Next index
    IsSha256 = True
End Function

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Template record must be sealed before use"
End Sub
