VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDataGroup"
Option Explicit

Private mKey As String
Private mDisplayValue As Variant
Private mCount As Long
Private mFirstSourceRow As Long
Private mLastSourceRow As Long
Private mSourceRows As Collection
Private mOriginalValues As Collection
Private mConfigured As Boolean
Private mIsSealed As Boolean

Private Sub Class_Initialize()
    Set mSourceRows = New Collection
    Set mOriginalValues = New Collection
End Sub

Public Sub Configure(ByVal key As String, ByVal displayValue As Variant, ByVal sourceRow As Long)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Data group can be configured only once"
    If Len(key) = 0 Then NxRaiseContractError "Data group key is required"
    If sourceRow < 1 Then NxRaiseContractError "Data group source row must be positive"

    mKey = key
    mDisplayValue = displayValue
    mCount = 1
    mFirstSourceRow = sourceRow
    mLastSourceRow = sourceRow
    mSourceRows.Add sourceRow
    mOriginalValues.Add displayValue
    mConfigured = True
End Sub

Public Sub AddOccurrence(ByVal sourceRow As Long, Optional ByVal originalValue As Variant)
    EnsureConfiguredMutable
    If sourceRow < 1 Then NxRaiseContractError "Data group source row must be positive"
    If sourceRow <= mLastSourceRow Then NxRaiseContractError "Data group occurrences must keep strict source order"

    mCount = mCount + 1
    mLastSourceRow = sourceRow
    mSourceRows.Add sourceRow
    If IsMissing(originalValue) Then
        mOriginalValues.Add mDisplayValue
    Else
        mOriginalValues.Add originalValue
    End If
End Sub

Public Sub Seal()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Data group is not configured"
    mIsSealed = True
End Sub

Public Property Get Key() As String: EnsureReadable: Key = mKey: End Property
Public Property Get DisplayValue() As Variant: EnsureReadable: DisplayValue = mDisplayValue: End Property
Public Property Get Count() As Long: EnsureReadable: Count = mCount: End Property
Public Property Get FirstSourceRow() As Long: EnsureReadable: FirstSourceRow = mFirstSourceRow: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property

Public Property Get OriginalValue(ByVal occurrence As Long) As Variant
    EnsureReadable
    If occurrence < 1 Or occurrence > mCount Then NxRaiseContractError "Data occurrence is outside the sealed group"
    OriginalValue = mOriginalValues(occurrence)
End Property

Public Property Get SourceRows() As Collection
    Dim snapshot As New Collection
    Dim item As Variant
    EnsureReadable
    For Each item In mSourceRows
        snapshot.Add CLng(item)
    Next item
    Set SourceRows = snapshot
End Property

Private Sub EnsureConfiguredMutable()
    EnsureMutable
    If Not mConfigured Then NxRaiseContractError "Data group is not configured"
End Sub

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed data group cannot be changed"
End Sub

Private Sub EnsureReadable()
    If Not mIsSealed Then NxRaiseContractError "Data group must be sealed before it is read"
End Sub
