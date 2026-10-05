VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxSideEffect"
Option Explicit

Private mKind As NxSideEffectKind
Private mTarget As String
Private mMutationMask As NxCellMutationMask
Private mRollbackStrategy As NxRollbackStrategy
Private mCompensatorMetadata As String
Private mPayloadKind As NxSideEffectPayloadKind
Private mPayload As String
Private mCapturedPriorText As String
Private mHasCapturedPriorText As Boolean
Private mMaximumCells As Long
Private mDeclaredItems As Long
Private mDirectoryItems As Collection
Private mOwner As String
Private mLimitDecision As Long
Private mIsSealed As Boolean
Private mConfigured As Boolean

Private Sub Class_Initialize()
    Set mDirectoryItems = New Collection
End Sub

Friend Sub ConfigureCell(ByVal canonicalTarget As String, ByVal mutationMask As NxCellMutationMask, ByVal maximumCells As Long)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Side effect can be configured only once"
    If Len(Trim$(canonicalTarget)) = 0 Then NxRaiseContractError "Side effect target is required"
    If mutationMask <= 0 Or (CLng(mutationMask) And Not CLng(NxCellAllSupportedMask)) <> 0 Then NxRaiseContractError "Cell mutation mask is unsupported"
    If maximumCells <= 0 Then NxRaiseContractError "Cell mutation requires a declared max_cells"
    mKind = NxCellMutation
    mTarget = canonicalTarget
    mMutationMask = mutationMask
    mRollbackStrategy = NxRollbackCellJournal
    mMaximumCells = maximumCells
    mConfigured = True
End Sub

Friend Sub ConfigureShape(ByVal canonicalTarget As String)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Side effect can be configured only once"
    If Len(Trim$(canonicalTarget)) = 0 Then NxRaiseContractError "Shape insertion target is required"
    mKind = NxShapeInsert
    mTarget = canonicalTarget
    mRollbackStrategy = NxRollbackTypedCompensator
    mCompensatorMetadata = "delete inserted shapes"
    mMaximumCells = 0
    mDeclaredItems = 1
    mConfigured = True
End Sub

Friend Sub ConfigureShapeBatch(ByVal canonicalTarget As String, ByVal declaredItems As Long, ByVal compensatorMetadata As String)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Side effect can be configured only once"
    If Len(Trim$(canonicalTarget)) = 0 Then NxRaiseContractError "Shape insertion batch target is required"
    If declaredItems < 1 Or declaredItems > 500 Then NxRaiseContractError "Shape insert batch must contain 1-500 items"
    If Len(Trim$(compensatorMetadata)) = 0 Then NxRaiseContractError "Shape insertion batch requires a compensator"
    mKind = NxShapeInsert
    mTarget = canonicalTarget
    mDeclaredItems = declaredItems
    mRollbackStrategy = NxRollbackTypedCompensator
    mCompensatorMetadata = "delete inserted shapes"
    mMaximumCells = 0
    mConfigured = True
End Sub

Friend Sub ConfigureDirectoryBatch(ByVal canonicalTarget As String, ByVal items As Variant)
    Dim item As Variant
    Dim itemText As String
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Side effect can be configured only once"
    If Len(Trim$(canonicalTarget)) = 0 Then NxRaiseContractError "Directory batch target is required"
    If Not IsObject(items) Then NxRaiseContractError "Directory batch requires a declared item collection"
    If TypeName(items) <> "Collection" Then NxRaiseContractError "Directory batch requires a declared item collection"
    If items.Count <= 0 Or items.Count > 500 Then NxRaiseContractError "Directory batch must declare between 1 and 500 items"
    For Each item In items
        itemText = Trim$(CStr(item))
        If Len(itemText) = 0 Then NxRaiseContractError "Directory batch item is required"
        mDirectoryItems.Add itemText
    Next item
    mKind = NxDirectoryBatchCreate
    mTarget = canonicalTarget
    mRollbackStrategy = NxRollbackTypedCompensator
    mCompensatorMetadata = "remove directories created by the batch"
    mMaximumCells = 0
    mDeclaredItems = mDirectoryItems.Count
    mConfigured = True
End Sub

Friend Sub ConfigureShapeMutation(ByVal canonicalTarget As String, ByVal compensatorMetadata As String)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Side effect can be configured only once"
    If Len(Trim$(canonicalTarget)) = 0 Or Len(Trim$(compensatorMetadata)) = 0 Then NxRaiseContractError "Picture mutation requires target and compensator"
    mKind = NxShapeMutation: mTarget = canonicalTarget
    mRollbackStrategy = NxRollbackTypedCompensator: mCompensatorMetadata = compensatorMetadata
    mMaximumCells = 0: mConfigured = True
End Sub

Friend Sub ConfigureCommandOwned(ByVal kind As NxSideEffectKind, ByVal canonicalTarget As String, _
    ByVal compensatorMetadata As String, ByVal declaredCells As Long)

    EnsureMutable
    If mConfigured Then NxRaiseContractError "Side effect can be configured only once"
    If kind <> NxBinaryPackageWrite And kind <> NxWorksheetCreate And kind <> NxColumnInsert Then NxRaiseContractError "Command-owned side effect kind is closed"
    If Len(Trim$(canonicalTarget)) = 0 Or Len(Trim$(compensatorMetadata)) = 0 Then NxRaiseContractError "Command-owned side effect requires target and compensator"
    If kind = NxBinaryPackageWrite And declaredCells <> 0 Then NxRaiseContractError "Binary package effect cannot declare cells"
    If (kind = NxWorksheetCreate Or kind = NxColumnInsert) And declaredCells <= 0 Then NxRaiseContractError "Worksheet creation requires an exact cell count"
    mKind = kind
    mTarget = canonicalTarget
    mRollbackStrategy = NxRollbackTypedCompensator
    mCompensatorMetadata = compensatorMetadata
    mMaximumCells = declaredCells
    mConfigured = True
End Sub

Friend Sub ConfigureTextAdapter(ByVal kind As NxSideEffectKind, ByVal canonicalTarget As String, ByVal payload As String, ByVal compensatorMetadata As String, ByVal capturedPriorText As String, ByVal hasCapturedPriorText As Boolean)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Side effect can be configured only once"
    If Len(Trim$(canonicalTarget)) = 0 Then NxRaiseContractError "Side effect target metadata is required"
    If kind <> NxFileCreate And kind <> NxFileReplace And kind <> NxClipboardWrite Then NxRaiseContractError "Text adapter kind is closed"
    If Len(Trim$(compensatorMetadata)) = 0 Then NxRaiseContractError "Text adapter compensator metadata is required"
    If kind = NxFileCreate And hasCapturedPriorText Then NxRaiseContractError "File create must not carry prior text"
    If (kind = NxFileReplace Or kind = NxClipboardWrite) And Not hasCapturedPriorText Then NxRaiseContractError "Text adapter requires captured prior text"
    mKind = kind: mTarget = canonicalTarget: mPayloadKind = NxPayloadText: mPayload = payload
    mCompensatorMetadata = compensatorMetadata: mCapturedPriorText = capturedPriorText: mHasCapturedPriorText = hasCapturedPriorText
    mRollbackStrategy = NxRollbackTypedCompensator: mMaximumCells = 0: mConfigured = True
End Sub

Friend Sub ConfigureManualAdapter(ByVal kind As NxSideEffectKind, ByVal canonicalTarget As String, ByVal payloadKind As NxSideEffectPayloadKind, ByVal payload As String, ByVal recoveryInstructions As String)
    EnsureMutable
    If mConfigured Then NxRaiseContractError "Side effect can be configured only once"
    If kind <> NxExternalProcess And kind <> NxBrowserOpen Then NxRaiseContractError "Manual adapter kind is closed"
    If Len(Trim$(canonicalTarget)) = 0 Or Len(Trim$(payload)) = 0 Or Len(Trim$(recoveryInstructions)) = 0 Then NxRaiseContractError "Manual adapter requires metadata, payload, and recovery instructions"
    If (kind = NxExternalProcess And payloadKind <> NxPayloadCommandLine) Or (kind = NxBrowserOpen And payloadKind <> NxPayloadUri) Then NxRaiseContractError "Manual adapter payload kind is invalid"
    mKind = kind: mTarget = canonicalTarget: mPayloadKind = payloadKind: mPayload = payload
    mCompensatorMetadata = recoveryInstructions: mRollbackStrategy = NxRollbackManualRecovery: mMaximumCells = 0: mConfigured = True
End Sub

Friend Sub BindDecision(ByVal owner As String, ByVal limitDecision As Long)
    EnsureMutable
    If Len(owner) = 0 Or limitDecision <= 0 Then NxRaiseContractError "Side effect requires exact owner and limit decision"
    If mKind = NxCellMutation And mMaximumCells > limitDecision Then NxRaiseContractError "Side effect exceeds exact feature limit"
    If (mKind = NxWorksheetCreate Or mKind = NxColumnInsert) And mMaximumCells > limitDecision Then NxRaiseContractError "Worksheet creation exceeds exact feature limit"
    mOwner = owner: mLimitDecision = limitDecision
End Sub

Public Sub Seal()
    EnsureMutable
    If Len(Trim$(mTarget)) = 0 Then NxRaiseContractError "Side effect target is required"
    If Len(mOwner) = 0 Or mLimitDecision <= 0 Then NxRaiseContractError "Side effect owner and limit decision are required"
    ValidatePayloadContract
    Call NxSideEffectWireName(mKind)
    mIsSealed = True
End Sub

Public Property Get Kind() As NxSideEffectKind: Kind = mKind: End Property
Public Property Get Target() As String: Target = mTarget: End Property
Public Property Get MutationMask() As NxCellMutationMask: MutationMask = mMutationMask: End Property
Public Property Get RollbackStrategy() As NxRollbackStrategy: RollbackStrategy = mRollbackStrategy: End Property
Public Property Get CompensatorMetadata() As String: CompensatorMetadata = mCompensatorMetadata: End Property
Public Property Get PayloadKind() As NxSideEffectPayloadKind: PayloadKind = mPayloadKind: End Property
Public Property Get Payload() As String: Payload = mPayload: End Property
Public Property Get CapturedPriorText() As String: CapturedPriorText = mCapturedPriorText: End Property
Public Property Get HasCapturedPriorText() As Boolean: HasCapturedPriorText = mHasCapturedPriorText: End Property
Public Property Get MaximumCells() As Long: MaximumCells = mMaximumCells: End Property
Public Property Get DeclaredItems() As Long: DeclaredItems = mDeclaredItems: End Property
Public Property Get DirectoryItems() As Collection
    Dim snapshot As New Collection
    Dim item As Variant
    For Each item In mDirectoryItems
        snapshot.Add CStr(item)
    Next item
    Set DirectoryItems = snapshot
End Property
Public Property Get IsSealed() As Boolean: IsSealed = mIsSealed: End Property
Public Property Get Owner() As String: Owner = mOwner: End Property
Public Property Get LimitDecision() As Long: LimitDecision = mLimitDecision: End Property

Public Function StableIdentity() As String
    If Not mIsSealed Then NxRaiseContractError "Side effect must be sealed"
    StableIdentity = "NXEFFECT2" & CanonicalField(NxSideEffectWireName(mKind)) & CanonicalField(mTarget) & CanonicalField(mOwner) & CanonicalField(NxInvariantUnsigned(CDbl(mLimitDecision))) & CanonicalField(NxInvariantUnsigned(CDbl(mMutationMask))) & CanonicalField(NxInvariantUnsigned(CDbl(mRollbackStrategy))) & CanonicalField(NxInvariantUnsigned(CDbl(mMaximumCells))) & CanonicalField(NxInvariantUnsigned(CDbl(mDeclaredItems))) & CanonicalField(NxInvariantUnsigned(CDbl(mPayloadKind))) & CanonicalField(mPayload) & CanonicalField(BooleanWire(mHasCapturedPriorText)) & CanonicalField(mCapturedPriorText) & CanonicalField(mCompensatorMetadata) & StructuredItemsIdentity
End Function

Private Sub ValidatePayloadContract()
    Select Case mKind
        Case NxCellMutation
            If mPayloadKind <> NxPayloadNone Or Len(mPayload) <> 0 Or mHasCapturedPriorText Or Len(mCapturedPriorText) <> 0 Then NxRaiseContractError "Cell effect cannot carry non-cell payload"
        Case NxFileCreate
            If mPayloadKind <> NxPayloadText Or mHasCapturedPriorText Then NxRaiseContractError "File create requires explicit text payload only"
        Case NxFileReplace, NxClipboardWrite
            If mPayloadKind <> NxPayloadText Or Not mHasCapturedPriorText Then NxRaiseContractError "Text replacement requires explicit payload and captured prior text"
        Case NxExternalProcess
            If mPayloadKind <> NxPayloadCommandLine Or mRollbackStrategy <> NxRollbackManualRecovery Then NxRaiseContractError "External process requires explicit command payload and manual recovery"
        Case NxBrowserOpen
            If mPayloadKind <> NxPayloadUri Or mRollbackStrategy <> NxRollbackManualRecovery Then NxRaiseContractError "Browser requires explicit URI payload and manual recovery"
        Case NxShapeInsert
            If mPayloadKind <> NxPayloadNone Or mDeclaredItems < 1 Or mDeclaredItems > 500 Or mRollbackStrategy <> NxRollbackTypedCompensator Or mCompensatorMetadata <> "delete inserted shapes" Then NxRaiseContractError "Shape insert batch requires a typed compensator and 1-500 items"
        Case NxShapeMutation
            If mPayloadKind <> NxPayloadNone Or mRollbackStrategy <> NxRollbackTypedCompensator Or mCompensatorMetadata <> "restore existing shape geometry and placement" Then NxRaiseContractError "Shape mutation requires typed rollback"
        Case NxBinaryPackageWrite
            If mPayloadKind <> NxPayloadNone Or mMaximumCells <> 0 Or mRollbackStrategy <> NxRollbackTypedCompensator Then NxRaiseContractError "Binary package write requires command-owned typed compensation"
        Case NxWorksheetCreate, NxColumnInsert
            If mPayloadKind <> NxPayloadNone Or mMaximumCells <= 0 Or mRollbackStrategy <> NxRollbackTypedCompensator Then NxRaiseContractError "Worksheet creation requires command-owned typed compensation"
        Case NxDirectoryBatchCreate
            If mPayloadKind <> NxPayloadNone Or mMaximumCells <> 0 Or mDeclaredItems <= 0 Or mDeclaredItems > 500 Or mRollbackStrategy <> NxRollbackTypedCompensator Then NxRaiseContractError "Directory batch creation requires a typed compensator and 1-500 declared items"
            If mDirectoryItems.Count <> mDeclaredItems Then NxRaiseContractError "Directory batch declared item count mismatch"
            If mCompensatorMetadata <> "remove directories created by the batch" Then NxRaiseContractError "Directory batch compensator is not exact"
        Case Else
            NxRaiseContractError "Side effect kind is closed"
    End Select
End Sub

Private Sub EnsureMutable()
    If mIsSealed Then NxRaiseContractError "Sealed side effect cannot be changed"
End Sub

Private Function CanonicalField(ByVal value As String) As String
    CanonicalField = NxInvariantUnsigned(CDbl(Len(value))) & ":" & value
End Function

Private Function BooleanWire(ByVal value As Boolean) As String
    If value Then BooleanWire = "1" Else BooleanWire = "0"
End Function

Private Function StructuredItemsIdentity() As String
    Dim item As Variant
    If mKind = NxDirectoryBatchCreate Then
        For Each item In mDirectoryItems
            StructuredItemsIdentity = StructuredItemsIdentity & CanonicalField(CStr(item))
        Next item
    End If
End Function
