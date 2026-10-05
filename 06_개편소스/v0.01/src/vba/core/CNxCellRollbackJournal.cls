VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxCellRollbackJournal"
Option Explicit

Private mRecords As Collection
Private mRecordKeys As Object
Private mPropertySets As Object
Private mLastProperties As Collection
Private mMatchingPrevious As Boolean
Private mMatchedPropertyCount As Long
Private mPreviousNames() As String
Private mPreviousTypes() As Integer
Private mPreviousValues() As Variant
Private mPreviousCount As Long
Private mTargetPreflightPassed As Boolean
Private mMergeRecords As Collection
Private mPlanIdentity As String
Private mCaptured As Boolean
Private mBorderEdges As Variant
Private mBorderPropertyNames(5 To 12, 0 To 5) As String
Private mRangeBackup As CNxRangeFormatBackup
Private mRangeMask As NxCellMutationMask
Private mRangeCount As Long

Private Sub Class_Initialize()
    Dim edge As Variant, prefix As String
    Set mRecords = New Collection
    Set mRecordKeys = CreateObject("Scripting.Dictionary")
    mRecordKeys.CompareMode = vbBinaryCompare
    Set mPropertySets = CreateObject("Scripting.Dictionary")
    mPropertySets.CompareMode = vbBinaryCompare
    Set mMergeRecords = New Collection
    mBorderEdges = Array(xlDiagonalDown, xlDiagonalUp, xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
    For Each edge In mBorderEdges
        prefix = "Border." & NxInvariantUnsigned(CDbl(edge))
        mBorderPropertyNames(edge, 0) = prefix & ".Weight"
        mBorderPropertyNames(edge, 1) = prefix & ".Color"
        mBorderPropertyNames(edge, 2) = prefix & ".ColorIndex"
        mBorderPropertyNames(edge, 3) = prefix & ".ThemeColor"
        mBorderPropertyNames(edge, 4) = prefix & ".TintAndShade"
        mBorderPropertyNames(edge, 5) = prefix & ".LineStyle"
    Next edge
End Sub

Friend Sub Capture(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan, ByVal selectionAuthority As Collection)
    Dim effect As Variant
    Dim area As Range
    Dim cell As Range
    Dim target As Range
    Dim captureLimit As Long
    If context Is Nothing Or approvedPlan Is Nothing Or selectionAuthority Is Nothing Then NxRaiseContractError "Journal requires context, plan, and selection authority"
    If mCaptured Then NxRaiseContractError "Journal is command-scoped and may be captured once"
    If Not approvedPlan.IsSealed Then NxRaiseContractError "Journal requires the approved sealed plan"
    mPlanIdentity = approvedPlan.PlanIdentity
    Set target = context.InternalSelectionRange(selectionAuthority)
    captureLimit = 200000
    For Each effect In approvedPlan.SideEffects
        If effect.Kind = NxCellMutation Then
            If effect.RollbackStrategy <> NxRollbackCellJournal Then NxRaiseContractError "Cell journal requires cell rollback strategy"
            If effect.Target <> context.SelectionTargetIdentity Then NxRaiseContractError "Approved cell target does not match current context"
            If effect.MaximumCells < captureLimit Then captureLimit = effect.MaximumCells
            If captureLimit <= 0 Then NxRaiseContractError "Cell journal requires a declared max_cells"
            AssertSupportedTarget target, effect.MutationMask
            If CanCaptureRange(target, approvedPlan, effect.MutationMask, captureLimit) Then
                Set mRangeBackup = New CNxRangeFormatBackup
                mRangeBackup.Capture target
                mRangeMask = effect.MutationMask
                mRangeCount = CLng(target.CountLarge)
                mCaptured = True
                Exit Sub
            End If
            For Each area In target.Areas
                For Each cell In area.Cells
                    AssertSupportedCell cell, target, effect.MutationMask
                    SnapshotUniqueCell cell, effect.MutationMask, captureLimit
                Next cell
            Next area
        End If
    Next effect
    mCaptured = True
End Sub

Friend Function Rollback() As String
    Dim record As Variant
    Dim failures As String
    If Not mRangeBackup Is Nothing Then
        On Error GoTo RangeFailure
        mRangeBackup.Restore
        Exit Function
    End If
    For Each record In mRecords
        RestoreRecord record, failures
    Next record
    ' Restore merge topology after every value and formatting property has been restored.
    RestoreMerges failures
    Rollback = failures
    Exit Function
RangeFailure:
    Rollback = "범위 서식 복구 실패: " & Err.Description
    Err.Clear
End Function

Public Function VerifyResidual() As String
    Dim record As Variant, cell As Range, item As Variant, properties As Collection, currentMerge As String
    If Not mRangeBackup Is Nothing Then
        VerifyResidual = VerifyRangeResidual()
        Exit Function
    End If
    For Each record In mRecords
        Set cell = record.Item(1)
        If Len(CStr(record.Item(4))) > 0 Then
            If Not FormulaMatches(cell, record.Item(3), CStr(record.Item(4))) Then VerifyResidual = NxJoinFailure(VerifyResidual, CellTarget(cell) & "|final-state formula")
        End If
        Set properties = record.Item(5)
        For Each item In properties
            If Not PropertyMatches(cell, CStr(item(0)), item(1)) Then VerifyResidual = NxJoinFailure(VerifyResidual, CellTarget(cell) & "|final-state " & CStr(item(0)))
        Next item
    Next record
    For Each record In mMergeRecords
        Set cell = record.Item(1): currentMerge = vbNullString
        If cell.MergeCells Then currentMerge = cell.MergeArea.Address
        If currentMerge <> CStr(record.Item(2)) Then VerifyResidual = NxJoinFailure(VerifyResidual, CellTarget(cell) & "|final-state MergeArea")
    Next record
End Function

Public Property Get Formula2FallbackContract() As String: Formula2FallbackContract = "438|1004": End Property

Friend Function DetectApprovedTargetChange() As Boolean
    Dim record As Variant
    Dim cell As Range
    Dim mode As String
    On Error GoTo Changed
    For Each record In mRecords
        mode = CStr(record.Item(4))
        If Len(mode) > 0 Then
            Set cell = record.Item(1)
            Select Case mode
                Case "Formula2": If Not VariantValuesMatch(NxReadFormula2OrFallback(cell, mode), record.Item(3)) Then DetectApprovedTargetChange = True: Exit Function
                Case "Formula": If Not VariantValuesMatch(cell.Formula, record.Item(3)) Then DetectApprovedTargetChange = True: Exit Function
                Case "Value2": If Not VariantValuesMatch(cell.Value2, record.Item(3)) Then DetectApprovedTargetChange = True: Exit Function
                Case Else: NxRaiseContractError "Unknown recorded formula mode"
            End Select
        End If
    Next record
    Exit Function
Changed:
    Err.Clear
    DetectApprovedTargetChange = True
End Function

Public Property Get Count() As Long: Count = mRecords.Count + mRangeCount: End Property
Public Property Get PlanIdentity() As String: PlanIdentity = mPlanIdentity: End Property

Private Function CanCaptureRange(ByVal target As Range, ByVal plan As CNxPlan, ByVal mask As NxCellMutationMask, ByVal cap As Long) As Boolean
    Dim shape As Shape, ownedMask As Long, usedCount As Double
    If plan.FeatureId <> "NX-DRAW-BUSINESS-TABLE" And plan.FeatureId <> "NX-DRAW-TITLE-TABLE" Then Exit Function
    If plan.SideEffects.Count <> 1 Or target.Areas.Count <> 1 Then Exit Function
    If target.CountLarge < 256 Or target.CountLarge > cap Then Exit Function
    ownedMask = NxCellFontMask Or NxCellInteriorMask Or NxCellAlignmentMask Or NxCellBorderMask
    If CLng(mask) <> ownedMask Then Exit Function
    usedCount = target.Worksheet.UsedRange.CountLarge
    ' Bound the native worksheet snapshot. Embedded pictures and charts are
    ' copied by Excel; linked objects and controls retain the cell journal.
    If usedCount > 1000000 Or target.Worksheet.Shapes.Count > 16 Then Exit Function
    For Each shape In target.Worksheet.Shapes
        Select Case shape.Type
            Case msoAutoShape, msoPicture, msoChart
            Case Else: Exit Function
        End Select
    Next shape
    ' Merge boundaries are still validated before any backup or mutation.
    AssertRangeMergeBounds target, target
    CanCaptureRange = True
End Function

Private Sub AssertRangeMergeBounds(ByVal part As Range, ByVal target As Range)
    Dim state As Variant, merged As Range, inside As Range, half As Long
    state = part.MergeCells
    If Not IsNull(state) Then
        If Not CBool(state) Then Exit Sub
        Set merged = part.Cells(1, 1).MergeArea
        Set inside = Intersect(merged, target)
        If inside Is Nothing Then NxRaiseContractError "Merge area crosses approved target"
        If inside.Address <> merged.Address Then NxRaiseContractError "Merge area crosses approved target"
        Set inside = Intersect(merged, part)
        If inside.CountLarge = part.CountLarge Then Exit Sub
    End If
    ' Split only mixed blocks; unmerged blocks require no per-cell dispatch.
    If part.Rows.Count > 1 Then
        half = part.Rows.Count \ 2
        AssertRangeMergeBounds part.Resize(half), target
        AssertRangeMergeBounds part.Cells(half + 1, 1).Resize(part.Rows.Count - half, part.Columns.Count), target
    ElseIf part.Columns.Count > 1 Then
        half = part.Columns.Count \ 2
        AssertRangeMergeBounds part.Resize(1, half), target
        AssertRangeMergeBounds part.Cells(1, half + 1).Resize(1, part.Columns.Count - half), target
    Else
        Set merged = part.MergeArea
        If Intersect(merged, target).Address <> merged.Address Then NxRaiseContractError "Merge area crosses approved target"
    End If
End Sub

Private Function VerifyRangeResidual() As String
    Dim original As Range, captured As Range, index As Long, properties As Collection, item As Variant
    On Error GoTo Failed
    Set original = mRangeBackup.SourceRange
    Set captured = mRangeBackup.BackupRange
    ' Full property comparison remains on the exceptional rollback path.
    ' Successful execution needs only Excel's native snapshot.
    For index = 1 To mRangeCount
        Set properties = New Collection
        mMatchingPrevious = False
        SnapshotFont captured.Cells(index), properties
        SnapshotInterior captured.Cells(index), properties
        SnapshotAlignment captured.Cells(index), properties
        SnapshotBorders captured.Cells(index), properties
        For Each item In properties
            If Not PropertyMatches(original.Cells(index), CStr(item(0)), item(1)) Then
                VerifyRangeResidual = NxJoinFailure(VerifyRangeResidual, CellTarget(original.Cells(index)) & "|final-state " & CStr(item(0)))
            End If
        Next item
    Next index
    Exit Function
Failed:
    VerifyRangeResidual = NxJoinFailure(VerifyRangeResidual, "범위 서식 복구 검증 실패: " & Err.Description)
    Err.Clear
End Function

Friend Sub ReleaseResources()
    If Not mRangeBackup Is Nothing Then mRangeBackup.Dispose
    Set mRangeBackup = Nothing
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    ReleaseResources
End Sub

Private Sub AssertSupportedCell(ByVal cell As Range, ByVal target As Range, ByVal mutationMask As NxCellMutationMask)
    Dim listObject As ListObject
    If mutationMask <= 0 Or (CLng(mutationMask) And Not CLng(NxCellAllSupportedMask)) <> 0 Then NxRaiseContractError "Unsupported cell mutation mask"
    If mTargetPreflightPassed Then
        If cell.MergeCells Then
            If Intersect(cell.MergeArea, target).Address <> cell.MergeArea.Address Then NxRaiseContractError "Merge area crosses approved target"
        End If
        Exit Sub
    End If
    If cell.FormatConditions.Count > 0 Then NxRaiseContractError "Conditional formatting is outside the cell journal contract"
    Set listObject = Nothing
    On Error Resume Next
    Set listObject = cell.ListObject
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        NxRaiseContractError "ListObject probe failed"
    End If
    On Error GoTo 0
    If Not listObject Is Nothing Then NxRaiseContractError "ListObject is outside the cell journal contract"
    If Not HasOnlyInputValidation(cell) Then NxRaiseContractError "Data validation is outside the cell journal contract"
    If cell.MergeCells Then
        If Intersect(cell.MergeArea, target).Address <> cell.MergeArea.Address Then NxRaiseContractError "Merge area crosses approved target"
    End If
End Sub

Private Sub AssertSupportedTarget(ByVal target As Range, ByVal mutationMask As NxCellMutationMask)
    Dim table As ListObject, area As Range, validated As Range, cell As Range
    Dim errorNumber As Long
    mTargetPreflightPassed = False
    If mutationMask <= 0 Or (CLng(mutationMask) And Not CLng(NxCellAllSupportedMask)) <> 0 Then NxRaiseContractError "Unsupported cell mutation mask"
    If target.FormatConditions.Count > 0 Then NxRaiseContractError "Conditional formatting is outside the cell journal contract"
    For Each table In target.Worksheet.ListObjects
        If Not Intersect(table.Range, target) Is Nothing Then NxRaiseContractError "ListObject is outside the cell journal contract"
    Next table
    For Each area In target.Areas
        Set validated = Nothing
        On Error Resume Next
        Set validated = area.SpecialCells(xlCellTypeAllValidation)
        errorNumber = Err.Number: Err.Clear
        On Error GoTo 0
        If errorNumber <> 0 And errorNumber <> 1004 Then NxRaiseContractError "Data validation probe failed"
        If Not validated Is Nothing Then
            ' Excel expands a single-cell SpecialCells search; confine it again.
            Set validated = Intersect(validated, area)
            If Not validated Is Nothing Then
                For Each cell In validated.Cells
                    If Not HasOnlyInputValidation(cell) Then NxRaiseContractError "Data validation is outside the cell journal contract"
                Next cell
            End If
        End If
    Next area
    mTargetPreflightPassed = True
End Sub

Private Sub SnapshotUniqueCell(ByVal cell As Range, ByVal mutationMask As NxCellMutationMask, ByVal captureLimit As Long)
    Dim key As String
    key = CellKey(cell)
    If mRecordKeys.Exists(key) Then Exit Sub
    If mRecords.Count >= captureLimit Then NxRaiseContractError "Cell journal hard cap or declared max_cells exceeded"
    SnapshotCell cell, mutationMask, key
    mRecordKeys.Add key, True
End Sub

Private Sub SnapshotCell(ByVal cell As Range, ByVal mutationMask As NxCellMutationMask, ByVal key As String)
    Dim record As New Collection
    Dim properties As New Collection
    Dim formulaValue As Variant
    Dim formulaMode As String
    ' Compare each freshly read property while capturing. Equal neighbouring
    ' cells need no duplicate arrays/collection entries; no Excel read is skipped.
    mMatchingPrevious = Not (mLastProperties Is Nothing)
    mMatchedPropertyCount = 0
    If (CLng(mutationMask) And CLng(NxCellValueMask)) <> 0 Then
        If cell.HasFormula Then
            formulaValue = ReadFormula2OrFallback(cell, formulaMode)
        Else
            formulaValue = cell.Value2
            formulaMode = "Value2"
        End If
    End If
    If (CLng(mutationMask) And CLng(NxCellNumberFormatMask)) <> 0 Then AddProperty properties, "NumberFormat", cell.NumberFormat
    If (CLng(mutationMask) And CLng(NxCellFontMask)) <> 0 Then SnapshotFont cell, properties
    If (CLng(mutationMask) And CLng(NxCellInteriorMask)) <> 0 Then SnapshotInterior cell, properties
    If (CLng(mutationMask) And CLng(NxCellAlignmentMask)) <> 0 Then SnapshotAlignment cell, properties
    If (CLng(mutationMask) And CLng(NxCellProtectionMask)) <> 0 Then SnapshotProtection cell, properties
    If (CLng(mutationMask) And CLng(NxCellBorderMask)) <> 0 Then SnapshotBorders cell, properties
    If (CLng(mutationMask) And CLng(NxCellMergeMask)) <> 0 Then SnapshotMerge cell, key
    record.Add cell
    record.Add cell.HasFormula
    record.Add formulaValue
    record.Add formulaMode
    record.Add ShareExactProperties(properties)
    mRecords.Add record
End Sub

Private Function ShareExactProperties(ByVal properties As Collection) As Collection
    Dim item As Variant, expected As Variant, key As String, valueText As String
    Dim candidate As Collection, index As Long
    If mMatchingPrevious Then
        If mMatchedPropertyCount = mLastProperties.Count Then
            Set ShareExactProperties = mLastProperties
            Exit Function
        End If
        ' A shorter vector (for example, a missing theme binding) is distinct.
        MaterializeMatchedProperties properties
    End If
    ' Immutable property vectors may be shared, but every actual value is compared
    ' before sharing. A string-key collision cannot merge different formats.
    Set ShareExactProperties = properties
    ' Adjacent cells commonly share formatting. Compare the complete vector first
    ' to avoid allocating the same dictionary key for every cell in a large area.
    If Not mLastProperties Is Nothing Then
        If ExactPropertyVectorsMatch(properties, mLastProperties) Then
            Set ShareExactProperties = mLastProperties
            Exit Function
        End If
    End If
    For Each item In properties
        If IsNull(item(1)) Then
            valueText = "null"
        ElseIf IsEmpty(item(1)) Then
            valueText = "empty"
        ElseIf IsError(item(1)) Then
            Exit Function
        Else
            valueText = CStr(item(1))
        End If
        key = key & CStr(Len(CStr(item(0)))) & ":" & CStr(item(0)) & "|" & CStr(VarType(item(1))) & ":" & CStr(Len(valueText)) & ":" & valueText
    Next item
    If mPropertySets.Exists(key) Then
        Set candidate = mPropertySets(key)
        If candidate.Count <> properties.Count Then Exit Function
        For index = 1 To properties.Count
            item = properties(index): expected = candidate(index)
            If StrComp(CStr(item(0)), CStr(expected(0)), vbBinaryCompare) <> 0 Then Exit Function
            If VarType(item(1)) <> VarType(expected(1)) Then Exit Function
            If Not VariantValuesMatch(item(1), expected(1)) Then Exit Function
        Next index
        Set ShareExactProperties = candidate
    ElseIf mPropertySets.Count < 4096 Then
        mPropertySets.Add key, properties
    End If
    Set mLastProperties = ShareExactProperties
    CachePreviousProperties
End Function

Private Sub CachePreviousProperties()
    Dim index As Long, item As Variant
    mPreviousCount = mLastProperties.Count
    If mPreviousCount = 0 Then Exit Sub
    ReDim mPreviousNames(1 To mPreviousCount)
    ReDim mPreviousTypes(1 To mPreviousCount)
    ReDim mPreviousValues(1 To mPreviousCount)
    For index = 1 To mPreviousCount
        item = mLastProperties(index)
        mPreviousNames(index) = CStr(item(0))
        mPreviousTypes(index) = VarType(item(1))
        mPreviousValues(index) = item(1)
    Next index
End Sub

Private Function ExactPropertyVectorsMatch(ByVal leftProperties As Collection, ByVal rightProperties As Collection) As Boolean
    Dim index As Long, item As Variant, expected As Variant
    If leftProperties.Count <> rightProperties.Count Then Exit Function
    For index = 1 To leftProperties.Count
        item = leftProperties(index): expected = rightProperties(index)
        If StrComp(CStr(item(0)), CStr(expected(0)), vbBinaryCompare) <> 0 Then Exit Function
        If VarType(item(1)) <> VarType(expected(1)) Then Exit Function
        If Not VariantValuesMatch(item(1), expected(1)) Then Exit Function
    Next index
    ExactPropertyVectorsMatch = True
End Function

Private Sub SnapshotMerge(ByVal cell As Range, ByVal cellKey As String)
    Dim mergeRecord As New Collection
    Dim mergeAddress As String
    Dim mergeKey As String
    If cell.MergeCells Then
        mergeAddress = cell.MergeArea.Address
        mergeKey = "A|" & cell.Worksheet.CodeName & "|" & mergeAddress
    Else
        mergeKey = "C|" & cellKey
    End If
    On Error Resume Next
    mergeRecord.Add cell
    mergeRecord.Add mergeAddress
    mMergeRecords.Add mergeRecord, mergeKey
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Sub

Private Sub SnapshotFont(ByVal cell As Range, ByVal properties As Collection)
    Dim fontState As Font
    Set fontState = cell.Font
    AddProperty properties, "Font.Name", fontState.Name
    AddProperty properties, "Font.Size", fontState.Size
    AddProperty properties, "Font.Bold", fontState.Bold
    AddProperty properties, "Font.Italic", fontState.Italic
    AddProperty properties, "Font.Underline", fontState.Underline
    AddProperty properties, "Font.Strikethrough", fontState.Strikethrough
    AddProperty properties, "Font.Color", fontState.Color
    AddProperty properties, "Font.ColorIndex", fontState.ColorIndex
    SnapshotTheme properties, "Font.ThemeColor", fontState
    AddProperty properties, "Font.TintAndShade", fontState.TintAndShade
End Sub

Private Sub SnapshotInterior(ByVal cell As Range, ByVal properties As Collection)
    Dim fillState As Interior
    Set fillState = cell.Interior
    AddProperty properties, "Interior.Pattern", fillState.Pattern
    AddProperty properties, "Interior.PatternColor", fillState.PatternColor
    AddProperty properties, "Interior.PatternColorIndex", fillState.PatternColorIndex
    AddProperty properties, "Interior.Color", fillState.Color
    AddProperty properties, "Interior.ColorIndex", fillState.ColorIndex
    SnapshotTheme properties, "Interior.ThemeColor", fillState
    AddProperty properties, "Interior.TintAndShade", fillState.TintAndShade
End Sub

Private Sub SnapshotAlignment(ByVal cell As Range, ByVal properties As Collection)
    AddProperty properties, "HorizontalAlignment", cell.HorizontalAlignment
    AddProperty properties, "VerticalAlignment", cell.VerticalAlignment
    AddProperty properties, "Orientation", cell.Orientation
    AddProperty properties, "AddIndent", cell.AddIndent
    AddProperty properties, "IndentLevel", cell.IndentLevel
    AddProperty properties, "ShrinkToFit", cell.ShrinkToFit
    AddProperty properties, "ReadingOrder", cell.ReadingOrder
    AddProperty properties, "WrapText", cell.WrapText
End Sub

Private Sub SnapshotProtection(ByVal cell As Range, ByVal properties As Collection)
    AddProperty properties, "Locked", cell.Locked
    AddProperty properties, "FormulaHidden", cell.FormulaHidden
End Sub

Private Sub SnapshotBorders(ByVal cell As Range, ByVal properties As Collection)
    Dim edge As Variant, borderState As Border, cellBorders As Borders
    Set cellBorders = cell.Borders
    For Each edge In mBorderEdges
        Set borderState = cellBorders(edge)
        AddProperty properties, mBorderPropertyNames(edge, 0), borderState.Weight
        AddProperty properties, mBorderPropertyNames(edge, 1), borderState.Color
        AddProperty properties, mBorderPropertyNames(edge, 2), borderState.ColorIndex
        SnapshotTheme properties, mBorderPropertyNames(edge, 3), borderState
        AddProperty properties, mBorderPropertyNames(edge, 4), borderState.TintAndShade
        ' Excel can reactivate an xlNone border when a dependent visual property is set.
        ' Apply LineStyle last so the recorded visible state remains authoritative.
        AddProperty properties, mBorderPropertyNames(edge, 5), borderState.LineStyle
    Next edge
End Sub

Private Sub SnapshotTheme(ByVal properties As Collection, ByVal propertyName As String, ByVal item As Object)
    Dim theme As Long
    ' RGB-only colors may not expose a theme. Restore valid theme bindings after
    ' ColorIndex, which otherwise replaces them with palette colors.
    On Error GoTo NoTheme
    theme = item.ThemeColor
    If theme > 0 Then AddProperty properties, propertyName, theme
    Exit Sub
NoTheme:
    Err.Clear
End Sub

Private Sub AddProperty(ByVal properties As Collection, ByVal propertyName As String, ByVal propertyValue As Variant)
    Dim nextIndex As Long
    If mMatchingPrevious Then
        If mMatchedPropertyCount < mPreviousCount Then
            nextIndex = mMatchedPropertyCount + 1
            If StrComp(propertyName, mPreviousNames(nextIndex), vbBinaryCompare) = 0 Then
                If VarType(propertyValue) = mPreviousTypes(nextIndex) Then
                    If VariantValuesMatch(propertyValue, mPreviousValues(nextIndex)) Then
                        mMatchedPropertyCount = mMatchedPropertyCount + 1
                        Exit Sub
                    End If
                End If
            End If
        End If
        MaterializeMatchedProperties properties
    End If
    ' One small array replaces a separate COM Collection for every property.
    properties.Add Array(propertyName, propertyValue)
End Sub

Private Sub MaterializeMatchedProperties(ByVal properties As Collection)
    Dim index As Long
    For index = 1 To mMatchedPropertyCount
        properties.Add mLastProperties(index)
    Next index
    mMatchingPrevious = False
End Sub

Private Sub RestoreRecord(ByVal record As Collection, ByRef failures As String)
    Dim cell As Range
    Dim properties As Collection
    Dim item As Variant
    Set cell = record.Item(1)
    If Len(CStr(record.Item(4))) > 0 Then RestoreFormula cell, record.Item(3), CStr(record.Item(4)), failures
    Set properties = record.Item(5)
    For Each item In properties
        RestoreProperty cell, CStr(item(0)), item(1), failures
    Next item
End Sub

Private Sub RestoreFormula(ByVal cell As Range, ByVal value As Variant, ByVal mode As String, ByRef failures As String)
    On Error GoTo Failed
    Select Case mode
        Case "Formula2": RestoreFormula2OrFallback cell, value, failures: Exit Sub
        Case "Formula": cell.Formula = value
        Case "Value2": cell.Value2 = value
        Case Else: NxRaiseContractError "Unknown recorded formula mode"
    End Select
    Exit Sub
Failed:
    AppendFailure failures, cell, mode, Err.Number, Err.Description
    Err.Clear
End Sub

Private Function ReadFormula2OrFallback(ByVal cell As Range, ByRef mode As String) As Variant
    ReadFormula2OrFallback = NxReadFormula2OrFallback(cell, mode)
End Function

Private Sub RestoreFormula2OrFallback(ByVal cell As Range, ByVal value As Variant, ByRef failures As String)
    Dim mode As String
    On Error GoTo Failed
    NxWriteFormulaByMode cell, value, "Formula2", mode
    Exit Sub
Failed:
    AppendFailure failures, cell, "Restore Formula2 failed", Err.Number, Err.Description
    Err.Clear
End Sub

Private Function FormulaMatches(ByVal cell As Range, ByVal expectedValue As Variant, ByVal mode As String) As Boolean
    On Error GoTo Different
    Select Case mode
        Case "Formula2": FormulaMatches = ErrorValuesMatch(NxReadFormula2OrFallback(cell, mode), expectedValue)
        Case "Formula": FormulaMatches = VariantValuesMatch(cell.Formula, expectedValue)
        Case "Value2": FormulaMatches = VariantValuesMatch(cell.Value2, expectedValue)
        Case Else: NxRaiseContractError "Unknown recorded formula mode"
    End Select
    Exit Function
Different:
    Err.Clear
End Function

Private Function PropertyMatches(ByVal cell As Range, ByVal propertyName As String, ByVal expectedValue As Variant) As Boolean
    Dim actual As Variant, parts() As String, edge As Long
    On Error GoTo Different
    Select Case propertyName
        Case "NumberFormat": actual = cell.NumberFormat
        Case "Font.Name": actual = cell.Font.Name
        Case "Font.Size": actual = cell.Font.Size
        Case "Font.Bold": actual = cell.Font.Bold
        Case "Font.Italic": actual = cell.Font.Italic
        Case "Font.Underline": actual = cell.Font.Underline
        Case "Font.Strikethrough": actual = cell.Font.Strikethrough
        Case "Font.Color": actual = cell.Font.Color
        Case "Font.ColorIndex": actual = cell.Font.ColorIndex
        Case "Font.ThemeColor": actual = cell.Font.ThemeColor
        Case "Font.TintAndShade": actual = cell.Font.TintAndShade
        Case "Interior.Pattern": actual = cell.Interior.Pattern
        Case "Interior.PatternColor": actual = cell.Interior.PatternColor
        Case "Interior.PatternColorIndex": actual = cell.Interior.PatternColorIndex
        Case "Interior.Color": actual = cell.Interior.Color
        Case "Interior.ColorIndex": actual = cell.Interior.ColorIndex
        Case "Interior.ThemeColor": actual = cell.Interior.ThemeColor
        Case "Interior.TintAndShade": actual = cell.Interior.TintAndShade
        Case "HorizontalAlignment": actual = cell.HorizontalAlignment
        Case "VerticalAlignment": actual = cell.VerticalAlignment
        Case "Orientation": actual = cell.Orientation
        Case "AddIndent": actual = cell.AddIndent
        Case "IndentLevel": actual = cell.IndentLevel
        Case "ShrinkToFit": actual = cell.ShrinkToFit
        Case "ReadingOrder": actual = cell.ReadingOrder
        Case "WrapText": actual = cell.WrapText
        Case "Locked": actual = cell.Locked
        Case "FormulaHidden": actual = cell.FormulaHidden
        Case Else
            parts = Split(propertyName, ".")
            If UBound(parts) = 2 And parts(0) = "Border" Then
                edge = CLng(parts(1))
                Select Case parts(2)
                    Case "LineStyle": actual = cell.Borders(edge).LineStyle
                    Case "Weight": actual = cell.Borders(edge).Weight
                    Case "Color": actual = cell.Borders(edge).Color
                    Case "ColorIndex": actual = cell.Borders(edge).ColorIndex
                    Case "ThemeColor": actual = cell.Borders(edge).ThemeColor
                    Case "TintAndShade": actual = cell.Borders(edge).TintAndShade
                    Case Else: Exit Function
                End Select
            Else
                PropertyMatches = True
                Exit Function
            End If
    End Select
    PropertyMatches = ErrorValuesMatch(actual, expectedValue)
    Exit Function
Different:
    Err.Clear
End Function

Private Function ErrorValuesMatch(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    If IsError(leftValue) Or IsError(rightValue) Then
        If IsError(leftValue) And IsError(rightValue) Then ErrorValuesMatch = (CLng(leftValue) = CLng(rightValue))
        Exit Function
    End If
    ErrorValuesMatch = VariantValuesMatch(leftValue, rightValue)
End Function

Private Sub RestoreProperty(ByVal cell As Range, ByVal propertyName As String, ByVal propertyValue As Variant, ByRef failures As String)
    Dim parts() As String
    Dim edge As Long
    On Error GoTo Failed
    Select Case propertyName
        Case "NumberFormat": cell.NumberFormat = propertyValue
        Case "Font.Name": cell.Font.Name = propertyValue
        Case "Font.Size": cell.Font.Size = propertyValue
        Case "Font.Bold": cell.Font.Bold = propertyValue
        Case "Font.Italic": cell.Font.Italic = propertyValue
        Case "Font.Underline": cell.Font.Underline = propertyValue
        Case "Font.Strikethrough": cell.Font.Strikethrough = propertyValue
        Case "Font.Color": cell.Font.Color = propertyValue
        Case "Font.ColorIndex": If cell.Font.ColorIndex <> propertyValue Then cell.Font.ColorIndex = propertyValue
        Case "Font.ThemeColor": cell.Font.ThemeColor = propertyValue
        Case "Font.TintAndShade": cell.Font.TintAndShade = propertyValue
        Case "Interior.Pattern": cell.Interior.Pattern = propertyValue
        Case "Interior.PatternColor": cell.Interior.PatternColor = propertyValue
        Case "Interior.PatternColorIndex": If cell.Interior.PatternColorIndex <> propertyValue Then cell.Interior.PatternColorIndex = propertyValue
        Case "Interior.Color": cell.Interior.Color = propertyValue
        Case "Interior.ColorIndex": If cell.Interior.ColorIndex <> propertyValue Then cell.Interior.ColorIndex = propertyValue
        Case "Interior.ThemeColor": cell.Interior.ThemeColor = propertyValue
        Case "Interior.TintAndShade": cell.Interior.TintAndShade = propertyValue
        Case "HorizontalAlignment": cell.HorizontalAlignment = propertyValue
        Case "VerticalAlignment": cell.VerticalAlignment = propertyValue
        Case "Orientation": cell.Orientation = propertyValue
        Case "AddIndent": cell.AddIndent = propertyValue
        Case "IndentLevel": cell.IndentLevel = propertyValue
        Case "ShrinkToFit": cell.ShrinkToFit = propertyValue
        Case "ReadingOrder": cell.ReadingOrder = propertyValue
        Case "WrapText": cell.WrapText = propertyValue
        Case "Locked": cell.Locked = propertyValue
        Case "FormulaHidden": cell.FormulaHidden = propertyValue
        Case Else
            parts = Split(propertyName, ".")
            If UBound(parts) <> 2 Or parts(0) <> "Border" Then NxRaiseContractError "Unsupported snapshot property"
            edge = CLng(parts(1))
            Select Case parts(2)
                Case "LineStyle": cell.Borders(edge).LineStyle = propertyValue
                Case "Weight": cell.Borders(edge).Weight = propertyValue
                Case "Color": cell.Borders(edge).Color = propertyValue
                Case "ColorIndex": If cell.Borders(edge).ColorIndex <> propertyValue Then cell.Borders(edge).ColorIndex = propertyValue
                Case "ThemeColor": cell.Borders(edge).ThemeColor = propertyValue
                Case "TintAndShade": cell.Borders(edge).TintAndShade = propertyValue
                Case Else: NxRaiseContractError "Unsupported border snapshot property"
            End Select
    End Select
    Exit Sub
Failed:
    AppendFailure failures, cell, propertyName, Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub RestoreMerges(ByRef failures As String)
    Dim mergeRecord As Variant
    For Each mergeRecord In mMergeRecords
        RestoreMerge mergeRecord.Item(1), CStr(mergeRecord.Item(2)), failures
    Next mergeRecord
End Sub

Private Sub RestoreMerge(ByVal cell As Range, ByVal mergeAddress As String, ByRef failures As String)
    On Error GoTo Failed
    If cell.MergeCells Then cell.MergeArea.UnMerge
    If Len(mergeAddress) > 0 Then cell.Worksheet.Range(mergeAddress).Merge
    Exit Sub
Failed:
    AppendFailure failures, cell, "Merge", Err.Number, Err.Description
    Err.Clear
End Sub

Private Function HasOnlyInputValidation(ByVal cell As Range) As Boolean
    Dim validationType As Long
    On Error GoTo OrdinaryCell
    validationType = cell.Validation.Type
    HasOnlyInputValidation = (validationType = xlValidateInputOnly)
    Exit Function
OrdinaryCell:
    ' Excel raises on a normal cell with no validation rule; that is safe to journal.
    Err.Clear
    HasOnlyInputValidation = True
End Function

Private Function VariantValuesMatch(ByVal leftValue As Variant, ByVal rightValue As Variant) As Boolean
    On Error GoTo Different
    If IsError(leftValue) Or IsError(rightValue) Then
        If Not (IsError(leftValue) And IsError(rightValue)) Then Exit Function
        VariantValuesMatch = (CLng(leftValue) = CLng(rightValue))
        Exit Function
    End If
    If IsEmpty(leftValue) Or IsEmpty(rightValue) Then
        VariantValuesMatch = (IsEmpty(leftValue) And IsEmpty(rightValue))
        Exit Function
    End If
    If IsNull(leftValue) Or IsNull(rightValue) Then
        VariantValuesMatch = (IsNull(leftValue) And IsNull(rightValue))
        Exit Function
    End If
    VariantValuesMatch = (leftValue = rightValue)
    Exit Function
Different:
    Err.Clear
End Function

Private Sub AppendFailure(ByRef failures As String, ByVal cell As Range, ByVal propertyName As String, ByVal errorNumber As Long, ByVal description As String)
    failures = NxJoinFailure(failures, CellTarget(cell) & "|" & propertyName & "|" & CStr(errorNumber) & "|" & description)
End Sub

Private Function CellKey(ByVal cell As Range) As String
    CellKey = cell.Worksheet.CodeName & "|" & NxInvariantUnsigned(CDbl(cell.Row)) & "," & NxInvariantUnsigned(CDbl(cell.Column))
End Function

Private Function CellTarget(ByVal cell As Range) As String
    CellTarget = cell.Worksheet.Parent.FullName & "|" & CellKey(cell)
End Function
