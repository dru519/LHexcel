VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxSideEffectDispatcher"
Option Explicit

Private mCompletedNonCellEffects As Collection
Private mCompletedShapes As Collection
Private mCompletedShapeEffects As Collection
Private mCompletedDirectoryPaths As Collection
Private mSelectionAuthority As Collection

Private Sub Class_Initialize()
    Set mCompletedNonCellEffects = New Collection
    Set mCompletedShapes = New Collection
    Set mCompletedShapeEffects = New Collection
    Set mCompletedDirectoryPaths = New Collection
End Sub

Friend Sub Configure(ByVal selectionAuthority As Collection)
    If selectionAuthority Is Nothing Then NxRaiseContractError "Dispatcher selection authority is required"
    If Not mSelectionAuthority Is Nothing Then NxRaiseContractError "Dispatcher can be configured only once"
    Set mSelectionAuthority = selectionAuthority
End Sub

Public Function Dispatch(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext) As String
    Dim effect As Variant
    RequireReady approvedPlan, context
    If approvedPlan Is Nothing Or context Is Nothing Then NxRaiseContractError "Dispatcher requires plan and context"
    For Each effect In approvedPlan.SideEffects
        Select Case effect.Kind
            Case NxCellMutation
                ' Cell changes remain command-owned and can only use ApplyCellValue.
            Case NxShapeInsert
                If approvedPlan.FeatureId = "NX-DRAW-INSERT-PICTURE" Then
                    DispatchPictureInsertBatch effect, approvedPlan
                Else
                    ' Shape insertion remains command-owned and can only use ApplyImageAsset.
                End If
            Case NxShapeMutation
                ' Existing shape mutation remains command-owned; rollback is typed by the feature command.
            Case NxBinaryPackageWrite, NxWorksheetCreate, NxColumnInsert
                ' Template transactions remain command-owned after exact plan validation.
            Case NxDirectoryBatchCreate
                DispatchDirectoryBatch effect
            Case NxFileCreate
                DispatchFileCreate effect
            Case NxFileReplace
                DispatchFileReplace effect
            Case NxClipboardWrite
                DispatchClipboardWrite effect
            Case NxExternalProcess
                DispatchExternalProcess effect
            Case NxBrowserOpen
                DispatchBrowserOpen effect
            Case Else
                NxRaiseContractError "Unsupported non-cell execution is blocked before command"
        End Select
    Next effect
End Function

Private Sub DispatchPictureInsertBatch(ByVal effect As CNxSideEffect, ByVal approvedPlan As CNxPlan)
    If approvedPlan.FeatureId <> "NX-DRAW-INSERT-PICTURE" Then NxRaiseContractError "Picture insert batch route is feature-specific"
    If effect Is Nothing Or Not effect.IsSealed Then NxRaiseContractError "Picture insert batch effect must be sealed"
    If effect.Kind <> NxShapeInsert Or effect.Owner <> "draw" Then NxRaiseContractError "Picture insert batch requires a draw shape effect"
    If effect.DeclaredItems < 1 Or effect.DeclaredItems > 500 Then NxRaiseContractError "Picture insert batch must contain 1-500 items"
    If Len(effect.Target) = 0 Or effect.CompensatorMetadata <> "delete inserted shapes" Then NxRaiseContractError "Picture insert batch requires a rollback target"
End Sub

Public Function ExecuteApproved(ByVal command As INxFeatureCommand, ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    If command Is Nothing Or context Is Nothing Or approvedPlan Is Nothing Then NxRaiseContractError "Typed dispatcher requires command, context, and approved plan"
    Set ExecuteApproved = command.Execute(context, approvedPlan)
End Function

Public Sub ApplyCellValue(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, ByVal value As Variant)
    Dim effect As Variant
    Dim target As Range
    RequireReady approvedPlan, context
    For Each effect In approvedPlan.SideEffects
        If effect.Kind = NxCellMutation Then
            If (CLng(effect.MutationMask) And CLng(NxCellValueMask)) = 0 Then NxRaiseContractError "Typed dispatcher requires NxCellValueMask"
            Set target = ResolveExactCellTarget(effect, context)
            target.Value2 = value
            Exit Sub
        End If
    Next effect
    NxRaiseContractError "Typed dispatcher rejected undeclared cell mutation"
End Sub

Public Sub ApplyCellValueVector(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByVal addresses As Variant, ByVal values As Variant)

    Dim effect As CNxSideEffect
    Dim target As Range
    Dim candidate As Range
    Dim targetArea As Range
    Dim seen As Object
    Dim addressIndex As Long
    Dim valueIndex As Long
    Dim addressCount As Long
    Dim valueCount As Long
    Dim targetCount As Double
    Dim addressText As String
    Dim canonicalAddress As String

    RequireReady approvedPlan, context
    Set effect = FindEffect(approvedPlan, NxCellMutation)
    If effect Is Nothing Then NxRaiseContractError "Typed dispatcher rejected undeclared visible-cell mutation"
    RequireMutationMask effect, NxCellValueMask
    Set target = ResolveExactCellTarget(effect, context)
    addressCount = OneDimensionalArrayCount(addresses, "Visible-cell address vector")
    valueCount = OneDimensionalArrayCount(values, "Visible-cell value vector")
    For Each targetArea In target.Areas
        targetCount = targetCount + CDbl(targetArea.Cells.CountLarge)
    Next targetArea
    If CDbl(addressCount) <> targetCount Then NxRaiseContractError "Visible-cell address vector does not cover the exact approved target"
    If valueCount <> 1 And valueCount <> addressCount Then NxRaiseContractError "Visible-cell value vector count mismatch"

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For addressIndex = LBound(addresses) To UBound(addresses)
        addressText = CStr(addresses(addressIndex))
        If Len(addressText) = 0 Or InStr(addressText, "!") > 0 Or InStr(addressText, "[") > 0 Or _
            InStr(addressText, "]") > 0 Or InStr(addressText, ":") > 0 Or InStr(addressText, ",") > 0 Then _
            NxRaiseContractError "Visible-cell address is not a single local A1 cell"
        Set candidate = target.Worksheet.Range(addressText)
        If candidate.Areas.Count <> 1 Or candidate.Cells.CountLarge <> 1 Then NxRaiseContractError "Visible-cell address must resolve to one cell"
        If Application.Intersect(candidate, target) Is Nothing Then NxRaiseContractError "Visible-cell address is outside the exact approved target"
        canonicalAddress = candidate.Address(True, True, xlA1, False)
        If seen.Exists(canonicalAddress) Then NxRaiseContractError "Visible-cell address vector contains a duplicate"
        seen.Add canonicalAddress, True
    Next addressIndex

    valueIndex = LBound(values)
    For addressIndex = LBound(addresses) To UBound(addresses)
        Set candidate = target.Worksheet.Range(CStr(addresses(addressIndex)))
        WriteLiteralCellValue candidate, values(valueIndex)
        If valueCount > 1 Then valueIndex = valueIndex + 1
    Next addressIndex
End Sub

Private Function OneDimensionalArrayCount(ByVal values As Variant, ByVal label As String) As Long
    On Error GoTo InvalidArray
    If Not IsArray(values) Then NxRaiseContractError label & " must be an array"
    OneDimensionalArrayCount = UBound(values) - LBound(values) + 1
    If OneDimensionalArrayCount <= 0 Then NxRaiseContractError label & " cannot be empty"
    Exit Function
InvalidArray:
    Err.Clear
    NxRaiseContractError label & " must be a one-dimensional array"
End Function

Private Sub WriteLiteralCellValue(ByVal target As Range, ByVal value As Variant)
    Dim textValue As String
    If IsError(value) Or IsObject(value) Then NxRaiseContractError "Visible-cell values must be clipboard text"
    If VarType(value) <> vbString Then
        target.Value2 = value
        Exit Sub
    End If
    textValue = CStr(value)
    If Len(textValue) > 0 And InStr(1, "=+-@", Left$(textValue, 1), vbBinaryCompare) > 0 Then
        target.Formula = "'" & textValue
    Else
        target.Value2 = textValue
    End If
End Sub

Public Sub ApplyCellValues(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByVal rowOffset As Long, ByVal columnOffset As Long, ByVal rowCount As Long, ByVal columnCount As Long, _
    ByVal matrix As Variant)
    Dim effect As CNxSideEffect
    Dim target As Range
    RequireReady approvedPlan, context
    Set effect = FindEffect(approvedPlan, NxCellMutation)
    If effect Is Nothing Then NxRaiseContractError "Typed dispatcher rejected undeclared cell mutation"
    If Not context.HasExactTarget(effect.Target) Then NxRaiseContractError "Typed dispatcher rejected a target outside the approved context"
    RequireMutationMask effect, NxCellValueMask
    Set target = context.InternalRelativeRange(mSelectionAuthority, rowOffset, columnOffset, rowCount, columnCount)
    target.Value2 = matrix
End Sub

Public Sub ApplyCellFormulas(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByVal rowOffset As Long, ByVal columnOffset As Long, ByVal rowCount As Long, ByVal columnCount As Long, _
    ByVal matrix As Variant)
    Dim effect As CNxSideEffect
    Dim target As Range
    RequireReady approvedPlan, context
    Set effect = FindEffect(approvedPlan, NxCellMutation)
    If effect Is Nothing Then NxRaiseContractError "Typed dispatcher rejected undeclared cell mutation"
    If Not context.HasExactTarget(effect.Target) Then NxRaiseContractError "Typed dispatcher rejected a target outside the approved context"
    RequireMutationMask effect, NxCellValueMask
    Set target = context.InternalRelativeRange(mSelectionAuthority, rowOffset, columnOffset, rowCount, columnCount)
    ApplyFormulaMatrix target, matrix
End Sub

Public Sub ApplyCellNumberFormats(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByVal rowOffset As Long, ByVal columnOffset As Long, ByVal rowCount As Long, ByVal columnCount As Long, _
    ByVal matrix As Variant)
    Dim effect As CNxSideEffect
    Dim target As Range
    RequireReady approvedPlan, context
    Set effect = FindEffect(approvedPlan, NxCellMutation)
    If effect Is Nothing Then NxRaiseContractError "Typed dispatcher rejected undeclared cell mutation"
    If Not context.HasExactTarget(effect.Target) Then NxRaiseContractError "Typed dispatcher rejected a target outside the approved context"
    RequireMutationMask effect, NxCellNumberFormatMask
    Set target = context.InternalRelativeRange(mSelectionAuthority, rowOffset, columnOffset, rowCount, columnCount)
    target.NumberFormat = matrix
End Sub

Public Sub ApplyImageAsset(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext, _
    ByVal rowOffset As Long, ByVal columnOffset As Long, ByVal rowCount As Long, ByVal columnCount As Long, _
    ByVal assetPath As String)

    Dim effect As CNxSideEffect
    Dim target As Range
    Dim shape As Shape
    RequireReady approvedPlan, context
    If Len(Dir$(assetPath)) = 0 Then NxRaiseContractError "Approved image asset is missing"
    Set effect = FindEffect(approvedPlan, NxShapeInsert)
    If effect Is Nothing Or Not context.HasExactTarget(effect.Target) Then NxRaiseContractError "Typed dispatcher rejected undeclared shape insertion"
    Set target = context.InternalRelativeRange(mSelectionAuthority, rowOffset, columnOffset, rowCount, columnCount)
    Set shape = target.Worksheet.Shapes.AddPicture(assetPath, False, True, target.Left, target.Top, -1, -1)
    mCompletedShapes.Add shape
    mCompletedShapeEffects.Add effect
End Sub

Private Function ResolveExactCellTarget(ByVal effect As CNxSideEffect, ByVal context As CNxExecutionContext) As Range
    If effect Is Nothing Or context Is Nothing Then NxRaiseContractError "Exact cell target is required"
    If Not context.HasExactTarget(effect.Target) Then NxRaiseContractError "Typed dispatcher rejected a target outside the approved context"
    Set ResolveExactCellTarget = context.InternalSelectionRange(mSelectionAuthority)
End Function

Public Function Recover() As String
    Recover = RemoveCompletedShapes()
    Recover = NxJoinFailure(Recover, RemoveCompletedDirectories())
    Recover = NxJoinFailure(Recover, CompensateCompleted())
End Function

Public Function CompensateCompleted() As String
    Dim index As Long
    For index = mCompletedNonCellEffects.Count To 1 Step -1
        CompensateEffect mCompletedNonCellEffects.Item(index), CompensateCompleted
    Next index
End Function

Public Property Get CompletedDescriptions() As Collection
    Dim result As New Collection
    Dim effect As Variant
    For Each effect In mCompletedNonCellEffects
        result.Add NxSideEffectWireName(effect.Kind) & ":" & effect.Target
    Next effect
    For Each effect In mCompletedShapeEffects
        result.Add NxSideEffectWireName(effect.Kind) & ":" & effect.Target
    Next effect
    Set CompletedDescriptions = result
End Property

Private Function FindEffect(ByVal approvedPlan As CNxPlan, ByVal kind As NxSideEffectKind) As CNxSideEffect
    Dim effect As CNxSideEffect
    For Each effect In approvedPlan.SideEffects
        If effect.Kind = kind Then Set FindEffect = effect: Exit Function
    Next effect
End Function

Private Sub RequireMutationMask(ByVal effect As CNxSideEffect, ByVal requiredMask As NxCellMutationMask)
    If (CLng(effect.MutationMask) And CLng(requiredMask)) = 0 Then NxRaiseContractError "Typed dispatcher mutation mask mismatch"
End Sub

Private Sub ApplyFormulaMatrix(ByVal target As Range, ByVal matrix As Variant)
    Dim mode As String
    NxWriteFormulaByMode target, matrix, "Formula2", mode
End Sub

Private Function RemoveCompletedShapes() As String
    Dim index As Long
    On Error GoTo Failed
    For index = mCompletedShapes.Count To 1 Step -1
        mCompletedShapes.Item(index).Delete
    Next index
    Exit Function
Failed:
    RemoveCompletedShapes = "shape_insert|compensate|" & CStr(Err.Number) & "|" & Err.Description
    Err.Clear
End Function

Private Sub DispatchFileCreate(ByVal effect As CNxSideEffect)
    RequireTextAdapter effect, NxFileCreate, False
    If Len(Dir$(effect.Target)) > 0 Then NxRaiseContractError "File-create target already exists"
    WriteTextFile effect.Target, effect.Payload
    mCompletedNonCellEffects.Add effect
End Sub

Private Sub DispatchFileReplace(ByVal effect As CNxSideEffect)
    RequireTextAdapter effect, NxFileReplace, True
    If Len(Dir$(effect.Target)) = 0 Then NxRaiseContractError "File-replace target is absent"
    If ReadTextFile(effect.Target) <> effect.CapturedPriorText Then NxRaiseContractError "File-replace prior-state drift"
    WriteTextFile effect.Target, effect.Payload
    mCompletedNonCellEffects.Add effect
End Sub

Private Sub DispatchClipboardWrite(ByVal effect As CNxSideEffect)
    RequireTextAdapter effect, NxClipboardWrite, True
    RunClipboardWrite effect.Payload
    mCompletedNonCellEffects.Add effect
End Sub

Private Sub RunClipboardWrite(ByVal value As String)
    Application.Run "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!NxClipboardWriteUnicode", value
End Sub

Private Sub RunClipboardRestore(ByVal value As String)
    Application.Run "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!NxClipboardRestoreUnicode", value
End Sub

Private Sub DispatchExternalProcess(ByVal effect As CNxSideEffect)
    If effect.RollbackStrategy <> NxRollbackManualRecovery Then NxRaiseContractError "External process requires explicit manual recovery"
    If effect.PayloadKind <> NxPayloadCommandLine Or Len(Trim$(effect.Payload)) = 0 Then NxRaiseContractError "External process payload is not typed"
    Call Shell(effect.Payload, vbNormalNoFocus)
    mCompletedNonCellEffects.Add effect
End Sub

Private Sub DispatchBrowserOpen(ByVal effect As CNxSideEffect)
    If effect.RollbackStrategy <> NxRollbackManualRecovery Then NxRaiseContractError "Browser open requires explicit manual recovery"
    If effect.PayloadKind <> NxPayloadUri Or Len(Trim$(effect.Payload)) = 0 Then NxRaiseContractError "Browser payload is not typed"
    ThisWorkbook.FollowHyperlink effect.Payload
    mCompletedNonCellEffects.Add effect
End Sub

Private Sub DispatchDirectoryBatch(ByVal effect As CNxSideEffect)
    Dim item As Variant
    Dim path As String
    Dim basePath As String
    Dim parentPath As String
    Dim fso As Object
    Dim seen As Object
    If effect Is Nothing Or Not effect.IsSealed Then NxRaiseContractError "Directory batch effect must be sealed"
    If effect.Kind <> NxDirectoryBatchCreate Or effect.DeclaredItems <= 0 Or effect.DeclaredItems > 500 Then NxRaiseContractError "Directory batch effect is invalid"
    If effect.DirectoryItems.Count <> effect.DeclaredItems Then NxRaiseContractError "Directory batch declared item count mismatch"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1
    basePath = NormalizeDirectoryPath(effect.Target)
    If Not fso.FolderExists(basePath) Then NxRaiseContractError "Directory batch base folder is absent"
    For Each item In effect.DirectoryItems
        path = NormalizeDirectoryPath(CStr(item))
        If Not IsDirectoryChildPath(basePath, path) Then NxRaiseContractError "Directory batch item is outside the approved base folder: " & path
        If seen.Exists(path) Then NxRaiseContractError "Directory batch contains a duplicate path: " & path
        seen.Add path, True
        If fso.FileExists(path) Then NxRaiseContractError "Directory batch target is an existing file: " & path
        If fso.FolderExists(path) Then NxRaiseContractError "Directory batch target already exists: " & path
        parentPath = fso.GetParentFolderName(path)
        If Len(parentPath) = 0 Or Not fso.FolderExists(parentPath) Then NxRaiseContractError "Directory batch parent is absent or out of declaration order: " & path
        MkDir path
        mCompletedDirectoryPaths.Add path
    Next item
    mCompletedNonCellEffects.Add effect
End Sub

Private Function RemoveCompletedDirectories() As String
    Dim index As Long
    Dim path As String
    Dim fso As Object
    Dim failures As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    For index = mCompletedDirectoryPaths.Count To 1 Step -1
        path = CStr(mCompletedDirectoryPaths.Item(index))
        On Error Resume Next
        Err.Clear
        If fso.FolderExists(path) Then
            If fso.GetFolder(path).Files.Count = 0 And fso.GetFolder(path).SubFolders.Count = 0 Then RmDir path
        End If
        If Err.Number <> 0 Then failures = NxJoinFailure(failures, "directory_batch_create|compensate|" & path & "|" & CStr(Err.Number) & "|" & Err.Description)
        Err.Clear
        On Error GoTo 0
    Next index
    RemoveCompletedDirectories = failures
End Function

Private Function NormalizeDirectoryPath(ByVal value As String) As String
    Dim normalized As String
    normalized = Replace(Trim$(value), "/", "\")
    Do While Len(normalized) > 3 And Right$(normalized, 1) = "\"
        normalized = Left$(normalized, Len(normalized) - 1)
    Loop
    NormalizeDirectoryPath = normalized
End Function

Private Function IsDirectoryChildPath(ByVal basePath As String, ByVal candidatePath As String) As Boolean
    Dim prefix As String
    prefix = LCase$(NormalizeDirectoryPath(basePath)) & "\"
    IsDirectoryChildPath = (Left$(LCase$(candidatePath), Len(prefix)) = prefix) And LCase$(candidatePath) <> LCase$(basePath)
End Function

Private Sub RequireTextAdapter(ByVal effect As CNxSideEffect, ByVal expectedKind As NxSideEffectKind, ByVal requirePriorText As Boolean)
    If effect Is Nothing Or Not effect.IsSealed Then NxRaiseContractError "Non-cell effect must be sealed"
    If effect.Kind <> expectedKind Or effect.PayloadKind <> NxPayloadText Then NxRaiseContractError "Text adapter kind or payload mismatch"
    If effect.RollbackStrategy <> NxRollbackTypedCompensator Then NxRaiseContractError "Non-cell effect requires a typed compensator"
    If requirePriorText <> effect.HasCapturedPriorText Then NxRaiseContractError "Text adapter prior-state contract mismatch"
End Sub

Private Sub CompensateEffect(ByVal effect As CNxSideEffect, ByRef failures As String)
    On Error GoTo Failed
    Select Case effect.Kind
        Case NxFileCreate
            If Len(Dir$(effect.Target)) > 0 Then Kill effect.Target
        Case NxFileReplace
            WriteTextFile effect.Target, effect.CapturedPriorText
        Case NxClipboardWrite
            RunClipboardRestore effect.CapturedPriorText
        Case NxExternalProcess, NxBrowserOpen
            NxRaiseContractError "Manual recovery is required: " & effect.CompensatorMetadata
        Case NxShapeInsert
            NxRaiseContractError "Shape insertion is recovered through the typed shape journal"
        Case NxShapeMutation
            ' FIT-Picture restores the existing shape through NxPictureLayout.
        Case NxDirectoryBatchCreate
            ' Directory batch paths are removed by the typed directory journal.
        Case Else
            NxRaiseContractError "Unsupported compensator kind"
    End Select
    Exit Sub
Failed:
    failures = NxJoinFailure(failures, NxSideEffectWireName(effect.Kind) & "|compensate|" & CStr(Err.Number) & "|" & Err.Description)
    Err.Clear
End Sub

Private Sub RequireReady(ByVal approvedPlan As CNxPlan, ByVal context As CNxExecutionContext)
    If approvedPlan Is Nothing Or context Is Nothing Then NxRaiseContractError "Dispatcher requires plan and context"
    If mSelectionAuthority Is Nothing Then NxRaiseContractError "Dispatcher selection authority was not configured"
    If Not approvedPlan.IsSealed Then NxRaiseContractError "Dispatcher requires a sealed plan"
End Sub

Private Sub WriteTextFile(ByVal path As String, ByVal value As String)
    Dim handle As Integer
    Dim opened As Boolean
    Dim errorNumber As Long, errorDescription As String
    On Error GoTo Failed
    handle = FreeFile
    Open path For Output Access Write As #handle
    opened = True
    Print #handle, value;
    Close #handle
    Exit Sub
Failed:
    errorNumber = Err.Number: errorDescription = Err.Description
    If opened Then Close #handle
    Err.Raise errorNumber, "CNxSideEffectDispatcher", errorDescription
End Sub

Private Function ReadTextFile(ByVal path As String) As String
    Dim handle As Integer
    Dim opened As Boolean
    Dim errorNumber As Long, errorDescription As String
    On Error GoTo Failed
    handle = FreeFile
    Open path For Binary Access Read As #handle
    opened = True
    If LOF(handle) > 0 Then ReadTextFile = Input$(LOF(handle), #handle)
    Close #handle
    Exit Function
Failed:
    errorNumber = Err.Number: errorDescription = Err.Description
    If opened Then Close #handle
    Err.Raise errorNumber, "CNxSideEffectDispatcher", errorDescription
End Function
