Attribute VB_Name = "NxModelMergeController"
Option Explicit

Private Const NX_MODEL_MERGE_MAX_CELLS As Long = 10000

Private mJournal As CNxModelMergeJournal
Private mForm As FNxModelMerge

Public Sub NxModelMergeOpen()
    On Error GoTo Failed
    If mForm Is Nothing Then Set mForm = New FNxModelMerge
    mForm.BindSelection
    If Not mForm.Visible Then mForm.Show vbModeless
    Exit Sub
Failed:
    ShowModelMergeError Err.Description
End Sub

Public Sub NxModelMergeFormClosed(ByVal view As FNxModelMerge)
    If mForm Is Nothing Then Exit Sub
    If view Is Nothing Then Exit Sub
    If mForm Is view Then Set mForm = Nothing
End Sub

Public Function NxModelMergeCanRestore(ByVal target As Range) As Boolean
    On Error GoTo NotAvailable
    ValidateTarget target
    NxModelMergeCanRestore = Journal().HasSnapshot(target)
    Exit Function
NotAvailable:
    Err.Clear
End Function

Public Function NxModelMergeCanApply(ByVal target As Range) As Boolean
    On Error GoTo NotAvailable
    ValidateTarget target
    NxModelMergeCanApply = True
    Exit Function
NotAvailable:
    Err.Clear
End Function

Public Sub NxModelMergeApply(ByVal target As Range, ByVal alignmentCode As String, ByVal fillAdjacent As Boolean)
    Dim originalState As New CNxModelMergeSnapshot
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim rollbackDescription As String

    ValidateTarget target
    alignmentCode = NormalizeAlignmentCode(alignmentCode)
    If Journal().HasSnapshot(target) Then NxModelMergeRestore target
    If Not fillAdjacent Then ValidateSingleSourcePerRow target
    originalState.Capture target

    On Error GoTo ApplyFailed
    ApplyModelMerge target, alignmentCode, fillAdjacent
    Journal().PutSnapshot originalState
    Exit Sub

ApplyFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Err.Clear
    On Error Resume Next
    originalState.Restore
    rollbackDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseTransactionFailure failureNumber, failureDescription, rollbackDescription
End Sub

Public Sub NxModelMergeRestore(ByVal target As Range)
    Dim originalState As CNxModelMergeSnapshot
    Dim currentState As New CNxModelMergeSnapshot
    Dim failureNumber As Long
    Dim failureDescription As String
    Dim rollbackDescription As String

    ValidateTarget target
    Set originalState = Journal().SnapshotFor(target)
    If originalState Is Nothing Then NxRaiseContractError "이 범위에 복구할 모형병합 원본이 없습니다."
    currentState.Capture target

    On Error GoTo RestoreFailed
    originalState.Restore
    Journal().RemoveSnapshot target
    Exit Sub

RestoreFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Err.Clear
    On Error Resume Next
    currentState.Restore
    rollbackDescription = Err.Description
    Err.Clear
    On Error GoTo 0
    RaiseTransactionFailure failureNumber, failureDescription, rollbackDescription
End Sub

Private Function Journal() As CNxModelMergeJournal
    If mJournal Is Nothing Then Set mJournal = New CNxModelMergeJournal
    Set Journal = mJournal
End Function

Private Sub ValidateTarget(ByVal target As Range)
    Dim mergeState As Variant
    Dim cell As Range
    If target Is Nothing Then NxRaiseContractError "모형병합을 적용할 범위를 선택하세요."
    If target.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 선택하세요."
    If target.Worksheet.Parent Is ThisWorkbook Then NxRaiseContractError "내엑셀 추가기능 파일에는 적용할 수 없습니다."
    If target.Worksheet.Parent.ReadOnly Then NxRaiseContractError "읽기 전용 통합문서에는 적용할 수 없습니다."
    If target.Worksheet.ProtectContents Or target.Worksheet.ProtectDrawingObjects Or target.Worksheet.ProtectScenarios Then _
        NxRaiseContractError "보호된 워크시트에는 적용할 수 없습니다."
    If target.Rows.Count = target.Worksheet.Rows.Count Then NxRaiseContractError "전체 열 범위에는 적용할 수 없습니다."
    If target.Columns.Count = target.Worksheet.Columns.Count Then NxRaiseContractError "전체 행 범위에는 적용할 수 없습니다."
    If target.Cells.CountLarge < 2 Then NxRaiseContractError "두 셀 이상을 선택하세요."
    If target.Cells.CountLarge > NX_MODEL_MERGE_MAX_CELLS Then NxRaiseContractError "한 번에 10,000셀까지 적용할 수 있습니다."

    mergeState = target.MergeCells
    If IsNull(mergeState) Then NxRaiseContractError "병합 셀이 포함된 범위에는 적용할 수 없습니다."
    If CBool(mergeState) Then NxRaiseContractError "병합 셀이 포함된 범위에는 적용할 수 없습니다."
    For Each cell In target.Cells
        If cell.HasArray Then NxRaiseContractError "배열 수식이 포함된 범위에는 적용할 수 없습니다."
    Next cell
End Sub

Private Sub ValidateSingleSourcePerRow(ByVal target As Range)
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim cell As Range
    Dim populatedCount As Long
    For rowIndex = 1 To target.Rows.Count
        populatedCount = 0
        For columnIndex = 1 To target.Columns.Count
            Set cell = target.Cells(rowIndex, columnIndex)
            If CBool(cell.HasFormula) Or Not IsEmpty(cell.Value2) Then populatedCount = populatedCount + 1
        Next columnIndex
        If populatedCount > 1 Then _
            NxRaiseContractError "한 행에 값이 둘 이상 있습니다. 같은 값 입력을 선택하거나 빈 셀만 포함해 선택하세요."
    Next rowIndex
End Sub

Private Sub ApplyModelMerge(ByVal target As Range, ByVal alignmentCode As String, ByVal fillAdjacent As Boolean)
    Dim rowIndex As Long
    Dim rowRange As Range
    Dim sourceCell As Range

    For rowIndex = 1 To target.Rows.Count
        Set rowRange = target.Rows(rowIndex)
        Set sourceCell = FindSourceCell(rowRange)
        If fillAdjacent Then
            ApplyRepeatingValueModelMerge rowRange, sourceCell, alignmentCode
        Else
            ApplyPreservingSourceModelMerge rowRange, sourceCell, alignmentCode
        End If
    Next rowIndex
End Sub

Private Sub ApplyPreservingSourceModelMerge(ByVal rowRange As Range, ByVal sourceCell As Range, ByVal alignmentCode As String)
    ' The default path is a visual simulation only: never relocate or erase a
    ' user value/formula. Excel only supports centre-across-selection for a
    ' value in the first cell of the selected row.
    If alignmentCode = "CENTER" Then
        ValidateCenterAcrossSource rowRange, sourceCell
        rowRange.HorizontalAlignment = xlCenterAcrossSelection
    Else
        rowRange.HorizontalAlignment = AlignmentValue(alignmentCode)
    End If
    rowRange.ShrinkToFit = False
End Sub

Private Sub ApplyRepeatingValueModelMerge(ByVal rowRange As Range, ByVal sourceCell As Range, ByVal alignmentCode As String)
    Dim columnIndex As Long
    Dim anchorColumn As Long
    Dim anchorCell As Range
    Dim cell As Range
    Dim sourceNumberFormat As Variant

    If sourceCell Is Nothing Then
        rowRange.HorizontalAlignment = AlignmentValue(alignmentCode)
        rowRange.ShrinkToFit = False
        Exit Sub
    End If

    sourceNumberFormat = sourceCell.NumberFormat
    anchorColumn = VisibleAnchorColumn(rowRange.Columns.Count, alignmentCode)
    Set anchorCell = rowRange.Cells(1, anchorColumn)
    For columnIndex = 1 To rowRange.Columns.Count
        Set cell = rowRange.Cells(1, columnIndex)
        CopyCellPayload sourceCell, cell
        If columnIndex <> anchorColumn Then cell.NumberFormat = ";;;"
    Next columnIndex

    anchorCell.NumberFormat = sourceNumberFormat
    rowRange.HorizontalAlignment = AlignmentValue(alignmentCode)
    rowRange.ShrinkToFit = False
    anchorCell.ShrinkToFit = True
End Sub

Private Sub ValidateCenterAcrossSource(ByVal rowRange As Range, ByVal sourceCell As Range)
    If sourceCell Is Nothing Then Exit Sub
    If sourceCell.Column <> rowRange.Cells(1, 1).Column Then _
        NxRaiseContractError "가운데 모형병합은 각 행의 선택 범위 첫 셀에 값 또는 수식이 있어야 합니다. 값과 수식은 이동하지 않습니다."
End Sub

Private Function FindSourceCell(ByVal rowRange As Range) As Range
    Dim candidate As Range
    For Each candidate In rowRange.Cells
        If CBool(candidate.HasFormula) Or Not IsEmpty(candidate.Value2) Then
            Set FindSourceCell = candidate
            Exit Function
        End If
    Next candidate
End Function

Private Function VisibleAnchorColumn(ByVal columnCount As Long, ByVal alignmentCode As String) As Long
    Select Case alignmentCode
        Case "LEFT": VisibleAnchorColumn = 1
        Case "CENTER": VisibleAnchorColumn = (columnCount + 1) \ 2
        Case "RIGHT": VisibleAnchorColumn = columnCount
        Case Else: NxRaiseContractError "정렬 방법을 선택하세요."
    End Select
End Function

Private Function AlignmentValue(ByVal alignmentCode As String) As Long
    Select Case alignmentCode
        Case "LEFT": AlignmentValue = xlLeft
        Case "CENTER": AlignmentValue = xlCenter
        Case "RIGHT": AlignmentValue = xlRight
        Case Else: NxRaiseContractError "정렬 방법을 선택하세요."
    End Select
End Function

Private Sub CopyCellPayload(ByVal sourceCell As Range, ByVal targetCell As Range)
    If CBool(sourceCell.HasFormula) Then
        On Error Resume Next
        Err.Clear
        targetCell.Formula2 = sourceCell.Formula2
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        Err.Clear
        On Error GoTo CopyFailed
        targetCell.Formula = sourceCell.Formula
        Exit Sub
    End If

    targetCell.Value2 = sourceCell.Value2
    Exit Sub

CopyFailed:
    Err.Raise Err.Number, "NxModelMergeController.CopyCellPayload", Err.Description
End Sub

Private Function NormalizeAlignmentCode(ByVal alignmentCode As String) As String
    NormalizeAlignmentCode = UCase$(Trim$(alignmentCode))
    Select Case NormalizeAlignmentCode
        Case "LEFT", "CENTER", "RIGHT"
        Case Else: NxRaiseContractError "정렬 방법을 선택하세요."
    End Select
End Function

Private Sub RaiseTransactionFailure(ByVal failureNumber As Long, ByVal failureDescription As String, ByVal rollbackDescription As String)
    If failureNumber = 0 Then failureNumber = vbObjectError + 2750
    If Len(rollbackDescription) > 0 Then failureDescription = failureDescription & vbCrLf & "원상복구 실패: " & rollbackDescription
    Err.Raise failureNumber, "NxModelMergeController", failureDescription
End Sub

Private Sub ShowModelMergeError(ByVal detail As String)
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 모형병합"
End Sub
