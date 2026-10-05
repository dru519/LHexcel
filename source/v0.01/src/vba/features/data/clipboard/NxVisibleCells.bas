Attribute VB_Name = "NxVisibleCells"
Option Explicit

Public Sub NxVisibleCellsCopyFromRibbon()
    Dim result As CNxResult
    Set result = NxVisibleCellsCopy(NxVisibleCellsSelectedRange())
    RequireRibbonSuccess result, "보이는 셀만 복사"
End Sub

Public Sub NxVisibleCellsPasteValuesFromRibbon()
    Dim result As CNxResult
    Set result = NxVisibleCellsPasteValues(NxVisibleCellsSelectedRange())
    If Not result Is Nothing Then
        If result.Outcome = NxCancelled Then Exit Sub
    End If
    RequireRibbonSuccess result, "보이는 셀에 값 붙여넣기"
End Sub

Public Function NxVisibleCellsCopy(ByVal source As Range) As CNxResult
    Dim commandObject As New CNxVisibleCellsCommand
    Dim command As INxFeatureCommand
    Dim payload As String
    Dim priorText As String
    Dim visibleCount As Long

    Set source = NxVisibleCellsRequireSource(source, NX_FEATURE_DATA_COPY_VISIBLE)
    visibleCount = NxVisibleCellsCount(source)
    payload = NxVisibleCellsCopyText(source)
    priorText = NxClipboardReadUnicode()
    source.Select
    commandObject.ConfigureCopy payload, priorText, visibleCount
    Set command = commandObject
    Set NxVisibleCellsCopy = NxDataRunCommand(command, "숨겨진 행과 열을 제외한 표시값을 클립보드에 복사합니다.")
End Function

Public Function NxVisibleCellsPasteValues(ByVal source As Range) As CNxResult
    Dim original As Range
    Dim visibleTarget As Range
    Dim commandObject As New CNxVisibleCellsCommand
    Dim command As INxFeatureCommand
    Dim result As CNxResult
    Dim addresses As Variant
    Dim values As Variant
    Dim clipboardText As String
    Dim visibleCount As Long
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    On Error GoTo Failed

    Set original = NxVisibleCellsRequireSource(source, NX_FEATURE_DATA_PASTE_VISIBLE_VALUES)
    NxVisibleCellsValidatePasteTarget original
    visibleCount = NxVisibleCellsCount(original)
    clipboardText = NxClipboardReadUnicode()
    If Len(clipboardText) = 0 Then NxRaiseContractError "클립보드에 붙여넣을 Unicode 텍스트가 없습니다."
    values = NxVisibleCellsFlattenClipboardValues(clipboardText)
    If NxVisibleCellsVariantCount(values) <> 1 And NxVisibleCellsVariantCount(values) <> visibleCount Then _
        NxRaiseContractError "클립보드 값 개수와 보이는 셀 개수가 다릅니다."
    addresses = NxVisibleCellsAddresses(original)
    Set visibleTarget = NxVisibleCellsTarget(original)
    visibleTarget.Select
    commandObject.ConfigurePaste addresses, values, visibleCount
    Set command = commandObject
    Set result = NxDataRunCommand(command, "숨겨진 행과 열은 건드리지 않고 보이는 셀에 값만 붙여넣습니다.")
    NxVisibleCellsRestoreSelection original
    Set NxVisibleCellsPasteValues = result
    Exit Function

Failed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Clear
    On Error Resume Next
    NxVisibleCellsRestoreSelection original
    On Error GoTo 0
    If errorNumber = 0 Then errorNumber = NX_CONTRACT_ERROR
    Err.Raise errorNumber, errorSource, errorDescription
End Function

Public Function NxVisibleCellsSelectedRange() As Range
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "셀 범위를 먼저 선택하세요."
    Set NxVisibleCellsSelectedRange = NxVisibleCellsRequireSource(Application.Selection, NX_FEATURE_DATA_COPY_VISIBLE)
End Function

Public Function NxVisibleCellsCopyText(ByVal source As Range) As String
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim rowText As String
    Dim rowsText As String
    Dim hasVisibleCell As Boolean
    Dim cell As Range

    Set source = NxVisibleCellsRequireSource(source, NX_FEATURE_DATA_COPY_VISIBLE)
    If NxVisibleCellsCount(source) <= 0 Then NxRaiseContractError "복사할 보이는 셀이 없습니다."
    For rowIndex = 1 To source.Rows.Count
        rowText = vbNullString
        hasVisibleCell = False
        For columnIndex = 1 To source.Columns.Count
            Set cell = source.Cells(rowIndex, columnIndex)
            If NxVisibleCellsIsVisible(cell) Then
                If hasVisibleCell Then rowText = rowText & vbTab
                rowText = rowText & CStr(cell.Text)
                hasVisibleCell = True
            End If
        Next columnIndex
        If hasVisibleCell Then
            If Len(rowsText) > 0 Then rowsText = rowsText & vbCrLf
            rowsText = rowsText & rowText
        End If
    Next rowIndex
    NxVisibleCellsCopyText = rowsText
End Function

Public Function NxVisibleCellsFlattenClipboardValues(ByVal clipboardText As String) As Variant
    clipboardText = Replace(clipboardText, vbCrLf, vbLf)
    clipboardText = Replace(clipboardText, vbCr, vbLf)
    Do While Len(clipboardText) > 0 And Right$(clipboardText, 1) = vbLf
        clipboardText = Left$(clipboardText, Len(clipboardText) - 1)
    Loop
    If Len(clipboardText) = 0 Then NxRaiseContractError "클립보드에 붙여넣을 값이 없습니다."
    NxVisibleCellsFlattenClipboardValues = Split(Replace(clipboardText, vbTab, vbLf), vbLf)
End Function

Public Function NxVisibleCellsAddresses(ByVal source As Range) As Variant
    Dim result As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim itemIndex As Long
    Dim visibleCount As Long
    Dim cell As Range

    visibleCount = NxVisibleCellsCount(source)
    ReDim result(0 To visibleCount - 1)
    For rowIndex = 1 To source.Rows.Count
        For columnIndex = 1 To source.Columns.Count
            Set cell = source.Cells(rowIndex, columnIndex)
            If NxVisibleCellsIsVisible(cell) Then
                result(itemIndex) = cell.Address(True, True, xlA1, False)
                itemIndex = itemIndex + 1
            End If
        Next columnIndex
    Next rowIndex
    NxVisibleCellsAddresses = result
End Function

Public Function NxVisibleCellsCount(ByVal source As Range) As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim count As Double
    Dim policy As New CNxLimitsPolicy
    Dim cell As Range
    If source Is Nothing Then NxRaiseContractError "보이는 셀 범위가 필요합니다."
    For rowIndex = 1 To source.Rows.Count
        For columnIndex = 1 To source.Columns.Count
            Set cell = source.Cells(rowIndex, columnIndex)
            If NxVisibleCellsIsVisible(cell) Then count = count + 1#
        Next columnIndex
    Next rowIndex
    If count <= 0# Then NxRaiseContractError "선택 범위에 보이는 셀이 없습니다."
    If count > policy.MaximumCellsFor(NX_FEATURE_DATA_PASTE_VISIBLE_VALUES) Then NxRaiseContractError "보이는 셀 수가 기능 한도를 초과합니다."
    NxVisibleCellsCount = CLng(count)
End Function

Public Function NxVisibleCellsTarget(ByVal source As Range) As Range
    Dim target As Range
    Dim expectedCount As Long
    Dim actualCount As Double
    Dim area As Range
    expectedCount = NxVisibleCellsCount(source)
    On Error Resume Next
    Set target = source.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If target Is Nothing Then NxRaiseContractError "선택 범위에 보이는 셀이 없습니다."
    For Each area In target.Areas
        actualCount = actualCount + CDbl(area.Cells.CountLarge)
    Next area
    If actualCount <> CDbl(expectedCount) Then NxRaiseContractError "Excel 보이는 셀 판정과 내엑셀 판정이 일치하지 않습니다."
    Set NxVisibleCellsTarget = target
End Function

Public Sub NxVisibleCellsValidatePasteTarget(ByVal source As Range)
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim cell As Range
    For rowIndex = 1 To source.Rows.Count
        For columnIndex = 1 To source.Columns.Count
            Set cell = source.Cells(rowIndex, columnIndex)
            If NxVisibleCellsIsVisible(cell) Then
                If cell.MergeCells Then NxRaiseContractError "병합 셀이 포함된 범위에는 보이는 셀 값 붙여넣기를 실행할 수 없습니다."
            End If
        Next columnIndex
    Next rowIndex
End Sub

Public Sub NxVisibleCellsRestoreSelection(ByVal original As Range)
    If original Is Nothing Then Exit Sub
    original.Parent.Parent.Activate
    original.Parent.Activate
    original.Select
End Sub

Public Function NxVisibleCellsVariantCount(ByVal values As Variant) As Long
    On Error GoTo InvalidArray
    If Not IsArray(values) Then NxRaiseContractError "Visible-cell values require a one-dimensional array"
    NxVisibleCellsVariantCount = UBound(values) - LBound(values) + 1
    If NxVisibleCellsVariantCount <= 0 Then NxRaiseContractError "Visible-cell values cannot be empty"
    Exit Function
InvalidArray:
    Err.Clear
    NxRaiseContractError "Visible-cell values require a one-dimensional array"
End Function

Private Function NxVisibleCellsRequireSource(ByVal source As Range, ByVal featureId As String) As Range
    Dim policy As New CNxLimitsPolicy
    If source Is Nothing Then NxRaiseContractError "연속된 셀 범위를 선택하세요."
    If source.Areas.Count <> 1 Then NxRaiseContractError "하나의 연속된 범위만 선택하세요."
    If TypeName(source.Parent) <> "Worksheet" Then NxRaiseContractError "일반 Excel 워크시트 범위가 필요합니다."
    If source.Parent.Parent Is ThisWorkbook Then NxRaiseContractError "일반 Excel 통합문서의 범위를 선택하세요."
    If source.Cells.CountLarge > policy.MaximumCellsFor(featureId) Then NxRaiseContractError "선택 셀 수가 기능 한도를 초과합니다."
    Set NxVisibleCellsRequireSource = source
End Function

Private Function NxVisibleCellsIsVisible(ByVal cell As Range) As Boolean
    NxVisibleCellsIsVisible = Not CBool(cell.EntireRow.Hidden) And Not CBool(cell.EntireColumn.Hidden)
End Function

Private Sub RequireRibbonSuccess(ByVal result As CNxResult, ByVal actionName As String)
    If result Is Nothing Then NxRaiseContractError actionName & " 결과가 없습니다."
    If result.Outcome <> NxSuccess Then NxRaiseContractError actionName & " 작업이 완료되지 않았습니다."
End Sub
