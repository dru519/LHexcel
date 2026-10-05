Attribute VB_Name = "NxRoleStyleEngine"
Option Explicit

Public Sub NxRoleStyleValidateTargetCore(ByVal target As Range, ByVal axisCode As String, ByVal relativeIndexes As String)
    Dim item As ListObject, cell As Range, mergeArea As Range
    If target Is Nothing Then NxRaiseContractError "role-style target is required"
    If target.Areas.Count <> 1 Then NxRaiseContractError "role-style target must be contiguous"
    If target.CountLarge > 100000 Then NxRaiseContractError "role-style target exceeds 100000 cells"
    If target.Worksheet.ProtectContents Then NxRaiseContractError "role-style target is protected"
    If target.Worksheet.Parent.ReadOnly Then NxRaiseContractError "role-style workbook is read-only"
    For Each item In target.Worksheet.ListObjects
        If Not Intersect(target, item.Range) Is Nothing Then NxRaiseContractError "role-style target conflicts with ListObjects"
    Next item
    ' Focus-cell highlighting is our own managed rule, not a user conflict.
    If target.FormatConditions.Count > NxFocusManagedConditionCount(target) Then NxRaiseContractError "role-style target conflicts with FormatConditions"
    For Each cell In target.Cells
        If cell.MergeCells Then
            Set mergeArea = cell.MergeArea
            If Intersect(target, mergeArea).Address(False, False) <> mergeArea.Address(False, False) Then
                NxRaiseContractError "partial merged area is not supported; fully contained merged area is required"
            End If
        End If
    Next cell
    NxRoleStyleNormalizeAxis axisCode
    If UCase$(axisCode) <> "ALL" Then NxRoleStyleParseRelativeIndexes relativeIndexes, IIf(UCase$(axisCode) = "ROWS", target.Rows.Count, target.Columns.Count)
End Sub

Public Sub NxRoleStyleApply(ByVal target As Range, ByVal roleCode As String, ByVal displayMode As String)
    Dim role As String, mode As String
    role = NxRoleStyleNormalizeRole(roleCode): mode = NxRoleStyleNormalizeMode(displayMode)
    Select Case role
        Case NX_ROLE_STYLE_TITLE: NxRoleStyleApplyTitle target, mode
        Case NX_ROLE_STYLE_SUBTITLE: NxRoleStyleApplySubtitle target, mode
        Case NX_ROLE_STYLE_TABLE_HEADER: NxRoleStyleApplyHeader target, mode
        Case NX_ROLE_STYLE_TABLE_BODY: NxRoleStyleApplyBody target, mode
        Case NX_ROLE_STYLE_EMPHASIS_CELL: NxRoleStyleApplyEmphasis target, mode
        Case NX_ROLE_STYLE_TOTAL_ROW: NxRoleStyleApplyTotal target, mode
    End Select
End Sub

Private Sub NxRoleStyleBase(ByVal target As Range, ByVal mode As String)
    With target.Font
        .Name = NxBodyFontName(): .Size = NxBodyFontSize(): .Bold = False: .Color = NxRoleStyleTokenColor("FONT_BODY", mode)
    End With
    target.Interior.Pattern = xlNone
    NxRoleStyleClearBorders target
End Sub

Private Sub NxRoleStyleApplyTitle(ByVal target As Range, ByVal mode As String)
    NxRoleStyleBase target, mode
    With target.Font: .Name = NxHeaderFontName(): .Size = 14#: .Bold = True: .Color = NxRoleStyleTokenColor("FONT", mode): End With
    target.HorizontalAlignment = xlLeft: target.VerticalAlignment = xlCenter
    NxRoleStyleSetEdge target, xlEdgeBottom, NxEmphasisLineWeight(), NxRoleStyleTokenColor("EMPHASIS", mode), xlContinuous
End Sub

Private Sub NxRoleStyleApplySubtitle(ByVal target As Range, ByVal mode As String)
    NxRoleStyleBase target, mode
    With target.Font: .Name = NxHeaderFontName(): .Size = 11#: .Bold = True: .Color = NxRoleStyleTokenColor("FONT", mode): End With
    target.Interior.Pattern = xlSolid: target.Interior.Color = NxRoleStyleTokenColor("FILL", mode)
    target.HorizontalAlignment = xlLeft: target.VerticalAlignment = xlCenter
    NxRoleStyleSetEdge target, xlEdgeBottom, NxBaseLineWeight(), NxRoleStyleTokenColor("EMPHASIS", mode), xlContinuous
End Sub

Private Sub NxRoleStyleApplyHeader(ByVal target As Range, ByVal mode As String)
    NxRoleStyleBase target, mode: target.Interior.Pattern = xlSolid: target.Interior.Color = NxRoleStyleTokenColor("FILL", mode)
    With target.Font: .Bold = True: .Color = NxRoleStyleTokenColor("FONT", mode): End With
    target.HorizontalAlignment = xlCenter: target.VerticalAlignment = xlCenter: target.WrapText = True
    NxRoleStyleSetOutline target, mode, NxBaseLineWeight()
End Sub

Private Sub NxRoleStyleApplyBody(ByVal target As Range, ByVal mode As String)
    Dim cell As Range, alignmentTarget As Range, value As Variant
    NxRoleStyleBase target, mode
    For Each cell In target.Cells
        Set alignmentTarget = Nothing
        If cell.MergeCells Then
            If cell.Address <> cell.MergeArea.Cells(1, 1).Address Then GoTo NextCell
            Set alignmentTarget = cell.MergeArea
        Else
            Set alignmentTarget = cell
        End If
        value = cell.Value2
        If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
            alignmentTarget.HorizontalAlignment = xlLeft
        ElseIf NxRoleStyleLooksLikeDate(cell, value) Then
            alignmentTarget.HorizontalAlignment = xlCenter
        ElseIf IsNumeric(value) Then
            alignmentTarget.HorizontalAlignment = xlRight
        Else
            alignmentTarget.HorizontalAlignment = xlLeft
        End If
        alignmentTarget.VerticalAlignment = xlCenter
NextCell:
    Next cell
End Sub

Private Sub NxRoleStyleApplyEmphasis(ByVal target As Range, ByVal mode As String)
    NxRoleStyleBase target, mode: target.Interior.Pattern = xlSolid: target.Interior.Color = NxRoleStyleTokenColor("EMPHASIS", mode)
    With target.Font: .Bold = True: .Color = RGB(255, 255, 255): End With
    NxRoleStyleSetOutline target, mode, NxBaseLineWeight()
End Sub

Private Sub NxRoleStyleApplyTotal(ByVal target As Range, ByVal mode As String)
    NxRoleStyleBase target, mode: target.Interior.Pattern = xlSolid: target.Interior.Color = NxRoleStyleTokenColor("TOTAL", mode)
    target.Font.Bold = True: target.HorizontalAlignment = xlRight: target.VerticalAlignment = xlCenter
    NxRoleStyleSetEdge target, xlEdgeTop, NxEmphasisLineWeight(), NxEmphasisLineColor(), xlContinuous
    NxRoleStyleSetEdge target, xlEdgeBottom, xlMedium, NxEmphasisLineColor(), xlDouble
End Sub

Private Function NxRoleStyleLooksLikeDate(ByVal cell As Range, ByVal value As Variant) As Boolean
    Dim formatText As String
    If Not IsDate(value) Then Exit Function
    formatText = LCase$(CStr(cell.NumberFormatLocal))
    NxRoleStyleLooksLikeDate = _
        (InStr(formatText, "y") > 0 Or InStr(formatText, "년") > 0) And _
        (InStr(formatText, "m") > 0 Or InStr(formatText, "월") > 0) And _
        (InStr(formatText, "d") > 0 Or InStr(formatText, "일") > 0)
End Function

Private Sub NxRoleStyleClearBorders(ByVal target As Range)
    Dim edge As Variant
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom, xlInsideHorizontal, xlInsideVertical)
        target.Borders(CLng(edge)).LineStyle = xlNone
    Next edge
End Sub

Private Sub NxRoleStyleSetOutline(ByVal target As Range, ByVal mode As String, ByVal weight As XlBorderWeight)
    Dim edge As Variant
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        NxRoleStyleSetEdge target, CLng(edge), weight, NxRoleStyleTokenColor("EMPHASIS", mode), xlContinuous
    Next edge
End Sub

Private Sub NxRoleStyleSetEdge(ByVal target As Range, ByVal edge As XlBordersIndex, ByVal weight As XlBorderWeight, ByVal colorValue As Long, ByVal lineStyle As XlLineStyle)
    With target.Borders(edge)
        .Weight = weight: .Color = colorValue: .LineStyle = lineStyle
    End With
End Sub
