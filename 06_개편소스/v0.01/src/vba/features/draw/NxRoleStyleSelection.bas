Attribute VB_Name = "NxRoleStyleSelection"
Option Explicit

Public Function NxRoleStyleNormalizeRole(ByVal roleCode As String) As String
    Dim value As String: value = UCase$(Trim$(roleCode))
    Select Case value
        Case NX_ROLE_STYLE_TITLE, NX_ROLE_STYLE_SUBTITLE, NX_ROLE_STYLE_TABLE_HEADER, NX_ROLE_STYLE_TABLE_BODY, NX_ROLE_STYLE_EMPHASIS_CELL, NX_ROLE_STYLE_TOTAL_ROW
            NxRoleStyleNormalizeRole = value
        Case Else: NxRaiseContractError "Unknown role-style role"
    End Select
End Function

Public Function NxRoleStyleNormalizeMode(ByVal displayMode As String) As String
    Dim value As String: value = UCase$(Trim$(displayMode))
    If value <> NX_ROLE_STYLE_MODE_MONO And value <> NX_ROLE_STYLE_MODE_COLOR Then NxRaiseContractError "Unknown role-style display mode"
    NxRoleStyleNormalizeMode = value
End Function

Public Function NxRoleStyleNormalizeAxis(ByVal axisCode As String) As String
    Dim value As String: value = UCase$(Trim$(axisCode))
    If value <> "ALL" And value <> "ROWS" And value <> "COLUMNS" Then NxRaiseContractError "Unknown role-style axis"
    NxRoleStyleNormalizeAxis = value
End Function

Public Function NxRoleStyleParseRelativeIndexes(ByVal textValue As String, ByVal axisLength As Long) As Collection
    ' seen.Exists prevents duplicate relative positions while retaining input order.
    Dim answer As New Collection, parts() As String, part As Variant, bounds() As String, firstIndex As Long, lastIndex As Long, i As Long
    If Len(Trim$(textValue)) = 0 Then NxRaiseContractError "relative indexes are required"
    parts = Split(Replace(Trim$(textValue), " ", vbNullString), ",")
    For Each part In parts
        If Len(CStr(part)) = 0 Then NxRaiseContractError "relative indexes must be integers"
        If InStr(1, CStr(part), "-", vbBinaryCompare) > 0 Then
            bounds = Split(CStr(part), "-")
            If UBound(bounds) <> 1 Then NxRaiseContractError "reversed ranges are invalid"
            firstIndex = NxRoleStyleParseInteger(bounds(0)): lastIndex = NxRoleStyleParseInteger(bounds(1))
            ' firstValue > lastValue is a reversed range and must be rejected.
            If firstIndex > lastIndex Then NxRaiseContractError "reversed ranges are invalid"
            For i = firstIndex To lastIndex: NxRoleStyleAddIndex answer, i, axisLength: Next i
        Else
            NxRoleStyleAddIndex answer, NxRoleStyleParseInteger(CStr(part)), axisLength
        End If
    Next part
    Set NxRoleStyleParseRelativeIndexes = answer
End Function

Private Function NxRoleStyleParseInteger(ByVal textValue As String) As Long
    If Len(textValue) = 0 Or textValue Like "*[!0-9]*" Then NxRaiseContractError "relative indexes must be integers"
    NxRoleStyleParseInteger = CLng(textValue)
End Function

Private Sub NxRoleStyleAddIndex(ByVal indexes As Collection, ByVal value As Long, ByVal axisLength As Long)
    If value < 1 Or value > axisLength Then NxRaiseContractError "relative index is outside the selected axis"
    On Error Resume Next: indexes.Add value, CStr(value): On Error GoTo 0
End Sub

Public Function NxRoleStyleResolveAffectedRange(ByVal target As Range, ByVal axisCode As String, ByVal relativeIndexes As String) As Range
    Dim answer As Range, index As Variant, indexes As Collection, part As Range
    Set answer = target
    If UCase$(axisCode) = "ALL" Then Set NxRoleStyleResolveAffectedRange = answer: Exit Function
    If UCase$(axisCode) = "ROWS" Then Set indexes = NxRoleStyleParseRelativeIndexes(relativeIndexes, target.Rows.Count) Else Set indexes = NxRoleStyleParseRelativeIndexes(relativeIndexes, target.Columns.Count)
    For Each index In indexes
        If UCase$(axisCode) = "ROWS" Then Set part = target.Rows(CLng(index)) Else Set part = target.Columns(CLng(index))
        If answer Is target Then Set answer = part Else Set answer = Application.Union(answer, part)
    Next index
    Set NxRoleStyleResolveAffectedRange = answer
End Function
