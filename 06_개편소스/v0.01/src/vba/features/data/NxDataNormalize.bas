Attribute VB_Name = "NxDataNormalize"
Option Explicit

Private Const NX_DATA_MAX_ITEMS As Long = 200000
Private Const NX_DATA_YIELD_INTERVAL As Long = 1024

Public Function NxDataKey(ByVal value As Variant, ByVal trimText As Boolean, ByVal caseSensitive As Boolean) As String
    If IsError(value) Then
        NxDataKey = "E:" & NxDataErrorCode(value)
    ElseIf IsEmpty(value) Then
        NxDataKey = "Z:"
    ElseIf VarType(value) = vbDate Then
        NxDataKey = "D:" & NxDataInvariantNumber(CDbl(value))
    ElseIf VarType(value) = vbString Then
        NxDataKey = "S:" & NxDataNormalizeText(CStr(value), trimText, caseSensitive)
    ElseIf VarType(value) = vbBoolean Then
        NxDataKey = "B:" & IIf(CBool(value), "1", "0")
    ElseIf IsNumeric(value) Then
        NxDataKey = "N:" & NxDataInvariantNumber(CDbl(value))
    Else
        NxRaiseContractError "Unsupported data value type"
    End If
End Function

Public Function NxDataBuildGroups( _
    ByVal values As Variant, _
    ByVal trimText As Boolean, _
    ByVal caseSensitive As Boolean, _
    Optional ByVal firstSourceRow As Long = 1 _
) As Collection
    Set NxDataBuildGroups = NxDataBuildGroupsCore(values, Empty, False, trimText, caseSensitive, firstSourceRow)
End Function

Public Function NxDataBuildGroupsAtRows( _
    ByVal values As Variant, _
    ByVal sourceRows As Variant, _
    ByVal trimText As Boolean, _
    ByVal caseSensitive As Boolean _
) As Collection
    Set NxDataBuildGroupsAtRows = NxDataBuildGroupsCore(values, sourceRows, True, trimText, caseSensitive, 1)
End Function

Private Function NxDataBuildGroupsCore( _
    ByVal values As Variant, _
    ByVal sourceRows As Variant, _
    ByVal hasSourceRows As Boolean, _
    ByVal trimText As Boolean, _
    ByVal caseSensitive As Boolean, _
    ByVal firstSourceRow As Long _
) As Collection
    Dim dimensions As Long
    Dim itemCount As Long
    Dim lowerRow As Long
    Dim lowerColumn As Long
    Dim itemIndex As Long
    Dim sourceRow As Long
    Dim value As Variant
    Dim key As String
    Dim priorSourceRow As Long
    Dim group As CNxDataGroup
    Dim groups As New Collection
    Dim byKey As Object

    If firstSourceRow < 1 Then NxRaiseContractError "First source row must be positive"
    dimensions = NxDataArrayDimensions(values)
    If dimensions <> 1 And dimensions <> 2 Then NxRaiseContractError "Data values must be a one-dimensional or single-column two-dimensional array"

    itemCount = NxDataArrayItemCount(values, dimensions, lowerRow, lowerColumn)
    If itemCount > NX_DATA_MAX_ITEMS Then NxRaiseContractError "Data input exceeds the 200,000 item limit"
    If hasSourceRows Then NxDataValidateSourceRows sourceRows, itemCount

    Set byKey = CreateObject("Scripting.Dictionary")
    byKey.CompareMode = vbBinaryCompare

    For itemIndex = 0 To itemCount - 1
        value = NxDataArrayItem(values, dimensions, lowerRow, lowerColumn, itemIndex)
        key = NxDataKey(value, trimText, caseSensitive)
        If hasSourceRows Then
            sourceRow = CLng(sourceRows(LBound(sourceRows, 1) + itemIndex))
            If itemIndex > 0 And sourceRow <= priorSourceRow Then NxRaiseContractError "Source rows must keep strict ascending order"
            priorSourceRow = sourceRow
        Else
            sourceRow = firstSourceRow + itemIndex
        End If

        If byKey.Exists(key) Then
            Set group = byKey.Item(key)
            group.AddOccurrence sourceRow, value
        Else
            Set group = New CNxDataGroup
            group.Configure key, value, sourceRow
            byKey.Add key, group
            groups.Add group
        End If

        If (itemIndex + 1) Mod NX_DATA_YIELD_INTERVAL = 0 Then DoEvents
    Next itemIndex

    For Each group In groups
        group.Seal
    Next group
    Set NxDataBuildGroupsCore = groups
End Function

Public Function NxDataGroupFingerprint( _
    ByVal values As Variant, _
    Optional ByVal trimText As Boolean = False, _
    Optional ByVal caseSensitive As Boolean = False, _
    Optional ByVal firstSourceRow As Long = 1 _
) As String
    Dim groups As Collection
    Dim group As CNxDataGroup
    Dim result As String

    Set groups = NxDataBuildGroups(values, trimText, caseSensitive, firstSourceRow)
    For Each group In groups
        If Len(result) > 0 Then result = result & ";"
        result = result & group.Key & "|" & CStr(group.Count) & "|" & CStr(group.FirstSourceRow)
    Next group
    NxDataGroupFingerprint = result
End Function

Private Function NxDataNormalizeText(ByVal value As String, ByVal trimText As Boolean, ByVal caseSensitive As Boolean) As String
    If trimText Then value = Trim$(value)
    If Not caseSensitive Then value = LCase$(value)
    NxDataNormalizeText = value
End Function

Private Function NxDataInvariantNumber(ByVal value As Double) As String
    Dim rendered As String
    rendered = Format$(value, "0.###############")
    If Application.International(xlDecimalSeparator) <> "." Then
        rendered = Replace(rendered, Application.International(xlDecimalSeparator), ".")
    End If
    NxDataInvariantNumber = rendered
End Function

Private Function NxDataErrorCode(ByVal value As Variant) As String
    Dim rendered As String
    Dim digits As String
    Dim index As Long
    Dim character As String

    rendered = CStr(value)
    For index = 1 To Len(rendered)
        character = Mid$(rendered, index, 1)
        If character >= "0" And character <= "9" Then digits = digits & character
    Next index
    If Len(digits) = 0 Then NxRaiseContractError "Unsupported worksheet error value"
    NxDataErrorCode = digits
End Function

Private Function NxDataArrayDimensions(ByRef values As Variant) As Long
    Dim ignored As Long
    If Not IsArray(values) Then Exit Function

    On Error Resume Next
    ignored = LBound(values, 1)
    If Err.Number <> 0 Then Err.Clear: On Error GoTo 0: Exit Function
    ignored = LBound(values, 2)
    If Err.Number = 0 Then
        NxDataArrayDimensions = 2
        ignored = LBound(values, 3)
        If Err.Number = 0 Then NxDataArrayDimensions = 3
        Err.Clear
    Else
        Err.Clear
        NxDataArrayDimensions = 1
    End If
    On Error GoTo 0
End Function

Private Function NxDataArrayItemCount( _
    ByRef values As Variant, _
    ByVal dimensions As Long, _
    ByRef lowerRow As Long, _
    ByRef lowerColumn As Long _
) As Long
    Dim upperRow As Long
    Dim upperColumn As Long

    On Error GoTo InvalidArray
    lowerRow = LBound(values, 1)
    upperRow = UBound(values, 1)
    If dimensions = 1 Then
        NxDataArrayItemCount = upperRow - lowerRow + 1
        Exit Function
    End If

    lowerColumn = LBound(values, 2)
    upperColumn = UBound(values, 2)
    On Error GoTo 0
    If upperColumn <> lowerColumn Then NxRaiseContractError "Data values must contain exactly one column"
    NxDataArrayItemCount = upperRow - lowerRow + 1
    Exit Function

InvalidArray:
    NxRaiseContractError "Data values array is empty or invalid"
End Function

Private Function NxDataArrayItem( _
    ByRef values As Variant, _
    ByVal dimensions As Long, _
    ByVal lowerRow As Long, _
    ByVal lowerColumn As Long, _
    ByVal itemIndex As Long _
) As Variant
    If dimensions = 1 Then
        NxDataArrayItem = values(lowerRow + itemIndex)
    Else
        NxDataArrayItem = values(lowerRow + itemIndex, lowerColumn)
    End If
End Function

Private Sub NxDataValidateSourceRows(ByRef sourceRows As Variant, ByVal expectedCount As Long)
    Dim lowerRow As Long
    Dim ignoredColumn As Long
    If NxDataArrayDimensions(sourceRows) <> 1 Then NxRaiseContractError "Source rows must be a one-dimensional array"
    If NxDataArrayItemCount(sourceRows, 1, lowerRow, ignoredColumn) <> expectedCount Then NxRaiseContractError "Source row count does not match data values"
End Sub
