Attribute VB_Name = "NxDataInput"
Option Explicit

Public Function NxDataValueIsBlank(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If IsEmpty(value) Or IsNull(value) Then NxDataValueIsBlank = True: Exit Function
    If VarType(value) = vbString Then NxDataValueIsBlank = (Len(Trim$(value)) = 0)
End Function

Public Function NxDataContentRange(ByVal source As Range) As Range
    Dim first As Range, last As Range
    If source Is Nothing Then NxRaiseContractError "변환할 범위를 선택하세요."
    If source.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위를 선택하세요."
    ' Find ignores empty/formatted-only tails; it never uses the inflated UsedRange.
    Set first = source.Find(What:="*", After:=source.Cells(source.Rows.Count, source.Columns.Count), _
        LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    If first Is Nothing Then NxRaiseContractError "변환할 값이 없습니다. 빈 셀은 제외합니다."
    Set last = source.Find(What:="*", After:=source.Cells(1, 1), LookIn:=xlValues, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    ' Preserve selected columns: a blank edge column still defines the insertion boundary.
    Set NxDataContentRange = source.Worksheet.Cells(first.Row, source.Column).Resize( _
        last.Row - first.Row + 1, source.Columns.Count)
End Function
