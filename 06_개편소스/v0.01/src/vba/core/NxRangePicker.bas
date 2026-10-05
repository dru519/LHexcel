Attribute VB_Name = "NxRangePicker"
Option Explicit

Public Function NxRangeSessionFromSelection() As CNxRangeSelectionSession
    Dim selected As Range
    Dim session As New CNxRangeSelectionSession
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "적용할 셀 범위를 먼저 선택하세요."
    Set selected = Application.Selection
    session.UpdateFromRange selected
    Set NxRangeSessionFromSelection = session
End Function

Public Sub NxRangeSessionUpdateFromText(ByVal session As CNxRangeSelectionSession, ByVal addressText As String)
    Dim selected As Range
    If session Is Nothing Then NxRaiseContractError "범위 선택 세션이 필요합니다."
    addressText = Trim$(addressText)
    If Len(addressText) = 0 Then NxRaiseContractError "적용할 셀 범위를 입력하세요."
    On Error Resume Next
    Set selected = Application.Range(addressText)
    If selected Is Nothing Then Set selected = Application.ActiveSheet.Range(addressText)
    On Error GoTo 0
    If selected Is Nothing Then NxRaiseContractError "입력한 셀 범위를 확인하세요."
    session.UpdateFromRange selected
End Sub

Public Function NxPickRange(ByVal form As Object, ByVal session As CNxRangeSelectionSession) As Boolean
    Dim selected As Range
    Dim previousAddress As String
    Dim failureNumber As Long
    Dim failureDescription As String
    If form Is Nothing Then NxRaiseContractError "범위 선택을 호출한 창이 필요합니다."
    If session Is Nothing Then NxRaiseContractError "범위 선택 세션이 필요합니다."
    If session.IsConfigured Then previousAddress = session.RangeAddress

    form.Hide
    On Error Resume Next
    Set selected = Application.InputBox(Prompt:="적용할 셀 범위를 드래그해 선택하세요.", _
        Title:="내엑셀 - 범위 선택", Default:=previousAddress, Type:=8)
    failureNumber = Err.Number
    failureDescription = Err.Description
    Err.Clear
    On Error GoTo RestoreFailed
    form.Show vbModeless
    If failureNumber <> 0 And selected Is Nothing Then Exit Function
    If selected Is Nothing Then Exit Function
    session.UpdateFromRange selected
    NxPickRange = True
    Exit Function

RestoreFailed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    Err.Clear
    On Error Resume Next
    form.Show vbModeless
    On Error GoTo 0
    Err.Raise failureNumber, "NxRangePicker", failureDescription
End Function
