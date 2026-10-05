Attribute VB_Name = "NxFormulaInterop"
Option Explicit

Public Function NxReadFormula2OrFallback(ByVal target As Object, ByRef mode As String) As Variant
    Dim number As Long, description As String
    If target Is Nothing Then NxRaiseContractError "Formula target is required"
    On Error GoTo Fallback
    NxReadFormula2OrFallback = CallByName(target, "Formula2", VbGet)
    mode = "Formula2"
    Exit Function
Fallback:
    number = Err.Number: description = Err.Description
    If number <> 438 And number <> 1004 Then Err.Raise number, "NxFormulaInterop", description
    Err.Clear
    On Error GoTo Failed
    NxReadFormula2OrFallback = CallByName(target, "Formula", VbGet)
    mode = "Formula"
    Exit Function
Failed:
    Err.Raise Err.Number, "NxFormulaInterop", "Formula source read failed: " & Err.Description
End Function

Public Sub NxWriteFormulaByMode(ByVal target As Object, ByVal value As Variant, ByVal expectedMode As String, ByRef actualMode As String)
    If target Is Nothing Then NxRaiseContractError "Formula target is required"
    Select Case expectedMode
        Case "Formula2": CallByName target, "Formula2", VbLet, value: actualMode = "Formula2"
        Case "Formula": CallByName target, "Formula", VbLet, value: actualMode = "Formula"
        Case Else: NxRaiseContractError "Formula write mode is outside the closed contract"
    End Select
End Sub
