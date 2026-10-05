Attribute VB_Name = "NxFormulaResolver"
Option Explicit

Public Function NxFormulaResolution(ByVal formulaText As String, ByVal sheetMap As Object) As String
    Dim pattern As Object, matches As Object, key As String
    If Len(Trim$(formulaText)) = 0 Then NxFormulaResolution = "literal": Exit Function
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Global = False
    pattern.IgnoreCase = True
    pattern.Pattern = "^='\[([^\]]+)\]([^']+)'!\$?[A-Z]{1,3}\$?[0-9]+$"
    Set matches = pattern.Execute(formulaText)
    If matches.Count = 1 Then
        key = matches(0).SubMatches(0) & "|" & matches(0).SubMatches(1)
        If Not sheetMap Is Nothing Then If sheetMap.Exists(key) Then NxFormulaResolution = "relink": Exit Function
    End If
    If InStr(1, formulaText, "[", vbTextCompare) > 0 Or InStr(1, formulaText, "!", vbTextCompare) > 0 Or _
       InStr(1, formulaText, "[", vbTextCompare) > 0 Then
        NxFormulaResolution = "decision"
    ElseIf InStr(1, formulaText, "=", vbBinaryCompare) = 1 Then
        NxFormulaResolution = "approved"
    Else
        NxFormulaResolution = "literal"
    End If
End Function

Public Function NxSafeSheetName(ByVal candidate As String, ByVal usedNames As Object) As String
    Dim base As String, index As Long, nextName As String
    base = Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(candidate, ":", "_"), "\\", "_"), "/", "_"), "?", "_"), "*", "_"), "[", "_"), "]", "_"), "'", "_")
    base = Left$(Trim$(base), 31)
    If Len(base) = 0 Then base = "Sheet"
    nextName = base: index = 1
    Do While Not usedNames Is Nothing And usedNames.Exists(nextName)
        index = index + 1
        nextName = Left$(base, 31 - Len(CStr(index)) - 3) & " (" & CStr(index) & ")"
    Loop
    If Not usedNames Is Nothing Then usedNames.Add nextName, True
    NxSafeSheetName = nextName
End Function
