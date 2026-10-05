Attribute VB_Name = "NxFocusRules"
Option Explicit

' These identifiers cover both legacy cleanup and the current internal-profile
' conditional-format backend. Only rules and names bearing these identifiers
' are owned and removed by v0.01.
Public Const NX_FOCUS_RULE_V001_ROW As String = "NX_FOCUS_RULE_V001_ROW"
Public Const NX_FOCUS_RULE_V001_COLUMN As String = "NX_FOCUS_RULE_V001_COLUMN"
Public Const NX_FOCUS_RULE_V002_ROW As String = "NX_FOCUS_RULE_V002_ROW"
Public Const NX_FOCUS_RULE_V002_COLUMN As String = "NX_FOCUS_RULE_V002_COLUMN"
Private Const NX_FOCUS_NAME_V002_TOP As String = "NX_FOCUS_TOP_V2"
Private Const NX_FOCUS_NAME_V002_BOTTOM As String = "NX_FOCUS_BOTTOM_V2"
Private Const NX_FOCUS_NAME_V002_LEFT As String = "NX_FOCUS_LEFT_V2"
Private Const NX_FOCUS_NAME_V002_RIGHT As String = "NX_FOCUS_RIGHT_V2"
Private Const NX_FOCUS_ROW_FORMULA As String = "=AND(N(""NX_FOCUS_RULE_V001_ROW"")=0,ROW()=CELL(""row""),COLUMN()<>CELL(""col""))"
Private Const NX_FOCUS_COLUMN_FORMULA As String = "=AND(N(""NX_FOCUS_RULE_V001_COLUMN"")=0,COLUMN()=CELL(""col""),ROW()<>CELL(""row""))"

Public Function NxFocusDeleteManagedRules(ByVal sheet As Worksheet) As Boolean
    On Error GoTo Failed
    Dim conditions As FormatConditions
    Dim index As Long
    Dim rule As Object

    If sheet Is Nothing Then Exit Function
    Set conditions = sheet.Cells.FormatConditions
    For index = conditions.Count To 1 Step -1
        Set rule = conditions(index)
        If NxFocusIsManagedRule(rule) Then rule.Delete
    Next index
    NxFocusDeleteSheetBridgeNames sheet
    NxFocusDeleteManagedRules = (NxFocusManagedRuleCount(sheet) = 0 And NxFocusManagedNameCount(sheet) = 0)
    Exit Function
Failed:
    NxFocusDeleteManagedRules = False
End Function

Public Function NxFocusIsManagedFormula(ByVal formula As String) As Boolean
    Dim normalized As String
    normalized = NxFocusNormalizeFormula(formula)
    NxFocusIsManagedFormula = (normalized = NxFocusNormalizeFormula(NX_FOCUS_ROW_FORMULA)) _
        Or (normalized = NxFocusNormalizeFormula(NX_FOCUS_COLUMN_FORMULA)) _
        Or (InStr(1, normalized, NX_FOCUS_RULE_V002_ROW, vbTextCompare) > 0) _
        Or (InStr(1, normalized, NX_FOCUS_RULE_V002_COLUMN, vbTextCompare) > 0)
End Function

Public Function NxFocusDeleteAddinCacheNames() As Boolean
    Dim item As Name
    Dim index As Long
    On Error GoTo Failed
    For index = ThisWorkbook.Names.Count To 1 Step -1
        Set item = ThisWorkbook.Names(index)
        If NxFocusIsManagedName(CStr(item.Name)) Then item.Delete
    Next index
    NxFocusDeleteAddinCacheNames = True
    Exit Function
Failed:
    NxFocusDeleteAddinCacheNames = False
End Function

Public Function NxFocusManagedConditionCount(ByVal target As Range) As Long
    Dim condition As Object
    Dim managed As Long
    On Error GoTo Done
    If target Is Nothing Then Exit Function
    For Each condition In target.FormatConditions
        If NxFocusIsManagedRule(condition) Then managed = managed + 1
    Next condition
Done:
    NxFocusManagedConditionCount = managed
End Function

Public Function NxFocusManagedRuleCount(ByVal sheet As Worksheet) As Long
    Dim conditions As FormatConditions
    Dim rule As Object
    On Error GoTo Done
    If sheet Is Nothing Then Exit Function
    Set conditions = sheet.Cells.FormatConditions
    For Each rule In conditions
        If NxFocusIsManagedRule(rule) Then NxFocusManagedRuleCount = NxFocusManagedRuleCount + 1
    Next rule
Done:
End Function

Private Function NxFocusIsManagedRule(ByVal rule As Object) As Boolean
    On Error GoTo NotManaged
    If rule Is Nothing Then Exit Function
    If rule.Type <> xlExpression Then Exit Function
    NxFocusIsManagedRule = NxFocusIsManagedFormula(CStr(rule.Formula1))
    Exit Function
NotManaged:
    NxFocusIsManagedRule = False
End Function

Private Function NxFocusNormalizeFormula(ByVal formula As String) As String
    Dim normalized As String
    normalized = UCase$(Trim$(formula))
    normalized = Replace(normalized, " ", vbNullString)
    normalized = Replace(normalized, vbTab, vbNullString)
    normalized = Replace(normalized, vbCr, vbNullString)
    normalized = Replace(normalized, vbLf, vbNullString)
    normalized = Replace(normalized, ";", ",")
    NxFocusNormalizeFormula = normalized
End Function

Private Sub NxFocusDeleteSheetBridgeNames(ByVal sheet As Worksheet)
    Dim item As Name
    Dim index As Long
    On Error Resume Next
    For index = sheet.Names.Count To 1 Step -1
        Set item = sheet.Names(index)
        If NxFocusIsManagedName(CStr(item.Name)) Then item.Delete
    Next index
    On Error GoTo 0
End Sub

Private Function NxFocusManagedNameCount(ByVal sheet As Worksheet) As Long
    Dim item As Name
    On Error GoTo Failed
    For Each item In sheet.Names
        If NxFocusIsManagedName(CStr(item.Name)) Then _
            NxFocusManagedNameCount = NxFocusManagedNameCount + 1
    Next item
    Exit Function
Failed:
    NxFocusManagedNameCount = -1
End Function

Private Function NxFocusIsManagedName(ByVal value As String) As Boolean
    value = UCase$(value)
    NxFocusIsManagedName = Right$(value, Len(NX_FOCUS_NAME_V002_TOP)) = NX_FOCUS_NAME_V002_TOP Or _
        Right$(value, Len(NX_FOCUS_NAME_V002_BOTTOM)) = NX_FOCUS_NAME_V002_BOTTOM Or _
        Right$(value, Len(NX_FOCUS_NAME_V002_LEFT)) = NX_FOCUS_NAME_V002_LEFT Or _
        Right$(value, Len(NX_FOCUS_NAME_V002_RIGHT)) = NX_FOCUS_NAME_V002_RIGHT
End Function
