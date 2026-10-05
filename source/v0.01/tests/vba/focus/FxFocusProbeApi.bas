Attribute VB_Name = "FxFocusProbeApi"
Option Explicit

Public Const NX_FOCUS_RULE_V002_ROW As String = "NX_FOCUS_RULE_V002_ROW"
Public Const NX_FOCUS_RULE_V002_COLUMN As String = "NX_FOCUS_RULE_V002_COLUMN"
Private Const NX_FOCUS_NAME_V002_TOP As String = "NX_FOCUS_TOP_V2"
Private Const NX_FOCUS_NAME_V002_BOTTOM As String = "NX_FOCUS_BOTTOM_V2"
Private Const NX_FOCUS_NAME_V002_LEFT As String = "NX_FOCUS_LEFT_V2"
Private Const NX_FOCUS_NAME_V002_RIGHT As String = "NX_FOCUS_RIGHT_V2"

Private mEvents As CFxFocusProbeEvents
Private mSheet As Worksheet
Private mTopRow As Long
Private mBottomRow As Long
Private mLeftColumn As Long
Private mRightColumn As Long
Private mHighlightSelectedArea As Boolean
Private mLastError As String

Public Sub FxFocusStart(ByVal app As Application)
    FxFocusStop
    Set mEvents = New CFxFocusProbeEvents
    mEvents.Start app
End Sub

Public Sub FxFocusStop()
    On Error Resume Next
    If Not mEvents Is Nothing Then mEvents.StopListening
    FxDeleteAddinCacheNames
    Set mEvents = Nothing
    Set mSheet = Nothing
    mTopRow = 0
    mBottomRow = 0
    mLeftColumn = 0
    mRightColumn = 0
    mLastError = vbNullString
    On Error GoTo 0
End Sub

Public Sub FxFocusResumeAfterEdit()
    If Not mEvents Is Nothing Then mEvents.ClearEditPending
End Sub

Public Function FxInstallRules(ByVal sheet As Worksheet, ByVal appliesTo As Range, _
                               ByVal focusColor As Long, ByVal highlightSelectedArea As Boolean) As Boolean
    Dim rowRule As FormatCondition
    Dim columnRule As FormatCondition
    Dim rowFormula As String
    Dim columnFormula As String

    On Error GoTo Rejected
    mLastError = vbNullString
    If sheet Is Nothing Or appliesTo Is Nothing Then Exit Function
    If Not appliesTo.Worksheet Is sheet Then Exit Function

    FxDeleteRules sheet
    If Not FxInstallAddinCacheNames() Then Exit Function
    If Not FxInstallSheetBridgeNames(sheet) Then Exit Function
    mHighlightSelectedArea = highlightSelectedArea
    FxCacheSelection sheet.Range("A1:C3")
    rowFormula = FxCanonicalFormula(NX_FOCUS_RULE_V002_ROW, True)
    columnFormula = FxCanonicalFormula(NX_FOCUS_RULE_V002_COLUMN, False)

    Set rowRule = appliesTo.FormatConditions.Add(Type:=xlExpression, Formula1:=rowFormula)
    rowRule.Interior.Color = focusColor
    rowRule.StopIfTrue = False
    Set columnRule = appliesTo.FormatConditions.Add(Type:=xlExpression, Formula1:=columnFormula)
    columnRule.Interior.Color = focusColor
    columnRule.StopIfTrue = False
    rowRule.SetLastPriority
    columnRule.SetLastPriority

    FxInstallRules = (FxManagedRuleCount(sheet) = 2)
    Exit Function
Rejected:
    mLastError = CStr(Err.Number) & "|" & Err.Description
    On Error Resume Next
    FxDeleteRules sheet
    On Error GoTo 0
    FxInstallRules = False
End Function

Public Function FxLastError() As String
    FxLastError = mLastError
End Function

Public Function FxRepairRules(ByVal sheet As Worksheet, ByVal appliesTo As Range, _
                              ByVal focusColor As Long, ByVal highlightSelectedArea As Boolean) As Boolean
    FxRepairRules = FxInstallRules(sheet, appliesTo, focusColor, highlightSelectedArea)
    If Not mEvents Is Nothing Then mEvents.ClearEditPending
End Function

Private Function FxCanonicalFormula(ByVal marker As String, ByVal rowAxis As Boolean) As String
    Dim topName As String
    Dim bottomName As String
    Dim leftName As String
    Dim rightName As String
    Dim axisFormula As String
    Dim selectedAreaFormula As String

    topName = NX_FOCUS_NAME_V002_TOP
    bottomName = NX_FOCUS_NAME_V002_BOTTOM
    leftName = NX_FOCUS_NAME_V002_LEFT
    rightName = NX_FOCUS_NAME_V002_RIGHT
    If rowAxis Then
        axisFormula = "OR(ROW()=" & topName & ",ROW()=" & bottomName & ")"
    Else
        axisFormula = "OR(COLUMN()=" & leftName & ",COLUMN()=" & rightName & ")"
    End If
    FxCanonicalFormula = "=AND(N(""" & marker & """)=0," & axisFormula
    If Not mHighlightSelectedArea Then
        selectedAreaFormula = "AND(ROW()>=" & topName & ",ROW()<=" & bottomName & _
            ",COLUMN()>=" & leftName & ",COLUMN()<=" & rightName & ")"
        FxCanonicalFormula = FxCanonicalFormula & ",NOT(" & selectedAreaFormula & ")"
    End If
    FxCanonicalFormula = FxCanonicalFormula & ")"
End Function

Private Function FxInstallAddinCacheNames() As Boolean
    Dim item As Name
    On Error GoTo Failed
    FxDeleteAddinCacheNames
    Set item = ThisWorkbook.Names.Add(Name:=NX_FOCUS_NAME_V002_TOP, RefersTo:="=1", Visible:=False)
    Set item = ThisWorkbook.Names.Add(Name:=NX_FOCUS_NAME_V002_BOTTOM, RefersTo:="=1", Visible:=False)
    Set item = ThisWorkbook.Names.Add(Name:=NX_FOCUS_NAME_V002_LEFT, RefersTo:="=1", Visible:=False)
    Set item = ThisWorkbook.Names.Add(Name:=NX_FOCUS_NAME_V002_RIGHT, RefersTo:="=1", Visible:=False)
    FxInstallAddinCacheNames = FxAddinCacheNamesReady()
    Exit Function
Failed:
    mLastError = CStr(Err.Number) & "|" & Err.Description
    FxDeleteAddinCacheNames
End Function

Private Sub FxDeleteAddinCacheNames()
    Dim item As Name
    Dim itemName As String
    On Error Resume Next
    For Each item In ThisWorkbook.Names
        itemName = UCase$(CStr(item.Name))
        If Right$(itemName, Len(NX_FOCUS_NAME_V002_TOP)) = NX_FOCUS_NAME_V002_TOP Or _
           Right$(itemName, Len(NX_FOCUS_NAME_V002_BOTTOM)) = NX_FOCUS_NAME_V002_BOTTOM Or _
           Right$(itemName, Len(NX_FOCUS_NAME_V002_LEFT)) = NX_FOCUS_NAME_V002_LEFT Or _
           Right$(itemName, Len(NX_FOCUS_NAME_V002_RIGHT)) = NX_FOCUS_NAME_V002_RIGHT Then item.Delete
    Next item
    On Error GoTo 0
End Sub

Private Function FxAddinCacheNamesReady() As Boolean
    On Error GoTo NotReady
    FxAddinCacheNamesReady = Not ThisWorkbook.Names(NX_FOCUS_NAME_V002_TOP) Is Nothing And _
        Not ThisWorkbook.Names(NX_FOCUS_NAME_V002_BOTTOM) Is Nothing And _
        Not ThisWorkbook.Names(NX_FOCUS_NAME_V002_LEFT) Is Nothing And _
        Not ThisWorkbook.Names(NX_FOCUS_NAME_V002_RIGHT) Is Nothing
    Exit Function
NotReady:
    FxAddinCacheNamesReady = False
End Function

Private Function FxQualifiedCacheName(ByVal cacheName As String) As String
    FxQualifiedCacheName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & cacheName
End Function

Private Function FxInstallSheetBridgeNames(ByVal sheet As Worksheet) As Boolean
    Dim item As Name
    On Error GoTo Failed
    FxDeleteSheetBridgeNames sheet
    Set item = sheet.Names.Add(Name:=NX_FOCUS_NAME_V002_TOP, _
        RefersTo:="=" & FxQualifiedCacheName(NX_FOCUS_NAME_V002_TOP), Visible:=False)
    Set item = sheet.Names.Add(Name:=NX_FOCUS_NAME_V002_BOTTOM, _
        RefersTo:="=" & FxQualifiedCacheName(NX_FOCUS_NAME_V002_BOTTOM), Visible:=False)
    Set item = sheet.Names.Add(Name:=NX_FOCUS_NAME_V002_LEFT, _
        RefersTo:="=" & FxQualifiedCacheName(NX_FOCUS_NAME_V002_LEFT), Visible:=False)
    Set item = sheet.Names.Add(Name:=NX_FOCUS_NAME_V002_RIGHT, _
        RefersTo:="=" & FxQualifiedCacheName(NX_FOCUS_NAME_V002_RIGHT), Visible:=False)
    FxInstallSheetBridgeNames = True
    Exit Function
Failed:
    mLastError = CStr(Err.Number) & "|" & Err.Description
    FxDeleteSheetBridgeNames sheet
End Function

Private Sub FxDeleteSheetBridgeNames(ByVal sheet As Worksheet)
    Dim item As Name
    Dim itemName As String
    On Error Resume Next
    For Each item In sheet.Names
        itemName = UCase$(CStr(item.Name))
        If Right$(itemName, Len(NX_FOCUS_NAME_V002_TOP)) = NX_FOCUS_NAME_V002_TOP Or _
           Right$(itemName, Len(NX_FOCUS_NAME_V002_BOTTOM)) = NX_FOCUS_NAME_V002_BOTTOM Or _
           Right$(itemName, Len(NX_FOCUS_NAME_V002_LEFT)) = NX_FOCUS_NAME_V002_LEFT Or _
           Right$(itemName, Len(NX_FOCUS_NAME_V002_RIGHT)) = NX_FOCUS_NAME_V002_RIGHT Then item.Delete
    Next item
    On Error GoTo 0
End Sub

Private Sub FxUpdateAddinCacheNames()
    If Not FxAddinCacheNamesReady() Then Exit Sub
    ThisWorkbook.Names(NX_FOCUS_NAME_V002_TOP).RefersTo = "=" & CStr(mTopRow)
    ThisWorkbook.Names(NX_FOCUS_NAME_V002_BOTTOM).RefersTo = "=" & CStr(mBottomRow)
    ThisWorkbook.Names(NX_FOCUS_NAME_V002_LEFT).RefersTo = "=" & CStr(mLeftColumn)
    ThisWorkbook.Names(NX_FOCUS_NAME_V002_RIGHT).RefersTo = "=" & CStr(mRightColumn)
End Sub

Public Sub FxCacheSelection(ByVal target As Range)
    If target Is Nothing Then Exit Sub
    If target.Areas.Count <> 1 Then Exit Sub
    Set mSheet = target.Worksheet
    mTopRow = target.Row
    mLeftColumn = target.Column
    mBottomRow = target.Row + target.Rows.Count - 1
    mRightColumn = target.Column + target.Columns.Count - 1
    FxUpdateAddinCacheNames
End Sub

Private Function FxAxisMatch(ByVal callerValue As Variant, ByVal rowAxis As Boolean) As Boolean
    Dim caller As Range
    On Error GoTo FailClosed
    If TypeName(callerValue) <> "Range" Then Exit Function
    Set caller = callerValue
    If mSheet Is Nothing Then Exit Function
    If Not caller.Worksheet Is mSheet Then Exit Function
    If Not mHighlightSelectedArea Then
        If caller.Row >= mTopRow And caller.Row <= mBottomRow And _
           caller.Column >= mLeftColumn And caller.Column <= mRightColumn Then Exit Function
    End If
    If rowAxis Then
        FxAxisMatch = (caller.Row = mTopRow Or caller.Row = mBottomRow)
    Else
        FxAxisMatch = (caller.Column = mLeftColumn Or caller.Column = mRightColumn)
    End If
    Exit Function
FailClosed:
    FxAxisMatch = False
End Function

Public Function FxProbePointMatches(ByVal sheet As Worksheet, ByVal rowNumber As Long, _
                                    ByVal columnNumber As Long, ByVal rowAxis As Boolean) As Boolean
    Dim probe As Range
    If sheet Is Nothing Then Exit Function
    Set probe = sheet.Cells(rowNumber, columnNumber)
    FxProbePointMatches = FxAxisMatch(probe, rowAxis)
End Function

Public Function FxManagedRuleCount(ByVal sheet As Worksheet) As Long
    Dim condition As Object
    On Error GoTo Done
    If sheet Is Nothing Then Exit Function
    For Each condition In sheet.Cells.FormatConditions
        If FxIsManagedFormula(CStr(condition.Formula1)) Then _
            FxManagedRuleCount = FxManagedRuleCount + 1
    Next condition
Done:
End Function

Public Function FxDeleteRules(ByVal sheet As Worksheet) As Boolean
    Dim conditions As FormatConditions
    Dim index As Long
    On Error GoTo Failed
    If sheet Is Nothing Then Exit Function
    Set conditions = sheet.Cells.FormatConditions
    For index = conditions.Count To 1 Step -1
        If FxIsManagedFormula(CStr(conditions(index).Formula1)) Then conditions(index).Delete
    Next index
    FxDeleteSheetBridgeNames sheet
    FxDeleteRules = (FxManagedRuleCount(sheet) = 0)
    Exit Function
Failed:
    FxDeleteRules = False
End Function

Private Function FxIsManagedFormula(ByVal formula As String) As Boolean
    FxIsManagedFormula = (InStr(1, formula, NX_FOCUS_RULE_V002_ROW, vbTextCompare) > 0) _
        Or (InStr(1, formula, NX_FOCUS_RULE_V002_COLUMN, vbTextCompare) > 0)
End Function

Public Function FxSelectionGeometry() As String
    FxSelectionGeometry = CStr(mTopRow) & ":" & CStr(mBottomRow) & "|" & _
        CStr(mLeftColumn) & ":" & CStr(mRightColumn)
End Function
