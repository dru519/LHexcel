VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFocusRuleEngine"
Option Explicit

Private Const NX_FOCUS_RULE_V002_ROW As String = "NX_FOCUS_RULE_V002_ROW"
Private Const NX_FOCUS_RULE_V002_COLUMN As String = "NX_FOCUS_RULE_V002_COLUMN"
Private Const NX_FOCUS_NAME_V002_TOP As String = "NX_FOCUS_TOP_V2"
Private Const NX_FOCUS_NAME_V002_BOTTOM As String = "NX_FOCUS_BOTTOM_V2"
Private Const NX_FOCUS_NAME_V002_LEFT As String = "NX_FOCUS_LEFT_V2"
Private Const NX_FOCUS_NAME_V002_RIGHT As String = "NX_FOCUS_RIGHT_V2"

Private mSettings As CNxFocusSettingsModel
Private mLastSheet As Worksheet
Private mLastRefreshKey As String
Private mVerifiedSheet As Worksheet
Private mVerifiedRange As Range
Private mVerifiedAt As Single

Public Sub Configure(ByVal settings As CNxFocusSettingsModel)
    If settings Is Nothing Then NxRaiseContractError "포커스셀 설정이 필요합니다."
    If Not settings.IsSealed Then NxRaiseContractError "포커스셀 설정이 필요합니다."
    If Not mSettings Is settings Then InvalidateInspection
    Set mSettings = settings
End Sub

Public Function ApplySelection(ByVal selection As CNxFocusSelection, Optional ByVal allowRepair As Boolean = True) As Boolean
    Dim priorPaint As Boolean
    Dim paintChanged As Boolean
    Dim sheet As Worksheet
    Dim repaint As Range
    Dim refreshKey As String
    Dim inspected As Boolean
    Dim installed As Boolean
    Dim rulesValid As Boolean
    Dim verifiedRange As Range
    On Error GoTo Failed
    If selection Is Nothing Then Exit Function
    If Not selection.IsValid Then Exit Function
    If mSettings Is Nothing Then NxRaiseContractError "포커스셀 엔진 설정이 없습니다."
    Set sheet = selection.Sheet
    If Not CanReuseInspection(sheet, selection.Target) Then
        inspected = True
        If Not EnsureAddinCacheNames() Then Exit Function
        rulesValid = RulesMatchSettings(sheet, selection.Target, verifiedRange)
        If Not rulesValid Or Not HasSheetBridgeNames(sheet) Then
            If Not allowRepair Then Exit Function
            If Not InstallRules(sheet, selection.Target) Then Exit Function
            installed = True
            Set verifiedRange = ManagedAppliesTo(sheet)
            If verifiedRange Is Nothing Then Exit Function
        End If
        RememberInspection sheet, verifiedRange
    End If
    Set repaint = VisibleRefreshRange(sheet, selection.Target)
    refreshKey = CStr(selection.TopRow) & ":" & CStr(selection.BottomRow) & ":" & _
        CStr(selection.LeftColumn) & ":" & CStr(selection.RightColumn) & "|" & repaint.Address
    If mLastSheet Is sheet Then
        If mLastRefreshKey = refreshKey Then
            ApplySelection = True
            Exit Function
        End If
    End If
    ' Batch intermediate paints caused by updating the four coordinate names.
    ' Keep the caller's paint state, including on a failed refresh.
    priorPaint = Application.ScreenUpdating
    paintChanged = True
    Application.ScreenUpdating = False
    UpdateAddinCache selection
    ' Excel can resolve DisplayFormat from the new names while the actual grid
    ' still paints the previous selection. Range.Calculate and ScreenUpdating
    ' do not invalidate that grid cache. Worksheet.Calculate does; it also
    ' recalculates offscreen formulas, including in manual calculation mode.
    sheet.Calculate
    Application.ScreenUpdating = priorPaint
    paintChanged = False
    ApplySelection = (Not inspected Or rulesValid Or installed)
    If ApplySelection Then
        Set mLastSheet = sheet
        mLastRefreshKey = refreshKey
    End If
    Exit Function
Failed:
    On Error Resume Next
    If paintChanged Then Application.ScreenUpdating = priorPaint
    InvalidateInspection
    ApplySelection = False
End Function

Private Sub InvalidateRefresh()
    Set mLastSheet = Nothing
    mLastRefreshKey = vbNullString
End Sub

Public Sub InvalidateInspection()
    Set mVerifiedSheet = Nothing
    Set mVerifiedRange = Nothing
    InvalidateRefresh
End Sub

Private Function CanReuseInspection(ByVal sheet As Worksheet, ByVal target As Range) As Boolean
    Dim elapsed As Single
    Dim overlap As Range
    If Not mVerifiedSheet Is sheet Then Exit Function
    If mVerifiedRange Is Nothing Then Exit Function
    elapsed = Timer - mVerifiedAt
    ' Recheck periodically: editing CF rules does not raise SheetChange.
    If elapsed < 0 Or elapsed >= 0.25 Then Exit Function
    Set overlap = Application.Intersect(mVerifiedRange, target)
    If overlap Is Nothing Then Exit Function
    CanReuseInspection = (CDbl(overlap.CountLarge) = CDbl(target.CountLarge))
End Function

Private Sub RememberInspection(ByVal sheet As Worksheet, ByVal verifiedRange As Range)
    Set mVerifiedRange = verifiedRange
    Set mVerifiedSheet = sheet
    mVerifiedAt = Timer
End Sub

Private Function VisibleRefreshRange(ByVal sheet As Worksheet, ByVal fallback As Range) As Range
    Dim pane As Pane
    Dim paneIndex As Long
    Dim visible As Range
    Dim combined As Range
    If Not Application.ActiveWindow Is Nothing Then
        If Application.ActiveSheet Is sheet Then
            For paneIndex = 1 To Application.ActiveWindow.Panes.Count
                Set pane = Application.ActiveWindow.Panes(paneIndex)
                Set visible = pane.VisibleRange
                If combined Is Nothing Then
                    Set combined = visible
                Else
                    Set combined = Application.Union(combined, visible)
                End If
            Next paneIndex
        End If
    End If
    If combined Is Nothing Then Set combined = fallback
    Set VisibleRefreshRange = combined
End Function

Public Function NxFocusManagedRuleCount(ByVal sheet As Worksheet) As Long
    NxFocusManagedRuleCount = ManagedRuleCount(sheet)
End Function

Public Function NxFocusDeleteManagedRules(ByVal sheet As Worksheet) As Boolean
    NxFocusDeleteManagedRules = DeleteManagedRules(sheet)
End Function

Public Function InstallRules(ByVal sheet As Worksheet, ByVal target As Range) As Boolean
    Dim appliesTo As Range
    Dim rowRule As FormatCondition
    Dim columnRule As FormatCondition
    On Error GoTo Failed
    If sheet Is Nothing Then Exit Function
    If target Is Nothing Then Exit Function
    If mSettings Is Nothing Then Exit Function
    If Not target.Worksheet Is sheet Then Exit Function
    If Not DeleteManagedRules(sheet) Then Exit Function
    If Not EnsureAddinCacheNames() Then Exit Function
    If Not InstallSheetBridgeNames(sheet) Then Exit Function
    Set appliesTo = PlanAppliesTo(sheet, target)
    Set rowRule = appliesTo.FormatConditions.Add(Type:=xlExpression, Formula1:=CanonicalFormula(True))
    rowRule.Interior.Color = NxFocusConditionalColor(mSettings)
    rowRule.StopIfTrue = False
    Set columnRule = appliesTo.FormatConditions.Add(Type:=xlExpression, Formula1:=CanonicalFormula(False))
    columnRule.Interior.Color = NxFocusConditionalColor(mSettings)
    columnRule.StopIfTrue = False
    rowRule.SetLastPriority
    columnRule.SetLastPriority
    InstallRules = (ManagedRuleCount(sheet) = 2)
    Exit Function
Failed:
    On Error Resume Next
    DeleteManagedRules sheet
    On Error GoTo 0
End Function

Public Function DeleteManagedRules(ByVal sheet As Worksheet) As Boolean
    Dim conditions As FormatConditions
    Dim rule As Object
    Dim index As Long
    On Error GoTo Failed
    InvalidateInspection
    If sheet Is Nothing Then Exit Function
    Set conditions = sheet.Cells.FormatConditions
    For index = conditions.Count To 1 Step -1
        Set rule = conditions(index)
        If IsManagedFormula(ExpressionFormula(rule)) Then rule.Delete
    Next index
    DeleteSheetBridgeNames sheet
    DeleteManagedRules = (ManagedRuleCount(sheet) = 0)
    Exit Function
Failed:
    DeleteManagedRules = False
End Function

Public Function ManagedRuleCount(ByVal sheet As Worksheet) As Long
    Dim rule As Object
    On Error GoTo Done
    If sheet Is Nothing Then Exit Function
    For Each rule In sheet.Cells.FormatConditions
        If IsManagedFormula(ExpressionFormula(rule)) Then ManagedRuleCount = ManagedRuleCount + 1
    Next rule
Done:
End Function

Public Function ManagedConditionCount(ByVal target As Range) As Long
    Dim rule As Object
    On Error GoTo Done
    If target Is Nothing Then Exit Function
    For Each rule In target.FormatConditions
        If IsManagedFormula(ExpressionFormula(rule)) Then ManagedConditionCount = ManagedConditionCount + 1
    Next rule
Done:
End Function

Public Sub DeleteAddinCacheNames()
    Dim item As Name
    Dim itemName As String
    On Error Resume Next
    For Each item In ThisWorkbook.Names
        itemName = UCase$(CStr(item.Name))
        If IsManagedName(itemName) Then item.Delete
    Next item
    On Error GoTo 0
End Sub

Private Function PlanAppliesTo(ByVal sheet As Worksheet, ByVal target As Range) As Range
    Dim topRow As Long
    Dim leftColumn As Long
    Dim bottomRow As Long
    Dim rightColumn As Long
    topRow = target.Row - 1024
    If topRow < 1 Then topRow = 1
    leftColumn = target.Column - 64
    If leftColumn < 1 Then leftColumn = 1
    bottomRow = topRow + NX_FOCUS_APPLIES_ROWS - 1
    rightColumn = leftColumn + NX_FOCUS_APPLIES_COLUMNS - 1
    If bottomRow > sheet.Rows.Count Then
        bottomRow = sheet.Rows.Count
        topRow = bottomRow - NX_FOCUS_APPLIES_ROWS + 1
        If topRow < 1 Then topRow = 1
    End If
    If rightColumn > sheet.Columns.Count Then
        rightColumn = sheet.Columns.Count
        leftColumn = rightColumn - NX_FOCUS_APPLIES_COLUMNS + 1
        If leftColumn < 1 Then leftColumn = 1
    End If
    Set PlanAppliesTo = sheet.Range(sheet.Cells(topRow, leftColumn), sheet.Cells(bottomRow, rightColumn))
End Function

Private Function RulesMatchSettings(ByVal sheet As Worksheet, ByVal target As Range, _
                                    ByRef verifiedRange As Range) As Boolean
    Dim rule As Object
    Dim appliesTo As Range
    Dim overlap As Range
    Dim formula As String
    Dim rowFormula As String
    Dim columnFormula As String
    Dim rowMatched As Boolean
    Dim columnMatched As Boolean
    Dim expectedColor As Long
    Dim managedCount As Long
    On Error GoTo Failed
    If sheet Is Nothing Then Exit Function
    If target Is Nothing Then Exit Function
    If mSettings Is Nothing Then Exit Function
    rowFormula = NormalizeFormula(CanonicalFormula(True))
    columnFormula = NormalizeFormula(CanonicalFormula(False))
    expectedColor = CLng(NxFocusConditionalColor(mSettings))
    For Each rule In sheet.Cells.FormatConditions
        formula = NormalizeFormula(ExpressionFormula(rule))
        If IsManagedFormula(formula) Then
            managedCount = managedCount + 1
            Set appliesTo = rule.AppliesTo
            Set overlap = Application.Intersect(appliesTo, target)
            If overlap Is Nothing Then Exit Function
            If CDbl(overlap.CountLarge) <> CDbl(target.CountLarge) Then Exit Function
            If verifiedRange Is Nothing Then Set verifiedRange = appliesTo
            If formula = rowFormula Then
                rowMatched = (CLng(rule.Interior.Color) = expectedColor And Not CBool(rule.StopIfTrue))
            ElseIf formula = columnFormula Then
                columnMatched = (CLng(rule.Interior.Color) = expectedColor And Not CBool(rule.StopIfTrue))
            End If
        End If
    Next rule
    RulesMatchSettings = (managedCount = 2 And rowMatched And columnMatched)
    Exit Function
Failed:
    RulesMatchSettings = False
End Function

Private Function ManagedAppliesTo(ByVal sheet As Worksheet) As Range
    Dim rule As Object
    On Error GoTo Failed
    For Each rule In sheet.Cells.FormatConditions
        If IsManagedFormula(ExpressionFormula(rule)) Then
            Set ManagedAppliesTo = rule.AppliesTo
            Exit Function
        End If
    Next rule
Failed:
End Function

Private Function NormalizeFormula(ByVal value As String) As String
    value = UCase$(Trim$(value))
    value = Replace(value, " ", vbNullString)
    value = Replace(value, vbTab, vbNullString)
    value = Replace(value, vbCr, vbNullString)
    value = Replace(value, vbLf, vbNullString)
    NormalizeFormula = Replace(value, ";", ",")
End Function

Private Function CanonicalFormula(ByVal rowAxis As Boolean) As String
    Dim axisFormula As String
    Dim selectedAreaFormula As String
    If rowAxis Then
        If mSettings.Shape = NX_FOCUS_SHAPE_VERTICAL Then
            axisFormula = "FALSE"
        Else
            axisFormula = "AND(ROW()>=" & NX_FOCUS_NAME_V002_TOP & ",ROW()<=" & NX_FOCUS_NAME_V002_BOTTOM & ")"
        End If
        CanonicalFormula = "=AND(N(""" & NX_FOCUS_RULE_V002_ROW & """)=0," & axisFormula
    Else
        If mSettings.Shape = NX_FOCUS_SHAPE_HORIZONTAL Then
            axisFormula = "FALSE"
        Else
            axisFormula = "AND(COLUMN()>=" & NX_FOCUS_NAME_V002_LEFT & ",COLUMN()<=" & NX_FOCUS_NAME_V002_RIGHT & ")"
        End If
        CanonicalFormula = "=AND(N(""" & NX_FOCUS_RULE_V002_COLUMN & """)=0," & axisFormula
    End If
    If Not mSettings.HighlightSelectedArea Then
        selectedAreaFormula = "AND(ROW()>=" & NX_FOCUS_NAME_V002_TOP & ",ROW()<=" & NX_FOCUS_NAME_V002_BOTTOM & _
            ",COLUMN()>=" & NX_FOCUS_NAME_V002_LEFT & ",COLUMN()<=" & NX_FOCUS_NAME_V002_RIGHT & ")"
        CanonicalFormula = CanonicalFormula & ",NOT(" & selectedAreaFormula & ")"
    End If
    CanonicalFormula = CanonicalFormula & ")"
End Function

Private Function EnsureAddinCacheNames() As Boolean
    Dim hasTop As Boolean
    Dim hasBottom As Boolean
    Dim hasLeft As Boolean
    Dim hasRight As Boolean
    On Error GoTo Failed
    hasTop = HasAddinName(NX_FOCUS_NAME_V002_TOP)
    hasBottom = HasAddinName(NX_FOCUS_NAME_V002_BOTTOM)
    hasLeft = HasAddinName(NX_FOCUS_NAME_V002_LEFT)
    hasRight = HasAddinName(NX_FOCUS_NAME_V002_RIGHT)
    If Not hasTop Or Not hasBottom Or Not hasLeft Or Not hasRight Then InvalidateRefresh
    If Not hasTop Then _
        ThisWorkbook.Names.Add Name:=NX_FOCUS_NAME_V002_TOP, RefersTo:="=1", Visible:=False
    If Not hasBottom Then _
        ThisWorkbook.Names.Add Name:=NX_FOCUS_NAME_V002_BOTTOM, RefersTo:="=1", Visible:=False
    If Not hasLeft Then _
        ThisWorkbook.Names.Add Name:=NX_FOCUS_NAME_V002_LEFT, RefersTo:="=1", Visible:=False
    If Not hasRight Then _
        ThisWorkbook.Names.Add Name:=NX_FOCUS_NAME_V002_RIGHT, RefersTo:="=1", Visible:=False
    EnsureAddinCacheNames = True
    Exit Function
Failed:
    EnsureAddinCacheNames = False
End Function

Private Sub UpdateAddinCache(ByVal selection As CNxFocusSelection)
    ThisWorkbook.Names(NX_FOCUS_NAME_V002_TOP).RefersTo = "=" & CStr(selection.TopRow)
    ThisWorkbook.Names(NX_FOCUS_NAME_V002_BOTTOM).RefersTo = "=" & CStr(selection.BottomRow)
    ThisWorkbook.Names(NX_FOCUS_NAME_V002_LEFT).RefersTo = "=" & CStr(selection.LeftColumn)
    ThisWorkbook.Names(NX_FOCUS_NAME_V002_RIGHT).RefersTo = "=" & CStr(selection.RightColumn)
End Sub

Private Function InstallSheetBridgeNames(ByVal sheet As Worksheet) As Boolean
    Dim item As Name
    On Error GoTo Failed
    DeleteSheetBridgeNames sheet
    Set item = sheet.Names.Add(Name:=NX_FOCUS_NAME_V002_TOP, RefersTo:="=" & QualifiedAddinName(NX_FOCUS_NAME_V002_TOP), Visible:=False)
    Set item = sheet.Names.Add(Name:=NX_FOCUS_NAME_V002_BOTTOM, RefersTo:="=" & QualifiedAddinName(NX_FOCUS_NAME_V002_BOTTOM), Visible:=False)
    Set item = sheet.Names.Add(Name:=NX_FOCUS_NAME_V002_LEFT, RefersTo:="=" & QualifiedAddinName(NX_FOCUS_NAME_V002_LEFT), Visible:=False)
    Set item = sheet.Names.Add(Name:=NX_FOCUS_NAME_V002_RIGHT, RefersTo:="=" & QualifiedAddinName(NX_FOCUS_NAME_V002_RIGHT), Visible:=False)
    InstallSheetBridgeNames = HasSheetBridgeNames(sheet)
    Exit Function
Failed:
    DeleteSheetBridgeNames sheet
End Function

Private Function HasSheetBridgeNames(ByVal sheet As Worksheet) As Boolean
    Dim found As Long
    Dim item As Name
    On Error GoTo Failed
    For Each item In sheet.Names
        If IsManagedName(UCase$(CStr(item.Name))) Then found = found + 1
    Next item
    HasSheetBridgeNames = (found = 4)
    Exit Function
Failed:
    HasSheetBridgeNames = False
End Function

Private Sub DeleteSheetBridgeNames(ByVal sheet As Worksheet)
    Dim item As Name
    On Error Resume Next
    For Each item In sheet.Names
        If IsManagedName(UCase$(CStr(item.Name))) Then item.Delete
    Next item
    On Error GoTo 0
End Sub

Private Function HasAddinName(ByVal cacheName As String) As Boolean
    Dim item As Name
    On Error GoTo Missing
    Set item = ThisWorkbook.Names(cacheName)
    HasAddinName = Not item Is Nothing
Missing:
End Function

Private Function QualifiedAddinName(ByVal cacheName As String) As String
    QualifiedAddinName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & cacheName
End Function

Private Function IsManagedName(ByVal value As String) As Boolean
    IsManagedName = Right$(value, Len(NX_FOCUS_NAME_V002_TOP)) = NX_FOCUS_NAME_V002_TOP Or _
        Right$(value, Len(NX_FOCUS_NAME_V002_BOTTOM)) = NX_FOCUS_NAME_V002_BOTTOM Or _
        Right$(value, Len(NX_FOCUS_NAME_V002_LEFT)) = NX_FOCUS_NAME_V002_LEFT Or _
        Right$(value, Len(NX_FOCUS_NAME_V002_RIGHT)) = NX_FOCUS_NAME_V002_RIGHT
End Function

Private Function ExpressionFormula(ByVal rule As Object) As String
    ' ColorScale, DataBar and IconSetCondition do not expose Formula1.
    ' Only expression rules can belong to this engine; leave all others alone.
    If rule.Type <> xlExpression Then Exit Function
    ExpressionFormula = CStr(rule.Formula1)
End Function

Private Function IsManagedFormula(ByVal formula As String) As Boolean
    IsManagedFormula = InStr(1, formula, NX_FOCUS_RULE_V002_ROW, vbTextCompare) > 0 Or _
        InStr(1, formula, NX_FOCUS_RULE_V002_COLUMN, vbTextCompare) > 0
End Function
