Attribute VB_Name = "T_R61RoleStylePreview"
Option Explicit

' Import into a disposable Product.xlam copy. Run only through the prepared Windows BAT.
' CommandButton.Value exercises the native Click event, not physical keyboard/mouse input.
Public Sub RunAll()
    Dim kind As Variant, result As String
    For Each kind In Array("palette", "axes", "drift", "conflict_cancel")
        result = NxR61RoleStylePreviewCase(CStr(kind))
        Require Left$(result, 5) = "PASS|", result
    Next kind
End Sub

Public Function NxR61RoleStylePreviewCase(ByVal kind As String) As String
    Dim book As Workbook, expectedBook As Workbook, panel As FNxRoleStyle
    Dim target As Range, expected As Range, cell As Range, label As Object, condition As FormatCondition
    Dim before As String, roleIndex As Long, modeIndex As Long, r As Long, c As Long, roles As Variant
    Dim expectedFill As Long, expectedLines As Long, loadedCount As Long
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1:J10")
    For Each cell In target.Cells
        cell.Value2 = cell.Row * 100 + cell.Column
    Next cell
    target.Cells(1, 1).Value2 = "fixture"
    target.Cells(1, 2).Formula = "=1+2"
    target.Cells(1, 3).Value2 = CDbl(DateSerial(2026, 9, 6))
    target.Cells(1, 3).NumberFormat = "yyyy-mm-dd"
    target.Cells(2, 1).Value = CVErr(xlErrNA)
    target.Font.Italic = True
    book.Worksheets(1).Shapes.AddShape msoShapeRectangle, 700, 20, 30, 30
    Set condition = book.Worksheets(1).Range("L1").FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
    condition.Interior.Color = RGB(220, 240, 220)
    Set expectedBook = Workbooks.Add(xlWBATWorksheet)
    book.Activate
    target.Select
    Set panel = CreatePanel()
    loadedCount = Workbooks.Count
    Require Not panel.cmdExecute.Enabled, "initial form must not be executable"

    Select Case kind
        Case "palette"
            Set target = book.Worksheets(1).Range("A1:C3")
            target.Select
            panel.txtRange.Value = target.Address(False, False)
            roles = Array("TITLE", "SUBTITLE", "TABLE_HEADER", "TABLE_BODY", "EMPHASIS_CELL", "TOTAL_ROW")
            For modeIndex = 0 To 1
                For roleIndex = 0 To 5
                    Set expected = expectedBook.Worksheets(1).Range("A1:C3")
                    expected.Clear
                    expected.Value2 = target.Value2
                    expected.Font.Italic = True
                    expected.Cells(1, 3).NumberFormat = "yyyy-mm-dd"
                    NxRoleStyleApply expected, CStr(roles(roleIndex)), IIf(modeIndex = 0, "MONO", "COLOR")
                    panel.cboRole.ListIndex = roleIndex
                    panel.cboDisplayMode.ListIndex = modeIndex
                    book.Saved = (roleIndex Mod 2 = 0)
                    before = SourceState(book)
                    panel.cmdPreview.Value = True
                    Require panel.cmdExecute.Enabled, "valid preview not sealed: " & panel.lblStatus.Caption
                    Require CountControls(panel, "nxRoleCell_") = 9, "3x3 sample count"
                    For r = 1 To 3
                        For c = 1 To 3
                            Set label = panel.Controls("nxRoleCell_" & r & "_" & c)
                            Set cell = expected.Cells(r, c)
                            expectedFill = RGB(255, 255, 255)
                            If cell.Interior.Pattern <> xlNone Then expectedFill = cell.Interior.Color
                            Require label.BackColor = expectedFill, "preview fill differs from real role engine"
                            Require label.ForeColor = cell.Font.Color, "preview font color differs from real role engine"
                            Require label.Font.Name = cell.Font.Name And label.Font.Size = cell.Font.Size, "preview typeface differs from real role engine"
                            Require label.Font.Bold = cell.Font.Bold And label.Font.Italic = cell.Font.Italic, "preview font emphasis differs from real role engine"
                        Next c
                    Next r
                    expectedLines = Choose(roleIndex + 1, 1, 1, 4, 0, 4, 3)
                    Require CountControls(panel, "nxRoleLine_") = expectedLines, "role outline/double-line mismatch"
                    AssertLineColors panel, expected, roleIndex
                    Require before = SourceState(book), "preview mutated source values, formats, CF, shapes or Saved"
                    Require Workbooks.Count = loadedCount, "preview created a workbook"
                Next roleIndex
            Next modeIndex
        Case "axes"
            panel.cboRole.ListIndex = 4
            panel.cboAxis.ListIndex = 1
            panel.txtRelativeIndexes.Value = "8,10"
            before = SourceState(book)
            panel.cmdPreview.Value = True
            Require panel.cmdExecute.Enabled, "relative rows not sealed"
            Require CountControls(panel, "nxRoleCell_") = 15, "relative rows sample must be 3x5"
            Require panel.Controls("nxRoleCell_1_1").Tag = "affected", "selected relative row not styled"
            Require panel.Controls("nxRoleCell_2_1").Tag = "unchanged", "unselected relative row styled"
            Require InStr(panel.txtPreview.Value, "8~10") > 0, "sample coordinates missing from summary"
            panel.cboAxis.ListIndex = 2
            AssertInvalidated panel
            panel.cmdPreview.Value = True
            Require panel.cmdExecute.Enabled, "relative columns not sealed"
            Require CountControls(panel, "nxRoleCell_") = 15, "relative columns sample must be 5x3"
            Require panel.Controls("nxRoleCell_1_2").Tag = "unchanged", "unselected relative column styled"
            panel.txtRelativeIndexes.Value = "2,4"
            AssertInvalidated panel
            panel.cmdPreview.Value = True
            panel.cboRole.ListIndex = 1
            AssertInvalidated panel
            panel.cmdPreview.Value = True
            panel.cboDisplayMode.ListIndex = 1
            AssertInvalidated panel
            panel.cmdPreview.Value = True
            panel.txtRange.Value = "A1:H8"
            AssertInvalidated panel
            panel.cmdPreview.Value = True
            panel.cmdResetDefaults.Value = True
            AssertInvalidated panel
            Require before = SourceState(book), "option changes mutated the source"
        Case "drift"
            panel.cmdPreview.Value = True
            Require panel.cmdExecute.Enabled, "initial preview not ready"
            target.Cells(1, 1).Value2 = "edited after preview"
            before = SourceState(book)
            panel.cmdExecute.Value = True
            AssertInvalidated panel
            Require InStr(panel.lblStatus.Caption, "원본 내용") > 0, "stale content must be rejected before opening the runner"
            Require before = SourceState(book), "stale preview changed source"
            panel.cmdPreview.Value = True
            expectedBook.Activate
            panel.cmdExecute.Value = True
            AssertInvalidated panel
            Require InStr(panel.lblStatus.Caption, "통합문서") > 0, "workbook switch must be rejected"
        Case "conflict_cancel"
            panel.cmdPreview.Value = True
            Require panel.cmdExecute.Enabled, "valid preview should enable Apply"
            Set condition = target.Cells(1, 1).FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
            before = SourceState(book)
            panel.cmdPreview.Value = True
            AssertInvalidated panel
            Require InStr(panel.lblStatus.Caption, "FormatConditions") > 0, "user CF conflict not retained"
            panel.cmdCancel.Value = True
            Set panel = Nothing
            Require before = SourceState(book), "cancel changed the source"
        Case Else
            NxRaiseContractError "unknown r61 role-style fixture case"
    End Select
    If Not panel Is Nothing Then
        before = SourceState(book)
        Unload panel
        Set panel = Nothing
        Require before = SourceState(book), "close changed the source"
    End If
    NxR61RoleStylePreviewCase = "PASS|" & kind & "|native controls; physical UI NOT_RUN"
    GoTo CleanUp
Failed:
    NxR61RoleStylePreviewCase = "FAIL|" & kind & "|" & Err.Number & "|" & Err.Description
CleanUp:
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    Set panel = Nothing
    If Not expectedBook Is Nothing Then expectedBook.Close SaveChanges:=False
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
End Function

Private Function CreatePanel() As FNxRoleStyle
    Dim panel As New FNxRoleStyle, context As New CNxFeatureDialogContext
    context.Configure "NX-DRAW-ROLE-STYLE", "NX-DLG-ROLE-STYLE", "role-style", False
    panel.BindFeatureContext context
    Set CreatePanel = panel
End Function

Private Sub AssertInvalidated(ByVal panel As FNxRoleStyle)
    Require Not panel.cmdExecute.Enabled, "changed/failed preview left Apply enabled"
    Require CountControls(panel, "nxRoleCell_") = 0, "invalid preview left stale visual cells"
End Sub

Private Function CountControls(ByVal panel As FNxRoleStyle, ByVal prefix As String) As Long
    Dim control As Object
    For Each control In panel.Controls
        If Left$(control.Name, Len(prefix)) = prefix Then CountControls = CountControls + 1
    Next control
End Function

Private Sub AssertLineColors(ByVal panel As FNxRoleStyle, ByVal expected As Range, ByVal roleIndex As Long)
    Dim control As Object, colorValue As Long
    If roleIndex = 3 Then Exit Sub
    colorValue = expected.Borders(xlEdgeBottom).Color
    For Each control In panel.Controls
        If Left$(control.Name, 11) = "nxRoleLine_" Then Require control.BackColor = colorValue, "preview border color differs from real engine"
    Next control
End Sub

Private Function SourceState(ByVal book As Workbook) As String
    Dim cell As Range, edge As Variant, state As String, shape As Shape
    state = CStr(book.Saved) & "|" & book.Sheets.Count & "|" & book.Worksheets(1).Cells.FormatConditions.Count
    For Each cell In book.Worksheets(1).Range("A1:J10")
        state = state & "|" & cell.Address & ":" & cell.Text & ":" & cell.NumberFormat
        If Not IsError(cell.Formula) Then state = state & ":" & CStr(cell.Formula)
        state = state & ":" & cell.Interior.Color & ":" & cell.Font.Name & ":" & cell.Font.Size & ":" & cell.Font.Color & ":" & cell.Font.Bold & ":" & cell.Font.Italic
        state = state & ":" & cell.Font.Underline & ":" & cell.Font.Strikethrough & ":" & cell.HorizontalAlignment & ":" & cell.RowHeight & ":" & cell.ColumnWidth
        For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
            state = state & ":" & cell.Borders(CLng(edge)).LineStyle & ":" & cell.Borders(CLng(edge)).Color
        Next edge
    Next cell
    For Each shape In book.Worksheets(1).Shapes
        state = state & "|shape=" & shape.Name & ":" & shape.Left & ":" & shape.Top & ":" & shape.Width & ":" & shape.Height
    Next shape
    SourceState = state
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then NxRaiseContractError detail
End Sub
