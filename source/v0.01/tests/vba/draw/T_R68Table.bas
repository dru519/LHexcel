Attribute VB_Name = "T_R68Table"
Option Explicit

' Native Excel fixture. Import into a disposable product copy and run from the prepared Windows BAT.
' Click events below are native-control coverage, not physical mouse/keyboard or DPI acceptance.
Public Function Names() As String
    Names = "defaults|no_header|open_grid|closed_grid|header_bottom_edges|total_options|autofit|preview_matrix|preview_merged|preview_offsample_merge|preview_vertical_merge|preview_wrapping|manual_range_apply|format_restore|title_defaults|apply_200_cells|runner_success|runner_stale"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, expectedBook As Workbook, priorBook As Workbook, priorSelection As Range
    Dim source As Range, expected As Range, cell As Range, panel As FNxDrawTable
    Dim request As New CNxDrawingRequest, before As String, contentBefore As String
    Dim index As Long, mode As Long, style As Long, finishing As Long, r As Long, c As Long, sourceRow As Long
    Dim label As Object, originalLabel As Object, fill As Long, loadedCount As Long, widthBefore As Double
    Dim displayMode As String, tableStyle As String, hasTotal As Boolean
    Dim startedAt As Double, elapsedSeconds As Double, previewRows As Long
    Dim formatBackup As CNxFormatSnapshot
    Dim commandObject As CNxDrawingFeatureCommand, command As INxFeatureCommand
    Dim approval As CNxApproval, result As CNxResult
    On Error GoTo Failed
    Set priorBook = ActiveWorkbook
    If TypeName(Selection) = "Range" Then Set priorSelection = Selection
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set source = book.Worksheets(1).Range("A1:C8")
    For Each cell In source.Cells
        cell.Value2 = cell.Row * 10 + cell.Column
    Next cell
    source.Cells(1, 1).Value2 = "항목"
    source.Cells(2, 1).Value2 = "문자열"
    source.Cells(2, 2).Formula = "=1+2"
    source.Cells(3, 2).Value2 = CDbl(DateSerial(2026, 9, 9))
    source.Cells(3, 2).NumberFormat = "yyyy-mm-dd"
    source.Cells(4, 1).NumberFormat = "@"
    source.Cells(4, 1).Value2 = "123"
    source.Cells(5, 1).Value = CVErr(xlErrNA)
    source.Cells(8, 1).Value2 = "마지막 합계 표본"
    source.Font.Italic = True
    source.Cells(2, 1).HorizontalAlignment = xlRight
    source.Columns(1).ColumnWidth = 24
    source.Columns(2).ColumnWidth = 12
    source.Columns(3).ColumnWidth = 8
    source.Select

    Select Case name
        Case "runner_success", "runner_stale"
            request.ConfigureRange NX_FEATURE_DRAW_BUSINESS_TABLE, source, False, True
            Set commandObject = New CNxDrawingFeatureCommand
            commandObject.Configure request
            Set command = commandObject
            Set approval = NxCommandRunner.Approve(command)
            If name = "runner_stale" Then source.Cells(2, 2).Value2 = "changed after approval"
            Set result = NxCommandRunner.Run(command, approval)
            If name = "runner_success" Then
                Require result.Outcome = NxSuccess And result.SourceChanged, "successful table must report mutation"
                Require source.Font.Size = 9, "approved table applied"
                Require source.Cells(2, 2).Formula = "=1+2", "formula preserved"
            Else
                Require result.Outcome = NxInputError, "changed source must reject approval"
                Require source.Cells(2, 2).Value2 = "changed after approval", "stale source value preserved"
                Require source.Font.Italic, "stale approval must not format source"
            End If
        Case "defaults"
            request.ConfigureRange NX_FEATURE_DRAW_BUSINESS_TABLE, source, False, True
            Require request.TableStyle = "열린 표", "request must default to OPEN"
            Require request.HeaderFontSize = 9 And request.BodyFontSize = 9, "request must default to 9pt"
            NxDrawApplyBusinessTable source
            AssertGrid source, True, 1
            Require source.Font.Size = 9, "business helper must use 9pt"
            Set panel = CreatePanel(NX_FEATURE_DRAW_BUSINESS_TABLE)
            Require panel.cboTableStyle.Value = "스타일: 열린 표", "form must default to OPEN"
            Require panel.cboHeaderSize.Value = "머리글 크기: 9", "header default 9"
            Require panel.cboBodySize.Value = "본문 크기: 9", "body default 9"
            panel.cboHeaderSize.Value = "머리글 크기: 12"
            panel.cboBodySize.Value = "본문 크기: 14"
            panel.cboTableStyle.ListIndex = 0
            panel.cboFinishing.ListIndex = 3
            panel.cmdResetDefaults.Value = True
            Require panel.cboTableStyle.ListIndex = 1 And panel.cboFinishing.ListIndex = 0, "reset style/finishing"
            Require panel.cboHeaderSize.Value = "머리글 크기: 9" And panel.cboBodySize.Value = "본문 크기: 9", "reset fonts"
        Case "open_grid", "closed_grid"
            tableStyle = IIf(name = "open_grid", "열린 표", "닫힌 표")
            For mode = 0 To 1
                displayMode = IIf(mode = 0, "MONO", "COLOR")
                source.Borders.LineStyle = xlDouble
                NxDrawApplyTableOptions source, NX_FEATURE_DRAW_BUSINESS_TABLE, displayMode, tableStyle, 2, 9, xlCenter, 9, xlGeneral
                AssertGrid source, (name = "open_grid"), 2
                Require source.Borders(xlEdgeTop).Color = NxRoleStyleTokenColor("LINE", displayMode), "grid mode color"
                Require source.Font.Size = 9, "grid fonts must remain 9"
                Require source.Cells(1, 1).Interior.Color = IIf(mode = 0, RGB(230, 230, 230), RGB(216, 226, 240)), "header one shade darker"
            Next mode
            Set source = book.Worksheets(1).Range("E1:E4")
            NxDrawApplyBusinessTable source
            Require source.Borders(xlInsideHorizontal).LineStyle = xlContinuous, "single-column open table horizontal grid"
            Set source = book.Worksheets(1).Range("G1:I1")
            NxDrawApplyBusinessTable source
            Require source.Borders(xlInsideVertical).LineStyle = xlContinuous, "single-row open table vertical grid"
        Case "header_bottom_edges"
            Set source = book.Worksheets(1).Range("E1:G2")
            For mode = 0 To 1
                displayMode = IIf(mode = 0, "MONO", "COLOR")
                For style = 0 To 1
                    tableStyle = IIf(style = 0, "열린 표", "닫힌 표")
                    For index = 0 To 2
                        source.ClearFormats
                        NxDrawApplyTableOptions source, NX_FEATURE_DRAW_BUSINESS_TABLE, displayMode, tableStyle, index, 9, xlCenter, 9, xlGeneral, True
                        Require source.Rows(2).Borders(xlEdgeBottom).Weight = xlThick, "last row thick with total/header overlap"
                        If index > 0 Then Require source.Rows(index).Borders(xlEdgeBottom).Weight = xlThick, "header thick with total overlap"
                    Next index
                Next style
            Next mode
            Set source = book.Worksheets(1).Range("A1:C8")
            source.Select
            Set panel = CreatePanel(NX_FEATURE_DRAW_BUSINESS_TABLE)
            Require panel.Controls("nxPrv_line_h1").Height = 3, "preview header line thickness"
            Require panel.Controls("nxPrv_line_bottom").Height = 3, "preview bottom line thickness"
        Case "total_options"
            contentBefore = ContentState(source)
            For mode = 0 To 1
                displayMode = IIf(mode = 0, "MONO", "COLOR")
                For finishing = 0 To 3
                    source.ClearFormats
                    hasTotal = (finishing = 1 Or finishing = 3)
                    NxDrawApplyTableOptions source, NX_FEATURE_DRAW_BUSINESS_TABLE, displayMode, "열린 표", 2, 9, xlCenter, 9, xlRight, hasTotal, (finishing >= 2)
                    Require source.Font.Size = 9, "finishing reset font size"
                    Require source.Cells(8, 2).Font.Bold = hasTotal, "last row bold must match total option"
                    Require source.Cells(8, 2).HorizontalAlignment = xlRight, "total must retain body alignment"
                    If hasTotal Then
                        Require source.Cells(8, 2).Borders(xlEdgeTop).Weight = xlMedium, "total separator overwritten"
                        Require source.Cells(8, 2).Interior.Color = IIf(mode = 0, RGB(242, 242, 242), NxTotalFillColor()), "total fill"
                        Require source.Cells(8, 2).Font.Color = IIf(mode = 0, RGB(32, 32, 32), NxBodyFontColor()), "total mode font color"
                    End If
                    Require source.Borders(xlInsideVertical).LineStyle = xlContinuous, "total finishing lost internal verticals"
                    Require ContentState(source) = contentBefore, "finishing changed values/formulas or added SUM"
                Next finishing
            Next mode
            NxDrawApplyTableOptions source, NX_FEATURE_DRAW_BUSINESS_TABLE, "COLOR", "열린 표", 1, 12, xlCenter, 11, xlLeft, True
            Require source.Cells(8, 2).Font.Size = 11, "custom total font must follow body"
        Case "autofit"
            source.Columns.ColumnWidth = 32
            widthBefore = source.Columns(3).ColumnWidth
            contentBefore = ContentState(source)
            NxDrawApplyTableOptions source, NX_FEATURE_DRAW_BUSINESS_TABLE, "MONO", "열린 표", 1, 9, xlCenter, 9, xlGeneral, False, False
            Require source.Columns(3).ColumnWidth = widthBefore, "default must not resize columns"
            NxDrawApplyTableOptions source, NX_FEATURE_DRAW_BUSINESS_TABLE, "MONO", "열린 표", 1, 9, xlCenter, 9, xlGeneral, False, True
            Require source.Columns(3).ColumnWidth <> widthBefore, "AutoFit did not resize"
            Require book.Worksheets(1).Range("C20").ColumnWidth = source.Columns(3).ColumnWidth, "column width scope"
            Require ContentState(source) = contentBefore, "AutoFit changed content"
        Case "preview_matrix"
            Set expectedBook = Workbooks.Add(xlWBATWorksheet)
            Set expected = expectedBook.Worksheets(1).Range("A1:C8")
            book.Activate
            source.Select
            book.Saved = True
            before = SourceState(book)
            Set panel = CreatePanel(NX_FEATURE_DRAW_BUSINESS_TABLE)
            Set originalLabel = panel.Controls("nxPrv_cell_1_1")
            loadedCount = Workbooks.Count
            panel.cboHeaderRows.ListIndex = 1
            For mode = 0 To 1
                panel.cboDisplayMode.ListIndex = mode
                displayMode = IIf(mode = 0, "MONO", "COLOR")
                For style = 0 To 1
                    panel.cboTableStyle.ListIndex = style
                    tableStyle = IIf(style = 0, "닫힌 표", "열린 표")
                    For finishing = 0 To 3
                        panel.cboFinishing.ListIndex = finishing
                        panel.cmdPreview.Value = True
                        Require panel.cmdExecute.Enabled, "preview failure: " & panel.lblPreviewStatus.Caption
                        expected.Clear
                        expected.Value2 = source.Value2
                        expected.Font.Italic = True
                        expected.Cells(2, 1).HorizontalAlignment = xlRight
                        hasTotal = (finishing = 1 Or finishing = 3)
                        NxDrawApplyTableOptions expected, NX_FEATURE_DRAW_BUSINESS_TABLE, displayMode, tableStyle, 2, 9, xlCenter, 9, xlGeneral, hasTotal
                        previewRows = 0
                        For r = 1 To 5
                            If VisibleControl(panel, "nxPrv_cell_" & r & "_1") Then previewRows = r
                        Next r
                        Require previewRows >= 2, "preview must show header and body"
                        For r = 1 To previewRows
                            sourceRow = r
                            If hasTotal And r = previewRows Then sourceRow = 8
                            For c = 1 To 3
                                Set label = panel.Controls("nxPrv_cell_" & r & "_" & c)
                                Set cell = expected.Cells(sourceRow, c)
                                Require label.Tag = source.Cells(sourceRow, c).Address(False, False), "wrong sample coordinate"
                                Require label.Font.Size = cell.Font.Size And label.Font.Bold = cell.Font.Bold, "preview font differs: row=" & r & " mode=" & mode & " style=" & style & " finishing=" & finishing
                                Require label.ForeColor = cell.Font.Color, "preview text color differs"
                                fill = RGB(255, 255, 255)
                                If cell.Interior.Pattern <> xlNone Then fill = cell.Interior.Color
                                Require label.BackColor = fill, "preview fill differs"
                            Next c
                        Next r
                        Require VisibleControl(panel, "nxPrv_line_v1") And VisibleControl(panel, "nxPrv_line_h1"), "preview missing internal grid"
                        Require VisibleControl(panel, "nxPrv_line_left") = (style = 0), "preview outside edge mismatch"
                        Require VisibleControl(panel, "nxPrv_line_total") = hasTotal, "preview total separator mismatch"
                        If hasTotal Then Require InStr(panel.lblPreview.Caption, "마지막 8행") > 0, "last-row sample not disclosed"
                        If finishing >= 2 Then Require InStr(panel.lblPreviewStatus.Caption, "표본 추정") > 0, "estimated AutoFit not disclosed"
                        Require SourceState(book) = before, "preview changed source, Saved, selection, widths or formats"
                        Require Workbooks.Count = loadedCount, "preview created a workbook"
                        Require ObjPtr(originalLabel) = ObjPtr(panel.Controls("nxPrv_cell_1_1")), "preview recreated controls"
                    Next finishing
                Next style
            Next mode
            panel.cboFinishing.ListIndex = 0
            Require panel.Controls("nxPrv_cell_1_1").Width > panel.Controls("nxPrv_cell_1_3").Width, "column ratios lost"
            Require panel.Controls("nxPrv_cell_4_1").TextAlign = 1, "numeric text must stay left aligned"
            Require panel.Controls("nxPrv_cell_2_1").TextAlign = 2, "header alignment must override source"
            panel.cboHeaderRows.ListIndex = 0
            Require panel.Controls("nxPrv_cell_2_1").TextAlign = 3, "automatic must preserve explicit body alignment"
        Case "preview_merged"
            source.Range("A1:B1").ClearContents
            source.Range("A1:B1").Merge
            source.Cells(1, 1).Value2 = "merged"
            before = SourceState(book)
            Set panel = CreatePanel(NX_FEATURE_DRAW_BUSINESS_TABLE)
            panel.cboFinishing.ListIndex = 1
            panel.cmdPreview.Value = True
            Require panel.cmdExecute.Enabled, "merged preview not ready"
            Require panel.Controls("nxPrv_cell_1_1").Caption = "merged", "merged anchor missing"
            Require panel.Controls("nxPrv_cell_1_2").Caption = vbNullString, "merged non-anchor duplicated"
            Require Not panel.Controls("nxPrv_cell_1_2").Visible, "merged non-anchor must not draw"
            Require Abs(panel.Controls("nxPrv_cell_1_1").Width - panel.Controls("nxPrv_cell_2_1").Width - panel.Controls("nxPrv_cell_2_2").Width) < 0.1, "merged width differs"
            Require Not VisibleControl(panel, "nxPrv_line_v1_1"), "line crosses merged header"
            Require VisibleControl(panel, "nxPrv_line_v1_2"), "unmerged line missing"
            Require Not VisibleControl(panel, "nxPrv_line_total"), "out-of-sample total must not be drawn"
            Require SourceState(book) = before, "merged preview mutated source"
        Case "preview_offsample_merge"
            source.Range("A6:B6").ClearContents
            source.Range("A6:B6").Merge
            before = SourceState(book)
            Set panel = CreatePanel(NX_FEATURE_DRAW_BUSINESS_TABLE)
            panel.cboFinishing.ListIndex = 1
            panel.cmdPreview.Value = True
            Require VisibleControl(panel, "nxPrv_line_total"), "off-sample merge hides last total"
            Require InStr(panel.lblPreview.Caption, "마지막 8행") > 0, "last row sample not explained"
            Require SourceState(book) = before, "off-sample preview changed source"
        Case "preview_vertical_merge"
            source.Range("A2:A3").ClearContents
            source.Range("A2:A3").Merge
            source.Cells(2, 1).Value2 = "세로" & vbLf & "병합"
            source.Cells(2, 1).WrapText = True
            before = SourceState(book)
            Set panel = CreatePanel(NX_FEATURE_DRAW_BUSINESS_TABLE)
            Require panel.cmdExecute.Enabled, "vertical preview not ready"
            Require panel.Controls("nxPrv_cell_2_1").Height >= 40, "vertical merge height missing"
            Require Not panel.Controls("nxPrv_cell_3_1").Visible, "vertical duplicate visible"
            Require Not VisibleControl(panel, "nxPrv_line_h2_1"), "line crosses vertical merge"
            Require SourceState(book) = before, "vertical preview mutated source"
        Case "preview_wrapping"
            source.Cells(2, 1).Value2 = "첫째 줄의 긴 설명입니다." & vbLf & "둘째 줄도 확인합니다."
            source.Cells(2, 1).WrapText = True
            before = SourceState(book)
            Set panel = CreatePanel(NX_FEATURE_DRAW_BUSINESS_TABLE)
            Set label = panel.Controls("nxPrv_cell_2_1")
            Require label.WordWrap, "wrapping disabled"
            Require label.Height > 20, "wrapped row height not increased"
            Require label.ControlTipText = source.Cells(2, 1).Text, "full text tooltip missing"
            Require panel.Controls("nxPrv_line_bottom").Top <= 372.1, "preview overlaps action/status area"
            Require SourceState(book) = before, "wrapped preview mutated source"
        Case "no_header"
            before = SourceState(book)
            Set panel = CreatePanel(NX_FEATURE_DRAW_BUSINESS_TABLE)
            Require panel.cboHeaderRows.ListIndex = 0, "default header must remain one row"
            panel.cboHeaderRows.Value = "없음 (0행)"
            Require Not panel.cboHeaderSize.Enabled And Not panel.cboHeaderAlign.Enabled, "zero header controls disabled"
            panel.cboHeaderRows.ListIndex = 1
            Require panel.cboHeaderSize.Enabled And panel.cboHeaderAlign.Enabled, "header controls restored"
            Require panel.cboHeaderSize.Value = "머리글 크기: 9", "header choice preserved"
            panel.cboHeaderRows.Value = "없음 (0행)"
            panel.cmdPreview.Value = True
            Require panel.cmdExecute.Enabled, "zero header preview rejected"
            Require SourceState(book) = before, "zero header preview mutated source"
            Require Not panel.Controls("nxPrv_cell_1_1").Font.Bold, "preview first row must be body"
            panel.cmdExecute.Value = True
            Require Not source.Cells(1, 1).Font.Bold, "first row must not be a header"
            Require source.Cells(1, 1).Font.Size = 9, "body font size"
            Require source.Cells(1, 1).Interior.Color = source.Cells(2, 1).Interior.Color, "first row body fill"
            Require source.Rows(1).Borders(xlEdgeBottom).Weight = xlThin, "no header separator"
            Require source.Cells(1, 1).Value2 = "항목", "first row value preserved"
            Require source.Cells(2, 2).Formula = "=1+2", "body formula preserved"
        Case "manual_range_apply"
            source.Font.Size = 11
            Set panel = CreatePanel(NX_FEATURE_DRAW_BUSINESS_TABLE)
            panel.txtRange.Value = "A2:B4"
            panel.cmdPreview.Value = True
            Require panel.cmdExecute.Enabled, "typed range preview rejected"
            Require Selection.Address = source.Address, "preview moved selection"
            panel.cmdExecute.Value = True
            Require source.Worksheet.Range("A2:B4").Font.Size = 9, "typed range not formatted"
            Require source.Cells(1, 1).Font.Size = 11 And source.Columns(3).Font.Size = 11, "formatted outside typed range"
            Require Selection.Address = source.Worksheet.Range("A2:B4").Address, "execution context not rebound"
        Case "format_restore"
            before = SourceState(book)
            Set formatBackup = New CNxFormatSnapshot
            formatBackup.Capture source
            NxDrawApplyTableOptions source, NX_FEATURE_DRAW_BUSINESS_TABLE, "COLOR", "닫힌 표", 2, 12, xlCenter, 11, xlRight, True, True
            formatBackup.Restore
            Require SourceState(book) = before, "format backup/restore differs after COM cache change"
        Case "title_defaults"
            before = SourceState(book)
            Set panel = CreatePanel(NX_FEATURE_DRAW_TITLE_TABLE)
            Require panel.cboHeaderSize.Value = "머리글 크기: 13", "title font default changed"
            ' MSForms converts the assigned StdFont to device units (13 -> 13.1 on this host).
            ' MSForms rounds 13pt to 12.75pt on the current native display.
            Require Abs(panel.Controls("nxPrv_cell_1_1").Font.Size - 13) <= 0.375, "title preview size=" & panel.Controls("nxPrv_cell_1_1").Font.Size & " expected13"
            Require Not panel.cboFinishing.Visible And Not panel.lblFinishingHelp.Visible, "title exposes irrelevant finishing"
            Require SourceState(book) = before, "title preview mutated source"
        Case "apply_200_cells"
            Set source = book.Worksheets(1).Range("A1:J20")
            source.Value2 = "fixture"
            source.Cells(2, 2).Formula = "=1+2"
            contentBefore = ContentState(source)
            startedAt = Timer
            NxDrawApplyTableOptions source, NX_FEATURE_DRAW_BUSINESS_TABLE, "MONO", "열린 표", 1, 9, xlCenter, 9, xlGeneral, True
            elapsedSeconds = Timer - startedAt
            If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400#
            Debug.Print "R68_TABLE_APPLY_200_CELLS_SECONDS=" & Format$(elapsedSeconds, "0.000")
            Require source.CountLarge = 200, "performance fixture must contain 200 cells"
            Require source.Font.Size = 9 And source.Cells(20, 1).Font.Bold, "200-cell apply styling failed"
            Require source.Borders(xlInsideVertical).LineStyle = xlContinuous, "200-cell apply lost internal grid"
            Require ContentState(source) = contentBefore, "200-cell apply changed content"
        Case Else
            NxRaiseContractError "Unknown r68 table fixture: " & name
    End Select
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & "|" & Err.Description
CleanUp:
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    Set panel = Nothing
    If Not expectedBook Is Nothing Then expectedBook.Close SaveChanges:=False
    If Not book Is Nothing Then book.Close SaveChanges:=False
    If Not priorBook Is Nothing Then priorBook.Activate
    If Not priorSelection Is Nothing Then priorSelection.Select
    On Error GoTo 0
End Function

Private Function CreatePanel(ByVal featureId As String) As FNxDrawTable
    Dim panel As New FNxDrawTable, context As CNxFeatureDialogContext, definition As CNxFeatureDefinition
    Set definition = NxDrawingFeatureDefinition(featureId)
    Set context = New CNxFeatureDialogContext
    context.Configure definition.FeatureId, definition.DialogId, definition.DialogVariant, definition.LaunchSurface = "workbench"
    panel.BindFeatureContext context
    Set CreatePanel = panel
End Function

Private Sub AssertGrid(ByVal source As Range, ByVal openTable As Boolean, ByVal headerRows As Long)
    Dim edge As Variant, r As Long
    For Each edge In Array(xlEdgeTop, xlInsideVertical)
        Require source.Borders(CLng(edge)).LineStyle = xlContinuous, "grid line missing: " & edge
        Require source.Borders(CLng(edge)).Weight = xlThin, "grid line not thin: " & edge
    Next edge
    For r = 1 To source.Rows.Count
        Require source.Rows(r).Borders(xlEdgeBottom).LineStyle = xlContinuous, "row bottom must be solid"
        Require source.Rows(r).Borders(xlEdgeBottom).Weight = IIf(r = headerRows Or r = source.Rows.Count, xlThick, xlThin), "row bottom weight: " & r
    Next r
    For Each edge In Array(xlEdgeLeft, xlEdgeRight)
        Require source.Borders(CLng(edge)).LineStyle = IIf(openTable, xlNone, xlContinuous), "outer edge mismatch: " & edge
    Next edge
End Sub

Private Function VisibleControl(ByVal panel As FNxDrawTable, ByVal name As String) As Boolean
    On Error Resume Next
    VisibleControl = panel.Controls(name).Visible
    On Error GoTo 0
End Function

Private Function ContentState(ByVal source As Range) As String
    Dim cell As Range, result As String
    For Each cell In source.Cells
        result = result & "|" & cell.Address & "="
        If IsError(cell.Formula) Then
            result = result & "#ERROR"
        Else
            result = result & CStr(cell.Formula)
        End If
        result = result & ":" & cell.MergeArea.Address
    Next cell
    ContentState = result
End Function

Private Function SourceState(ByVal book As Workbook) As String
    Dim cell As Range, edge As Variant, result As String
    result = CStr(book.Saved) & "|" & book.Sheets.Count & "|" & book.Worksheets(1).Shapes.Count
    result = result & "|" & ActiveWorkbook.Name & "|" & Selection.Address(External:=True)
    result = result & ContentState(book.Worksheets(1).Range("A1:C8"))
    For Each cell In book.Worksheets(1).Range("A1:C8")
        result = result & "|" & cell.Text & ":" & cell.NumberFormat & ":" & cell.Interior.Pattern & ":" & cell.Interior.Color
        result = result & ":" & cell.Font.Name & ":" & cell.Font.Size & ":" & cell.Font.Bold & ":" & cell.Font.Italic & ":" & cell.Font.Color
        result = result & ":" & cell.Font.Underline & ":" & cell.Font.Strikethrough & ":" & cell.HorizontalAlignment & ":" & cell.VerticalAlignment
        result = result & ":" & cell.ColumnWidth & ":" & cell.RowHeight & ":" & cell.WrapText
        For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
            result = result & ":" & cell.Borders(CLng(edge)).LineStyle & ":" & cell.Borders(CLng(edge)).Weight & ":" & cell.Borders(CLng(edge)).Color
        Next edge
    Next cell
    SourceState = result
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then NxRaiseContractError detail
End Sub
