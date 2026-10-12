Attribute VB_Name = "T_HangulFocusOptions"
Option Explicit

Public Function Names() As String
    Names = "hangul_settings|hangul_payload|focus_cross|focus_horizontal|focus_vertical|focus_preview|window_restore|table_cf"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, sheet As Worksheet, settings As CNxHangulSettings
    Dim model As CNxFocusSettingsModel, engine As CNxFocusRuleEngine, selection As CNxFocusSelection
    Dim form As FNxHangulSettings, focusForm As FNxFocusSettings, payload As String, shape As String
    Dim firstWindow As Window, secondWindow As Window, guard As CNxStateGuard
    Dim rule As FormatCondition, backup As CNxRangeFormatBackup
    Dim request As CNxDrawingRequest, drawing As CNxDrawingFeatureCommand, command As INxFeatureCommand
    Dim approval As CNxApproval, result As CNxResult
    On Error GoTo Failed
    Select Case name
        Case "hangul_settings"
            Set settings = New CNxHangulSettings
            settings.Configure "중고딕", "HFT", 13, True, True, "공문형", 2, True
            settings.Seal
            NxHangulSaveSettings settings
            Set settings = NxHangulLoadSettings()
            Require settings.HeaderRows = 2 And settings.TotalRow, "settings roundtrip"
            Set form = New FNxHangulSettings
            form.BindSettings settings
            Require form.Controls("cboHeaderRows").Value = "2", "header choice"
            Require Not form.Controls("chkTotalRow").Value, "total row unchecked even with saved true"
            Require Not form.Controls("chkIncludeHidden").Value, "hidden rows excluded by default"
            Require form.Controls("cboStyle").Value = "스타일: 열린 표", "table style label"
            Require form.Controls("txtRange").Top < form.Controls("cboStyle").Top, "range is first"
            form.Controls("cboStyle").ListIndex = 1
            Require form.Controls("cboStyle").Value = "스타일: 닫힌 표", "closed choice"
            Unload form: Set form = New FNxHangulSettings
            Set settings = New CNxHangulSettings
            settings.Configure "중고딕", "HFT", 9, True, True, "공문형", 6, False
            settings.Seal
            form.BindSettings settings
            Require form.Controls("cboSize").Value = "9", "valid saved size outside suggestions"
            Require form.Controls("cboHeaderRows").Value = "6", "valid saved header outside suggestions"
            form.SetTransferMode
            form.Show vbModeless
            Require form.Controls("txtRange").Visible, "unified range visible"
        Case "hangul_payload"
            Set book = Workbooks.Add: Set sheet = book.Worksheets(1)
            sheet.Range("A1:B5").Value = "sample"
            Set settings = New CNxHangulSettings
            settings.Configure "중고딕", "HFT", 13, True, False, "공문형", 2, True
            settings.Seal
            payload = NxHwpxSerializeWithSettings(sheet.Range("A1:B5"), settings)
            Require InStr(payload, """header_rows"":2") > 0 And InStr(payload, """total_row"":true") > 0, "payload roles"
            sheet.Rows(1).Hidden = True: sheet.Rows(5).Hidden = True
            payload = NxHwpxSerializeWithSettings(sheet.Range("A1:B5"), settings)
            Require InStr(payload, """header_rows"":1") > 0 And InStr(payload, """total_row"":false") > 0, "hidden original roles"
        Case "focus_cross", "focus_horizontal", "focus_vertical"
            shape = NX_FOCUS_SHAPE_CROSS
            If name = "focus_horizontal" Then shape = NX_FOCUS_SHAPE_HORIZONTAL
            If name = "focus_vertical" Then shape = NX_FOCUS_SHAPE_VERTICAL
            Set book = Workbooks.Add: Set sheet = book.Worksheets(1)
            sheet.Range("A1:J10").Interior.Color = RGB(240, 230, 210)
            Set model = NxFocusCreateSettings(shape, NX_FOCUS_STYLE_OUTER_CELLS, "#5FC8D8", 80, False, True)
            Set engine = New CNxFocusRuleEngine: engine.Configure model
            Set selection = New CNxFocusSelection
            Require selection.Configure(sheet.Range("D4:E5")), "selection"
            Require engine.ApplySelection(selection), "apply"
            Require engine.ManagedRuleCount(sheet) = 4, "four edge rules"
            Require sheet.Range("A4").DisplayFormat.Interior.Color = RGB(240, 230, 210), "no fill overwrite"
            If shape <> NX_FOCUS_SHAPE_VERTICAL Then Require sheet.Range("A3").DisplayFormat.Interior.Color = NxFocusConditionalColor(model), "row above shading"
            If shape <> NX_FOCUS_SHAPE_HORIZONTAL Then Require sheet.Range("C1").DisplayFormat.Interior.Color = NxFocusConditionalColor(model), "column left shading"
            Require sheet.Range("D4").DisplayFormat.Interior.Color = RGB(240, 230, 210), "selected cells clear"
            If shape = NX_FOCUS_SHAPE_CROSS Then
                Require sheet.Range("D3").DisplayFormat.Interior.Color = NxFocusConditionalColor(model), "above selection shaded"
                Require sheet.Range("C4").DisplayFormat.Interior.Color = NxFocusConditionalColor(model), "left of selection shaded"
                Set selection = New CNxFocusSelection
                Require selection.Configure(sheet.Rows("4:5"), True), "whole row selection"
                Require engine.ApplySelection(selection), "whole row apply"
                Require sheet.Range("D3").DisplayFormat.Interior.Color = NxFocusConditionalColor(model), "whole row above shaded"
                Require sheet.Range("D6").DisplayFormat.Interior.Color = NxFocusConditionalColor(model), "whole row below shaded"
            End If
            Set selection = New CNxFocusSelection
            Require selection.Configure(sheet.Range("G7")), "move selection"
            Require engine.ApplySelection(selection), "move"
            Require engine.ManagedRuleCount(sheet) = 4, "no duplicate rules"
            Require engine.DeleteManagedRules(sheet), "remove"
            Require sheet.Range("A4").Interior.Color = RGB(240, 230, 210), "original preserved"
            Set model = NxFocusCreateSettings(shape, NX_FOCUS_STYLE_WIDE, "#5FC8D8", 80, False, True)
            engine.Configure model
            Require engine.ApplySelection(selection), "existing stripe retained"
            Require engine.ManagedRuleCount(sheet) = 2, "existing two rules retained"
        Case "focus_preview"
            Set focusForm = New FNxFocusSettings
            focusForm.Controls("cboStyle").ListIndex = 1
            Require Not focusForm.Controls("chkSelectedArea").Enabled, "outline area disabled"
            focusForm.Controls("cboShape").ListIndex = 1
            focusForm.Controls("cboShape").ListIndex = 2
            focusForm.Controls("cboStyle").ListIndex = 0
        Case "window_restore"
            Set book = Workbooks.Add: Set sheet = book.Worksheets(1)
            Set firstWindow = book.Windows(1): Set secondWindow = book.NewWindow
            secondWindow.Activate
            sheet.Range("D20").Select: secondWindow.ScrollRow = 12
            Set guard = New CNxStateGuard
            firstWindow.Activate
            guard.Restore
            Require Application.ActiveWindow.Hwnd = secondWindow.Hwnd, "exact second window restored: active=" & Application.ActiveWindow.Hwnd & ", expected=" & secondWindow.Hwnd & ", recovery=" & guard.Recovery
            Require Application.Selection.Address = "$D$20", "selection preserved"
            Require secondWindow.ScrollRow = 12, "scroll preserved"
            NxResetCommandRegistry
            NxRouteCommand "NX-CMD-STYLE-FONT-SIZE-9"
            Require Application.ActiveWindow.Hwnd = secondWindow.Hwnd, "font action keeps second window"
            NxRouteCommand "NX-CMD-STYLE-FILL-LIGHT-GREEN"
            Require Application.ActiveWindow.Hwnd = secondWindow.Hwnd, "fill action keeps second window"
            NxQuickFormatUndoLast
            Require Application.ActiveWindow.Hwnd = secondWindow.Hwnd, "undo keeps second window"
        Case "table_cf"
            Set book = Workbooks.Add: Set sheet = book.Worksheets(1)
            sheet.Range("A1:D100").Value = 1
            sheet.Range("A1:D100").Font.Size = 12
            Set rule = sheet.Range("A1:D100").FormatConditions.Add(xlExpression, , "=A1>0")
            rule.Interior.Color = vbYellow
            Set backup = New CNxRangeFormatBackup
            backup.Capture sheet.Range("A1:D80")
            NxDrawApplyBusinessTable sheet.Range("A1:D80")
            Require sheet.Range("A1").Font.Size = 9, "table applied with CF"
            Require rule.AppliesTo.Address = "$A$1:$D$100", "CF scope preserved"
            Require sheet.Range("A1").DisplayFormat.Interior.Color = vbYellow, "CF display retained"
            backup.Restore
            Require sheet.Range("A1").Font.Size = 12, "format restored"
            Require sheet.Cells.FormatConditions.Count = 1, "CF not duplicated"
            Require rule.AppliesTo.Address = "$A$1:$D$100", "CF scope preserved after restore"
            backup.Dispose
            sheet.Range("A1:D80").Select
            Set request = New CNxDrawingRequest
            request.ConfigureRange NX_FEATURE_DRAW_BUSINESS_TABLE, sheet.Range("A1:D80"), False, True
            Set drawing = New CNxDrawingFeatureCommand: drawing.Configure request
            Set command = drawing
            Set approval = NxCommandRunner.Approve(command)
            Set result = NxCommandRunner.Run(command, approval)
            Require result.Outcome = NxSuccess, "runner accepts CF table: " & CStr(result.Outcome)
        Case Else: Err.Raise 5, , "unknown test"
    End Select
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & ":" & Err.Description
Clean:
    On Error Resume Next
    If Not engine Is Nothing Then engine.DeleteManagedRules sheet
    If Not form Is Nothing Then Unload form
    If Not focusForm Is Nothing Then Unload focusForm
    If Not backup Is Nothing Then backup.Dispose
    Set guard = Nothing
    If Not book Is Nothing Then book.Close False
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 106, , detail
End Sub
