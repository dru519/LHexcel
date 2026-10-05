Attribute VB_Name = "T_R74Features"
Option Explicit

Public Function Names() As String
    Names = "emphasis_number|emphasis_percent|emphasis_invalid|emphasis_protected|formula_notes|presets_toggle|presets_windows|presets_settings"
End Function

Private Sub Require(ByVal ok As Boolean, ByVal detail As String)
    If Not ok Then Err.Raise 5, "R74", detail
End Sub

Public Function RunCase(ByVal name As String) As String
    Dim wb As Workbook, ws As Worksheet, second As Worksheet, window As Excel.Window
    Dim panel As Object, r As Range, values As Variant, oldSettings As String, text As String, i As Long
    Dim percentage As Boolean, detail As String, saved As Boolean
    On Error GoTo Failed
    oldSettings = GetSetting("LHExcel", "ViewPresets", "Zooms", "__missing__")
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    Set r = ws.Range("B2:B5")
    r.Cells(1).Value2 = 1234.567: r.Cells(2).Value2 = -1234.567: r.Cells(3).Value2 = 0: r.Cells(4).Formula = "=6/4"
    r.NumberFormat = "General": ws.Columns("B").ColumnWidth = 35: r.Select
    Select Case name
        Case "emphasis_number", "emphasis_percent", "emphasis_invalid", "emphasis_protected"
            percentage = name = "emphasis_percent"
            Set panel = New FNxNumberFormat
            panel.BindTarget r, percentage, True
            panel.Controls("txtPlaces").Value = "2"
            panel.Controls("chkPercentage").Value = percentage
            If name = "emphasis_invalid" Then panel.Controls("txtPlaces").Value = "16"
            If name = "emphasis_protected" Then ws.Protect "r74"
            panel.NxProbeExecute
            If name = "emphasis_invalid" Or name = "emphasis_protected" Then
                Require r.Cells(1).NumberFormat = "General", "invalid/protected formatting mutated cells"
            Else
                Require Left$(r.Cells(1).Text, 2) = ChrW(&H25B2) & " ", "positive triangle missing"
                Require Left$(r.Cells(2).Text, 2) = ChrW(&H25BC) & " ", "negative triangle missing"
                Require InStr(r.Cells(2).Text, "-") = 0, "negative sign duplicated"
                Require InStr(r.Cells(3).Text, ChrW(&H25B2)) = 0 And InStr(r.Cells(3).Text, ChrW(&H25BC)) = 0, "zero has triangle"
                Require InStr(r.Cells(1).NumberFormat, "[Red]") > 0 And InStr(r.Cells(1).NumberFormat, "[Blue]") > 0, "colors missing"
                Require ((InStr(r.Cells(1).Text, "%") > 0) = percentage), "percent option mismatch"
            End If
            Require r.Cells(1).Value2 = 1234.567 And r.Cells(2).Value2 = -1234.567, "numeric value changed"
            Require r.Cells(4).Formula = "=6/4", "formula changed"
        Case "formula_notes"
            ws.Range("D4").Formula = "=3+4": ws.Range("D4").AddComment "outside"
            NxFormulaCommentsApply r, True
            Require r.Cells(4).Comment.Text = "=6/4", "formula note prefix/content"
            Require r.Cells(4).Comment.Visible, "show note failed"
            Require r.Cells(1).Comment Is Nothing, "constant acquired note"
            NxFormulaCommentsApply r, False
            Require Not r.Cells(4).Comment.Visible, "hide note failed"
            Require ws.Range("D4").Comment.Text = "outside", "outside selection changed"
        Case "presets_toggle", "presets_windows"
            SaveSetting "LHExcel", "ViewPresets", "Zooms", "75,100,125,150,200"
            ActiveWindow.Zoom = 90: wb.Saved = True
            NxCmdView "NX_VIEW_PRESET_3"
            Require ActiveWindow.Zoom = 125, "preset zoom"
            NxCmdView "NX_VIEW_PRESET_3"
            Require ActiveWindow.Zoom = 90, "repeat did not restore"
            NxCmdView "NX_VIEW_PRESET_1": NxCmdView "NX_VIEW_PRESET_5": NxCmdView "NX_VIEW_PRESET_5"
            Require ActiveWindow.Zoom = 90, "switch preset lost original"
            Require wb.Saved, "preset changed Saved state"
            If name = "presets_windows" Then
                Set window = wb.NewWindow: window.Zoom = 110
                NxViewPresetApply window, 2
                Require window.Zoom = 100, "second window not applied"
                NxViewPresetApply window, 2
                Require window.Zoom = 110, "second window not restored"
                window.Close
                Require ActiveWindow.Zoom = 90, "other window changed"
            End If
        Case "presets_settings"
            Set panel = New FNxViewPresets: panel.BindSettings
            For i = 1 To 5: panel.Controls("txtZoom" & CStr(i)).Value = CStr(60 + i * 10): Next i
            panel.NxProbeExecute
            values = NxViewPresetValues()
            Require values(1) = 70 And values(5) = 110, "settings failed"
            Set panel = New FNxViewPresets: panel.BindSettings
            panel.Controls("txtZoom3").Value = "401"
            panel.NxProbeExecute
            values = NxViewPresetValues()
            Require values(3) = 90, "invalid partially saved"
            SaveSetting "LHExcel", "ViewPresets", "Zooms", "broken"
            values = NxViewPresetValues(): Require values(1) = 75 And values(5) = 200, "corrupt settings fallback failed"
        Case Else: Err.Raise 5, , "Unknown case"
    End Select
    RunCase = "PASS|" & name
    GoTo Cleanup
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
Cleanup:
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    If Not wb Is Nothing Then NxViewPresetForgetWorkbook wb: wb.Close False
    If oldSettings = "__missing__" Then DeleteSetting "LHExcel", "ViewPresets", "Zooms" Else SaveSetting "LHExcel", "ViewPresets", "Zooms", oldSettings
End Function
