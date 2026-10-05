Attribute VB_Name = "T_R69UserFlows"
Option Explicit

Public Function Names() As String
    Names = "date_inputs|date_invalid|special_engines|money_popup|mask_popup|age_popup|date_popup|number_popup|resize_popup|file_output_popup|consolidate_popup|calculator_route|focus_cf"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, other As Workbook, checkBook As Workbook, target As Range, panel As Object
    Dim item As Variant, parsed As Date, expected As Date, enabled As Variant
    Dim output As String, inputText As String, before As String, detail As String, color As Long
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    book.SaveAs Environ$("LHEXCEL_PROFILE_ROOT") & "\" & name & ".xlsx", xlOpenXMLWorkbook
    Set target = book.Worksheets(1).Range("B2:C3")
    target.Value2 = 1234.56789: target.Select
    Select Case name
        Case "date_inputs"
            expected = DateSerial(2024, 6, 17)
            For Each item In Array("240617", 240617, "20240617", 20240617, "2024.06.17", "2024.06.17.", _
                "2024.6.17", "2024-06-17", "24-06-17", "2024/06/17", "2024년 6월 17일", "240617-3000000", "240617-3******", "240617-*******")
                Check NxDateTryParseValue(item, DateSerial(2026, 9, 10), parsed), "Rejected supported date " & CStr(item)
                Check parsed = expected, "Date mismatch " & CStr(item)
            Next item
            Check NxDateTryParseValue(CDbl(expected) + 0.5, Date, parsed), "Serial time rejected"
            Check parsed = expected + TimeSerial(12, 0, 0), "Serial time lost"
            Check NxDateTryParseValue("2024-06-17 12:34:56", Date, parsed), "Timestamp rejected"
            Check parsed = expected + TimeSerial(12, 34, 56), "Timestamp changed"
            Check NxDateTryParseValue("2024-02-29", Date, parsed), "Leap day rejected"
        Case "date_invalid"
            For Each item In Array("2024-02-30", "230229", "20241301", "13/6/2024", "240617x", "2024-6", "2024-06-17 24:00", "240617-abc0000", "240617-*1*****")
                Check Not NxDateTryParseValue(item, DateSerial(2026, 9, 10), parsed), "Invalid date accepted " & CStr(item)
            Next item
        Case "special_engines"
            For Each item In T_R62DataSpecial.TestNames()
                T_R62DataSpecial.RunCase CStr(item)
            Next item
        Case "money_popup", "mask_popup", "age_popup"
            target.NumberFormat = "@"
            Select Case name
                Case "money_popup": inputText = "1,234원": output = "금1,234원정(금일천이백삼십사원정)": before = NX_FEATURE_DATA_KOREAN_MONEY
                Case "mask_popup": inputText = "010-1234-5678": output = "010-****-5678": before = NX_FEATURE_DATA_PRIVACY_MASK
                Case "age_popup": inputText = "900520": output = "만 36세": before = NX_FEATURE_DATA_AGE
            End Select
            target.Value2 = inputText
            ' Exercise the actual feature launch surface, not just a direct helper.
            NxRouteFeatureSurface NxDataFeatureDefinition(before)
            If name = "age_popup" Then Set panel = FindForm("FNxAgeCalculator") Else Set panel = FindForm("FNxDataSpecial")
            Check Not panel Is Nothing, "Data popup route missing"
            panel.cboOutputMode.Value = "원본 변경"
            If name = "age_popup" Then
                panel.txtReferenceDate.Value = "2026-09-10"
                panel.chkAgePrefix.Value = True: panel.chkAgeSuffix.Value = True
            ElseIf name = "money_popup" Then
                panel.chkPrefix.Value = True: panel.chkSuffix.Value = True: panel.chkMixed.Value = True
            End If
            If name = "mask_popup" Then panel.cboMaskType.Value = "휴대폰번호"
            Check panel.cmdExecute.Enabled, "Execution still requires Preview"
            panel.NxProbeExecute
            Check target.Cells(1, 1).Value2 = output And target.Cells(2, 2).Value2 = output, "Popup did not transform current range"
        Case "date_popup"
            target.NumberFormat = "@": target.Value2 = "240617"
            NxCmdSelection "NX_DATA_DATE_CONVERT"
            Set panel = FindForm("FNxDataNormalize")
            Check Not panel Is Nothing, "Date popup missing"
            Check panel.cboPreset.Value = "날짜 정리", "Date preset not selected"
            Check panel.cboOutputMode.Value = "새 통합문서", "Default output changed"
            panel.cboDateFormat.Value = "yymmdd"
            panel.NxProbeDefaults
            Check panel.cboPreset.Value = "날짜 정리" And Not panel.cboPreset.Enabled, "Reset changed date conversion mode"
            Check panel.cboDateFormat.Value = "yyyy-mm-dd", "Reset lost default date format"
            panel.cboDateFormat.Value = "yyyy.mm.dd"
            panel.cboOutputMode.Value = "새 시트 만들기"
            Check panel.cmdExecute.Enabled, "Normalization still requires Preview"
            panel.NxProbeExecute
            Check book.Worksheets.Count = 2, "New date sheet missing"
            Check target.Cells(1, 1).Value2 = "240617", "Date conversion changed source"
            Check ActiveSheet.Range("A1").Value2 = CDbl(DateSerial(2024, 6, 17)), "Date output is not an Excel date"
            Check ActiveSheet.Range("A1").NumberFormat = "yyyy.mm.dd", "Date output format missing"
        Case "number_popup"
            Set panel = New FNxNumberFormat: panel.BindTarget target, True
            panel.txtPlaces.Value = "3": panel.NxProbeExecute
            Check target.NumberFormat = "0.000%" And target.Cells(1, 1).Value2 = 1234.56789, "Percent mutated value or used wrong precision"
            Set panel = New FNxNumberFormat: panel.BindTarget target, False
            panel.txtPlaces.Value = "0": panel.NxProbeExecute
            Check target.NumberFormat = "#,##0", "Decimal precision zero failed"
            Set panel = New FNxNumberFormat: panel.BindTarget target, False
            panel.txtPlaces.Value = "16": panel.NxProbeExecute
            Check target.NumberFormat = "#,##0", "Invalid precision changed format"
            Unload panel
            target.Value2 = 123
            NxCmdNumberFormat "RB_EDIT_NUMBERFORMAT_KOREAN"
            Check target.Cells(1, 1).Text = "일백이십삼", "Korean display: " & target.Cells(1, 1).Text
            NxCmdNumberFormat "RB_EDIT_NUMBERFORMAT_HANMOON"
            Check target.Cells(1, 1).Text = "一百二十三", "Chinese display: " & target.Cells(1, 1).Text
        Case "resize_popup"
            Set panel = New FNxResize: panel.BindTarget target
            panel.chkWidth.Value = True: panel.txtWidth.Value = "16"
            panel.chkHeight.Value = True: panel.txtHeight.Value = "30"
            panel.chkFont.Value = True: panel.txtFont.Value = "12"
            panel.NxProbeExecute
            Check Abs(target.ColumnWidth - 16) < 0.2 And Abs(target.RowHeight - 30) < 0.5 And target.Font.Size = 12, "Resize values wrong"
            Check target.Cells(1, 1).Value2 = 1234.56789, "Resize changed cell contents"
        Case "file_output_popup"
            target.Interior.Color = RGB(255, 230, 128): target.Font.Color = vbBlack
            target.ColumnWidth = 16: target.RowHeight = 25
            Set panel = New FNxFileOutput: panel.BindFeature NX_FEATURE_FILE_RANGE_PNG
            output = Environ$("LHEXCEL_PROFILE_ROOT") & "\output_direct.png"
            panel.txtOutputPath.Value = output
            Check panel.cmdExecute.Enabled, "Image output requires Preview"
            panel.NxProbeExecute
            Check Len(Dir$(output)) > 0, "Image output missing"
        Case "consolidate_popup"
            Set other = Workbooks.Add(xlWBATWorksheet)
            other.Worksheets(1).Range("A1").Value2 = "통합 시험"
            inputText = Environ$("LHEXCEL_PROFILE_ROOT") & "\input.xlsx"
            other.SaveAs inputText, xlOpenXMLWorkbook: other.Close False: Set other = Nothing
            book.Activate: target.Select
            Set panel = New FNxFileConsolidate: panel.BindFeature
            panel.lstInputFiles.AddItem inputText
            output = Environ$("LHEXCEL_PROFILE_ROOT") & "\consolidated.xlsx"
            panel.txtOutputPath.Value = output
            Check panel.cmdExecute.Enabled, "Consolidation requires Preview"
            panel.NxProbeExecute
            Check Len(Dir$(output)) > 0, "Consolidation output missing"
            Set checkBook = Workbooks.Open(output, UpdateLinks:=0, ReadOnly:=True)
            Check checkBook.Worksheets(1).Range("A1").Value2 = "통합 시험", "Consolidation content changed"
            checkBook.Close False: Set checkBook = Nothing
        Case "calculator_route"
            Set panel = New FNxNumberFormat
            panel.Tag = "nx1|feature|NX-UTIL-CALCULATOR"
            NxRibbonGetEnabled panel, enabled
            Check CBool(enabled), "Static calculator control disabled"
            Unload panel
            NxRouteFeatureSurface NxCreateProductRegistry().FeatureById(NX_FEATURE_UTIL_CALCULATOR)
            Set panel = FindForm("FNxCalculator")
            Check Not panel Is Nothing, "Calculator route missing"
        Case "focus_cf"
            Check Not NxHostBridgeCompatible(), "No-DLL test requires unavailable compatible DLL"
            target.Value2 = 1
            target.Select: NxFocusControllerEnable
            Check NxFocusControllerBackend() = "cf", "CF not selected: " & NxFocusControllerBackendDiagnostic()
            color = NxFocusConditionalColor(NxFocusLoadSettings())
            Check book.Worksheets(1).Range("A2").DisplayFormat.Interior.Color = color, "CF row not painted"
            Check book.Worksheets(1).Range("B1").DisplayFormat.Interior.Color = color, "CF column not painted"
            Set target = book.Worksheets(1).Range("F6:G7")
            target.Select: NxFocusControllerSelectionChanged target
            Check book.Worksheets(1).Range("A6:A7").DisplayFormat.Interior.Color = color, "Moved two-row band not painted"
            Check book.Worksheets(1).Range("F1:G1").DisplayFormat.Interior.Color = color, "Moved two-column band not painted"
            NxFocusControllerDisable
            Check book.Worksheets(1).Cells.FormatConditions.Count = 0, "Residual CF after Off"
    End Select
    GoTo Cleanup
Failed:
    detail = CStr(Err.Number) & "|" & Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    NxFocusControllerDisable
    For Each panel In VBA.UserForms: Unload panel: Next panel
    If Not other Is Nothing Then other.Close False
    If Not checkBook Is Nothing Then checkBook.Close False
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    If Len(detail) = 0 Then RunCase = "PASS|" & name Else RunCase = "FAIL|" & name & "|" & detail
End Function

Private Function FindForm(ByVal name As String) As Object
    Dim panel As Object
    For Each panel In VBA.UserForms
        If TypeName(panel) = name Then Set FindForm = panel: Exit Function
    Next panel
End Function

Private Sub Check(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise 5, "T_R69UserFlows", detail
End Sub
