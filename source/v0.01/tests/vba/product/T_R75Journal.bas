Attribute VB_Name = "T_R75Journal"
Option Explicit
Public FailDraw As Boolean
Public DrawFaultObserved As Boolean
Public Function Names() As String
    Names = "plain|validation_inside|validation_outside|input_message|conditional_inside|conditional_outside|table_inside|table_outside|merged_selection|multi_area|value_only|range_backup|range_backup_merged|range_backup_mixed|range_backup_execute_error|range_backup_execute_error_dirty|range_backup_execute_error_shape|range_backup_execute_error_inflated|range_backup_execute_error_chart|range_backup_execute_error_picture|range_backup_execute_error_million|range_backup_fallback_control|range_backup_fallback_objects|range_backup_fallback_size"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim wb As Workbook, ws As Worksheet, target As Range, journal As New CNxCellRollbackJournal
    Dim context As CNxExecutionContext, plan As CNxPlan, authority As New Collection
    Dim captured As Boolean, reject As Boolean, detail As String, mask As NxCellMutationMask
    Dim rangeMode As Boolean, feature As String, beforeBooks As Long
    Dim execution As CNxResult
    Dim chart As ChartObject, picture As Shape, series As String
    Dim fallback As Boolean, index As Long
    On Error GoTo Failed
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    Set target = ws.Range("B2:D5"): target.Value2 = "original"
    rangeMode = (Left$(name, 12) = "range_backup")
    fallback = (Left$(name, 21) = "range_backup_fallback")
    If rangeMode Then
        Set target = ws.Range("B2:U21"): target.Value2 = "original": target.Font.Size = 9
        If name = "range_backup_merged" Then ws.Range("C3:D4").ClearContents: ws.Range("C3:D4").Merge
        If name = "range_backup_mixed" Then
            target.Cells(7).Font.ThemeColor = xlThemeColorAccent2
            target.Cells(8).Interior.Pattern = xlPatternGray25
            target.Cells(9).Orientation = 45
        End If
        If name = "range_backup_execute_error_shape" Then
            With ws.Shapes.AddShape(1, 800, 800, 40, 30)
                .Name = "preserve-shape"
                .Fill.ForeColor.RGB = RGB(20, 60, 180)
            End With
        End If
        If name = "range_backup_execute_error_inflated" Then ws.Range("A1:AX8000").Interior.Color = vbYellow
        If name = "range_backup_execute_error_million" Then ws.Range("A1:AX20000").Interior.Color = vbYellow
        If name = "range_backup_fallback_size" Then ws.Range("A1:AX20001").Interior.Color = vbYellow
        If name = "range_backup_fallback_control" Then ws.Shapes.AddFormControl xlButtonControl, 900, 100, 100, 30
        If name = "range_backup_fallback_objects" Then
            For index = 1 To 17
                ws.Shapes.AddShape 1, 900, 100 + index * 10, 10, 10
            Next index
        End If
        If name = "range_backup_execute_error_chart" Then
            ws.Range("BA1:BA3").Value2 = 7
            Set chart = ws.ChartObjects.Add(900, 100, 150, 100)
            chart.Placement = xlFreeFloating
            chart.Chart.SetSourceData ws.Range("BA1:BA3")
            series = chart.Chart.SeriesCollection(1).Formula
        End If
        If name = "range_backup_execute_error_picture" Then
            ws.Range("B2:D4").CopyPicture Appearance:=xlScreen, Format:=xlPicture
            ws.Paste
            Set picture = ws.Shapes(ws.Shapes.Count)
            picture.Placement = xlFreeFloating
            picture.Left = 900: picture.Top = 100
            If picture.Type <> msoPicture Then Err.Raise 5, , "Picture fixture missing"
            Application.CutCopyMode = False
        End If
    End If
    Select Case name
        Case "validation_inside": target.Cells(12).Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, 1, 9: reject = True
        Case "validation_outside": Set target = ws.Range("B2"): ws.Range("Z50").Validation.Add xlValidateWholeNumber, xlValidAlertStop, xlBetween, 1, 9
        Case "input_message": target.Cells(12).Validation.Add xlValidateInputOnly: target.Cells(12).Validation.InputMessage = "help"
        Case "conditional_inside": target.Cells(12).FormatConditions.Add xlExpression, Formula1:="=TRUE": reject = True
        Case "conditional_outside": ws.Range("Z50").FormatConditions.Add xlExpression, Formula1:="=TRUE"
        Case "table_inside": ws.ListObjects.Add xlSrcRange, ws.Range("D5:F8"), , xlYes: reject = True
        Case "table_outside": ws.ListObjects.Add xlSrcRange, ws.Range("Z50:AB53"), , xlYes
        Case "merged_selection": ws.Range("A1:B2").ClearContents: ws.Range("A1:B2").Merge
        Case "multi_area": Set target = Union(target, ws.Range("G8:H10"))
    End Select
    target.Select
    If name = "merged_selection" Then
        If Application.Selection.Address <> "$A$1:$D$5" Then Err.Raise 5, , "Excel merge selection did not expand as observed"
        Set target = Application.Selection
    End If
    mask = NxCellFontMask Or NxCellBorderMask
    If rangeMode Then mask = mask Or NxCellInteriorMask Or NxCellAlignmentMask
    If name = "value_only" Then mask = NxCellValueMask
    Set context = NxContextFactory.CaptureCurrent()
    context.BindSelectionAuthority authority
    feature = "NX-TEST-CORE"
    If rangeMode Then feature = "NX-DRAW-BUSINESS-TABLE"
    Set plan = NxCreatePlan(feature, context.Fingerprint, context.SourceStateDigest)
    If rangeMode Then
        plan.AddSideEffect NxCreateFeatureCellEffect(feature, context.SelectionTargetIdentity, mask, CLng(target.CountLarge))
    Else
        plan.AddSideEffect NxCreateCellEffect(context.SelectionTargetIdentity, mask)
    End If
    plan.Seal
    beforeBooks = Workbooks.Count
    wb.Saved = (name <> "range_backup_execute_error_dirty")
    On Error Resume Next
    journal.Capture context, plan, authority
    captured = (Err.Number = 0): detail = Err.Description: Err.Clear
    On Error GoTo Failed
    If captured = reject Then Err.Raise 5, , "capture decision mismatch: " & detail & "; requested=" & target.Address & "; selected=" & Application.Selection.Address
    If captured Then
        If rangeMode Then
            If fallback Then
                If Workbooks.Count <> beforeBooks Then Err.Raise 5, , "Fallback guard bypassed"
            Else
                If Workbooks.Count <> beforeBooks + 1 Then Err.Raise 5, , "range backup path not used"
            End If
        End If
        If journal.Count <> target.CountLarge Then Err.Raise 5, , "capture cell count mismatch"
        If Left$(name, 26) = "range_backup_execute_error" Then
            FailDraw = True: DrawFaultObserved = False
            Set execution = NxDrawRun(NX_FEATURE_DRAW_BUSINESS_TABLE, target)
            FailDraw = False
            If Not DrawFaultObserved Then Err.Raise 5, , "drawing fault was not reached"
            If execution.Outcome <> NxEnvironmentError Or execution.SourceChanged Then Err.Raise 5, , "runner failed to restore source: outcome=" & CStr(execution.Outcome) & "; changed=" & CStr(execution.SourceChanged) & "; " & execution.Recovery
            ' Format restoration is an Excel edit. Do not mark a user workbook
            ' saved merely because the owned range was restored.
            If wb.Saved Then Err.Raise 5, , "rollback incorrectly suppressed the save prompt"
            If WorksheetFunction.CountIf(target, "original") <> target.CountLarge Then Err.Raise 5, , "rollback changed source values"
        Else
            If name = "value_only" Then target.Value2 = "changed" Else target.Font.Size = 25
            detail = journal.Rollback()
        End If
        If Len(detail) > 0 Or Len(journal.VerifyResidual()) > 0 Then Err.Raise 5, , "rollback mismatch"
        journal.ReleaseResources
        If Workbooks.Count <> beforeBooks Then Err.Raise 5, , "backup cleanup leaked"
        If name = "range_backup_execute_error_shape" Then
            If ws.Shapes.Count <> 1 Then Err.Raise 5, , "Shape count changed"
            With ws.Shapes("preserve-shape")
                If .Left <> 800 Or .Top <> 800 Or .Width <> 40 Or .Height <> 30 Or .Fill.ForeColor.RGB <> RGB(20, 60, 180) Then Err.Raise 5, , "Shape state changed"
            End With
        End If
        If name = "range_backup_execute_error_inflated" Then
            If ws.Range("AX8000").Interior.Color <> vbYellow Then Err.Raise 5, , "Outside format changed"
        End If
        If name = "range_backup_execute_error_million" Then
            If ws.Range("AX20000").Interior.Color <> vbYellow Then Err.Raise 5, , "Outside format changed"
        End If
        If name = "range_backup_execute_error_chart" Then
            If ws.ChartObjects.Count <> 1 Or chart.Chart.SeriesCollection(1).Formula <> series Then Err.Raise 5, , "Chart changed"
            If chart.Left <> 900 Or chart.Top <> 100 Or chart.Width <> 150 Or chart.Height <> 100 Then Err.Raise 5, , "Chart geometry changed"
        End If
        If name = "range_backup_execute_error_picture" Then
            If ws.Shapes.Count <> 1 Or picture.Left <> 900 Or picture.Top <> 100 Then Err.Raise 5, , "Picture changed"
        End If
    End If
    RunCase = "PASS|" & name
    GoTo Cleanup
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
Cleanup:
    On Error Resume Next
    FailDraw = False
    journal.ReleaseResources
    If Not wb Is Nothing Then wb.Close False
End Function
