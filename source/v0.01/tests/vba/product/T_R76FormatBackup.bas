Attribute VB_Name = "T_R76FormatBackup"
Option Explicit

Public Function Names() As String
    Names = "plain|theme|latent_borders|mixed|merged|large"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim wb As Workbook, ws As Worksheet, target As Range, authority As New Collection
    Dim expected As New CNxCellRollbackJournal, backup As New CNxRangeFormatBackup
    Dim context As CNxExecutionContext, plan As CNxPlan, mask As NxCellMutationMask
    Dim baselineBooks As Long, captured As Double, restored As Double, started As Double
    Dim beforeSaved As Boolean, detail As String, cell As Range
    On Error GoTo Failed
    baselineBooks = Workbooks.Count
    Set wb = Workbooks.Add(xlWBATWorksheet): Set ws = wb.Worksheets(1)
    Set target = ws.Range("B2:F10")
    If name = "large" Then Set target = ws.Range("A1:AX1800")
    target.Value2 = "original": target.Font.Size = 9
    target.Cells(1).Formula = "=1+2"
    Select Case name
        Case "theme": target.Font.ThemeColor = xlThemeColorAccent2: target.Font.TintAndShade = 0.4: target.Interior.ThemeColor = xlThemeColorAccent3
        Case "latent_borders"
            With target.Cells(5).Borders(xlEdgeRight)
                .Color = RGB(130, 40, 60): .Weight = xlThick: .LineStyle = xlNone
            End With
        Case "mixed"
            target.Cells(5).Font.Name = "Arial": target.Cells(8).Font.Size = 17
            target.Cells(9).Interior.Pattern = xlPatternGray25
            target.Cells(9).Interior.PatternColor = RGB(0, 100, 200)
            target.Cells(9).Interior.TintAndShade = 0.2
            target.Cells(10).NumberFormat = "0.000"
            target.Cells(11).Orientation = 45
        Case "merged": ws.Range("C3:D4").ClearContents: ws.Range("C3:D4").Merge
    End Select
    target.Select
    mask = NxCellFontMask Or NxCellInteriorMask Or NxCellBorderMask Or NxCellAlignmentMask Or NxCellNumberFormatMask Or NxCellProtectionMask
    If name <> "large" Then
        Set context = NxContextFactory.CaptureCurrent()
        context.BindSelectionAuthority authority
        Set plan = NxCreatePlan("NX-TEST-CORE", context.Fingerprint, context.SourceStateDigest)
        plan.AddSideEffect NxCreateCellEffect(context.SelectionTargetIdentity, mask)
        plan.Seal
        expected.Capture context, plan, authority
    End If
    wb.Saved = True: beforeSaved = wb.Saved
    started = Timer: backup.Capture target: captured = Timer - started
    If wb.Saved <> beforeSaved Then Err.Raise 5, , "capture dirtied source"
    If Not ActiveWorkbook Is wb Then Err.Raise 5, , "capture changed active workbook"
    target.Font.Size = 22: target.Font.Color = vbRed
    target.Interior.Color = vbYellow: target.Borders.LineStyle = xlContinuous
    target.HorizontalAlignment = xlCenter
    started = Timer: backup.Restore: restored = Timer - started
    If name <> "large" Then
        detail = expected.VerifyResidual()
        If Len(detail) > 0 Then Err.Raise 5, , Left$(detail, 500)
    Else
        If target.Font.Size <> 9 Or target.Cells(1).Formula <> "=1+2" Or target.Cells(90000).Value2 <> "original" Then Err.Raise 5, , "large restoration mismatch"
    End If
    backup.Dispose
    wb.Close False: Set wb = Nothing
    If Workbooks.Count <> baselineBooks Then Err.Raise 5, , "backup workbook leaked"
    RunCase = "PASS|" & name & "|capture_s=" & CStr(captured) & "|restore_s=" & CStr(restored)
    Exit Function
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
    On Error Resume Next
    backup.Dispose
    If Not wb Is Nothing Then wb.Close False
End Function
