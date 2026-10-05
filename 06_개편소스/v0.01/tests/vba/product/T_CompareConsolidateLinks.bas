Attribute VB_Name = "T_CompareConsolidateLinks"
Option Explicit

Public Function Names() As String
    Names = "sheet_links|sheet_append|sheet_many|sheet_equal|sheet_formulas|sheets_contents|both_contents|rows_no_contents|both_partial|mode_controls"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim a As Workbook, b As Workbook, report As Workbook, destination As Workbook
    Dim panel As FNxFileConsolidate, result As CNxResult, detail As Worksheet, sheet As Worksheet
    Dim paths(0 To 1) As String, root As String, output As String, mode As String, i As Long
    Dim oldAlerts As Boolean, oldEvents As Boolean, first As Long, expected As Long, link As Hyperlink, stage As String
    On Error GoTo Failed
    oldAlerts = Application.DisplayAlerts: oldEvents = Application.EnableEvents
    Application.DisplayAlerts = False: Application.EnableEvents = False
    root = Environ$("LHEXCEL_PROFILE_ROOT") & "\" & name: MkDir root
    If name = "mode_controls" Then
        Set panel = New FNxFileConsolidate
        Require panel.cboMode.ListCount = 3, "three output modes"
        Require panel.chkContents.Enabled And Not panel.chkSkipHeaders.Enabled, "sheet controls"
        panel.cboMode.ListIndex = 1
        Require Not panel.chkContents.Enabled And panel.chkSkipHeaders.Enabled, "row controls"
        panel.cboMode.ListIndex = 2
        Require panel.chkContents.Enabled And panel.chkSkipHeaders.Enabled, "both controls"
        GoTo Passed
    End If
    Set a = Workbooks.Add(xlWBATWorksheet): Set b = Workbooks.Add(xlWBATWorksheet)
    a.Worksheets(1).Name = "목차": b.Worksheets(1).Name = "목차"
    a.Worksheets(1).Range("A1:B2").Value2 = 1: b.Worksheets(1).Range("A1:B2").Value2 = 1
    b.Worksheets(1).Range("B2").Value2 = 2
    If Left$(name, 6) = "sheet_" Then
        If name = "sheet_equal" Then b.Worksheets(1).Range("B2").Value2 = 1
        If name = "sheet_formulas" Then
            a.Worksheets(1).Range("B2").Formula = "=2"
            b.Worksheets(1).Range("B2").Formula = "=1+1"
        End If
        If name = "sheet_many" Then
            a.Worksheets(1).Range("A1:J7000").Value2 = 1
            b.Worksheets(1).Range("A1:J7000").Value2 = 2
        ElseIf name <> "sheet_equal" Then
            a.Worksheets(1).Range("A3").NumberFormat = "@"
            a.Worksheets(1).Range("A3").Value2 = "=literal"
        End If
        a.Saved = True: b.Saved = True
        stage = "create"
        Set result = NxDataCompareCreate("NX-FILE-SHEET-COMPARE", NxDataCompareUsedRange(a.Worksheets(1)), NxDataCompareUsedRange(b.Worksheets(1)), name = "sheet_formulas")
        stage = "finish"
        NxDataCompareFinish True: Set report = ActiveWorkbook
        Require report.Worksheets.Count = 3, "three result sheets"
        Set detail = report.Worksheets("비교 결과")
        Require a.Saved And b.Saved, "source state preserved"
        If name = "sheet_formulas" Then
            Require detail.Range("B6").Value2 = "=2" And detail.Range("C6").Value2 = "=1+1", "formula differences shown"
            Require Not detail.Range("B6").HasFormula And Not detail.Range("C6").HasFormula, "formula text not executed"
            GoTo Passed
        End If
        If name = "sheet_equal" Then
            Require detail.Range("A6").Value2 = "차이가 없습니다.", "no-difference report"
            GoTo Passed
        End If
        If name = "sheet_many" Then
            Require detail.Range("A70005").Value2 = "J7000", "all 70000 differences retained"
            Require detail.Range("D70005").HasFormula, "large report link"
            stage = "large follow; links=" & detail.Range("E70005").Hyperlinks.Count & "; formula=" & detail.Range("E70005").Formula
            FollowFormulaTarget detail.Range("E70005")
            Require ActiveSheet Is report.Worksheets(2), "large link sheet"
            Require ActiveCell.Address = "$J$7004", "large link address"
            GoTo Passed
        End If
        Require detail.Range("A6").Value2 = "B2" And detail.Range("A7").Value2 = "A3", "only actual differences"
        Require detail.Range("B7").Value2 = "=literal" And Not detail.Range("B7").HasFormula, "literal text preserved"
        stage = "base follow; links=" & detail.Range("D6").Hyperlinks.Count & "; formula=" & detail.Range("D6").Formula
        FollowFormulaTarget detail.Range("D6")
        Require ActiveSheet Is report.Worksheets(1), "base snapshot link"
        Require ActiveCell.Address = "$B$6", "base snapshot row offset"
        ActiveSheet.Range("A4").Hyperlinks(1).Follow
        Require ActiveSheet Is detail, "back to result"
        report.SaveAs root & "\report.xlsx", xlOpenXMLWorkbook
        report.Close False: Set report = Workbooks.Open(root & "\report.xlsx", 0)
        Set detail = report.Worksheets("비교 결과")
        FollowFormulaTarget detail.Range("E7")
        Require ActiveSheet Is report.Worksheets(2), "reopened link"
        Require ActiveCell.Address = "$A$7", "one-sided difference destination"
        If name = "sheet_append" Then
            Set destination = Workbooks.Add(xlWBATWorksheet)
            destination.Worksheets(1).Name = "기준 데이터": destination.Worksheets(1).Range("A1").Value2 = "기존 내용"
            NxCompareAppendResultSheets report, destination: Set report = Nothing
            Set detail = destination.Worksheets(4)
            FollowFormulaTarget detail.Range("E6")
            Require ActiveSheet Is destination.Worksheets(3), "renamed destination link"
            Require ActiveCell.Address = "$B$6", "renamed destination address"
            ActiveSheet.Range("A4").Hyperlinks(1).Follow
            Require ActiveSheet Is detail, "renamed back link"
            Require destination.Worksheets(1).Range("A1").Value2 = "기존 내용", "destination preserved"
        End If
        GoTo Passed
    End If
    paths(0) = root & "\first.xlsx": paths(1) = root & "\second.xlsx"
    If name = "both_partial" Then b.Worksheets(1).Range("C1").Value2 = "different width"
    a.SaveAs paths(0), xlOpenXMLWorkbook: b.SaveAs paths(1), xlOpenXMLWorkbook
    a.Close False: b.Close False: Set a = Nothing: Set b = Nothing
    mode = "SHEETS"
    If name = "both_contents" Or name = "both_partial" Then mode = "BOTH"
    If name = "rows_no_contents" Then mode = "ROWS"
    output = root & "\output.xlsx"
    Set result = NxRunConsolidation(paths, output, name = "both_partial", name = "both_partial", mode, False, True)
    If name = "both_partial" Then
        Require result.Outcome = NxPartialFailure, result.Recovery
        output = root & "\output_부분결과.xlsx"
    Else
        Require result.Outcome = NxSuccess, result.Recovery
    End If
    Set report = Workbooks.Open(output, 0, True)
    If mode = "ROWS" Then
        Require report.Worksheets.Count = 1, "row output only"
        Require report.Worksheets(1).Hyperlinks.Count = 0, "no row links"
        Require report.Worksheets(1).Range("B4").Value2 = 2, "appended row data"
    Else
        Set detail = report.Worksheets("목차")
        expected = IIf(name = "both_partial", 1, 2)
        Require detail.Hyperlinks.Count = expected, "contents count"
        For i = 1 To expected
            detail.Cells(i + 3, 2).Hyperlinks(1).Follow
            Set sheet = ActiveSheet
            Require Not sheet Is detail, "contents opens data"
            Require ActiveCell.Address = "$A$2" And sheet.Range("A2").Value2 = 1, "original A1 preserved"
            sheet.Range("A1").Hyperlinks(1).Follow
            Require ActiveSheet Is detail And ActiveCell.Row = i + 3, "return to matching contents entry"
        Next i
        If mode = "BOTH" Then
            Require report.Worksheets("통합데이터").Hyperlinks.Count = 0, "combined sheet has no links"
            Require report.Worksheets("통합데이터").Cells(2 * expected, 2).Value2 = expected, "both rows match accepted sheets"
            Require IsEmpty(report.Worksheets("통합데이터").Cells(2 * expected + 1, 1).Value2), "failed input leaves no row data"
        End If
    End If
    Set a = Workbooks.Open(paths(0), 0, True)
    Require a.Worksheets(1).Range("A1").Value2 = 1 And a.Worksheets(1).Hyperlinks.Count = 0, "source file unchanged"
Passed:
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & stage & "|" & Err.Source & "|" & Err.Number & ":" & Err.Description
Clean:
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    If Not report Is Nothing Then report.Close False
    If Not destination Is Nothing Then destination.Close False
    If Not a Is Nothing Then a.Close False
    If Not b Is Nothing Then b.Close False
    NxDataCompareFinish False
    Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 104, "CompareConsolidateLinks", detail
End Sub

Private Sub FollowFormulaTarget(ByVal cell As Range)
    Dim formula As String, target As String, delimiter As Long, sheetName As String, address As String
    formula = CStr(cell.Formula)
    Require Left$(formula, 13) = "=HYPERLINK(""#", "internal formula hyperlink"
    delimiter = InStr(14, formula, """")
    Require delimiter > 14, "hyperlink destination"
    target = Mid$(formula, 14, delimiter - 14)
    ' Formula hyperlinks are not exposed by Range.Hyperlinks in Excel 2024.
    ' Resolve the generated destination through Excel; UI clicking is checked separately.
    delimiter = InStrRev(target, "'!")
    Require delimiter > 1, "quoted worksheet destination"
    sheetName = Replace$(Mid$(target, 2, delimiter - 2), "''", "'")
    address = Mid$(target, delimiter + 2)
    Application.Goto cell.Worksheet.Parent.Worksheets(sheetName).Range(address), True
End Sub
