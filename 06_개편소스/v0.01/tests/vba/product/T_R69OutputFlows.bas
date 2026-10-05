Attribute VB_Name = "T_R69OutputFlows"
Option Explicit

Public FailAfterCreate As Boolean

Public Function Names() As String
    Names = "normalize_runs|normalize_text|date_right|date_original|date_sheet|date_book|date_popup|age_right|age_original|age_sheet|age_book|age_options|money_blank|mask_blank|whole_column|column_rollback|copy_save|image_png|image_jpg|consolidate_popup|compare_unsaved|blank_edge_columns|money_popup|mask_popup"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, other As Workbook, report As Workbook, target As Range, cell As Range, panel As Object
    Dim result As CNxResult, options As CNxDataSpecialOptions, workflow As Object, context As CNxFeatureDialogContext
    Dim mode As String, root As String, path As String, firstPath As String, secondPath As String
    Dim detail As String, item As Variant, prefix As Long, suffix As Long, actual As Variant
    Dim sourcePath As String, sheetName As String, cellAddress As String, issue As String
    On Error GoTo Failed
    root = Environ$("LHEXCEL_PROFILE_ROOT")
    Set book = Workbooks.Add(xlWBATWorksheet)
    book.SaveAs root & "\" & name & ".xlsx", xlOpenXMLWorkbook
    Set target = book.Worksheets(1).Range("B2:B6")
    target.NumberFormat = "@"
    target.Cells(1).Value2 = "240617"
    target.Cells(3).Value2 = "2024-02-30"
    target.Cells(4).Value2 = CVErr(xlErrValue)
    target.Cells(5).Value2 = "20000229"
    target.Select
    mode = "오른쪽 새 열 삽입"
    If InStr(name, "original") > 0 Then mode = "원본 변경"
    If InStr(name, "sheet") > 0 Then mode = "새 시트 만들기"
    If InStr(name, "book") > 0 Then mode = "새 통합문서"
    Select Case name
        Case "normalize_runs", "normalize_text"
            Set target = book.Worksheets(1).Range("B2:G3")
            target.NumberFormat = "@": target.Value2 = "  sample  "
            target.Cells(1, 1).Value2 = "=1+2"
            target.Cells(1, 3).Value2 = 42
            target.Cells(1, 4).Value2 = "+1+2"
            target.Cells(2, 6).Value2 = "=9+9"
            target.Select
            mode = "기본 정리"
            If name = "normalize_text" Then mode = "텍스트 변환"
            Set result = NxDataRunNormalization(target, mode, "yyyy-mm-dd", "새 시트 만들기")
            Check result.Outcome = NxSuccess, result.Recovery
            Set cell = ActiveSheet.Range("A1")
            Check Not cell.HasFormula And cell.Value2 = "=1+2", "Leading literal executed"
            Check Not cell.Offset(0, 3).HasFormula And cell.Offset(0, 3).Value2 = "+1+2", "Second text run executed"
            Check Not cell.Offset(1, 5).HasFormula And cell.Offset(1, 5).Value2 = "=9+9", "Last run lost"
            If name = "normalize_runs" Then
                Check VarType(cell.Offset(0, 2).Value2) = vbDouble, "Number converted to text"
            Else
                Check cell.Offset(0, 2).NumberFormat = "@", "Text preset format lost"
            End If
        Case "date_right", "date_original", "date_sheet", "date_book", "whole_column", "column_rollback"
            book.Worksheets(1).Range("C2").Value2 = "인접 원본"
            book.Worksheets(1).Range("D2").Formula = "=C2"
            If name = "whole_column" Then Set target = book.Worksheets(1).Columns("B")
            FailAfterCreate = (name = "column_rollback")
            Set result = NxDataRunNormalization(target, "날짜 정리", "yyyy-mm-dd", mode)
            FailAfterCreate = False
            If name = "column_rollback" Then
                Check result.Outcome <> NxSuccess, "Injected failure succeeded"
                Check book.Worksheets(1).Range("C2").Value2 = "인접 원본", "Inserted column not removed"
                Check book.Worksheets(1).Range("D2").Formula = "=C2", "Reference not restored"
            Else
                Check result.Outcome = NxSuccess, result.Recovery
                If mode = "원본 변경" Then
                    Set cell = book.Worksheets(1).Range("B2")
                ElseIf mode = "오른쪽 새 열 삽입" Then
                    Set cell = book.Worksheets(1).Range("C2")
                    Check book.Worksheets(1).Range("D2").Value2 = "인접 원본", "Adjacent data overwritten"
                    Check book.Worksheets(1).Range("E2").Formula = "=D2", "Inserted column reference incorrect"
                Else
                    Set cell = ActiveSheet.Range("A1")
                    If mode = "새 통합문서" Then
                        Set report = ActiveWorkbook
                        Check Not report Is book, "New workbook missing"
                        Check Len(report.Path) = 0, "New result unexpectedly saved"
                    End If
                End If
                Check cell.Value2 = CDbl(DateSerial(2024, 6, 17)), "Date was not converted"
                Check cell.NumberFormat = "yyyy-mm-dd", "Date format missing"
                Check Len(CStr(cell.Offset(1).Value2)) = 0, "Blank cell converted"
                Check Len(CStr(cell.Offset(2).Value2)) = 0, "Invalid date not blank"
                Check Len(CStr(cell.Offset(3).Value2)) = 0, "Error input not blank"
                Check cell.Offset(4).Value2 = CDbl(DateSerial(2000, 2, 29)), "Leap date incorrect"
            End If
        Case "blank_edge_columns"
            Set target = book.Worksheets(1).Range("A1:C10")
            book.Worksheets(1).Range("D2").Value2 = "오른쪽 경계"
            Set result = NxDataRunNormalization(target, "날짜 정리", "yyyy-mm-dd", mode)
            Check result.Outcome = NxSuccess, result.Recovery
            Check book.Worksheets(1).Range("E2").Value2 = CDbl(DateSerial(2024, 6, 17)), "Selection width lost"
            Check book.Worksheets(1).Range("G2").Value2 = "오른쪽 경계", "Inserted before selection boundary"
        Case "date_popup"
            NxOpenDateConversion
            For Each item In VBA.UserForms
                If TypeName(item) = "FNxDataNormalize" Then Set panel = item
            Next item
            Check Not panel Is Nothing, "Date dialog missing"
            Check panel.cboOutputMode.Value = "새 통합문서", "Date default"
            Check InStr(panel.Caption, "날짜/생년월일") > 0, "Date caption"
            panel.NxProbeExecute
            Set report = ActiveWorkbook
            Check Not report Is book, "Date default new workbook"
            Check report.Worksheets(1).Range("A1").Value2 = CDbl(DateSerial(2024, 6, 17)), "Real date handler failed"
        Case "age_right", "age_original", "age_sheet", "age_book"
            target.Cells(1).Value2 = "900520"
            Set panel = New FNxAgeCalculator
            panel.BindFeature NX_FEATURE_DATA_AGE
            panel.cboOutputMode.Value = mode
            panel.txtReferenceDate.Value = "2026-05-19"
            panel.chkHeader.Value = False
            panel.NxProbeExecute
            If mode = "오른쪽 새 열 삽입" Then
                Set cell = book.Worksheets(1).Range("C2")
            ElseIf mode = "원본 변경" Then
                Set cell = book.Worksheets(1).Range("B2")
            Else
                Set cell = ActiveSheet.Range("A1")
                If mode = "새 통합문서" Then Set report = ActiveWorkbook
            End If
            Check cell.Value2 = "만 35세", "Age result incorrect"
            Check Len(CStr(cell.Offset(1).Value2)) = 0, "Age blank converted"
            Check Len(CStr(cell.Offset(2).Value2)) = 0, "Age invalid not blank"
        Case "age_options"
            For prefix = 0 To 1
                For suffix = 0 To 1
                    Set options = New CNxDataSpecialOptions
                    options.Configure "오른쪽 새 열 삽입", False, "2026-05-20", "사용자 지정", False, False, False, "빈칸", "자동", "*", CBool(prefix), CBool(suffix)
                    actual = options.TransformValue(NX_FEATURE_DATA_AGE, "900520")
                    Check actual = IIf(prefix = 1, "만 ", "") & "36" & IIf(suffix = 1, "세", ""), "Age prefix/suffix"
                Next suffix
            Next prefix
        Case "money_blank", "mask_blank"
            Set options = New CNxDataSpecialOptions
            options.Configure "원본 변경", False, "", "숫자", False, False, False, "빈칸", "휴대폰번호", "*"
            If name = "money_blank" Then firstPath = NX_FEATURE_DATA_KOREAN_MONEY Else firstPath = NX_FEATURE_DATA_PRIVACY_MASK
            Check options.TransformValue(firstPath, Empty) = "", "Empty input not blank"
            Check options.TransformValue(firstPath, CVErr(xlErrValue)) = "", "Error input not blank"
            Check options.TransformValue(firstPath, "invalid") = "", "Invalid conversion not blank"
            Set panel = New FNxDataSpecial: panel.BindFeature firstPath
            Check panel.Controls("lblMoney").Visible = (name = "money_blank"), "Money section not separated"
            Check panel.Controls("lblMask").Visible = (name = "mask_blank"), "Mask section not separated"
        Case "money_popup", "mask_popup"
            target.Cells(1).Value2 = IIf(name = "money_popup", "12345", "01012345678")
            If name = "money_popup" Then firstPath = NX_FEATURE_DATA_KOREAN_MONEY Else firstPath = NX_FEATURE_DATA_PRIVACY_MASK
            Set panel = New FNxDataSpecial: panel.BindFeature firstPath
            panel.chkPrefix.Value = False: panel.chkSuffix.Value = False: panel.chkMixed.Value = False
            panel.cboMaskType.Value = "휴대폰번호"
            panel.NxProbeExecute
            Set report = ActiveWorkbook
            Check Not report Is book, "Default new workbook"
            actual = report.Worksheets(1).Range("A1").Value2
            If name = "money_popup" Then
                Check InStr(actual, "만") > 0 And InStr(actual, "삼백") > 0, "Money popup result"
            Else
                Check InStr(actual, "*") > 0 And InStr(actual, "1234") = 0, "Mask popup result"
            End If
            Check Len(CStr(report.Worksheets(1).Range("A2").Value2)) = 0, "Popup blank converted"
        Case "copy_save"
            path = NxFileVacantCopyPath(root, "사본시험")
            Set result = NxFileRun(NX_FEATURE_FILE_RANGE_COPY_SAVE, path, selectedTarget:=target)
            Check result.Outcome = NxSuccess, result.Recovery
            Check Len(Dir$(path)) > 0, "Copy file missing"
            Check NxFileVacantCopyPath(root, "사본시험") <> path, "Copy would overwrite"
            Set other = Workbooks.Open(path, 0, True)
            Check other.Worksheets(1).Range("A1").Value2 = "240617", "Copy content"
        Case "image_png", "image_jpg"
            path = root & "\" & name & "." & Right$(name, 3)
            target.Interior.Color = RGB(255, 230, 128): target.Font.Color = vbBlack
            target.ColumnWidth = 18: target.RowHeight = 26
            Set panel = New FNxFileOutput: panel.BindFeature NX_FEATURE_FILE_RANGE_PNG
            panel.txtOutputPath.Value = path
            panel.NxProbeExecute
            Check Len(Dir$(path)) > 0, "Image output missing"
            Check FileLen(path) > 300, "Image output empty"
        Case "consolidate_popup"
            Set other = Workbooks.Add(xlWBATWorksheet)
            other.Worksheets(1).Range("A1").Value2 = "통합 데이터"
            path = root & "\consolidate_input.xlsx"
            other.SaveAs path, xlOpenXMLWorkbook: other.Close False: Set other = Nothing
            book.Activate: target.Select
            Set panel = New FNxFileConsolidate: panel.BindFeature
            panel.lstInputFiles.AddItem path
            panel.txtOutputPath.Value = root & "\consolidate_output.xlsx"
            panel.NxProbeExecute
            Check Len(Dir$(root & "\consolidate_output.xlsx")) > 0, "Consolidation failed"
        Case "compare_unsaved"
            book.Worksheets(1).Name = "비교'시트"
            book.Save
            firstPath = book.FullName
            Set other = Workbooks.Add(xlWBATWorksheet)
            other.Worksheets(1).Name = "비교'시트"
            other.Worksheets(1).Range("B2").Value2 = "다른 값"
            secondPath = root & "\compare_other.xlsx"
            other.SaveAs secondPath, xlOpenXMLWorkbook: other.Close False: Set other = Nothing
            book.Activate: target.Select
            Set panel = New FNxWorkbookCompare: panel.BindFeature
            panel.txtBasePath.Value = firstPath: panel.txtComparePath.Value = secondPath
            panel.NxProbeExecute
            Set report = ActiveWorkbook
            Check Not report Is book, "Comparison report not opened"
            Check Len(report.Path) = 0 And Not report.Saved, "Comparison unexpectedly saved"
            Check report.Worksheets("비교요약").Range("B4").Value2 > 0, "Comparison differences missing"
            Check report.Worksheets("비교요약").Range("D6").Hyperlinks.Count = 1, "Summary detail link missing"
            Check report.Worksheets("셀차이").Range("D2:E2").Hyperlinks.Count = 2, "Snapshot links missing"
            Check report.Worksheets("셀차이").Range("J1").Value2 = "기준 표시 형식", "Detailed comparison missing"
            Check NxWorkbookCompareNavigateSnapshot(report.Worksheets("셀차이"), 2, 0), "Snapshot link invalid"
    End Select
    GoTo Cleanup
Failed:
    detail = CStr(Err.Number) & "|" & Err.Description: Err.Clear
Cleanup:
    FailAfterCreate = False
    On Error Resume Next
    For Each item In VBA.UserForms: Unload item: Next item
    If Not report Is Nothing Then report.Close False
    If Not other Is Nothing Then other.Close False
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    If Len(detail) = 0 Then RunCase = "PASS|" & name Else RunCase = "FAIL|" & name & "|" & detail
End Function

Private Sub Check(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise 5, "R69OutputFlows", message
End Sub
