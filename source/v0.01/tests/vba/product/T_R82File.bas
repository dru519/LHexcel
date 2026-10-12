Attribute VB_Name = "T_R82File"
Option Explicit
Public Function Names() As String
    Names = "direct_90000|routed_90000|formulas_90000|layout_picture|existing_output|missing_folder|cf_range|cf_sheet|focus_range|focus_sheet"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, saved As Workbook, target As Range, result As CNxResult
    Dim path As String, started As Double, elapsed As Double, handle As Integer
    Dim events As Boolean, alerts As Boolean, books As Long
    Dim picture As Shape
    Dim engine As CNxFocusRuleEngine, focusSettings As CNxFocusSettingsModel, focusSelection As CNxFocusSelection
    On Error GoTo Failed
    events = Application.EnableEvents: alerts = Application.DisplayAlerts
    Application.EnableEvents = False: Application.DisplayAlerts = False
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1:AX1800")
    target.Value2 = 7: target.Font.Size = 9
    target.Cells(2, 2).Formula = "=A2+1"
    If name = "cf_range" Or name = "cf_sheet" Then
        With target.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="5")
            .Interior.Color = RGB(255, 200, 200)
        End With
    End If
    If name = "focus_range" Or name = "focus_sheet" Then
        Set engine = New CNxFocusRuleEngine
        Set focusSettings = NxFocusCreateSettings(NX_FOCUS_SHAPE_CROSS, NX_FOCUS_STYLE_OUTER_CELLS, "#5FC8D8", 80, False, True)
        engine.Configure focusSettings
        Set focusSelection = New CNxFocusSelection
        If Not focusSelection.Configure(target.Range("D4")) Then Err.Raise 5, , "Focus selection failed"
        If Not engine.ApplySelection(focusSelection) Then Err.Raise 5, , "Focus setup failed"
    End If
    If name = "formulas_90000" Then target.FormulaR1C1 = "=ROW()*COLUMN()": target.Cells(2, 2).Formula = "=A2+1"
    If name = "layout_picture" Then
        target.Columns.ColumnWidth = 15: target.Rows.RowHeight = 21
        target.Columns(3).ColumnWidth = 24: target.Rows(4).RowHeight = 33
        target.Range("D5:E5").Merge
        Set picture = book.Worksheets(1).Shapes.AddPicture(Environ$("LHEXCEL_PROFILE_ROOT") & "\copy-save-picture.png", _
            False, True, target.Range("G7").Left, target.Range("G7").Top, 48, 48)
        picture.Left = target.Range("G7").Left: picture.Top = target.Range("G7").Top
        Application.CutCopyMode = False
    End If
    target.Select: book.Saved = True: books = Workbooks.Count
    path = Environ$("LHEXCEL_PROFILE_ROOT") & "\" & name & ".xlsx"
    If name = "existing_output" Then
        handle = FreeFile: Open path For Output As #handle: Print #handle, "preserve": Close #handle
    End If
    If name = "missing_folder" Then path = Environ$("LHEXCEL_PROFILE_ROOT") & "\absent\result.xlsx"
    started = Timer
    If name = "direct_90000" Or name = "existing_output" Or name = "missing_folder" Then
        Set result = NxRunRangeCopySave(target, path)
    ElseIf name = "cf_sheet" Or name = "focus_sheet" Then
        Set result = NxFileRun(NX_FEATURE_FILE_SHEET_COPY_SAVE, path, Empty, False, book.Worksheets(1))
    Else
        Set result = NxFileRun(NX_FEATURE_FILE_RANGE_COPY_SAVE, path, Empty, False, target)
    End If
    elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400
    If name = "existing_output" Or name = "missing_folder" Then
        If result.Outcome = NxSuccess Then Err.Raise 5, , "Invalid output accepted"
        If Workbooks.Count <> books Or Not book.Saved Then Err.Raise 5, , "Failure changed source"
        If name = "existing_output" Then
            Dim line As String
            handle = FreeFile: Open path For Input As #handle: Line Input #handle, line: Close #handle
            If line <> "preserve" Then Err.Raise 5, , "Existing output overwritten"
        End If
        RunCase = "PASS|" & name
        GoTo Clean
    End If
    If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
    If Workbooks.Count <> books Then Err.Raise 5, , "Workbook leaked"
    If Not book.Saved Then Err.Raise 5, , "Source dirty"
    If Not engine Is Nothing Then
        If engine.ManagedRuleCount(book.Worksheets(1)) <> 4 Then Err.Raise 5, , "Source focus rules changed"
    End If
    Set saved = Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True)
    If saved.Worksheets(1).Range("B2").Formula <> "=A2+1" Then Err.Raise 5, , "Formula changed"
    If name <> "formulas_90000" And saved.Worksheets(1).Range("AX1800").Value2 <> 7 Then Err.Raise 5, , "Value changed"
    If name = "formulas_90000" And saved.Worksheets(1).Range("AX1800").Formula <> "=ROW()*COLUMN()" Then Err.Raise 5, , "Bulk formula changed"
    If saved.Worksheets(1).Range("AX1800").Font.Size <> 9 Then Err.Raise 5, , "Format changed"
    If name = "cf_range" Or name = "cf_sheet" Then
        If target.FormatConditions.Count <> 1 Then Err.Raise 5, , "Source conditional format changed"
        If saved.Worksheets(1).Range("AX1800").FormatConditions.Count <> 1 Then Err.Raise 5, , "Conditional format lost"
        If saved.Worksheets(1).Range("AX1800").DisplayFormat.Interior.Color <> RGB(255, 200, 200) Then Err.Raise 5, , "Conditional format display changed"
    End If
    If name = "layout_picture" Then
        If saved.Worksheets(1).Columns(3).ColumnWidth <> 24 Or saved.Worksheets(1).Rows(4).RowHeight <> 33 Then Err.Raise 5, , "Dimensions changed"
        If saved.Worksheets(1).Range("D5").MergeArea.Address <> "$D$5:$E$5" Then Err.Raise 5, , "Merge changed"
        If saved.Worksheets(1).Shapes.Count <> 1 Then Err.Raise 5, , "Picture lost or duplicated"
    End If
    saved.Close False: Set saved = Nothing
    handle = FreeFile
    Open Environ$("LHEXCEL_PROFILE_ROOT") & "\copy-save.tsv" For Append As #handle
    Print #handle, name & vbTab & Format$(elapsed, "0.000")
    Close #handle
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
Clean:
    On Error Resume Next
    If Not saved Is Nothing Then saved.Close False
    If Not book Is Nothing Then book.Close False
    Application.EnableEvents = events: Application.DisplayAlerts = alerts
End Function
