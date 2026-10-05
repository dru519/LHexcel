Attribute VB_Name = "T_R69SnapshotProbe"
Option Explicit
Public UseDuplicateBackup As Boolean
Private mFault As Boolean, mEvidence As String
Public Sub Configure(ByVal folder As String)
    mEvidence = folder
End Sub
Public Sub MaybeFail()
    If mFault Then Err.Raise vbObjectError + 769, "snapshot probe", "injected after table and autofit"
End Sub
Public Function Names() As String
    Names = "direct_body|direct_title|legacy_body|legacy_title|single_body|single_title|single_mixed|single_alternating|snapshot_200|snapshot_1000"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, result As CNxResult, cell As Range
    Dim feature As String, style As String, before As String, after As String, caught As Long, count As Long
    Dim iteration As Long, phase As Long, started As Double, elapsed As Double, handle As Integer, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    count = 18
    If Left$(name, 9) = "snapshot_" Then count = CLng(Split(name, "_")(1))
    If Left$(name, 9) = "snapshot_" Then
        Set target = book.Worksheets(1).Range("A1").Resize(count / 5, 5)
    Else
        Set target = book.Worksheets(1).Range("A1").Resize(count / 3, 3)
    End If
    target.Value2 = "long sample text"
    target.Cells(2, 2).Formula = "=12+30"
    target.Font.Name = "맑은 고딕": target.Font.Size = 11
    target.Font.ThemeColor = xlThemeColorAccent2: target.Font.TintAndShade = 0.2
    target.Interior.ThemeColor = xlThemeColorAccent3: target.Interior.TintAndShade = 0.3
    target.Borders.LineStyle = xlContinuous
    target.Borders.ThemeColor = xlThemeColorAccent4
    target.Borders.TintAndShade = 0.1
    target.Columns.ColumnWidth = 14: target.RowHeight = 18
    If name = "single_mixed" Then
        target.Cells(2, 1).Font.Bold = True
        target.Cells(2, 2).Font.Size = 14
        target.Cells(3, 1).Interior.Color = RGB(255, 200, 150)
        target.Cells(3, 2).HorizontalAlignment = xlRight
        target.Cells(4, 1).Borders(xlEdgeBottom).Weight = xlThick
        target.Cells(4, 2).NumberFormat = "0.000"
        target.Cells(5, 1).Font.Color = RGB(123, 87, 213)
        target.Cells(5, 2).Borders(xlEdgeLeft).Color = RGB(41, 153, 201)
    End If
    If name = "single_alternating" Then
        ' Equal runs alternate with shorter RGB-only property vectors and back.
        ' This exercises lazy prefix materialization and omission of theme keys.
        For iteration = 1 To target.Cells.Count
            Set cell = target.Cells(iteration)
            If (iteration Mod 4) < 2 Then
                cell.Font.Color = RGB(123, 87, 213)
                cell.Interior.Color = RGB(250, 240, 220)
                cell.Borders(xlEdgeBottom).Color = RGB(41, 153, 201)
            End If
            If iteration = 9 Then
                cell.Font.Size = 17
                cell.Borders(xlDiagonalUp).LineStyle = xlContinuous
                cell.Borders(xlDiagonalUp).Color = vbRed
            End If
        Next iteration
    End If
    target.Select
    feature = NX_FEATURE_DRAW_BUSINESS_TABLE
    If InStr(name, "title") > 0 Then feature = NX_FEATURE_DRAW_TITLE_TABLE
    style = "열린 표"
    If feature = NX_FEATURE_DRAW_TITLE_TABLE Then style = "제목"
    If Left$(name, 9) = "snapshot_" Then
        mFault = False
        book.SaveAs mEvidence & Application.PathSeparator & name & ".xlsx", xlOpenXMLWorkbook
        For iteration = 1 To 4
            For phase = 0 To 1
                UseDuplicateBackup = ((iteration + phase) Mod 2 = 1)
                started = Timer
                Set result = NxDrawRun(feature, target)
                elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400#
                If result.Outcome <> NxSuccess Or Not result.SourceChanged Then Err.Raise 5, , "apply failed"
                handle = FreeFile
                Open mEvidence & Application.PathSeparator & "snapshot.tsv" For Append As #handle
                Print #handle, count & vbTab & iteration & vbTab & IIf(UseDuplicateBackup, "duplicate", "single") & vbTab & Format$(elapsed * 1000#, "0.000")
                Close #handle
            Next phase
        Next iteration
    Else
        target.Columns(2).Hidden = True
        before = State(target)
        UseDuplicateBackup = (Left$(name, 7) = "legacy_")
        mFault = True
        If Left$(name, 7) = "direct_" Then
            On Error Resume Next
            NxDrawApplyTableOptions target, feature, "MONO", style, 1, 9, xlCenter, 9, xlGeneral, False, True
            caught = Err.Number: Err.Clear
            On Error GoTo Failed
            If caught = 0 Then Err.Raise 5, , "fault was not raised"
        Else
            Set result = NxDrawRun(feature, target, autoFitColumns:=True, tableStyle:=style)
            If result.Outcome = NxSuccess Then Err.Raise 5, , "fault returned success"
        End If
        after = State(target)
        If before <> after Then
            handle = FreeFile
            Open mEvidence & Application.PathSeparator & name & "-state.txt" For Output As #handle
            Print #handle, before
            Print #handle, after
            Close #handle
            Err.Raise 5, , "rollback state mismatch"
        End If
    End If
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    mFault = False: UseDuplicateBackup = False
    If Not book Is Nothing Then book.Close False
End Function
Private Function State(ByVal target As Range) As String
    Dim cell As Range, edge As Variant, result As String, color As Variant
    For Each cell In target
        result = result & "|" & cell.Formula & "|" & cell.Font.Name & "|" & cell.Font.Size & "|" & cell.Font.Bold & "|" & cell.Font.Italic
        result = result & "|" & cell.Font.Color & "|" & cell.Font.TintAndShade & "|" & Theme(cell.Font)
        result = result & "|" & cell.Interior.Pattern & "|" & cell.Interior.Color & "|" & cell.Interior.TintAndShade & "|" & Theme(cell.Interior)
        result = result & "|" & cell.HorizontalAlignment & "|" & cell.VerticalAlignment & "|" & cell.WrapText & "|" & cell.NumberFormat
        result = result & "|" & cell.ColumnWidth & "|" & cell.EntireColumn.Hidden & "|" & cell.RowHeight
        For Each edge In Array(xlDiagonalDown, xlDiagonalUp, xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom, xlInsideVertical, xlInsideHorizontal)
            result = result & "|" & cell.Borders(edge).LineStyle & "|" & cell.Borders(edge).Color & "|" & cell.Borders(edge).Weight
            result = result & "|" & Theme(cell.Borders(edge)) & "|" & cell.Borders(edge).TintAndShade
        Next edge
    Next cell
    State = result & "|" & target.Worksheet.Shapes.Count & "|" & target.Worksheet.Parent.Sheets.Count
End Function
Private Function Theme(ByVal item As Object) As String
    On Error GoTo None
    Theme = CStr(item.ThemeColor)
    Exit Function
None:
    Err.Clear: Theme = "none"
End Function
