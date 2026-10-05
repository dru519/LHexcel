Attribute VB_Name = "T_R83Files"
Option Explicit
Public CurrentCase As String, LockHandle As Integer, FaultObserved As Boolean
Public Sub Mark(ByVal stage As String)
    Dim h As Integer
    h = FreeFile
    Open Environ$("LHEXCEL_PROFILE_ROOT") & "\stages.tsv" For Append As #h
    Print #h, CurrentCase & vbTab & stage & vbTab & CStr(Timer)
    Close #h
End Sub
Public Sub AfterSave()
    If CurrentCase = "save_failure" Then FaultObserved = True: Err.Raise 5, , "Injected post-save failure"
End Sub
Public Sub BeforeCommit(ByVal path As String)
    If CurrentCase = "locked_temp" Then
        LockHandle = FreeFile
        Open path For Binary Access Read Write Lock Read Write As #LockHandle
        FaultObserved = True
    End If
End Sub
Public Sub ReleaseLock()
    If LockHandle <> 0 Then Close #LockHandle: LockHandle = 0
End Sub
Public Function Names() As String
    Names = "copy_stages|save_failure|locked_temp|picker_cancel|compare_1000"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, other As Workbook, output As Workbook, target As Range, result As CNxResult
    Dim path As String, temp As String, p As String, q As String, books As Long, started As Double
    Dim events As Boolean, alerts As Boolean, issue As String
    On Error GoTo Failed
    CurrentCase = name: FaultObserved = False
    events = Application.EnableEvents: alerts = Application.DisplayAlerts
    Application.EnableEvents = False: Application.DisplayAlerts = False
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1:AX1800")
    target.Value2 = 7: target.Font.Size = 9
    If name <> "copy_stages" Then Set target = book.Worksheets(1).Range("A1:J100")
    target.Select: book.Saved = True: books = Workbooks.Count
    path = Environ$("LHEXCEL_PROFILE_ROOT") & "\" & name & ".xlsx"
    temp = Replace(path, ".xlsx", ".nx-copy-save.tmp.xlsx")
    If name = "picker_cancel" Then
        NxFileRunCopySaveFromRibbon NX_FEATURE_FILE_RANGE_COPY_SAVE
        If Workbooks.Count <> books Or Not book.Saved Then Err.Raise 5, , "Cancel changed source"
    ElseIf name = "compare_1000" Then
        book.Worksheets(1).Cells.Clear
        Set target = book.Worksheets(1).Range("A1:J100")
        target.NumberFormat = "@": target.Value2 = "=literal": book.Worksheets(1).Name = "비교'시트"
        Set other = Workbooks.Add(xlWBATWorksheet)
        other.Worksheets(1).Name = "비교'시트"
        other.Worksheets(1).Range("A1:J100").Value2 = "different"
        p = Environ$("LHEXCEL_PROFILE_ROOT") & "\base.xlsx": q = Environ$("LHEXCEL_PROFILE_ROOT") & "\compare.xlsx"
        book.SaveAs p, xlOpenXMLWorkbook: other.SaveAs q, xlOpenXMLWorkbook
        book.Close False: other.Close False: Set book = Nothing: Set other = Nothing
        Mark "compare-start"
        Set result = NxWorkbookCompareReport(p, q, vbNullString, False)
        Mark "compare-end"
        If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
        NxWorkbookCompareFinishReport True
        Set output = ActiveWorkbook
        If Len(output.Path) <> 0 Or output.Saved Then Err.Raise 5, , "Report must remain unsaved"
        If output.Worksheets("비교요약").Range("B4").Value2 <> 1000 Then Err.Raise 5, , "Difference count"
        With output.Worksheets("셀차이")
            If .Range("H2:I1001").Hyperlinks.Count <> 2000 Then Err.Raise 5, , "Links missing"
            If .Range("B1001").Value2 <> "J100" Or .Range("D2").Value2 <> "8:=literal" Or .Range("D2").HasFormula Then Err.Raise 5, , "Literal or address mismatch"
            If .Range("N2").Value2 <> "=literal" Or .Range("N2").HasFormula Then Err.Raise 5, , "Display text became formula"
            If .Range("J2").Value2 <> "@" Or .Range("K2").Value2 <> "General" Or .Range("O2").Value2 <> "different" Then Err.Raise 5, , "Report metadata mismatch"
            If Not NxWorkbookCompareNavigate(output.Worksheets("셀차이"), .Range("H2").Hyperlinks(1), issue) Then Err.Raise 5, , issue
            Set book = ActiveWorkbook
            If book.FullName <> p Or Not book.ReadOnly Or ActiveCell.Address <> "$A$1" Then Err.Raise 5, , "Base navigation"
            book.Close False: Set book = Nothing
            If Not NxWorkbookCompareNavigate(output.Worksheets("셀차이"), .Range("I1001").Hyperlinks(1), issue) Then Err.Raise 5, , issue
            Set other = ActiveWorkbook
            If other.FullName <> q Or Not other.ReadOnly Or ActiveCell.Address <> "$J$100" Then Err.Raise 5, , "Compare navigation"
            other.Close False: Set other = Nothing
        End With
    Else
        Mark "total-start"
        Set result = NxRunRangeCopySave(target, path)
        Mark "total-end"
        If Workbooks.Count <> books Or Not book.Saved Then Err.Raise 5, , "Source or workbook cleanup changed"
        If name = "copy_stages" Then
            If result.Outcome <> NxSuccess Or Len(Dir$(path)) = 0 Then Err.Raise 5, , result.Recovery
        Else
            If Not FaultObserved Or result.Outcome = NxSuccess Then Err.Raise 5, , "Fault did not stop operation"
            If Len(Dir$(path)) > 0 Or Len(Dir$(temp)) > 0 Then Err.Raise 5, , "Partial file remains"
        End If
    End If
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
Clean:
    On Error Resume Next
    ReleaseLock
    If Not output Is Nothing Then output.Close False
    If Not other Is Nothing Then other.Close False
    If Not book Is Nothing Then book.Close False
    Application.EnableEvents = events: Application.DisplayAlerts = alerts
End Function
