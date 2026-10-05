Attribute VB_Name = "T_R67CompareResults"
Option Explicit

Public Function Names() As String
    Names = "snapshot|navigate|source_closed|missing_sheet|row_changed|path_changed|formula_field|external_link|hidden_row|closed_report|replaced_sheet|released|invalid_index|zero_rows|unsaved_source"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim baseBook As Workbook, otherBook As Workbook, report As Workbook, reopened As Workbook, book As Workbook
    Dim baseSheet As Worksheet, otherSheet As Worksheet, detail As Worksheet, extra As Worksheet
    Dim session As CNxCompareResultsSession, result As CNxResult, rows As Variant, message As String
    Dim folder As String, a As String, b As String, output As String, startSecurity As Long, startEvents As Boolean
    Dim beforeValue As Variant, beforeSaved As Boolean, fso As Object, index As Long, fail As String
    Dim ownedBooks As New Collection
    On Error GoTo Failed
    folder = Environ$("LHEXCEL_PROFILE_ROOT") & "\" & name
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(folder) Then Err.Raise 5, , "fresh case folder required"
    fso.CreateFolder folder
    a = folder & "\base.xlsx": b = folder & "\compare.xlsx": output = folder & "\report.xlsx"
    Set baseBook = Workbooks.Add(xlWBATWorksheet): ownedBooks.Add baseBook
    Set baseSheet = baseBook.Worksheets(1): baseSheet.Name = "검토"
    baseSheet.Range("B2").Value2 = 1: baseSheet.Range("B3").Value2 = 2
    If name = "missing_sheet" Then
        Set extra = baseBook.Worksheets.Add: extra.Name = "기준만"
    End If
    baseBook.SaveAs Filename:=a, FileFormat:=xlOpenXMLWorkbook
    Set otherBook = Workbooks.Add(xlWBATWorksheet): ownedBooks.Add otherBook
    Set otherSheet = otherBook.Worksheets(1): otherSheet.Name = "검토"
    otherSheet.Range("B2").Value2 = 4: otherSheet.Range("B3").Value2 = 3: otherSheet.Range("B3").Font.Bold = True
    If name = "zero_rows" Then
        otherSheet.Range("B2").Value2 = 1: otherSheet.Range("B3").Value2 = 2: otherSheet.Range("B3").Font.Bold = False
    End If
    otherBook.SaveAs Filename:=b, FileFormat:=xlOpenXMLWorkbook
    Set result = NxWorkbookCompareReport(a, b, output, True)
    Require result.Outcome = NxSuccess, "report creation failed"
    Set report = Workbooks.Open(output, UpdateLinks:=0): ownedBooks.Add report
    Set detail = report.Worksheets("셀차이")
    Set session = New CNxCompareResultsSession
    session.Configure report
    rows = session.Rows
    Require session.IsAlive, "session not alive"
    If name <> "zero_rows" Then Require CStr(rows(1, 2)) = IIf(name = "missing_sheet", "A1", "B2"), "snapshot order mismatch"
    If name = "unsaved_source" Then baseSheet.Range("B2").Value2 = 99
    startSecurity = Application.AutomationSecurity: startEvents = Application.EnableEvents
    beforeValue = baseSheet.Range("B2").Value2: beforeSaved = baseBook.Saved
    Select Case name
        Case "snapshot"
            message = session.Navigate(1, 0)
            Require InStr(message, "이동했습니다") > 0 And ActiveWorkbook Is report, message
            Require ActiveSheet Is report.Worksheets(1), "wrong baseline snapshot"
            Require Selection.Address = "$B$2", "wrong snapshot cell"
            message = session.Navigate(2, 1)
            Require InStr(message, "이동했습니다") > 0 And ActiveSheet Is report.Worksheets(2), message
            Require Selection.Address = "$B$3", "wrong comparison snapshot cell"
        Case "navigate"
            message = session.Navigate(1, 0)
            Require InStr(message, "이동했습니다") > 0, message
            Require ActiveWorkbook Is report And ActiveSheet Is report.Worksheets(1), "wrong baseline snapshot"
            Require Selection.Address = "$B$2", "wrong source cell"
            message = session.Navigate(2, 1)
            Require InStr(message, "이동했습니다") > 0 And ActiveSheet Is report.Worksheets(2), message
            Require Selection.Address = "$B$3", "wrong second source cell"
        Case "source_closed"
            baseBook.Close SaveChanges:=False: Set baseBook = Nothing: Set baseSheet = Nothing
            message = session.Navigate(1, 0)
            Require InStr(message, "이동했습니다") > 0, message
            Require ActiveWorkbook Is report And ActiveSheet Is report.Worksheets(1), "closed source should use snapshot"
            For Each book In Application.Workbooks
                Require book.Name <> "base.xlsx", "snapshot navigation reopened source"
            Next book
        Case "missing_sheet"
            Require CStr(rows(1, 1)) = "기준만", "missing row absent"
            message = session.Navigate(1, 1): Require InStr(message, "없습니다") > 0, "missing side allowed"
            message = session.Navigate(1, 0): Require InStr(message, "이동했습니다") > 0, message
            Require ActiveSheet.Name = "기준만", "wrong missing-sheet source"
        Case "row_changed"
            detail.Range("D2").Value2 = "changed"
            AssertBlocked session, 1, 0
        Case "path_changed"
            report.Worksheets("비교요약").Range("B2").Value2 = b
            AssertBlocked session, 1, 0
        Case "formula_field"
            detail.Range("D2").NumberFormat = "General": detail.Range("D2").Formula = "=1"
            AssertBlocked session, 1, 0
        Case "external_link"
            detail.Range("D2").Hyperlinks.Delete
            detail.Hyperlinks.Add Anchor:=detail.Range("D2"), Address:="https://example.invalid", TextToDisplay:=CStr(rows(1, 4))
            AssertBlocked session, 1, 0
        Case "hidden_row"
            baseSheet.Rows(2).Hidden = True: beforeSaved = baseBook.Saved
            message = session.Navigate(1, 0)
            Require InStr(message, "이동했습니다") > 0 And baseSheet.Rows(2).Hidden, "hidden source was changed"
            Require ActiveWorkbook Is report And ActiveSheet Is report.Worksheets(1), "hidden source should use visible snapshot"
        Case "closed_report"
            report.Close SaveChanges:=False
            Require Not session.IsAlive, "closed report remained alive"
            Set reopened = Workbooks.Open(output, UpdateLinks:=0)
            ownedBooks.Add reopened
            Require Not session.IsAlive, "reopened report reused old identity"
            AssertBlocked session, 1, 0
        Case "replaced_sheet"
            detail.Delete
            Set extra = report.Worksheets.Add: extra.Name = "셀차이"
            AssertBlocked session, 1, 0
        Case "released"
            session.ReleaseSession
            Require Not session.IsAlive, "released session alive"
            AssertBlocked session, 1, 0
        Case "invalid_index"
            AssertBlocked session, 0, 0
            AssertBlocked session, 999, 0
            AssertBlocked session, 1, 2
        Case "zero_rows"
            Require IsEmpty(rows), "zero result not empty"
            AssertBlocked session, 1, 0
        Case "unsaved_source"
            message = session.Navigate(1, 0)
            Require InStr(message, "이동했습니다") > 0 And ActiveWorkbook Is report, message
            Require ActiveCell.Value2 = 1, "snapshot changed with unsaved source"
        Case Else: Err.Raise 5, , "Unknown r67 case"
    End Select
    If Not baseBook Is Nothing Then
        Require baseSheet.Range("B2").Value2 = beforeValue, "source value changed"
        Require baseBook.Saved = beforeSaved, "source Saved changed"
    End If
    Require Application.AutomationSecurity = startSecurity And Application.EnableEvents = startEvents, "application state changed"
    RunCase = "PASS|" & name
CleanUp:
    On Error Resume Next
    If Not session Is Nothing Then session.ReleaseSession
    ' Only actual objects opened or created by this case are owned.
    For index = ownedBooks.Count To 1 Step -1
        Set book = ownedBooks(index)
        book.Close SaveChanges:=False
    Next index
    On Error GoTo 0
    Exit Function
Failed:
    fail = CStr(Err.Number) & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & fail
    Resume CleanUp
End Function

Private Sub AssertBlocked(ByVal session As CNxCompareResultsSession, ByVal index As Long, ByVal side As Long)
    Dim beforeBook As Workbook, beforeSheet As Worksheet, beforeAddress As String, message As String
    Set beforeBook = ActiveWorkbook: Set beforeSheet = ActiveSheet: beforeAddress = Selection.Address
    message = session.Navigate(index, side)
    Require InStr(message, "이동했습니다") = 0, "blocked request reported success"
    Require ActiveWorkbook Is beforeBook, "blocked request changed workbook"
    Require ActiveSheet Is beforeSheet, "blocked request changed sheet"
    Require Selection.Address = beforeAddress, "blocked request changed selection"
End Sub

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 1670, "T_R67CompareResults", message
End Sub
