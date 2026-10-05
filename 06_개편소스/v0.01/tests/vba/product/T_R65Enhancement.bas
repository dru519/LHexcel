Attribute VB_Name = "T_R65Enhancement"
Option Explicit

Public Function NxR65Probe(ByVal kind As String) As String
    Dim target As Range, rows() As Variant, item As Object, book As Workbook, sheet As Worksheet
    Dim beforeSaved As Boolean, beforeShapes As Long, beforeValue As Variant
    Dim request As CNxPictureInsertRequest, preview As CNxPictureInsertPreview, paths As New Collection
    Dim root As String, side As Variant, row As Long, column As Long, data() As Variant, result As CNxResult
    Dim started As Double, status As Variant, cancelKey As XlEnableCancelKey, beforeCount As Long, height As Long
    Dim cancellation As CNxR65CancelCompare, failure As Long
    On Error GoTo Failed
    root = Environ$("LHEXCEL_PROFILE_ROOT")
    Select Case kind
        Case "document-prepare"
            Set book = ActiveWorkbook
            For row = 1 To 80
                Set sheet = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
                sheet.Name = "Sheet " & Format$(row, "000")
                If row = 2 Then sheet.Visible = xlSheetHidden
                If row = 3 Then sheet.Visible = xlSheetVeryHidden
            Next row
            book.Worksheets(1).Activate
            Set target = ActiveSheet.Range("B2"): target.Value2 = "keep": target.Select
            NxR65Probe = "PASS|document-prepare"
        Case "picture"
            Set target = ActiveSheet.Range("B2")
            paths.Add root & "\preview.png"
            Set request = NxPictureInsertCreateRequest(target, paths, "FIT")
            Set preview = NxPictureInsertBuildPreview(request)
            beforeSaved = target.Parent.Parent.Saved: beforeShapes = target.Parent.Shapes.Count: beforeValue = target.Value2
            If Not NxPictureInsertNativePreview(request, preview) Then Err.Raise vbObjectError + 965, , "Picture preview not accepted"
            If target.Parent.Parent.Saved <> beforeSaved Or target.Parent.Shapes.Count <> beforeShapes Or target.Value2 <> beforeValue Then Err.Raise vbObjectError + 965, , "Preview mutated workbook"
            NxR65Probe = "PASS|picture"
        Case "compare-prepare", "compare-large-prepare"
            height = 500
            If kind = "compare-large-prepare" Then
                height = 10000: root = root & "\large": MkDir root
            End If
            Set book = ActiveWorkbook
            For Each side In Array("base", "other")
                Set item = Workbooks.Add(xlWBATWorksheet)
                Set sheet = item.Worksheets(1)
                ReDim data(1 To height, 1 To 20)
                For row = 1 To height
                    For column = 1 To 20: data(row, column) = CDbl((row - 1) * 20 + column - 1): Next column
                Next row
                sheet.Range("A1:T" & CStr(height)).Value2 = data
                If CStr(side) = "other" Then
                    sheet.Range("B2").Value2 = 42: sheet.Range("D4").Formula = "=1+2"
                End If
                item.SaveAs Filename:=root & "\" & CStr(side) & ".xlsx", FileFormat:=xlOpenXMLWorkbook
                item.Close False
            Next side
            book.Activate
            NxR65Probe = "PASS|" & kind
        Case "compare-values", "compare-formats", "compare-large"
            If kind = "compare-large" Then root = root & "\large"
            NxHostHideNavigator
            If kind = "compare-values" Then Application.StatusBar = False Else Application.StatusBar = "R65 preserve status"
            Set item = NxHostCreateWorkbookCompare()
            If item Is Nothing Then Err.Raise vbObjectError + 965, , "Native compare unavailable"
            status = Application.StatusBar: cancelKey = Application.EnableCancelKey: beforeCount = Workbooks.Count
            started = Timer
            Set result = NxWorkbookCompareReport(root & "\base.xlsx", root & "\other.xlsx", root & "\" & kind & ".xlsx", kind = "compare-formats")
            If result.Outcome <> NxSuccess Then Err.Raise vbObjectError + 965, , "Compare outcome"
            If Application.StatusBar <> status Or Application.EnableCancelKey <> cancelKey Or Workbooks.Count <> beforeCount Then Err.Raise vbObjectError + 965, , "Compare state leak: status=" & CStr(status) & "(" & CStr(VarType(status)) & ")/" & CStr(Application.StatusBar) & "(" & CStr(VarType(Application.StatusBar)) & "); cancel=" & CStr(cancelKey) & "/" & CStr(Application.EnableCancelKey) & "; books=" & CStr(beforeCount) & "/" & CStr(Workbooks.Count)
            NxR65Probe = "PASS|" & kind & "|" & Format$((Timer - started) * 1000, "0")
        Case "compare-cancel"
            beforeCount = Workbooks.Count: status = Application.StatusBar: cancelKey = Application.EnableCancelKey
            Set cancellation = New CNxR65CancelCompare
            NxWorkbookCompareBindProgress cancellation
            On Error Resume Next
            Set result = NxWorkbookCompareReport(root & "\base.xlsx", root & "\other.xlsx", root & "\cancelled.xlsx", False)
            failure = Err.Number: Err.Clear
            On Error GoTo Failed
            If failure <> 18 Or Not cancellation.Triggered Then Err.Raise vbObjectError + 965, , "Cancellation was not observed"
            If Len(Dir$(root & "\cancelled.xlsx")) <> 0 Or Workbooks.Count <> beforeCount Then Err.Raise vbObjectError + 965, , "Cancelled report residue"
            If Application.StatusBar <> status Or Application.EnableCancelKey <> cancelKey Then Err.Raise vbObjectError + 965, , "Cancelled compare state leak"
            NxR65Probe = "PASS|compare-cancel"
        Case Else
            Err.Raise vbObjectError + 965, , "Unknown case"
    End Select
    Exit Function
Failed:
    NxR65Probe = "FAIL|" & kind & "|" & CStr(Err.Number) & "|" & Err.Description
End Function
