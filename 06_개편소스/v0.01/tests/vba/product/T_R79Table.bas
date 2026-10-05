Attribute VB_Name = "T_R79Table"
Option Explicit
Public Function Names() As String
    Names = "sparse_merge_90000|dense_merge|crossing_merge|shape_outside|sheet_edge"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, journal As New CNxCellRollbackJournal
    Dim context As CNxExecutionContext, plan As CNxPlan, authority As New Collection
    Dim mask As NxCellMutationMask, started As Double, elapsed As Double
    Dim oldAlerts As Boolean, oldEvents As Boolean, countBefore As Long, n As Long, detail As String, handle As Integer
    On Error GoTo Failed
    oldAlerts = Application.DisplayAlerts: oldEvents = Application.EnableEvents
    Application.DisplayAlerts = False: Application.EnableEvents = False
    countBefore = Workbooks.Count
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1:AX1800")
    If name <> "sparse_merge_90000" Then Set target = book.Worksheets(1).Range("A1:T20")
    If name = "sheet_edge" Then Set target = book.Worksheets(1).Range("XEJ1048557:XFC1048576")
    target.Font.Size = 9
    target.Cells(1).Value2 = "header"
    target.Cells(1, 1).Resize(1, 2).Merge
    If name = "dense_merge" Then
        For n = 2 To 20: book.Worksheets(1).Range("A" & n & ":B" & n).Merge: Next n
    ElseIf name = "shape_outside" Then
        book.Worksheets(1).Shapes.AddShape 1, 800, 800, 20, 20
    End If
    target.Select
    Set context = NxContextFactory.CaptureCurrent()
    context.BindSelectionAuthority authority
    If name = "crossing_merge" Then book.Worksheets(1).Range("T20:U20").Merge
    mask = NxCellFontMask Or NxCellInteriorMask Or NxCellBorderMask Or NxCellAlignmentMask
    Set plan = NxCreatePlan("NX-DRAW-BUSINESS-TABLE", context.Fingerprint, context.SourceStateDigest)
    plan.AddSideEffect NxCreateFeatureCellEffect("NX-DRAW-BUSINESS-TABLE", context.SelectionTargetIdentity, mask, CLng(target.CountLarge))
    plan.Seal
    book.Saved = True
    started = Timer
    On Error Resume Next
    journal.Capture context, plan, authority
    n = Err.Number: detail = Err.Description: Err.Clear
    On Error GoTo Failed
    elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400
    If name = "crossing_merge" Then
        If n = 0 Or InStr(detail, "Merge") = 0 Then Err.Raise 5, , "Crossing merge accepted: " & detail
    Else
        If n <> 0 Then Err.Raise n, , detail
        If Not book.Saved Then Err.Raise 5, , "Backup changed Saved state"
        target.Font.Size = 16
        detail = journal.Rollback()
        If Len(detail) > 0 Then Err.Raise 5, , detail
        If target.Font.Size <> 9 Then Err.Raise 5, , "Font not restored"
        If Not target.Cells(1, 1).MergeCells Then Err.Raise 5, , "Merge changed"
        If name = "shape_outside" And book.Worksheets(1).Shapes.Count <> 1 Then Err.Raise 5, , "Shape changed"
    End If
    journal.ReleaseResources
    handle = FreeFile
    Open Environ$("LHEXCEL_PROFILE_ROOT") & "\table-timing.tsv" For Append As #handle
    Print #handle, name & vbTab & Format$(elapsed, "0.000")
    Close #handle
    book.Close False: Set book = Nothing
    If Workbooks.Count <> countBefore Then Err.Raise 5, , "Workbook leaked"
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & "|" & Err.Description
Clean:
    On Error Resume Next
    journal.ReleaseResources
    If Not book Is Nothing Then book.Close False
    Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents
End Function
