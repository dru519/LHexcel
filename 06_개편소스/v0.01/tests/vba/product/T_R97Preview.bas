Attribute VB_Name = "T_R97Preview"
Option Explicit
Private mEvidence As String
Public Sub Configure(ByVal folder As String)
    mEvidence = folder
End Sub
Public Function Names() As String
    Names = "preview_90000|execute_90000"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, target As Range, panel As FNxDrawTable
    Dim context As CNxFeatureDialogContext, definition As CNxFeatureDefinition
    Dim started As Double, iteration As Long, result As CNxResult, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1:AX1800")
    target.Value2 = "측정 자료"
    target.Font.Name = "맑은 고딕": target.Font.Size = 9
    target.RowHeight = 15: target.Columns.ColumnWidth = 12
    target.Cells(2, 2).Formula = "=1+2"
    target.Cells(1800, 50).Value2 = "마지막 값"
    target.Select
    If name = "preview_90000" Then
        For iteration = 1 To 3
            Set panel = New FNxDrawTable
            Set definition = NxDrawingFeatureDefinition(NX_FEATURE_DRAW_BUSINESS_TABLE)
            Set context = New CNxFeatureDialogContext
            context.Configure definition.FeatureId, definition.DialogId, definition.DialogVariant, definition.LaunchSurface = "workbench"
            started = Timer
            panel.BindFeatureContext context
            Metric "initial_preview", iteration, started
            started = Timer
            panel.cboDisplayMode.ListIndex = 1
            panel.cboDisplayMode.ListIndex = 0
            panel.cboTableStyle.ListIndex = 0
            panel.cboTableStyle.ListIndex = 1
            Metric "four_option_changes", iteration, started
            If Not panel.cmdExecute.Enabled Then Err.Raise 5, , "Preview blocked execution"
            If Not panel.Controls("nxPrv_cell_1_1").Visible Then Err.Raise 5, , "Preview missing"
            Unload panel: Set panel = Nothing
        Next iteration
        If target.Cells(1, 1).Interior.ColorIndex <> xlColorIndexNone Then Err.Raise 5, , "Preview mutated formatting"
    Else
        started = Timer
        Set result = NxDrawRun(NX_FEATURE_DRAW_BUSINESS_TABLE, target)
        Metric "full_execution", 1, started
        If result.Outcome <> NxSuccess Then Err.Raise 5, , result.Recovery
        If target.Font.Size <> 9 Then Err.Raise 5, , "Unexpected font"
    End If
    If target.Cells(2, 2).Formula <> "=1+2" Or target.Cells(1800, 50).Value2 <> "마지막 값" Then Err.Raise 5, , "Content changed"
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    If Not book Is Nothing Then book.Close False
End Function
Private Sub Metric(ByVal stage As String, ByVal iteration As Long, ByVal started As Double)
    Dim elapsed As Double, handle As Integer
    elapsed = Timer - started: If elapsed < 0 Then elapsed = elapsed + 86400#
    handle = FreeFile
    Open mEvidence & Application.PathSeparator & "preview-90000.tsv" For Append As #handle
    Print #handle, stage & vbTab & iteration & vbTab & Format$(elapsed * 1000#, "0.000")
    Close #handle
End Sub
