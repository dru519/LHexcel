Attribute VB_Name = "T_R100FocusPreview"
Option Explicit
Public Sub Configure(ByVal folder As String)
End Sub
Public Function Names() As String
    Names = "options|invalid_input|sheet_edge|picture_options|product_registry"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, panel As FNxFocusSettings, picturePanel As FNxHangulPictures, detail As String
    Dim originalSaved As Boolean, originalRules As Long, originalSelection As String
    Dim registry As CNxFeatureRegistry, route As Variant, checked As Long
    On Error GoTo Failed
    If name = "product_registry" Then
        ' Ribbon clicks parse the tag through the full product registry.
        Set registry = NxCreateProductRegistry()
        Assert registry.FeatureCount = NxProductSurfaceFeatureCount(), "Registry count"
        Assert registry.FeatureById(NX_FEATURE_HANGUL_PICTURE_SEND).ReuseMode = "clean-room", "Picture reuse mode"
        For Each route In NxGeneratedNavigationItems()
            If CStr(route(1)) = "feature" Then
                Assert NxParseRibbonTag("nx1|feature|" & CStr(route(2))).Identifier = CStr(route(2)), "Ribbon " & CStr(route(2))
                checked = checked + 1
            End If
        Next route
        Assert checked > 30, "Ribbon feature routes"
        RunCase = "PASS|" & name
        GoTo CleanUp
    End If
    If name = "picture_options" Then
        Set picturePanel = New FNxHangulPictures
        Load picturePanel
        Assert picturePanel.txtWidth.Value = "7", "Default width"
        Assert picturePanel.txtHeight.Value = "6", "Default height"
        Assert picturePanel.cboColumns.Value = "2", "Default columns"
        Assert picturePanel.cboRows.Value = "1", "Default rows"
        Assert picturePanel.cboSort.ListCount = 2, "Sort choices"
        Assert picturePanel.cboSort.ListIndex = 0, "Default title order"
        picturePanel.cboSort.ListIndex = 1
        Assert InStr(picturePanel.lblStatus.Caption, picturePanel.cboSort.Value) > 0, "Capture time option"
        Assert Not picturePanel.cmdExecute.Enabled, "No files must not execute"
        picturePanel.chkTitles.Value = True
        Assert InStr(picturePanel.Controls("nxPhoto_1").Caption, "제목행") > 0, "Caption preview"
        RunCase = "PASS|" & name
        GoTo CleanUp
    End If
    Set book = Workbooks.Add(xlWBATWorksheet)
    book.Worksheets(1).Range("A1:E6").Value2 = "표본"
    book.Worksheets(1).Range("C3").Select
    If name = "sheet_edge" Then book.Worksheets(1).Cells(1048576, 16384).Select
    originalSaved = book.Saved: originalRules = book.Worksheets(1).Cells.FormatConditions.Count
    originalSelection = Selection.Address
    Set panel = New FNxFocusSettings
    Load panel
    panel.txtColor.Value = "#5FC8D8": panel.txtIntensity.Value = "100"
    panel.cboShape.ListIndex = 0
    If name = "options" Then
        Assert panel.Controls("nxFocus_cell_3_1").BackColor = RGB(95, 200, 216), "Cross row"
        Assert panel.Controls("nxFocus_cell_1_3").BackColor = RGB(95, 200, 216), "Cross column"
        Assert panel.Controls("nxFocus_cell_3_3").BackColor = vbWhite, "Selected cell excluded"
        panel.cboShape.ListIndex = 1
        Assert panel.Controls("nxFocus_cell_1_3").BackColor = vbWhite, "Horizontal only"
        panel.cboShape.ListIndex = 2
        Assert panel.Controls("nxFocus_cell_3_1").BackColor = vbWhite, "Vertical only"
        panel.txtIntensity.Value = "0"
        Assert panel.Controls("nxFocus_cell_1_3").BackColor = vbWhite, "Zero intensity"
        panel.txtIntensity.Value = "100": panel.chkSelectedArea.Value = True
        Assert panel.Controls("nxFocus_top").BackColor = RGB(95, 200, 216), "Selection option"
        Assert panel.Controls("nxFocus_column_3").Caption = "C", "Column address"
        Assert panel.Controls("nxFocus_row_3").Caption = "3", "Row address"
        Assert panel.Controls("nxFocus_cell_1_1").Caption = "표본", "Actual sample"
    ElseIf name = "invalid_input" Then
        panel.txtColor.Value = "#NO"
        Assert Not panel.Controls("nxFocus_cell_1_1").Visible, "Invalid sample hidden"
        panel.txtColor.Value = "#5FC8D8"
        Assert panel.Controls("nxFocus_cell_1_1").Visible, "Valid sample recovered"
    Else
        Assert panel.Controls("nxFocus_column_5").Caption = "XFD", "Last column"
        Assert panel.Controls("nxFocus_row_6").Caption = "1048576", "Last row"
    End If
    Unload panel: Set panel = Nothing
    Assert book.Saved = originalSaved, "Saved state changed"
    Assert Selection.Address = originalSelection, "Selection changed"
    Assert book.Worksheets(1).Cells.FormatConditions.Count = originalRules, "Focus rules changed"
    Assert book.Worksheets(1).Range("A1").Value2 = "표본", "Content changed"
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    If Not picturePanel Is Nothing Then Unload picturePanel
    If Not panel Is Nothing Then Unload panel
    If Not book Is Nothing Then book.Close False
End Function
Private Sub Assert(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise 5, , detail
End Sub
