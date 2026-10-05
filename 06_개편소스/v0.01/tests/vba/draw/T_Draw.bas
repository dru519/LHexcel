Attribute VB_Name = "T_Draw"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestStyleTokensAreExact": names.Add "TestHeaderContrastFixture"
    names.Add "TestTitleBusinessStylesPreserveContent": names.Add "TestClearInnerPreservesOutline"
    names.Add "TestTableBorderHelperPreservesInnerBorders": names.Add "TestPictureLayoutContract"
    names.Add "TestPictureFitUsesSelectedShape": names.Add "TestDrawOptionsReachController"
    names.Add "TestDrawingRegistryKeepsClosedFeatureIds": names.Add "TestDrawingFormHasRequiredControls"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestStyleTokensAreExact": TestStyleTokensAreExact
        Case "TestHeaderContrastFixture": TestHeaderContrastFixture
        Case "TestTitleBusinessStylesPreserveContent": TestTitleBusinessStylesPreserveContent
        Case "TestClearInnerPreservesOutline": TestClearInnerPreservesOutline
        Case "TestTableBorderHelperPreservesInnerBorders": TestTableBorderHelperPreservesInnerBorders
        Case "TestPictureLayoutContract": TestPictureLayoutContract
        Case "TestPictureFitUsesSelectedShape": TestPictureFitUsesSelectedShape
        Case "TestDrawOptionsReachController": TestDrawOptionsReachController
        Case "TestDrawingRegistryKeepsClosedFeatureIds": TestDrawingRegistryKeepsClosedFeatureIds
        Case "TestDrawingFormHasRequiredControls": TestDrawingFormHasRequiredControls
        Case Else: NxRaiseContractError "Unknown Draw test"
    End Select
End Sub

Private Sub TestStyleTokensAreExact()
    NxTestHarness.AssertTrue NxBodyFontColor() = RGB(31, 41, 51), "body font token changed"
    NxTestHarness.AssertTrue NxHeaderFontColor() = RGB(23, 50, 77), "header font token changed"
    NxTestHarness.AssertTrue NxHeaderFillColor() = RGB(232, 238, 247), "header fill token changed"
    NxTestHarness.AssertTrue NxBaseLineWeight() = xlThin And NxEmphasisLineWeight() = xlMedium, "line weight token changed"
End Sub

Private Sub TestHeaderContrastFixture()
    NxTestHarness.AssertTrue ContrastRatio(NxHeaderFontColor(), NxHeaderFillColor()) >= 4.5, "header contrast is below WCAG threshold"
End Sub

Private Sub TestTitleBusinessStylesPreserveContent()
    Dim book As Workbook, source As Range, before As String, after As String
    Set book = Fixture(source)
    source.Cells(2, 2).Formula = "=1+1"
    source.Range("A1:B1").ClearContents
    source.Range("A1:B1").Merge
    source.Range("A1").Value2 = "fixture"
    before = ContentFingerprint(source)
    NxDrawApplyBusinessTable source, True, True
    after = ContentFingerprint(source)
    NxTestHarness.AssertTrue after = before, "table style changed content or merge"
    book.Close False
End Sub

Private Sub TestClearInnerPreservesOutline()
    Dim book As Workbook, source As Range, before As Long
    Set book = Fixture(source): NxDrawApplyOutline source: before = source.Borders(xlEdgeLeft).LineStyle
    NxDrawClearInner source
    NxTestHarness.AssertTrue source.Borders(xlEdgeLeft).LineStyle = before, "inner clear changed outline"
    book.Close False
End Sub

Private Sub TestTableBorderHelperPreservesInnerBorders()
    Dim book As Workbook, source As Range, before As Long
    Set book = Fixture(source): source.Borders(xlInsideVertical).LineStyle = xlContinuous: before = source.Borders(xlInsideVertical).LineStyle
    NxDrawApplyOutline source
    NxTestHarness.AssertTrue source.Borders(xlInsideVertical).LineStyle = before, "outline changed inner border"
    book.Close False
End Sub

Private Sub TestPictureLayoutContract()
    NxTestHarness.AssertTrue InStr(NxPictureLayoutContractText(), "LockAspectRatio") > 0, "picture layout must lock ratio"
End Sub

Private Sub TestPictureFitUsesSelectedShape()
    Dim book As Workbook, source As Range, picture As Shape, oldWidth As Double
    Set book = Fixture(source)
    Set picture = source.Worksheet.Shapes.AddShape(msoShapeRectangle, source.Left, source.Top, 80, 40)
    oldWidth = picture.Width
    NxDrawFitPicture picture, source, 2, True
    NxTestHarness.AssertTrue picture.Width > 0 And picture.Width <= source.Width, "picture must fit inside target"
    NxTestHarness.AssertTrue picture.Height > 0 And picture.Height <= source.Height, "picture height must fit inside target"
    NxTestHarness.AssertTrue picture.Placement = xlMoveAndSize, "picture placement option was not applied"
    NxTestHarness.AssertTrue picture.Width <> oldWidth, "picture fit did not resize the selected shape"
    picture.Delete: book.Close False
End Sub

Private Sub TestDrawOptionsReachController()
    NxTestHarness.AssertTrue NxDrawingFormContractHas("chkHeader"), "header option is not exposed"
    NxTestHarness.AssertTrue NxDrawingFormContractHas("chkAutoFitColumns"), "autofit option is not exposed"
    NxTestHarness.AssertTrue NxDrawingFormContractHas("txtPictureMargin"), "picture margin option is not exposed"
    NxTestHarness.AssertTrue NxDrawingFormContractHas("chkMoveAndSize"), "move and size option is not exposed"
End Sub

Private Sub TestDrawingRegistryKeepsClosedFeatureIds()
    Dim definition As CNxFeatureDefinition
    Set definition = NxDrawingFeatureDefinition(NX_FEATURE_DRAW_BUSINESS_TABLE)
    NxTestHarness.AssertTrue definition.CategoryId = "NX-CAT-DRAW", "drawing category changed"
End Sub

Private Sub TestDrawingFormHasRequiredControls()
    Dim item As Variant
    For Each item In Array("cboFeature", "lblRange", "chkHeader", "chkTotalRow", "chkPreserveAlignment", "chkAutoFitColumns", "txtPictureMargin", "chkMoveAndSize", "lblConflicts", "cmdPreview", "cmdExecute", "cmdCancel")
        NxTestHarness.AssertTrue NxDrawingFormContractHas(CStr(item)), "drawing form control missing"
    Next item
End Sub

Private Function Fixture(ByRef source As Range) As Workbook
    Dim book As Workbook
    Set book = Application.Workbooks.Add: Set source = book.Worksheets(1).Range("A1:D4")
    source.Value2 = "fixture": source.Cells(2, 2).Value2 = 12
    Set Fixture = book
End Function

Private Function ContentFingerprint(ByVal source As Range) As String
    Dim cell As Range, result As String, mergeArea As String
    For Each cell In source.Cells
        mergeArea = vbNullString
        If cell.MergeCells Then mergeArea = cell.MergeArea.Address(False, False)
        result = result & cell.Address(False, False) & "=" & ScalarFormula(cell.Formula) & ";merged=" & CStr(cell.MergeCells) & ";area=" & mergeArea & "|"
    Next cell
    ContentFingerprint = result
End Function

Private Function ScalarFormula(ByVal value As Variant) As String
    On Error GoTo Failed
    If IsError(value) Then
        ScalarFormula = "#ERROR"
    ElseIf IsNull(value) Then
        ScalarFormula = "#NULL"
    ElseIf IsEmpty(value) Then
        ScalarFormula = "#EMPTY"
    Else
        ScalarFormula = CStr(value)
    End If
    Exit Function
Failed:
    ScalarFormula = "#UNREADABLE"
End Function

Private Function ContrastRatio(ByVal foreground As Long, ByVal background As Long) As Double
    Dim a As Double, b As Double
    a = RelativeLuminance(foreground): b = RelativeLuminance(background)
    ContrastRatio = (WorksheetFunction.Max(a, b) + 0.05) / (WorksheetFunction.Min(a, b) + 0.05)
End Function

Private Function RelativeLuminance(ByVal color As Long) As Double
    Dim channels(0 To 2) As Double, i As Long, value As Double
    channels(0) = (color And &HFF&): channels(1) = (color \ &H100&) And &HFF&: channels(2) = (color \ &H10000) And &HFF&
    For i = 0 To 2
        value = channels(i) / 255#: channels(i) = IIf(value <= 0.04045, value / 12.92, ((value + 0.055) / 1.055) ^ 2.4)
    Next i
    RelativeLuminance = 0.2126 * channels(0) + 0.7152 * channels(1) + 0.0722 * channels(2)
End Function

Private Function NxPictureLayoutContractText() As String
    NxPictureLayoutContractText = "LockAspectRatio;contain;center;rollback"
End Function
