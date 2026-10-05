Attribute VB_Name = "T_VisibleCellsClipboard"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestVisibleCopySerializesRowsAndColumns"
    names.Add "TestVisibleCopyWritesUnicodeClipboard"
    names.Add "TestVisiblePasteWritesExactLiteralValues"
    names.Add "TestVisiblePasteRejectsInvalidTargets"
    names.Add "TestVisiblePasteRollbackRestoresEarlierWrites"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestVisibleCopySerializesRowsAndColumns": TestVisibleCopySerializesRowsAndColumns
        Case "TestVisibleCopyWritesUnicodeClipboard": TestVisibleCopyWritesUnicodeClipboard
        Case "TestVisiblePasteWritesExactLiteralValues": TestVisiblePasteWritesExactLiteralValues
        Case "TestVisiblePasteRejectsInvalidTargets": TestVisiblePasteRejectsInvalidTargets
        Case "TestVisiblePasteRollbackRestoresEarlierWrites": TestVisiblePasteRollbackRestoresEarlierWrites
        Case Else: NxRaiseContractError "Unknown visible-cell clipboard test"
    End Select
End Sub

Private Sub TestVisibleCopySerializesRowsAndColumns()
    Dim fixtureBook As Workbook
    Dim source As Range
    Dim addresses As Variant
    Dim expectedText As String
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    Set fixtureBook = NewVisibleFixture(source)
    expectedText = "A1" & vbTab & "C1" & vbCrLf & "A3" & vbTab & "C3" & vbCrLf & "A4" & vbTab & "C4"
    NxTestHarness.AssertTrue NxVisibleCellsCount(source) = 6, "Visible-cell count changed"
    NxTestHarness.AssertTrue NxVisibleCellsCopyText(source) = expectedText, "Visible-cell row-major serialization changed"
    addresses = NxVisibleCellsAddresses(source)
    NxTestHarness.AssertTrue Join(addresses, ",") = "$A$1,$C$1,$A$3,$C$3,$A$4,$C$4", "Visible-cell address order changed"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_VisibleCellsClipboard.TestVisibleCopySerializesRowsAndColumns", failureDescription
End Sub

Private Sub TestVisibleCopyWritesUnicodeClipboard()
    Dim fixtureBook As Workbook
    Dim source As Range
    Dim result As CNxResult
    Dim expectedText As String
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    NxClipboardSetTestFailure False
    NxClipboardSetTestRead "기존 클립보드"
    Set fixtureBook = NewVisibleFixture(source)
    expectedText = NxVisibleCellsCopyText(source)
    Set result = NxVisibleCellsCopy(source)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Visible-cell clipboard copy did not succeed"
    NxTestHarness.AssertTrue result.Stage = "complete", "Visible-cell clipboard copy result was not final"
    NxTestHarness.AssertTrue Len(result.Recovery) = 0, "Visible-cell clipboard copy unexpectedly required recovery"
    NxTestHarness.AssertTrue NxClipboardTestCallCount() = 1, "Unicode clipboard adapter was not called exactly once"
    NxTestHarness.AssertTrue NxClipboardTestLastValue() = expectedText, "Unicode clipboard payload changed"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    NxClipboardResetTestMode
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_VisibleCellsClipboard.TestVisibleCopyWritesUnicodeClipboard", failureDescription
End Sub

Private Sub TestVisiblePasteWritesExactLiteralValues()
    Dim fixtureBook As Workbook
    Dim source As Range
    Dim result As CNxResult
    Dim values As Variant
    Dim expectedAddresses As Variant
    Dim index As Long
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    Set fixtureBook = NewVisibleFixture(source)
    values = Array("=1+1", "+SUM", "-9", "@tag", "plain", "last")
    expectedAddresses = Array("A1", "C1", "A3", "C3", "A4", "C4")
    Set result = ExecutePlannedPaste(source, values)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Visible-cell planned paste did not succeed"
    NxTestHarness.AssertTrue result.Stage = "complete", "Visible-cell planned paste result was not final"
    NxTestHarness.AssertTrue Len(result.Recovery) = 0, "Visible-cell planned paste unexpectedly required recovery"
    For index = 0 To UBound(expectedAddresses)
        With source.Worksheet.Range(CStr(expectedAddresses(index)))
            NxTestHarness.AssertTrue CStr(.Value2) = CStr(values(index)), "Visible-cell literal value changed"
            NxTestHarness.AssertTrue Not .HasFormula, "Formula-like clipboard text was executed"
        End With
    Next index
    NxTestHarness.AssertTrue CStr(source.Worksheet.Range("B1").Value2) = "B1", "Hidden column value changed"
    NxTestHarness.AssertTrue CStr(source.Worksheet.Range("A2").Value2) = "A2", "Hidden row value changed"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_VisibleCellsClipboard.TestVisiblePasteWritesExactLiteralValues", failureDescription
End Sub

Private Sub TestVisiblePasteRejectsInvalidTargets()
    Dim fixtureBook As Workbook
    Dim source As Range
    Dim mergedSource As Range
    Dim ignored As CNxResult
    Dim before As String
    Dim mismatchError As Long
    Dim mergedError As Long
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    Set fixtureBook = NewVisibleFixture(source)
    before = RangeFingerprint(source)
    NxClipboardSetTestRead "one" & vbTab & "two"
    On Error Resume Next
    Set ignored = NxVisibleCellsPasteValues(source)
    mismatchError = Err.Number
    Err.Clear
    On Error GoTo Failed
    NxTestHarness.AssertTrue mismatchError <> 0, "Visible-cell paste accepted a mismatched clipboard count"
    NxTestHarness.AssertTrue RangeFingerprint(source) = before, "Rejected visible-cell paste changed the source"
    NxTestHarness.AssertTrue Application.Selection.Address = source.Address, "Rejected visible-cell paste did not restore selection"

    Set mergedSource = source.Worksheet.Range("A1:A2")
    mergedSource.ClearContents
    mergedSource.Merge
    mergedSource.Select
    NxClipboardSetTestRead "one"
    On Error Resume Next
    Set ignored = NxVisibleCellsPasteValues(mergedSource)
    mergedError = Err.Number
    Err.Clear
    On Error GoTo Failed
    NxTestHarness.AssertTrue mergedError <> 0, "Visible-cell paste accepted a merged target"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    NxClipboardResetTestMode
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_VisibleCellsClipboard.TestVisiblePasteRejectsInvalidTargets", failureDescription
End Sub

Private Sub TestVisiblePasteRollbackRestoresEarlierWrites()
    Dim fixtureBook As Workbook
    Dim source As Range
    Dim result As CNxResult
    Dim values As Variant
    Dim before As String
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    Set fixtureBook = Application.Workbooks.Add
    Set source = fixtureBook.Worksheets(1).Range("A1:A3")
    source.Value2 = Application.Transpose(Array("old1", "old2", "old3"))
    source.Select
    before = RangeFingerprint(source)
    values = Array("new1", CVErr(xlErrValue), "new3")
    Set result = ExecutePlannedPaste(source, values)
    NxTestHarness.AssertTrue result.Outcome = NxEnvironmentError, "Injected visible-cell paste failure changed outcome"
    NxTestHarness.AssertTrue RangeFingerprint(source) = before, "Visible-cell paste rollback left an earlier write"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_VisibleCellsClipboard.TestVisiblePasteRollbackRestoresEarlierWrites", failureDescription
End Sub

Private Function NewVisibleFixture(ByRef source As Range) As Workbook
    Dim fixtureBook As Workbook
    Dim sheet As Worksheet
    Dim values(1 To 4, 1 To 3) As Variant
    Dim rowIndex As Long
    Dim columnIndex As Long
    Set fixtureBook = Application.Workbooks.Add
    Set sheet = fixtureBook.Worksheets(1)
    For rowIndex = 1 To 4
        For columnIndex = 1 To 3
            values(rowIndex, columnIndex) = Chr$(64 + columnIndex) & CStr(rowIndex)
        Next columnIndex
    Next rowIndex
    sheet.Range("A1:C4").Value2 = values
    Set source = sheet.Range("A1:C4")
    source.Rows(2).EntireRow.Hidden = True
    source.Columns(2).EntireColumn.Hidden = True
    source.Select
    Set NewVisibleFixture = fixtureBook
End Function

Private Function ExecutePlannedPaste(ByVal source As Range, ByVal values As Variant) As CNxResult
    Dim visibleTarget As Range
    Dim addresses As Variant
    Dim commandObject As New CNxVisibleCellsCommand
    Dim command As INxFeatureCommand
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim definition As CNxFeatureDefinition
    addresses = NxVisibleCellsAddresses(source)
    Set visibleTarget = NxVisibleCellsTarget(source)
    visibleTarget.Select
    commandObject.ConfigurePaste addresses, values, NxVisibleCellsCount(source)
    Set command = commandObject
    Set definition = NxDataFeatureDefinition(NX_FEATURE_DATA_PASTE_VISIBLE_VALUES)
    Set ticket = router.Prepare(definition, command)
    Set ExecutePlannedPaste = NxDataRunPlannedFrameForTest(command, ticket, router, "visible-cell test")
End Function

Private Function RangeFingerprint(ByVal source As Range) As String
    Dim cell As Range
    For Each cell In source.Cells
        RangeFingerprint = RangeFingerprint & cell.Address(False, False) & ":" & CStr(cell.Formula) & "|"
    Next cell
End Function
