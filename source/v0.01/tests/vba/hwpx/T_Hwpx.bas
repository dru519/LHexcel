Attribute VB_Name = "T_Hwpx"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestOnlyOneContiguousRangeIsSerialized"
    names.Add "TestSerializedRowsColumnsAndCells"
    names.Add "TestSerializedMergeIsRelative"
    names.Add "TestFormulaAndWorkbookPathAreAbsent"
    names.Add "TestHiddenRowsAreExplicitlyExcluded"
    names.Add "TestPreviewLeavesSourceUnchanged"
    names.Add "TestHwpxRegistryHasFastAndGuardedBases"
    names.Add "TestExportResolvesToPlannedForFileMutation"
    names.Add "TestHwpxFormHasRequiredControls"
    names.Add "TestHwpxFormUsesGeneratedSources"
    names.Add "TestPowerShellArgumentsRejectUnsafeQuotes"
    names.Add "TestEmbeddedRuntimeUsesVersionedTempRoot"
    names.Add "TestEmbeddedRuntimeHasExpectedResourceNames"
    names.Add "TestHwpxRequestUsesLocalPlannedContract"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestOnlyOneContiguousRangeIsSerialized": TestOnlyOneContiguousRangeIsSerialized
        Case "TestSerializedRowsColumnsAndCells": TestSerializedRowsColumnsAndCells
        Case "TestSerializedMergeIsRelative": TestSerializedMergeIsRelative
        Case "TestFormulaAndWorkbookPathAreAbsent": TestFormulaAndWorkbookPathAreAbsent
        Case "TestHiddenRowsAreExplicitlyExcluded": TestHiddenRowsAreExplicitlyExcluded
        Case "TestPreviewLeavesSourceUnchanged": TestPreviewLeavesSourceUnchanged
        Case "TestHwpxRegistryHasFastAndGuardedBases": TestHwpxRegistryHasFastAndGuardedBases
        Case "TestExportResolvesToPlannedForFileMutation": TestExportResolvesToPlannedForFileMutation
        Case "TestHwpxFormHasRequiredControls": TestHwpxFormHasRequiredControls
        Case "TestHwpxFormUsesGeneratedSources": TestHwpxFormUsesGeneratedSources
        Case "TestPowerShellArgumentsRejectUnsafeQuotes": TestPowerShellArgumentsRejectUnsafeQuotes
        Case "TestEmbeddedRuntimeUsesVersionedTempRoot": TestEmbeddedRuntimeUsesVersionedTempRoot
        Case "TestEmbeddedRuntimeHasExpectedResourceNames": TestEmbeddedRuntimeHasExpectedResourceNames
        Case "TestHwpxRequestUsesLocalPlannedContract": TestHwpxRequestUsesLocalPlannedContract
        Case Else: NxRaiseContractError "Unknown HWPX test"
    End Select
End Sub

Private Sub TestOnlyOneContiguousRangeIsSerialized()
    Dim book As Workbook, rejected As Boolean
    On Error GoTo Expected
    Set book = Application.Workbooks.Add
    Call NxHwpxSerialize(Union(book.Worksheets(1).Range("A1:B2"), book.Worksheets(1).Range("D1:E2")), True, True, "display")
    GoTo Cleanup
Expected:
    rejected = (InStr(1, Err.Description, "연속된 한 범위", vbBinaryCompare) > 0)
    Err.Clear
Cleanup:
    On Error Resume Next
    book.Close False
    On Error GoTo 0
    NxTestHarness.AssertTrue rejected, "Multiple areas were serialized"
End Sub

Private Sub TestSerializedRowsColumnsAndCells()
    Dim book As Workbook, source As Range, payload As String
    Dim number As Long, description As String
    On Error GoTo Failed
    Set book = CreateFixture(source)
    payload = NxHwpxSerialize(source, True, True, "display")
    NxTestHarness.AssertTrue InStr(payload, """rows"":4") > 0, "Serialized row count changed"
    NxTestHarness.AssertTrue InStr(payload, """columns"":3") > 0, "Serialized column count changed"
    NxTestHarness.AssertTrue CountText(payload, """row"":") = 12, "Dense cell count changed"
    GoTo Cleanup
Failed:
    number = Err.Number: description = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    book.Close False
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "T_Hwpx.TestSerializedRowsColumnsAndCells", description
End Sub

Private Sub TestSerializedMergeIsRelative()
    Dim book As Workbook, source As Range, payload As String
    Dim number As Long, description As String
    On Error GoTo Failed
    Set book = CreateFixture(source)
    source.Range("B2:C2").Merge
    source.Range("B2").Value2 = "병합"
    payload = NxHwpxSerialize(source, True, True, "display")
    NxTestHarness.AssertTrue InStr(payload, """row"":2,""column"":2,""row_span"":1,""column_span"":2") > 0, "Merge coordinates changed"
    GoTo Cleanup
Failed:
    number = Err.Number: description = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    book.Close False
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "T_Hwpx.TestSerializedMergeIsRelative", description
End Sub

Private Sub TestFormulaAndWorkbookPathAreAbsent()
    Dim book As Workbook, source As Range, payload As String
    Dim number As Long, description As String
    On Error GoTo Failed
    Set book = CreateFixture(source)
    source.Cells(2, 2).Formula = "=1+1"
    payload = NxHwpxSerialize(source, True, True, "display")
    NxTestHarness.AssertTrue InStr(payload, "=1+1") = 0, "Formula source leaked"
    NxTestHarness.AssertTrue InStr(payload, book.FullName) = 0, "Workbook path leaked"
    NxTestHarness.AssertTrue InStr(payload, """source_path""") = 0, "Source metadata leaked"
    GoTo Cleanup
Failed:
    number = Err.Number: description = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    book.Close False
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "T_Hwpx.TestFormulaAndWorkbookPathAreAbsent", description
End Sub

Private Sub TestHiddenRowsAreExplicitlyExcluded()
    Dim book As Workbook, source As Range, payload As String
    Dim number As Long, description As String
    On Error GoTo Failed
    Set book = CreateFixture(source)
    source.Rows(3).EntireRow.Hidden = True
    payload = NxHwpxSerialize(source, True, False, "value")
    NxTestHarness.AssertTrue InStr(payload, """rows"":3") > 0, "Hidden row was not excluded"
    NxTestHarness.AssertTrue CountText(payload, """row"":") = 9, "Hidden-row dense cell count changed"
    GoTo Cleanup
Failed:
    number = Err.Number: description = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    book.Close False
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "T_Hwpx.TestHiddenRowsAreExplicitlyExcluded", description
End Sub

Private Sub TestPreviewLeavesSourceUnchanged()
    Dim book As Workbook, source As Range, before As String
    Dim number As Long, description As String
    On Error GoTo Failed
    Set book = CreateFixture(source)
    before = Fingerprint(source)
    Call NxHwpxPreviewText(source, True, True, "display")
    NxTestHarness.AssertTrue Fingerprint(source) = before, "HWPX preview changed source"
    GoTo Cleanup
Failed:
    number = Err.Number: description = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    book.Close False
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "T_Hwpx.TestPreviewLeavesSourceUnchanged", description
End Sub

Private Sub TestHwpxRegistryHasFastAndGuardedBases()
    Dim preview As CNxFeatureDefinition, exportOpen As CNxFeatureDefinition
    Set preview = NxHwpxFeatureDefinition("NX-HWPX-PREVIEW")
    Set exportOpen = NxHwpxFeatureDefinition("NX-HWPX-EXPORT-OPEN")
    NxTestHarness.AssertTrue preview.ExecutionGrade = NxExecutionFast, "HWPX preview must start Fast"
    NxTestHarness.AssertTrue Not preview.ExternalDataTransfer, "Preview invoked external transfer"
    NxTestHarness.AssertTrue exportOpen.ExecutionGrade = NxExecutionPlanned, "HWPX export must start Planned"
    NxTestHarness.AssertTrue Not exportOpen.ExternalDataTransfer, "HWPX export must remain local"
End Sub

Private Sub TestExportResolvesToPlannedForFileMutation()
    NxTestHarness.AssertTrue NxExecutionGradePolicy.ResolveGrade(NxExecutionGuarded, 12, 5000, False, True, True, False, False) = NxExecutionPlanned, "HWPX file creation did not resolve to Planned"
End Sub

Private Sub TestHwpxFormHasRequiredControls()
    Dim controlName As Variant
    For Each controlName In Array("txtRangeSummary", "cboTitleRow", "cboHidden", "cboSourceMode", _
        "txtExcludedStyles", "txtOutputSafety", "txtOutputFolder", _
        "txtPreview", "cmdPreview", "cmdCreateOpen", "cmdCancel")
        NxTestHarness.AssertTrue NxHwpxFormContractHas(CStr(controlName)), "Required HWPX form control is missing"
    Next controlName
End Sub

Private Sub TestHwpxFormUsesGeneratedSources()
    NxTestHarness.AssertTrue NxHwpxUsesGeneratedFormSources(), "HWPX form must use JSON plus VBA sources"
End Sub

Private Sub TestPowerShellArgumentsRejectUnsafeQuotes()
    Dim rejected As Boolean
    On Error GoTo Expected
    Call NxPowerShellQuoted("C:\unsafe""path")
    GoTo AssertResult
Expected:
    rejected = True
    Err.Clear
AssertResult:
    NxTestHarness.AssertTrue rejected, "Unsafe PowerShell argument was accepted"
End Sub

Private Sub TestEmbeddedRuntimeUsesVersionedTempRoot()
    Dim root As String
    root = NxHwpxEmbeddedRoot()
    NxTestHarness.AssertTrue InStr(1, root, "LHexcel\embedded_support\v2026_08_05_", vbBinaryCompare) > 0, "Embedded runtime root is not versioned"
End Sub

Private Sub TestEmbeddedRuntimeHasExpectedResourceNames()
    Dim names As String
    names = NxHwpxEmbeddedResourceNames()
    NxTestHarness.AssertTrue InStr(1, names, "tools\lhexcel_hwpx_table_export.ps1", vbBinaryCompare) > 0, "Embedded PowerShell resource is missing"
    NxTestHarness.AssertTrue InStr(1, names, "templates\표.hwpx", vbBinaryCompare) > 0, "Embedded HWPX template is missing"
End Sub

Private Sub TestHwpxRequestUsesLocalPlannedContract()
    Dim book As Workbook, source As Range, request As New CNxHwpxRequest
    Dim number As Long, description As String
    On Error GoTo Failed
    Set book = CreateFixture(source)
    source.Cells(2, 2).Value2 = "010-1234-5678"
    source.Select
    request.Configure "NX-HWPX-EXPORT-OPEN", source, True, True, "value"
    NxTestHarness.AssertTrue InStr(1, request.OutputSafetySummary, "출력 안전", vbBinaryCompare) > 0, "HWPX request lacks output safety summary"
    GoTo Cleanup
Failed:
    number = Err.Number: description = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    book.Close False
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "T_Hwpx.TestHwpxRequestUsesLocalPlannedContract", description
End Sub

Private Function CreateFixture(ByRef source As Range) As Workbook
    Dim book As Workbook, sheet As Worksheet
    Set book = Application.Workbooks.Add
    Set sheet = book.Worksheets(1)
    sheet.Range("A1").Value2 = "구분": sheet.Range("B1").Value2 = "내용": sheet.Range("C1").Value2 = "비고"
    sheet.Range("A2").Value2 = "A": sheet.Range("B2").Value2 = "값1": sheet.Range("C2").Value2 = ""
    sheet.Range("A3").Value2 = "B": sheet.Range("B3").Value2 = 10: sheet.Range("C3").Value2 = "확인"
    sheet.Range("A4").Value2 = "C": sheet.Range("B4").Value2 = "끝": sheet.Range("C4").Value2 = ""
    Set source = sheet.Range("A1:C4")
    source.Select
    Set CreateFixture = book
End Function

Private Function Fingerprint(ByVal source As Range) As String
    Dim cell As Range
    For Each cell In source.Cells
        Fingerprint = Fingerprint & CStr(cell.Value2) & "|" & CStr(cell.Formula) & ";"
    Next cell
End Function

Private Function CountText(ByVal source As String, ByVal token As String) As Long
    Dim offset As Long
    offset = 1
    Do
        offset = InStr(offset, source, token, vbBinaryCompare)
        If offset = 0 Then Exit Do
        CountText = CountText + 1
        offset = offset + Len(token)
    Loop
End Function
