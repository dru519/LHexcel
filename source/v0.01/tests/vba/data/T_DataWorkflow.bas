Attribute VB_Name = "T_DataWorkflow"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestUniquePreviewLeavesSourceUnchanged"
    names.Add "TestDuplicateResultDoesNotChangeSource"
    names.Add "TestDuplicateResultUsesClosedHeadersAndNextName"
    names.Add "TestObservedCellsEscalateFastToPlanned"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestUniquePreviewLeavesSourceUnchanged": TestUniquePreviewLeavesSourceUnchanged
        Case "TestDuplicateResultDoesNotChangeSource": TestDuplicateResultDoesNotChangeSource
        Case "TestDuplicateResultUsesClosedHeadersAndNextName": TestDuplicateResultUsesClosedHeadersAndNextName
        Case "TestObservedCellsEscalateFastToPlanned": TestObservedCellsEscalateFastToPlanned
        Case Else: NxRaiseContractError "Unknown Data workflow test"
    End Select
End Sub

Private Sub TestUniquePreviewLeavesSourceUnchanged()
    Dim fixtureBook As Workbook
    Dim source As Range
    Dim before As String
    Dim result As CNxResult
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    Set fixtureBook = NxDataFixtureBook(source)
    before = NxDataRangeFingerprint(source)
    Set result = NxDataRunForTest("NX-DATA-UNIQUE-COUNT", source, 1, True, False, False)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Unique preview did not succeed"
    NxTestHarness.AssertTrue Not result.SourceChanged, "Unique preview reported a workbook change"
    NxTestHarness.AssertTrue before = NxDataRangeFingerprint(source), "Unique preview changed the source"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_DataWorkflow.TestUniquePreviewLeavesSourceUnchanged", failureDescription
End Sub

Private Sub TestDuplicateResultDoesNotChangeSource()
    Dim fixtureBook As Workbook
    Dim source As Range
    Dim before As String
    Dim result As CNxResult
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    Set fixtureBook = NxDataFixtureBook(source)
    before = NxDataRangeFingerprint(source)
    Set result = NxDataRunForTest("NX-DATA-DUPLICATE-LIST", source, 1, True, False, False)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Duplicate result did not succeed"
    NxTestHarness.AssertTrue result.SourceChanged, "New command-owned sheet was not reported"
    NxTestHarness.AssertTrue before = NxDataRangeFingerprint(source), "Duplicate result changed the source"
    NxTestHarness.AssertTrue fixtureBook.Worksheets(fixtureBook.Worksheets.Count).Name = "중복목록", "Duplicate sheet name is not stable"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_DataWorkflow.TestDuplicateResultDoesNotChangeSource", failureDescription
End Sub

Private Sub TestDuplicateResultUsesClosedHeadersAndNextName()
    Dim fixtureBook As Workbook
    Dim source As Range
    Dim result As CNxResult
    Dim outputSheet As Worksheet
    Dim expected As Variant
    Dim index As Long
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    Set fixtureBook = NxDataFixtureBook(source)
    fixtureBook.Worksheets.Add(After:=fixtureBook.Worksheets(fixtureBook.Worksheets.Count)).Name = "중복목록"
    source.Parent.Activate
    source.Select
    Set result = NxDataRunForTest("NX-DATA-DUPLICATE-LIST", source, 1, True, False, False)
    Set outputSheet = fixtureBook.Worksheets("중복목록 (2)")
    expected = Array("중복 그룹 ID", "정규화 비교값", "원본 값", "발생 건수", "원본 시트", "원본 행")
    For index = 0 To UBound(expected)
        NxTestHarness.AssertTrue CStr(outputSheet.Cells(1, index + 1).Value2) = CStr(expected(index)), "Duplicate header contract changed"
    Next index
    NxTestHarness.AssertTrue outputSheet.Range("A1:F5").SpecialCells(xlCellTypeConstants).Count = 30, "Duplicate result contains non-value cells"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_DataWorkflow.TestDuplicateResultUsesClosedHeadersAndNextName", failureDescription
End Sub

Private Sub TestObservedCellsEscalateFastToPlanned()
    Dim fixtureBook As Workbook
    Dim source As Range
    Dim request As New CNxDataRequest
    Dim commandObject As New CNxDataFeatureCommand
    Dim command As INxFeatureCommand
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim definition As New CNxFeatureDefinition
    Dim failureNumber As Long
    Dim failureDescription As String
    On Error GoTo Failed
    Set fixtureBook = Application.Workbooks.Add
    Set source = fixtureBook.Worksheets(1).Range("A1:A2")
    source.Value2 = "same"
    source.Select
    request.Configure "NX-DATA-UNIQUE-COUNT", source, False, 1, False, False, True
    commandObject.Configure request
    Set command = commandObject
    definition.Configure "NX-DATA-UNIQUE-COUNT", "NX-CAT-DATA", "data_unique_count", _
        "data_unique_count_description", NxCapabilityRead, NxRiskL0, NxRiskL1, "data_result", _
        10, NxOfficeCapabilityNone, "disable_feature", "disable_feature", NxExecutionFast, _
        False, False, 1
    definition.ConfigureLaunch "direct", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "direct"
    definition.Seal
    Set ticket = router.Prepare(definition, command)
    NxTestHarness.AssertTrue ticket.Decision.BaseGrade = NxExecutionFast, "Data base grade changed"
    NxTestHarness.AssertTrue ticket.Decision.ActualCellCount = 2, "Observed cell count did not reach the execution decision"
    NxTestHarness.AssertTrue ticket.Decision.ResolvedGrade = NxExecutionPlanned, "Observed cells did not raise the execution grade"
    GoTo Cleanup
Failed:
    failureNumber = Err.Number: failureDescription = Err.Description: Err.Clear
Cleanup:
    On Error Resume Next
    fixtureBook.Close False
    On Error GoTo 0
    If failureNumber <> 0 Then Err.Raise failureNumber, "T_DataWorkflow.TestObservedCellsEscalateFastToPlanned", failureDescription
End Sub

Private Function NxDataFixtureBook(ByRef source As Range) As Workbook
    Dim workbook As Workbook
    Dim sheet As Worksheet
    Set workbook = Application.Workbooks.Add
    Set sheet = workbook.Worksheets(1)
    sheet.Range("A1").Value2 = "이름"
    sheet.Range("A2").Value2 = "가"
    sheet.Range("A3").Value2 = "나"
    sheet.Range("A4").Value2 = "가"
    sheet.Range("A5").Value2 = "나"
    Set source = sheet.Range("A1:A5")
    source.Select
    Set NxDataFixtureBook = workbook
End Function

Private Function NxDataRangeFingerprint(ByVal source As Range) As String
    Dim cell As Range
    Dim value As Variant
    For Each cell In source.Cells
        value = cell.Value
        If IsError(value) Then
            NxDataRangeFingerprint = NxDataRangeFingerprint & "E|"
        Else
            NxDataRangeFingerprint = NxDataRangeFingerprint & CStr(VarType(value)) & ":" & CStr(value) & "|"
        End If
    Next cell
End Function
