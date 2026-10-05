Attribute VB_Name = "T_TemplatePolicy"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestRetainedAreasMustStayInsideSource"
    names.Add "TestRetainedConstantsRoundTrip"
    names.Add "TestConstantsOutsideRetainedAreasAreBlank"
    names.Add "TestInternalFormulaR1C1IsAllowed"
    names.Add "TestVolatileFormulaRequiresConsent"
    names.Add "TestDynamicFormulaRequiresConsent"
    names.Add "TestEnvironmentFormulaRequiresConsent"
    names.Add "TestExternalAndOtherSheetFormulaIsBlocked"
    names.Add "TestExecutableOrExternalResourceIsBlocked"
    names.Add "TestUnsupportedDisplayResourceRequiresExclusionAck"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestRetainedAreasMustStayInsideSource": TestRetainedAreasMustStayInsideSource
        Case "TestRetainedConstantsRoundTrip": TestRetainedConstantsRoundTrip
        Case "TestConstantsOutsideRetainedAreasAreBlank": TestConstantsOutsideRetainedAreasAreBlank
        Case "TestInternalFormulaR1C1IsAllowed": TestInternalFormulaR1C1IsAllowed
        Case "TestVolatileFormulaRequiresConsent": TestVolatileFormulaRequiresConsent
        Case "TestDynamicFormulaRequiresConsent": TestDynamicFormulaRequiresConsent
        Case "TestEnvironmentFormulaRequiresConsent": TestEnvironmentFormulaRequiresConsent
        Case "TestExternalAndOtherSheetFormulaIsBlocked": TestExternalAndOtherSheetFormulaIsBlocked
        Case "TestExecutableOrExternalResourceIsBlocked": TestExecutableOrExternalResourceIsBlocked
        Case "TestUnsupportedDisplayResourceRequiresExclusionAck": TestUnsupportedDisplayResourceRequiresExclusionAck
        Case Else: NxRaiseContractError "Unknown Template policy test"
    End Select
End Sub

Private Sub TestRetainedAreasMustStayInsideSource()
    Dim book As Workbook, sheet As Worksheet, retained As Collection
    Dim failed As Boolean
    Set book = NewPolicyFixture(sheet)
    Set retained = OneRange(sheet.Range("C3"))
    On Error Resume Next
    Call NxTemplateAnalyze(sheet.Range("A1:B2"), retained, False, False)
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Retained area outside source was accepted"
    ClosePolicyFixture book
End Sub

Private Sub TestRetainedConstantsRoundTrip()
    Dim book As Workbook, sheet As Worksheet, retained As Collection
    Dim analysis As CNxTemplateAnalysis
    Set book = NewPolicyFixture(sheet)
    sheet.Range("A1").Value2 = "keep-1"
    sheet.Range("B2").Value2 = 42
    Set retained = TwoRanges(sheet.Range("A1"), sheet.Range("B2"))
    Set analysis = NxTemplateAnalyze(sheet.Range("A1:B2"), retained, False, False)
    NxTestHarness.AssertTrue analysis.StoredConstantCount = 2, "Retained constant count mismatch"
    NxTestHarness.AssertTrue analysis.ExcludedConstantCount = 0, "Retained constants were excluded"
    ClosePolicyFixture book
End Sub

Private Sub TestConstantsOutsideRetainedAreasAreBlank()
    Dim book As Workbook, sheet As Worksheet, retained As Collection
    Dim analysis As CNxTemplateAnalysis
    Set book = NewPolicyFixture(sheet)
    sheet.Range("A1:B2").Value2 = "constant"
    Set retained = OneRange(sheet.Range("A1"))
    Set analysis = NxTemplateAnalyze(sheet.Range("A1:B2"), retained, False, False)
    NxTestHarness.AssertTrue analysis.StoredConstantCount = 1, "Stored constant count mismatch"
    NxTestHarness.AssertTrue analysis.ExcludedConstantCount = 3, "Constants outside retained areas were not excluded"
    ClosePolicyFixture book
End Sub

Private Sub TestInternalFormulaR1C1IsAllowed()
    Dim book As Workbook, sheet As Worksheet
    Set book = NewPolicyFixture(sheet)
    sheet.Range("A1").Value2 = 1
    sheet.Range("B1").FormulaR1C1 = "=RC[-1]+1"
    NxTestHarness.AssertTrue NxTemplateFormulaDisposition(sheet.Range("B1"), sheet.Range("A1:B2")) = "allow", "Internal FormulaR1C1 was rejected"
    ClosePolicyFixture book
End Sub

Private Sub TestVolatileFormulaRequiresConsent()
    Dim book As Workbook, sheet As Worksheet
    Set book = NewPolicyFixture(sheet)
    sheet.Range("A1").Formula = "=NOW()"
    NxTestHarness.AssertTrue NxTemplateFormulaDisposition(sheet.Range("A1"), sheet.Range("A1:B2")) = "warn:VolatileFormula", "NOW warning mismatch"
    ClosePolicyFixture book
End Sub

Private Sub TestDynamicFormulaRequiresConsent()
    Dim book As Workbook, sheet As Worksheet
    Set book = NewPolicyFixture(sheet)
    sheet.Range("A1").Formula = "=INDIRECT(""B1"")"
    NxTestHarness.AssertTrue NxTemplateFormulaDisposition(sheet.Range("A1"), sheet.Range("A1:B2")) = "warn:DynamicFormula", "INDIRECT warning mismatch"
    ClosePolicyFixture book
End Sub

Private Sub TestEnvironmentFormulaRequiresConsent()
    Dim book As Workbook, sheet As Worksheet
    Set book = NewPolicyFixture(sheet)
    sheet.Range("A1").Formula = "=CELL(""filename"")"
    NxTestHarness.AssertTrue NxTemplateFormulaDisposition(sheet.Range("A1"), sheet.Range("A1:B2")) = "warn:EnvironmentFormula", "CELL warning mismatch"
    ClosePolicyFixture book
End Sub

Private Sub TestExternalAndOtherSheetFormulaIsBlocked()
    Dim book As Workbook, sheet As Worksheet
    Set book = NewPolicyFixture(sheet)
    book.Worksheets.Add.Name = "Other"
    sheet.Range("A1").Formula = "='Other'!A1"
    NxTestHarness.AssertTrue NxTemplateFormulaDisposition(sheet.Range("A1"), sheet.Range("A1:B2")) = "block:ExternalReference", "Other-sheet formula was accepted"
    ClosePolicyFixture book
End Sub

Private Sub TestExecutableOrExternalResourceIsBlocked()
    Dim book As Workbook, sheet As Worksheet, retained As Collection
    Dim failed As Boolean
    Set book = NewPolicyFixture(sheet)
    sheet.Shapes.AddFormControl xlButtonControl, sheet.Range("A1").Left, sheet.Range("A1").Top, 40, 20
    Set retained = OneRange(sheet.Range("A1"))
    On Error Resume Next
    Call NxTemplateAnalyze(sheet.Range("A1:B2"), retained, False, True)
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Executable form control was accepted"
    ClosePolicyFixture book
End Sub

Private Sub TestUnsupportedDisplayResourceRequiresExclusionAck()
    Dim book As Workbook, sheet As Worksheet, retained As Collection
    Dim analysis As CNxTemplateAnalysis, failed As Boolean
    Set book = NewPolicyFixture(sheet)
    sheet.Shapes.AddShape msoShapeRectangle, sheet.Range("A1").Left, sheet.Range("A1").Top, 40, 20
    Set retained = OneRange(sheet.Range("A1"))
    On Error Resume Next
    Call NxTemplateAnalyze(sheet.Range("A1:B2"), retained, False, False)
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Display resource exclusion was not confirmed"
    Set analysis = NxTemplateAnalyze(sheet.Range("A1:B2"), retained, False, True)
    NxTestHarness.AssertTrue analysis.HasExcludedDisplayResource, "Display resource exclusion was not recorded"
    ClosePolicyFixture book
End Sub

Private Function NewPolicyFixture(ByRef sheet As Worksheet) As Workbook
    Dim book As Workbook
    Set book = Application.Workbooks.Add
    Set sheet = book.Worksheets(1)
    sheet.Name = "Source"
    Set NewPolicyFixture = book
End Function

Private Function OneRange(ByVal first As Range) As Collection
    Dim result As New Collection
    result.Add first
    Set OneRange = result
End Function

Private Function TwoRanges(ByVal first As Range, ByVal second As Range) As Collection
    Dim result As New Collection
    result.Add first
    result.Add second
    Set TwoRanges = result
End Function

Private Sub ClosePolicyFixture(ByVal book As Workbook)
    On Error Resume Next
    book.Close False
    On Error GoTo 0
End Sub
