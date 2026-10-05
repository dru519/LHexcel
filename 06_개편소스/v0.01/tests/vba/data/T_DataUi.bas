Attribute VB_Name = "T_DataUi"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestDataFormHasRequiredControls"
    names.Add "TestDataFormUsesGeneratedSources"
    names.Add "TestDataRegistryKeepsLocalGrades"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestDataFormHasRequiredControls": TestDataFormHasRequiredControls
        Case "TestDataFormUsesGeneratedSources": TestDataFormUsesGeneratedSources
        Case "TestDataRegistryKeepsLocalGrades": TestDataRegistryKeepsLocalGrades
        Case Else: NxRaiseContractError "Unknown Data UI test"
    End Select
End Sub

Private Sub TestDataFormHasRequiredControls()
    Dim controlName As Variant
    For Each controlName In Array("cboFeature", "txtRangeSummary", "cboHeader", "cboKeyColumn", _
        "txtOutputColumns", "cboTrim", "cboCase", "cboHidden", "txtPreview", _
        "cmdPreview", "cmdExecute", "cmdCancel")
        NxTestHarness.AssertTrue NxDataFormContractHas(CStr(controlName)), "Required Data form control is missing"
    Next controlName
End Sub

Private Sub TestDataFormUsesGeneratedSources()
    NxTestHarness.AssertTrue NxDataUsesGeneratedFormSources(), "Data form must use JSON plus VBA sources"
End Sub

Private Sub TestDataRegistryKeepsLocalGrades()
    Dim definition As CNxFeatureDefinition
    Set definition = NxDataFeatureDefinition("NX-DATA-UNIQUE-COUNT")
    NxTestHarness.AssertTrue definition.ExecutionGrade = NxExecutionFast, "Unique count must start Fast"
    NxTestHarness.AssertTrue Not definition.AiProcessing And Not definition.ExternalDataTransfer, "Local Data feature invoked privacy gate"
    Set definition = NxDataFeatureDefinition("NX-DATA-PASTE-VISIBLE-VALUES")
    NxTestHarness.AssertTrue definition.ExecutionGrade = NxExecutionPlanned, "Visible-value paste must be Planned"
End Sub
