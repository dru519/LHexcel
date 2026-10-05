Attribute VB_Name = "T_G005Grade"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestExecutionGradeWireValues"
    names.Add "TestCellBoundary"
    names.Add "TestAiRaisesGuarded"
    names.Add "TestFileRaisesPlanned"
    names.Add "TestManualRecoveryRaisesPlanned"
    names.Add "TestGradeNeverDowngrades"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestExecutionGradeWireValues": TestExecutionGradeWireValues
        Case "TestCellBoundary": TestCellBoundary
        Case "TestAiRaisesGuarded": TestAiRaisesGuarded
        Case "TestFileRaisesPlanned": TestFileRaisesPlanned
        Case "TestManualRecoveryRaisesPlanned": TestManualRecoveryRaisesPlanned
        Case "TestGradeNeverDowngrades": TestGradeNeverDowngrades
        Case Else: NxRaiseContractError "Unknown G005 grade test"
    End Select
End Sub

Private Sub TestExecutionGradeWireValues()
    NxTestHarness.AssertTrue NxExecutionFast = 0, "Fast wire value mismatch"
    NxTestHarness.AssertTrue NxExecutionGuarded = 1, "Guarded wire value mismatch"
    NxTestHarness.AssertTrue NxExecutionPlanned = 2, "Planned wire value mismatch"
    NxTestHarness.AssertTrue NxExecutionGradeWireName(NxExecutionGuarded) = "guarded", "Guarded wire name mismatch"
End Sub

Private Sub TestCellBoundary()
    NxTestHarness.AssertTrue ResolveGrade(NxExecutionFast, 9999, 10000, False, False, False, False, False) = NxExecutionFast, "9,999 must remain Fast"
    NxTestHarness.AssertTrue ResolveGrade(NxExecutionFast, 10000, 10000, False, False, False, False, False) = NxExecutionFast, "10,000 must remain Fast"
    NxTestHarness.AssertTrue ResolveGrade(NxExecutionFast, 10001, 10000, False, False, False, False, False) = NxExecutionPlanned, "10,001 must be Planned"
End Sub

Private Sub TestAiRaisesGuarded()
    NxTestHarness.AssertTrue ResolveGrade(NxExecutionFast, 1, 10000, True, False, False, False, False) = NxExecutionGuarded, "AI processing must be Guarded"
    NxTestHarness.AssertTrue ResolveGrade(NxExecutionFast, 1, 10000, False, True, False, False, False) = NxExecutionGuarded, "External transfer must be Guarded"
End Sub

Private Sub TestFileRaisesPlanned()
    NxTestHarness.AssertTrue ResolveGrade(NxExecutionFast, 0, 10000, False, False, True, False, False) = NxExecutionPlanned, "File mutation must be Planned"
End Sub

Private Sub TestManualRecoveryRaisesPlanned()
    NxTestHarness.AssertTrue ResolveGrade(NxExecutionFast, 0, 10000, False, False, False, True, True) = NxExecutionPlanned, "Manual external mutation must be Planned"
End Sub

Private Sub TestGradeNeverDowngrades()
    NxTestHarness.AssertTrue ResolveGrade(NxExecutionPlanned, 0, 10000, False, False, False, False, False) = NxExecutionPlanned, "Planned must not downgrade"
    NxTestHarness.AssertTrue RaiseGrade(NxExecutionGuarded, NxExecutionFast) = NxExecutionGuarded, "Guarded must not downgrade"
End Sub
