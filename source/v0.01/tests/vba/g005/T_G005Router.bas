Attribute VB_Name = "T_G005Router"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestFastRunsDirectly"
    names.Add "TestGuardedRequiresConfirmation"
    names.Add "TestPlannedReturnsApproval"
    names.Add "TestTicketCannotBeReused"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestFastRunsDirectly": TestFastRunsDirectly
        Case "TestGuardedRequiresConfirmation": TestGuardedRequiresConfirmation
        Case "TestPlannedReturnsApproval": TestPlannedReturnsApproval
        Case "TestTicketCannotBeReused": TestTicketCannotBeReused
        Case Else: NxRaiseContractError "Unknown G005 router test"
    End Select
End Sub

Private Sub TestFastRunsDirectly()
    Dim fakeObject As New CFakeFeatureCommand
    Dim command As INxFeatureCommand
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim result As CNxResult
    Set command = fakeObject
    Set ticket = router.Prepare(TestDefinition(NxExecutionFast), command)
    Set result = router.ExecuteFast(ticket)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess, "Fast execution did not succeed"
    NxTestHarness.AssertTrue fakeObject.ExecuteCount = 1 And ticket.IsConsumed, "Fast execution did not consume one ticket"
End Sub

Private Sub TestGuardedRequiresConfirmation()
    Dim fakeObject As New CFakeFeatureCommand
    Dim command As INxFeatureCommand
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim result As CNxResult
    Set command = fakeObject
    Set ticket = router.Prepare(TestDefinition(NxExecutionGuarded), command)
    AssertProtectedRejected router, ticket
    Set result = router.ExecuteGuarded(ticket, True)
    NxTestHarness.AssertTrue result.Outcome = NxSuccess And fakeObject.ExecuteCount = 1, "Confirmed Guarded execution failed"
End Sub

Private Sub TestPlannedReturnsApproval()
    Dim fakeObject As New CFakeFeatureCommand
    Dim command As INxFeatureCommand
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim approval As CNxApproval
    Set command = fakeObject
    Set ticket = router.Prepare(TestDefinition(NxExecutionPlanned), command)
    Set approval = router.TakePlannedApproval(ticket, True)
    NxTestHarness.AssertTrue Not approval Is Nothing, "Planned execution did not return its sealed approval"
    NxTestHarness.AssertTrue ticket.IsConsumed And fakeObject.ExecuteCount = 0, "Planned handoff must not execute immediately"
End Sub

Private Sub TestTicketCannotBeReused()
    Dim fakeObject As New CFakeFeatureCommand
    Dim command As INxFeatureCommand
    Dim router As New CNxExecutionRouter
    Dim ticket As CNxExecutionTicket
    Dim result As CNxResult
    Set command = fakeObject
    Set ticket = router.Prepare(TestDefinition(NxExecutionFast), command)
    Set result = router.ExecuteFast(ticket)
    On Error GoTo Expected
    Set result = router.ExecuteFast(ticket)
    Err.Raise vbObjectError + 880, "T_G005Router", "Consumed ticket was accepted"
Expected:
    NxTestHarness.AssertTrue InStr(1, Err.Description, "only once", vbTextCompare) > 0, "Wrong ticket reuse rejection"
    Err.Clear
End Sub

Private Sub AssertProtectedRejected(ByVal router As CNxExecutionRouter, ByVal ticket As CNxExecutionTicket)
    Dim result As CNxResult
    On Error GoTo Expected
    Set result = router.ExecuteGuarded(ticket, False)
    Err.Raise vbObjectError + 881, "T_G005Router", "Unconfirmed protected execution was accepted"
Expected:
    NxTestHarness.AssertTrue InStr(1, Err.Description, "explicit confirmation", vbTextCompare) > 0, "Wrong confirmation rejection"
    Err.Clear
End Sub

Private Function TestDefinition(ByVal grade As NxExecutionGrade) As CNxFeatureDefinition
    Dim definition As New CNxFeatureDefinition
    definition.Configure "NX-TEST-CORE", "NX-CAT-AI", "test", "test_description", _
        NxCapabilityRead, NxRiskL0, NxRiskL4, "test_result", 1, NxOfficeCapabilityNone, _
        "none", "disable_feature", grade, False, False, 10000
    definition.ConfigureLaunch "direct", vbNullString, vbNullString, vbNullString, vbNullString, vbNullString, "direct"
    definition.Seal
    Set TestDefinition = definition
End Function
