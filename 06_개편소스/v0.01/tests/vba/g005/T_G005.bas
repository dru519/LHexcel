Attribute VB_Name = "T_G005"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    Dim part As Collection
    Set part = T_G005Grade.TestNames
    AppendNames names, part
    Set part = T_G005Router.TestNames
    AppendNames names, part
    If names.Count <> 10 Then NxRaiseContractError "G005 suite inventory must be exactly 10"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestExecutionGradeWireValues", "TestCellBoundary", "TestAiRaisesGuarded", _
            "TestFileRaisesPlanned", "TestManualRecoveryRaisesPlanned", "TestGradeNeverDowngrades"
            T_G005Grade.RunCase name
        Case "TestFastRunsDirectly", "TestGuardedRequiresConfirmation", _
            "TestPlannedReturnsApproval", "TestTicketCannotBeReused"
            T_G005Router.RunCase name
        Case Else
            NxRaiseContractError "Unknown G005 test"
    End Select
End Sub

Private Sub AppendNames(ByVal target As Collection, ByVal source As Collection)
    Dim item As Variant
    For Each item In source
        target.Add CStr(item)
    Next item
End Sub
