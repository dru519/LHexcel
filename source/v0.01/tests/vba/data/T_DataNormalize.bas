Attribute VB_Name = "T_DataNormalize"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestTypedKeysRemainDistinct"
    names.Add "TestTextOptionsAreExplicit"
    names.Add "TestUniqueGroupsUseFirstAppearance"
    names.Add "TestSourceRowsAreSnapshotted"
    names.Add "TestUnsupportedValueTypeIsRejected"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestTypedKeysRemainDistinct": TestTypedKeysRemainDistinct
        Case "TestTextOptionsAreExplicit": TestTextOptionsAreExplicit
        Case "TestUniqueGroupsUseFirstAppearance": TestUniqueGroupsUseFirstAppearance
        Case "TestSourceRowsAreSnapshotted": TestSourceRowsAreSnapshotted
        Case "TestUnsupportedValueTypeIsRejected": TestUnsupportedValueTypeIsRejected
        Case Else: NxRaiseContractError "Unknown Data normalize test"
    End Select
End Sub

Private Sub TestTypedKeysRemainDistinct()
    NxTestHarness.AssertTrue NxDataKey(1, False, False) <> NxDataKey("1", False, False), "Number and text keys collided"
    NxTestHarness.AssertTrue NxDataKey(DateSerial(2026, 7, 10), False, False) <> NxDataKey(CDbl(DateSerial(2026, 7, 10)), False, False), "Date and serial keys collided"
    NxTestHarness.AssertTrue NxDataKey(Empty, False, False) <> NxDataKey(CVErr(xlErrNA), False, False), "Empty and error keys collided"
    NxTestHarness.AssertTrue NxDataKey(True, False, False) <> NxDataKey(-1, False, False), "Boolean and number keys collided"
End Sub

Private Sub TestTextOptionsAreExplicit()
    NxTestHarness.AssertTrue NxDataKey(" A ", False, True) = "S: A ", "Text was trimmed without consent"
    NxTestHarness.AssertTrue NxDataKey(" A ", True, False) = "S:a", "Text trim and case options were not applied"
End Sub

Private Sub TestUniqueGroupsUseFirstAppearance()
    NxTestHarness.AssertTrue NxDataGroupFingerprint(Array("b", "a", "b", "c")) = "S:b|2|1;S:a|1|2;S:c|1|4", "Group order is unstable"
End Sub

Private Sub TestSourceRowsAreSnapshotted()
    Dim groups As Collection
    Dim group As CNxDataGroup
    Dim rows As Collection
    Set groups = NxDataBuildGroups(Array("a", "a"), False, False, 10)
    Set group = groups(1)
    Set rows = group.SourceRows
    rows.Add 99
    NxTestHarness.AssertTrue group.SourceRows.Count = 2, "Caller mutated sealed source rows"
    NxTestHarness.AssertTrue CLng(group.SourceRows(1)) = 10 And CLng(group.SourceRows(2)) = 11, "Source row identities changed"
End Sub

Private Sub TestUnsupportedValueTypeIsRejected()
    Dim rejected As Boolean
    Dim unsupported As New Collection
    On Error GoTo Expected
    Call NxDataKey(unsupported, False, False)
    GoTo Done
Expected:
    rejected = (Err.Number = NX_CONTRACT_ERROR)
    Err.Clear
Done:
    NxTestHarness.AssertTrue rejected, "Unsupported object value was accepted"
End Sub
