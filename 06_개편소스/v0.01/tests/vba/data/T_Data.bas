Attribute VB_Name = "T_Data"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    AppendNames names, T_DataNormalize.TestNames
    AppendNames names, T_DataWorkflow.TestNames
    AppendNames names, T_DataUi.TestNames
    AppendNames names, T_VisibleCellsClipboard.TestNames
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestTypedKeysRemainDistinct", "TestTextOptionsAreExplicit", _
             "TestUniqueGroupsUseFirstAppearance", "TestSourceRowsAreSnapshotted", _
             "TestUnsupportedValueTypeIsRejected"
            T_DataNormalize.RunCase name
        Case "TestUniquePreviewLeavesSourceUnchanged", "TestDuplicateResultDoesNotChangeSource", _
             "TestDuplicateResultUsesClosedHeadersAndNextName", "TestObservedCellsEscalateFastToPlanned"
            T_DataWorkflow.RunCase name
        Case "TestDataFormHasRequiredControls", "TestDataFormUsesGeneratedSources", _
             "TestDataRegistryKeepsLocalGrades", "TestDataColumnProjectionIsExact"
            T_DataUi.RunCase name
        Case "TestVisibleCopySerializesRowsAndColumns", "TestVisibleCopyWritesUnicodeClipboard", _
             "TestVisiblePasteWritesExactLiteralValues", "TestVisiblePasteRejectsInvalidTargets", _
             "TestVisiblePasteRollbackRestoresEarlierWrites"
            T_VisibleCellsClipboard.RunCase name
        Case Else: NxRaiseContractError "Unknown Data test"
    End Select
End Sub

Private Sub AppendNames(ByVal target As Collection, ByVal source As Collection)
    Dim item As Variant
    For Each item In source
        target.Add CStr(item)
    Next item
End Sub
