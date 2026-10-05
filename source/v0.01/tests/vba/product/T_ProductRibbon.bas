Attribute VB_Name = "T_ProductRibbon"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestValidTags"
    names.Add "TestMalformedTagsFailClosed"
    names.Add "TestUnknownIdsFailClosed"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal testName As String)
    Select Case testName
        Case "TestValidTags": TestValidTags
        Case "TestMalformedTagsFailClosed": TestMalformedTagsFailClosed
        Case "TestUnknownIdsFailClosed": TestUnknownIdsFailClosed
        Case Else: Err.Raise vbObjectError + 710, "T_ProductRibbon", "Unknown ProductRibbon test"
    End Select
End Sub

Private Sub TestValidTags()
    Dim target As CNxRibbonTarget
    Set target = NxParseRibbonTag("nx1|entry|NX-ENTRY-AI")
    NxTestHarness.AssertTrue target.IsSealed And target.Kind = "entry" And target.Identifier = "NX-ENTRY-AI", "Valid entry tag rejected"
    Set target = NxParseRibbonTag("nx1|category|NX-CAT-DATA")
    NxTestHarness.AssertTrue target.Kind = "category", "Valid category tag rejected"
End Sub

Private Sub TestMalformedTagsFailClosed()
    AssertTagRejected vbNullString
    AssertTagRejected "nx1|entry|NX-ENTRY-AI|extra"
    AssertTagRejected "nx1|entry|"
    AssertTagRejected "nx2|entry|NX-ENTRY-AI"
    AssertTagRejected "nx1|entry|Application.Run"
End Sub

Private Sub TestUnknownIdsFailClosed()
    AssertTagRejected "nx1|feature|NX-NOT-REGISTERED"
    AssertTagRejected "nx1|category|NX-CAT-NOT-REGISTERED"
    AssertTagRejected "nx1|entry|NX-ENTRY-NOT-REGISTERED"
End Sub

Private Sub AssertTagRejected(ByVal rawTag As String)
    Dim target As CNxRibbonTarget
    On Error Resume Next
    Set target = NxParseRibbonTag(rawTag)
    NxTestHarness.AssertTrue Err.Number <> 0 And target Is Nothing, "Malformed ribbon tag was accepted: " & rawTag
    Err.Clear
    On Error GoTo 0
End Sub
