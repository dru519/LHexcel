Attribute VB_Name = "T_R76Policy"
Option Explicit

Public Sub SetupMetadata(ByVal blockFrom As String, ByVal validUntil As String, ByVal profile As String, ByVal tampered As Boolean)
    Dim props As Object
    Set props = ThisWorkbook.CustomDocumentProperties
    On Error Resume Next
    props("DistributionBlockFrom").Delete
    props("DistributionValidUntil").Delete
    props("DistributionProfile").Delete
    On Error GoTo 0
    If tampered Then blockFrom = "2099-12-31"
    props.Add "DistributionBlockFrom", False, 4, blockFrom
    props.Add "DistributionValidUntil", False, 4, validUntil
    props.Add "DistributionProfile", False, 4, profile
End Sub

Public Function Names() As String
    Names = "active|expires_today|expired|tampered_date|tampered_signature|tampered_metadata"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim expected As String, allowed As Boolean, enabled As Variant
    Dim control As New T_R76Control, errorNumber As Long, book As Workbook
    On Error GoTo Failed
    Select Case name
        Case "active": expected = "active"
        Case "expired", "expires_today": expected = "expired"
        Case Else: expected = "invalid"
    End Select
    allowed = (expected = "active")
    If NxDistributionState() <> expected Then Err.Raise 5, , "policy state:" & NxDistributionState() & "|" & NxDistributionProbeDetail()
    If NxDistributionCanExecute() <> allowed Then Err.Raise 5, , "execution permission"
    control.Tag = "nx1|feature|NX-UTIL-CALCULATOR"
    NxRibbonGetEnabled control, enabled
    If CBool(enabled) <> allowed Then Err.Raise 5, , "calculator ribbon exemption"
    NxRibbonGetDistributionEnabled control, enabled
    If CBool(enabled) <> allowed Then Err.Raise 5, , "container ribbon gate"
    If Not NxDistributionEntryAllowed("entry", "NX-ENTRY-MANAGEMENT") Then Err.Raise 5, , "management inaccessible"
    If NxDistributionEntryAllowed("entry", "NX-ENTRY-ALL") <> allowed Then Err.Raise 5, , "all functions accessible after expiry"
    If InStr(NxDistributionMessage(), "AI") = 0 Then Err.Raise 5, , "modification notice missing"
    If Not allowed Then
        Set book = Application.Workbooks.Add(xlWBATWorksheet)
        book.Worksheets(1).Range("A1").Value2 = "unchanged"
        book.Saved = True
        On Error Resume Next
        NxRouteFeature "NX-UTIL-CALCULATOR"
        errorNumber = Err.Number: Err.Clear
        On Error GoTo Failed
        If errorNumber = 0 Then Err.Raise 5, , "feature route executed"
        On Error Resume Next
        NxRunDirectFeature Nothing
        errorNumber = Err.Number: Err.Clear
        On Error GoTo Failed
        If errorNumber = 0 Then Err.Raise 5, , "direct surface executed"
        If book.Worksheets(1).Range("A1").Value2 <> "unchanged" Or Not book.Saved Then Err.Raise 5, , "source state changed"
        If InStr(NxRibbonManagementMenuXml(), "NX-MGMT-SHORTCUTS") > 0 Then Err.Raise 5, , "expired management exposes shortcuts"
        book.Close False: Set book = Nothing
    End If
    NxDistributionStart
    NxDistributionStop
    If NxDistributionProbeTimer() Then Err.Raise 5, , "residual timer"
    RunCase = "PASS|" & name
    Exit Function
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
    On Error Resume Next
    NxDistributionStop
    If Not book Is Nothing Then book.Close False
End Function
