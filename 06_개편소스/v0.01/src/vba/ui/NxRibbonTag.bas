Attribute VB_Name = "NxRibbonTag"
Option Explicit

Public Function NxParseRibbonTag(ByVal rawTag As String) As CNxRibbonTarget
    Dim parts As Variant
    Dim target As New CNxRibbonTarget
    Dim registry As CNxFeatureRegistry
    Dim feature As CNxFeatureDefinition
    Dim category As CNxCategoryDefinition
    Dim command As CNxCommandDefinition
    Dim identifier As String

    If Len(rawTag) = 0 Then NxRaiseContractError "Ribbon tag is required"
    parts = Split(rawTag, "|")
    If UBound(parts) <> 2 Then NxRaiseContractError "Ribbon tag must contain exactly two pipes"
    If parts(0) <> "nx1" Then NxRaiseContractError "Unknown ribbon tag version"
    If Len(parts(1)) = 0 Or Len(parts(2)) = 0 Then NxRaiseContractError "Ribbon tag fields cannot be empty"

    identifier = CStr(parts(2))
    If Not NxRibbonIdentifierIsSafe(identifier) Then NxRaiseContractError "Ribbon tag identifier is invalid"
    Select Case CStr(parts(1))
        Case "feature"
            Set registry = NxCreateProductRegistry()
            Set feature = registry.FeatureById(identifier)
        Case "category"
            Set registry = NxCreateProductRegistry()
            Set category = registry.CategoryById(identifier)
        Case "entry"
            If Not NxRibbonEntryIsAllowed(identifier) Then NxRaiseContractError "Unknown ribbon entry"
        Case "command"
            Set command = NxCommandDefinition(identifier)
        Case Else
            NxRaiseContractError "Unknown ribbon tag kind"
    End Select

    target.Configure CStr(parts(1)), identifier
    target.Seal
    Set NxParseRibbonTag = target
End Function

Public Function NxRibbonParseTag(ByVal rawTag As String) As CNxRibbonTarget
    Set NxRibbonParseTag = NxParseRibbonTag(rawTag)
End Function

Public Function NxRibbonIdentifierIsSafe(ByVal identifier As String) As Boolean
    Dim index As Long
    Dim code As Long
    If Len(identifier) = 0 Then Exit Function
    For index = 1 To Len(identifier)
        code = AscW(Mid$(identifier, index, 1))
        If Not ((code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or code = 45) Then Exit Function
    Next index
    NxRibbonIdentifierIsSafe = True
End Function

Public Function NxRibbonEntryIsAllowed(ByVal identifier As String) As Boolean
    Select Case identifier
        Case "NX-ENTRY-START", "NX-ENTRY-AI", "NX-ENTRY-TEMPLATE", "NX-ENTRY-ALL", _
             "NX-ENTRY-FAVORITES", "NX-ENTRY-MANAGEMENT", "NX-ENTRY-FOCUS-SETTINGS", _
             "NX-ENTRY-FOCUS-RESET", "NX-MGMT-INSTALL", _
             "NX-MGMT-REMOVE", "NX-MGMT-SAVED-FOLDER", "NX-MGMT-INSTALL-FOLDER", "NX-MGMT-QAT", _
             "NX-MGMT-EXCEL-ADDINS", "NX-MGMT-FILE-INFO", "NX-MGMT-SHORTCUTS", _
             "NX-ENTRY-HANGUL-SETTINGS"
            NxRibbonEntryIsAllowed = True
    End Select
End Function
