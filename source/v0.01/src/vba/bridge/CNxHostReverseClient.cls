VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxHostReverseClient"
Option Explicit

Public Function GetDocumentSnapshot() As Variant
    GetDocumentSnapshot = NxDocumentNavigatorDllSnapshot()
End Function

Public Function MoveDocument(ByVal token As String) As Boolean
    On Error GoTo Failed
    If Len(token) = 0 Or Len(token) > 10 Then Exit Function
    MoveDocument = NxDocumentNavigatorDllMove(token)
Failed:
End Function

' Explicit reverse callback surface. The host may request exactly one typed v0.01 identifier.
Public Function ExecuteRequest(ByVal featureId As String, ByVal commandId As String, ByVal payload As String) As Boolean
    Dim feature As CNxFeatureDefinition
    Dim command As CNxCommandDefinition
    Dim registry As CNxFeatureRegistry

    On Error GoTo Failed
    If Len(payload) > 512 Or InStr(payload, vbCr) > 0 Or InStr(payload, vbLf) > 0 Or InStr(payload, Chr$(0)) > 0 Then Exit Function
    If (Len(featureId) > 0) = (Len(commandId) > 0) Then Exit Function

    If Len(featureId) > 0 Then
        Set registry = NxCreateProductRegistry()
        Set feature = registry.FeatureById(featureId)
        If feature Is Nothing Then Exit Function
        NxRouteFeature feature.FeatureId
    Else
        Set command = NxCommandDefinition(commandId)
        If command Is Nothing Then Exit Function
        NxRouteCommand command.CommandId
    End If

    ExecuteRequest = True
    Exit Function
Failed:
    ExecuteRequest = False
End Function

Public Function GetRouteAvailability(ByVal routeKey As String) As String
    On Error GoTo Failed
    If Len(routeKey) = 0 Or Len(routeKey) > 160 Then Exit Function
    If InStr(routeKey, vbCr) > 0 Or InStr(routeKey, vbLf) > 0 Or InStr(routeKey, Chr$(0)) > 0 Then Exit Function
    If Not NxGeneratedNavigationRouteExists(routeKey) Then Exit Function
    GetRouteAvailability = NxRouteAvailabilitySummary(routeKey)
    Exit Function
Failed:
    Err.Clear
    GetRouteAvailability = NX_ROUTE_UNAVAILABLE & "|현재 Excel 상태를 확인할 수 없습니다."
End Function
