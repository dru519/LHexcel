VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxHostBridgeClient"
Option Explicit

Private mService As Object
Private mReverseClient As Object
Public Function Connect() As Boolean
    On Error GoTo Failed
    Set mService = NxHostVerifiedCreate("LH.NxHost.BridgeService")
    Set mReverseClient = New CNxHostReverseClient
    Connect = Not (mService Is Nothing)
    Exit Function
Failed:
    Set mService = Nothing
End Function
Public Function Ping(ByVal productTitle As String, ByVal interfaceVersion As String, ByVal bridgeHash As String, ByVal featureHash As String, ByVal commandHash As String, ByVal excelBitness As String) As Boolean
    On Error GoTo Failed
    If mService Is Nothing Then Exit Function
    Ping = CBool(mService.Ping(productTitle, interfaceVersion, bridgeHash, featureHash, commandHash, excelBitness, mReverseClient))
    Exit Function
Failed:
    Set mService = Nothing
End Function
Public Function ExecuteRequest(ByVal featureId As String, ByVal commandId As String, ByVal payload As String) As Boolean
    On Error GoTo Failed
    If featureId <> "NX-DATA-FOCUS-CELL" And featureId <> "NX-UTIL-NAVIGATOR" And featureId <> "NX-UTIL-DOCUMENT-NAVIGATOR" Then Exit Function
    If Len(payload) > 512 Or InStr(payload, vbCr) > 0 Or InStr(payload, vbLf) > 0 Or InStr(payload, Chr$(0)) > 0 Then Exit Function
    If mService Is Nothing Then Exit Function
    ExecuteRequest = CBool(mService.ExecuteRequest(featureId, commandId, payload))
    Exit Function
Failed:
    Set mService = Nothing
End Function
Public Sub Disconnect()
    On Error Resume Next
    If Not mService Is Nothing Then Call mService.ExecuteRequest("NX-HOST-BRIDGE", "DISCONNECT", "disconnect=1")
    On Error GoTo 0
    Set mReverseClient = Nothing: Set mService = Nothing
End Sub
