Attribute VB_Name = "NxHostBridge"
Option Explicit
Private Const NX_HOST_BRIDGE_VERSION As String = "nx-bridge-v1"
Private Const NX_HOST_PRODUCT_TITLE As String = "내엑셀 v0.01_r86"
Private Const NX_HOST_MAX_PAYLOAD As Long = 512
Private Const bridge_sha256 As String = "79027b72f740152d898b2b5a954f7f27bba9e85ed4f3317a49666cd9969f5401"
Private Const feature_sha256 As String = "fb4c29845c44d115061adae4f375d7ede8b6ce87717973120fe9ce80743686db"
Private Const command_sha256 As String = "8bf73de7e93494322cf58f2a9e5648ea983e3ce0cbba40b0821364152eb3e809"
Private mClient As CNxHostBridgeClient
Private mCompatibilityKnown As Boolean, mCompatible As Boolean, mFocusActive As Boolean
Private mHwpxConnectionStatus As String
Private mCompareResults As Object

Public Function NxHostCreateWorkbookCompareResults() As Object
    Dim service As Object
    On Error GoTo Unavailable
    Set service = NxHostVerifiedCreate("LH.NxHost.WorkbookCompareResultsService")
    If CBool(service.Connect(NX_HOST_PRODUCT_TITLE, "nx-compare-results-v1", bridge_sha256, feature_sha256, command_sha256, NxHostExcelBitness())) Then Set NxHostCreateWorkbookCompareResults = service
Unavailable:
End Function

Public Sub NxHostRetainWorkbookCompareResults(ByVal service As Object)
    Set mCompareResults = service
End Sub

Public Function NxHostCreateHwpxService() As Object
    Dim service As Object
    On Error GoTo Unavailable
    mHwpxConnectionStatus = "not-installed"
    Set service = NxHostVerifiedCreate("LH.NxHost.HwpxExportService")
    If Not CBool(service.Connect(NX_HOST_PRODUCT_TITLE, "nx-hwpx-v1", bridge_sha256, feature_sha256, command_sha256, NxHostExcelBitness())) Then
        mHwpxConnectionStatus = CStr(service.ConnectionStatus)
        Exit Function
    End If
    mHwpxConnectionStatus = "ready"
    Set NxHostCreateHwpxService = service
    Exit Function
Unavailable:
    mHwpxConnectionStatus = "unavailable"
End Function

Public Function NxHostHwpxConnectionStatus() As String
    NxHostHwpxConnectionStatus = mHwpxConnectionStatus
End Function

Public Function NxHostCreatePicturePreview() As Object
    Dim service As Object
    On Error GoTo Unavailable
    Set service = NxHostVerifiedCreate("LH.NxHost.PicturePreviewService")
    If CBool(service.Connect(NX_HOST_PRODUCT_TITLE, "nx-picture-preview-v1", bridge_sha256, feature_sha256, command_sha256, NxHostExcelBitness())) Then Set NxHostCreatePicturePreview = service
Unavailable:
End Function

Public Function NxHostCreateWorkbookCompare() As Object
    Dim service As Object
    On Error GoTo Unavailable
    Set service = NxHostVerifiedCreate("LH.NxHost.WorkbookCompareService")
    If CBool(service.Connect(NX_HOST_PRODUCT_TITLE, "nx-compare-v1", bridge_sha256, feature_sha256, command_sha256, NxHostExcelBitness())) Then Set NxHostCreateWorkbookCompare = service
Unavailable:
End Function

Public Function NxHostBridgeCompatible() As Boolean
    If mCompatibilityKnown Then NxHostBridgeCompatible = mCompatible: Exit Function
    mCompatibilityKnown = True
    On Error GoTo Unavailable
    Set mClient = New CNxHostBridgeClient
    If Not mClient.Connect() Then NxHostBridgeReset: Exit Function
    If mClient.Ping(NX_HOST_PRODUCT_TITLE, NX_HOST_BRIDGE_VERSION, bridge_sha256, feature_sha256, command_sha256, NxHostExcelBitness()) Then mCompatible = True Else NxHostBridgeReset
    NxHostBridgeCompatible = mCompatible
    Exit Function
Unavailable:
    NxHostBridgeReset
End Function

Public Function NxHostShowNavigator() As Boolean
    On Error GoTo Failed
    If Not NxHostBridgeCompatible() Then Exit Function
    NxHostShowNavigator = mClient.ExecuteRequest("NX-UTIL-NAVIGATOR", "SHOW_NAVIGATOR", "show=1")
    If Not NxHostShowNavigator Then NxHostBridgeReset
    Exit Function
Failed:
    NxHostBridgeReset
End Function
Public Function NxHostShowDocumentNavigator() As Boolean
    On Error GoTo Failed
    If Not NxHostBridgeCompatible() Then Exit Function
    NxHostShowDocumentNavigator = mClient.ExecuteRequest("NX-UTIL-DOCUMENT-NAVIGATOR", "SHOW_DOCUMENT_NAVIGATOR", "show=1")
Failed:
End Function
Public Sub NxHostHideDocumentNavigator()
    On Error Resume Next
    If Not mClient Is Nothing Then Call mClient.ExecuteRequest("NX-UTIL-DOCUMENT-NAVIGATOR", "HIDE_DOCUMENT_NAVIGATOR", "show=0")
End Sub
Public Function NxHostHideNavigator() As Boolean
    On Error GoTo Failed
    If Not NxHostBridgeCompatible() Then Exit Function
    NxHostHideNavigator = mClient.ExecuteRequest("NX-UTIL-NAVIGATOR", "HIDE_NAVIGATOR", "show=0")
    Exit Function
Failed:
    NxHostBridgeReset
End Function
Public Function NxHostFocusApply(ByVal settings As CNxFocusSettingsModel) As Boolean
    Dim operationId As String
    If settings Is Nothing Or Not settings.IsSealed Then Exit Function
    If Not NxHostBridgeCompatible() Then Exit Function
    If mFocusActive Then operationId = "FOCUS_UPDATE" Else operationId = "FOCUS_START"
    NxHostFocusApply = mClient.ExecuteRequest("NX-DATA-FOCUS-CELL", operationId, "version=1;" & NxHostFocusPayload(settings))
    If NxHostFocusApply Then mFocusActive = True Else NxHostBridgeReset
End Function
Public Function NxHostFocusStop() As Boolean
    Dim stopped As Boolean
    If Not mFocusActive Then NxHostFocusStop = True: Exit Function
    stopped = mClient.ExecuteRequest("NX-DATA-FOCUS-CELL", "FOCUS_STOP", "version=1;stop=1")
    NxHostFocusStop = stopped
    NxHostBridgeReset
End Function
Public Sub NxHostBridgeReset()
    On Error Resume Next
    If Not mCompareResults Is Nothing Then mCompareResults.Close
    Set mCompareResults = Nothing
    On Error GoTo 0
    If Not mClient Is Nothing Then mClient.Disconnect
    Set mClient = Nothing: mCompatibilityKnown = False: mCompatible = False: mFocusActive = False
End Sub
Private Function NxHostFocusPayload(ByVal settings As CNxFocusSettingsModel) As String
    NxHostFocusPayload = "shape=" & settings.Shape & ";style=" & settings.Style & ";color=" & settings.ColorHex & ";intensity=" & CStr(settings.IntensityPercent) & ";selected=" & CStr(Abs(CLng(settings.HighlightSelectedArea)))
End Function
Private Function NxHostExcelBitness() As String
    #If Win64 Then
        NxHostExcelBitness = "x64"
    #Else
        NxHostExcelBitness = "x86"
    #End If
End Function
