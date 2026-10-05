Attribute VB_Name = "NxHostBridgeStub"
Option Explicit

Public Function NxHostBridgeCompatible() As Boolean
    NxHostBridgeCompatible = False
End Function

Public Function NxHostFocusApply(ByVal settings As CNxFocusSettingsModel) As Boolean
    NxHostFocusApply = False
End Function

Public Function NxHostFocusStop() As Boolean
    NxHostFocusStop = True
End Function

Public Sub NxHostBridgeReset()
End Sub
