Attribute VB_Name = "NxFocusPalette"
Option Explicit

Public Sub NxFocusOpenPalette(ByVal context As CNxFeatureDialogContext)
    If context Is Nothing Then NxRaiseContractError "Focus palette context is required"
    NxFocusOpenSettings
End Sub

Public Function NxFocusApplySettings(ByVal colors As String) As Boolean
    colors = NxFocusNormalizeColors(colors)
    If Len(colors) = 0 Then Exit Function
    NxFocusSaveColors colors
    NxFocusControllerApplySettings NxFocusLoadSettings()
    NxFocusApplySettings = True
End Function
