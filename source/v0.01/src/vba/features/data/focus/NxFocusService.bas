Attribute VB_Name = "NxFocusService"
Option Explicit

Public Function NxFocusIsEnabled() As Boolean
    NxFocusIsEnabled = NxFocusControllerIsEnabled()
End Function

Public Sub NxFocusSetEnabled(ByVal enabled As Boolean)
    NxFocusControllerSetEnabled enabled
End Sub
Public Sub NxFocusToggle()
    NxFocusControllerToggle
End Sub

Public Sub NxFocusEnable(ByVal selectedColor As Long)
    Dim current As CNxFocusSettingsModel
    Set current = NxFocusLoadSettings()
    NxFocusControllerCommitSettings NxFocusCreateSettings(current.Shape, current.Style, _
        NxFocusColorHexFromLong(selectedColor), current.IntensityPercent, _
        current.HighlightSelectedArea, True)
    NxFocusControllerSetEnabled True
End Sub

Public Sub NxFocusDisable()
    NxFocusControllerSetEnabled False
End Sub

Public Function NxFocusTryDisable() As Boolean
    On Error GoTo Failed
    NxFocusDisable
    NxFocusTryDisable = Not NxFocusIsEnabled()
    Exit Function
Failed:
    NxFocusTryDisable = False
End Function

Public Sub NxFocusSuspendForOperation()
    NxFocusControllerSuspend
End Sub

Public Sub NxFocusResumeAfterOperation()
    NxFocusControllerResume
End Sub

Public Sub NxFocusOpenSettings()
    Dim form As New FNxFocusSettings
    form.Show vbModal
End Sub

Public Sub NxFocusResetDefaults()
    Dim settings As CNxFocusSettingsModel
    Set settings = NxFocusDefaultSettings()
    NxFocusControllerCommitSettings settings
End Sub
