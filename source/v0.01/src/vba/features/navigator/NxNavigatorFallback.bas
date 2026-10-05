Attribute VB_Name = "NxNavigatorFallback"
Option Explicit

Private mNavigator As FNxNavigator

Public Sub NxNavigatorOpenVbaFallback()
    If mNavigator Is Nothing Then Set mNavigator = New FNxNavigator
    If mNavigator.Visible Then
        mNavigator.ActivateNavigator
    Else
        mNavigator.Show vbModeless
    End If
End Sub

Public Sub NxNavigatorRefreshAvailability()
    If mNavigator Is Nothing Then Exit Sub
    mNavigator.RefreshAvailability
End Sub

Public Sub NxNavigatorShutdown()
    On Error Resume Next
    If Not mNavigator Is Nothing Then Unload mNavigator
    Set mNavigator = Nothing
    On Error GoTo 0
End Sub

Public Sub NxNavigatorRelease(ByVal navigator As FNxNavigator)
    If mNavigator Is navigator Then Set mNavigator = Nothing
End Sub
