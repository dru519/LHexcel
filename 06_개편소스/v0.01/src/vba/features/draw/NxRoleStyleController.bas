Attribute VB_Name = "NxRoleStyleController"
Option Explicit

Public Function NxRoleStyleRun(ByVal target As Range, ByVal roleCode As String, ByVal displayMode As String, ByVal axisCode As String, ByVal relativeIndexes As String) As CNxResult
    Dim request As New CNxRoleStyleRequest, commandObject As New CNxRoleStyleCommand, command As INxFeatureCommand
    request.Configure target, roleCode, displayMode, axisCode, relativeIndexes
    commandObject.Configure request: Set command = commandObject
    Set NxRoleStyleRun = NxRunDrawingFeatureCommand(command, NxDrawingTargetSummary(request.AffectedTarget))
End Function

Public Sub NxRoleStyleValidateTarget(ByVal target As Range, ByVal axisCode As String, ByVal relativeIndexes As String)
    NxRoleStyleValidateTargetCore target, axisCode, relativeIndexes
End Sub

Public Function NxRoleStyleBuildPreview(ByVal target As Range, ByVal roleCode As String, ByVal displayMode As String, ByVal axisCode As String, ByVal relativeIndexes As String) As CNxRoleStylePreview
    Dim request As New CNxRoleStyleRequest
    request.Configure target, roleCode, displayMode, axisCode, relativeIndexes
    Set NxRoleStyleBuildPreview = request.BuildPreview
End Function
