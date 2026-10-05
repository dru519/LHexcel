VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxRoleStyleRequest"
Option Explicit

Private mTarget As Range, mAffectedTarget As Range, mRoleCode As String, mDisplayMode As String, mAxisCode As String, mRelativeIndexes As String, mTargetDigest As String

Public Sub Configure(ByVal target As Range, ByVal roleCode As String, ByVal displayMode As String, ByVal axisCode As String, ByVal relativeIndexes As String)
    NxRoleStyleValidateTargetCore target, axisCode, relativeIndexes
    Set mTarget = target: mRoleCode = NxRoleStyleNormalizeRole(roleCode): mDisplayMode = NxRoleStyleNormalizeMode(displayMode): mAxisCode = NxRoleStyleNormalizeAxis(axisCode): mRelativeIndexes = relativeIndexes
    Set mAffectedTarget = NxRoleStyleResolveAffectedRange(target, mAxisCode, relativeIndexes)
    mTargetDigest = NxRoleStyleRequestDigest(target)
End Sub

Public Sub RequireContext(ByVal context As CNxExecutionContext)
    If mTarget Is Nothing Then NxRaiseContractError "role-style request is not configured"
    If context Is Nothing Then NxRaiseContractError "role-style execution context is required"
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "role-style selection drift detected"
    If Application.Selection.Address(External:=True) <> mTarget.Address(External:=True) Then NxRaiseContractError "role-style selection drift detected"
    NxRoleStyleValidateTargetCore mTarget, mAxisCode, mRelativeIndexes
    If NxRoleStyleRequestDigest(mTarget) <> mTargetDigest Then NxRaiseContractError "role-style target changed after preview"
End Sub

Public Function BuildPreview() As CNxRoleStylePreview
    Dim preview As New CNxRoleStylePreview
    preview.Capture mTarget, mRoleCode, mDisplayMode, mAxisCode, mRelativeIndexes, mTargetDigest: preview.Seal
    Set BuildPreview = preview
End Function

Public Property Get FeatureId() As String: FeatureId = "NX-DRAW-ROLE-STYLE": End Property
Public Property Get Target() As Range: Set Target = mTarget: End Property
Public Property Get AffectedTarget() As Range: Set AffectedTarget = mAffectedTarget: End Property
Public Property Get RoleCode() As String: RoleCode = mRoleCode: End Property
Public Property Get DisplayMode() As String: DisplayMode = mDisplayMode: End Property
Public Property Get AxisCode() As String: AxisCode = mAxisCode: End Property
Public Property Get RelativeIndexes() As String: RelativeIndexes = mRelativeIndexes: End Property
Public Property Get ActualCellCount() As Long: ActualCellCount = CLng(mAffectedTarget.CountLarge): End Property

Private Function NxRoleStyleRequestDigest(ByVal target As Range) As String
    Dim cell As Range, result As String, mergeAddress As String
    For Each cell In target.Cells
        mergeAddress = vbNullString
        If cell.MergeCells Then mergeAddress = cell.MergeArea.Address(False, False)
        result = result & cell.Address(False, False) & _
            "|formula=" & NxRoleStyleDigestScalar(cell.Formula) & _
            "|text=" & NxRoleStyleDigestScalar(cell.Text) & _
            "|number=" & NxRoleStyleDigestScalar(cell.NumberFormat) & _
            "|merge=" & mergeAddress & ";"
    Next cell
    ' Managed focus rules are excluded so caret movement cannot drift the digest.
    NxRoleStyleRequestDigest = target.Address(External:=True) & "|" & CStr(target.CountLarge) & "|" & CStr(target.FormatConditions.Count - NxFocusManagedConditionCount(target)) & "|" & result
End Function

Private Function NxRoleStyleDigestScalar(ByVal value As Variant) As String
    If IsError(value) Then
        NxRoleStyleDigestScalar = "#ERROR"
    ElseIf IsNull(value) Then
        NxRoleStyleDigestScalar = "#NULL"
    ElseIf IsEmpty(value) Then
        NxRoleStyleDigestScalar = "#EMPTY"
    Else
        NxRoleStyleDigestScalar = CStr(value)
    End If
End Function
