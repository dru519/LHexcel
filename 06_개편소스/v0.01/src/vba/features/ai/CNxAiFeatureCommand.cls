VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxAiFeatureCommand"
Option Explicit
Implements INxFeatureCommand

Private mFeatureId As String
Private mPrompt As String
Private mConfigured As Boolean

Public Sub Configure(ByVal featureId As String, ByVal prompt As String)
    If mConfigured Then NxRaiseContractError "AI command can be configured only once"
    If Not IsAiFeature(featureId) Then NxRaiseContractError "AI command feature is outside the closed contract"
    If Len(Trim$(prompt)) = 0 Then NxRaiseContractError "AI command requires a prompt"
    Dim policy As New CNxLimitsPolicy
    If Len(prompt) > policy.MaximumPromptCharsFor(featureId) Then NxRaiseContractError "AI prompt exceeds the exact feature limit"
    mFeatureId = featureId
    mPrompt = prompt
    mConfigured = True
End Sub

Public Property Get DeliveryMode() As String: EnsureConfigured: DeliveryMode = "CLIPBOARD_ONLY": End Property

Private Property Get INxFeatureCommand_FeatureId() As String
    EnsureConfigured
    INxFeatureCommand_FeatureId = mFeatureId
End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    EnsureConfigured
    If context Is Nothing Then NxRaiseContractError "AI command requires an execution context"
    Dim plan As CNxPlan
    Set plan = NxCreatePlan(mFeatureId, context.Fingerprint, context.SourceStateDigest)
    plan.Seal
    Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    EnsureConfigured
    If context Is Nothing Or approvedPlan Is Nothing Then NxRaiseContractError "AI execution requires context and approved plan"
    If approvedPlan.FeatureId <> mFeatureId Then NxRaiseContractError "AI approved plan identity mismatch"
    Set INxFeatureCommand_Execute = NxAiCopyPrompt(mPrompt, mFeatureId)
End Function

Private Function IsAiFeature(ByVal featureId As String) As Boolean
    Select Case featureId
        Case "NX-AI-SUMMARY", "NX-AI-CLEAN", "NX-AI-FORMULA", "NX-AI-WRITE", "NX-AI-IMAGE"
            IsAiFeature = True
    End Select
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "AI command is not configured"
End Sub
