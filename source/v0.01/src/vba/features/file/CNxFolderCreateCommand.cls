VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFolderCreateCommand"
Option Explicit
Implements INxFeatureCommand
Private mBaseFolder As String, mItems As Collection, mPreview As CNxFolderCreatePreview, mConfigured As Boolean, mPlanBuilt As Boolean
Public Sub Configure(ByVal baseFolder As String, ByVal items As Collection, ByVal preview As CNxFolderCreatePreview)
    If mConfigured Or Len(baseFolder) = 0 Or items Is Nothing Or preview Is Nothing Then NxRaiseContractError "Folder command requires base, items and preview"
    If Not preview.IsValid Then NxRaiseContractError "Folder command requires a valid preview"
    Set mItems = items: Set mPreview = preview: mBaseFolder = baseFolder: mConfigured = True
End Sub
Private Property Get INxFeatureCommand_FeatureId() As String: INxFeatureCommand_FeatureId = "NX-FILE-FOLDER-CREATE": End Property
Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan: If Not mConfigured Or context Is Nothing Or Not mPreview.IsValid Then NxRaiseContractError "Folder preview snapshot invalidated"
    If mPlanBuilt Then NxRaiseContractError "Folder plan can be built only once"
    Set plan = NxCreatePlan("NX-FILE-FOLDER-CREATE", context.Fingerprint, context.SourceStateDigest)
    plan.AddSideEffect NxCreateDirectoryBatchEffect("NX-FILE-FOLDER-CREATE", mBaseFolder, mPreview.DirectoryPaths): plan.Seal: mPlanBuilt = True: Set INxFeatureCommand_BuildPlan = plan
End Function
Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    If Not mPlanBuilt Or context Is Nothing Or approvedPlan Is Nothing Or Not mPreview.IsValid Then NxRaiseContractError "Folder preview snapshot invalidated"
    If approvedPlan.FeatureId <> "NX-FILE-FOLDER-CREATE" Then NxRaiseContractError "Folder plan identity mismatch"
    Set INxFeatureCommand_Execute = NxCreateResult("NX-FILE-FOLDER-CREATE", NxSuccess, "complete", context.WorkbookIdentity, False, vbNullString, "folder_created")
End Function
