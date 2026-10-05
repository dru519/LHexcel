VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxTemplateFeatureCommand"
Option Explicit
Implements INxFeatureCommand

Private mFeatureId As String
Private mRequest As CNxTemplateRequest
Private mTemplateId As String
Private mDisplayName As String
Private mDescription As String
Private mTargetBook As Workbook
Private mRiskConfirmed As Boolean
Private mActualCellCount As Long
Private mHasRiskyFormula As Boolean
Private mConfigured As Boolean
Private mOpenCopy As Boolean
Private mBuildCatalog As Boolean
Private mCategory As Variant
Private mQuarantineId As String
Private mMoveDirection As Long

Public Sub ConfigureRegistration(ByVal request As CNxTemplateRequest)
    If request Is Nothing Then NxRaiseContractError "Template command request is required"
    ConfigureBase request.FeatureId
    Set mRequest = request
    mActualCellCount = request.ActualCellCount
    mHasRiskyFormula = request.HasRiskyFormula
End Sub

Public Sub ConfigureCopy(ByVal templateId As String, ByVal targetBook As Workbook, ByVal riskConfirmed As Boolean, ByVal actualCellCount As Long, ByVal hasRiskyFormula As Boolean)
    ConfigureLoad templateId, targetBook, riskConfirmed, actualCellCount, hasRiskyFormula
    mOpenCopy = True
End Sub

Public Sub ConfigureCatalog(ByVal targetBook As Workbook, ByVal recordCount As Long)
    ConfigureBase "NX-TPL-LOAD"
    If targetBook Is Nothing Or recordCount < 0 Then NxRaiseContractError "Template catalog target is invalid"
    Set mTargetBook = targetBook
    mActualCellCount = (recordCount + 1) * 6
    mBuildCatalog = True
End Sub

Public Sub ConfigureList()
    ConfigureBase "NX-TPL-LIST"
End Sub

Public Sub ConfigureLoad(ByVal templateId As String, ByVal targetBook As Workbook, ByVal riskConfirmed As Boolean, ByVal actualCellCount As Long, ByVal hasRiskyFormula As Boolean)
    ConfigureBase "NX-TPL-LOAD"
    If Not NxTemplateIdIsValid(templateId) Or targetBook Is Nothing Or actualCellCount <= 0 Then NxRaiseContractError "Template load command is invalid"
    mTemplateId = templateId: Set mTargetBook = targetBook: mRiskConfirmed = riskConfirmed
    mActualCellCount = actualCellCount: mHasRiskyFormula = hasRiskyFormula
End Sub

Public Sub ConfigureRename(ByVal templateId As String, ByVal displayName As String, ByVal description As String, Optional ByVal category As Variant)
    ConfigureBase "NX-TPL-RENAME"
    If Not NxTemplateIdIsValid(templateId) Or Len(Trim$(displayName)) = 0 Then NxRaiseContractError "Template rename command is invalid"
    mTemplateId = templateId: mDisplayName = displayName: mDescription = description
    If Not IsMissing(category) Then mCategory = category
End Sub

Public Sub ConfigureRestore(ByVal templateId As String, ByVal quarantineId As String)
    ConfigureDelete templateId
    If Not NxTemplateIdIsValid(quarantineId) Then NxRaiseContractError "Template quarantine ID is invalid"
    mQuarantineId = quarantineId
End Sub

Public Sub ConfigureMove(ByVal templateId As String, ByVal direction As Long)
    ConfigureBase "NX-TPL-RENAME"
    If Not NxTemplateIdIsValid(templateId) Or (direction <> -1 And direction <> 1) Then NxRaiseContractError "Template move is invalid"
    mTemplateId = templateId: mMoveDirection = direction
End Sub

Public Sub ConfigureDelete(ByVal templateId As String)
    ConfigureBase "NX-TPL-DELETE"
    If Not NxTemplateIdIsValid(templateId) Then NxRaiseContractError "Template delete command is invalid"
    mTemplateId = templateId
End Sub

Private Property Get INxFeatureCommand_FeatureId() As String
    EnsureConfigured
    INxFeatureCommand_FeatureId = mFeatureId
End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan
    EnsureConfigured
    Set plan = NxCreatePlan(mFeatureId, context.Fingerprint, context.SourceStateDigest)
    Select Case mFeatureId
        Case "NX-TPL-REGISTER-SHEET", "NX-TPL-RENAME", "NX-TPL-DELETE"
            plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxBinaryPackageWrite, context.WorkbookIdentity, "remove only command-owned template assets")
        Case "NX-TPL-LOAD"
            If Not mBuildCatalog Then plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxBinaryPackageWrite, NxTemplateMetadataPath(mTemplateId), "restore previous template metadata on write failure")
            If mOpenCopy Then
                plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxWorksheetCreate, context.WorkbookIdentity, "close only the newly created unsaved template-copy workbook", mActualCellCount)
            Else
                plan.AddSideEffect NxCreateCommandOwnedEffect(mFeatureId, NxWorksheetCreate, context.WorkbookIdentity, "delete only the newly created template or catalog sheets", mActualCellCount)
            End If
        Case "NX-TPL-LIST"
            ' Local read-only command has no side effect.
        Case Else
            NxRaiseContractError "Template feature is outside the closed contract"
    End Select
    If mFeatureId = "NX-TPL-DELETE" Then plan.DeclareCapability NxCapabilityTemplateQuarantine
    If mHasRiskyFormula Then plan.DeclareCapability NxCapabilityRiskyFormula
    plan.Seal
    Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    Dim serviceResult As CNxResult, records As Collection
    EnsureConfigured
    If context Is Nothing Or approvedPlan Is Nothing Or approvedPlan.FeatureId <> mFeatureId Then NxRaiseContractError "Template approved plan identity mismatch"
    Select Case mFeatureId
        Case "NX-TPL-REGISTER-SHEET": Set serviceResult = NxTemplateRegisterSheet(mRequest)
        Case "NX-TPL-LIST": Set records = NxTemplateList(): Set serviceResult = NxCreateResult(mFeatureId, NxSuccess, "complete", context.WorkbookIdentity, False, vbNullString, "result_success")
        Case "NX-TPL-LOAD"
            If mBuildCatalog Then
                Set serviceResult = NxTemplateBuildCatalog(mTargetBook)
            ElseIf mOpenCopy Then
                Set serviceResult = NxTemplateOpenCopy(mTemplateId, mRiskConfirmed)
            Else
                Set serviceResult = NxTemplateLoad(mTemplateId, mTargetBook, mRiskConfirmed)
            End If
        Case "NX-TPL-RENAME"
            If mMoveDirection <> 0 Then
                Set serviceResult = NxTemplateMove(mTemplateId, mMoveDirection)
            Else
                Set serviceResult = NxTemplateRename(mTemplateId, mDisplayName, mDescription, mCategory)
            End If
        Case "NX-TPL-DELETE"
            If Len(mQuarantineId) > 0 Then
                Set serviceResult = NxTemplateRestore(mTemplateId, mQuarantineId)
            Else
                Set serviceResult = NxTemplateDeleteToQuarantine(mTemplateId)
            End If
    End Select
    If mFeatureId = "NX-TPL-LOAD" And Not mBuildCatalog Then
        Set INxFeatureCommand_Execute = RebindResultTarget(serviceResult, NxTemplateMetadataPath(mTemplateId))
    Else
        Set INxFeatureCommand_Execute = RebindResultTarget(serviceResult, context.WorkbookIdentity)
    End If
End Function

Private Function RebindResultTarget(ByVal source As CNxResult, ByVal target As String) As CNxResult
    Set RebindResultTarget = NxCreateResult(source.FeatureId, source.Outcome, source.Stage, target, source.SourceChanged, source.Recovery, source.MessageKey)
End Function

Private Sub ConfigureBase(ByVal featureId As String)
    If mConfigured Then NxRaiseContractError "Template command can be configured only once"
    Select Case featureId
        Case "NX-TPL-REGISTER-SHEET", "NX-TPL-LIST", "NX-TPL-LOAD", "NX-TPL-RENAME", "NX-TPL-DELETE"
            mFeatureId = featureId
        Case Else
            NxRaiseContractError "Template feature is outside the closed contract"
    End Select
    mConfigured = True
End Sub

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "Template command is not configured"
End Sub
