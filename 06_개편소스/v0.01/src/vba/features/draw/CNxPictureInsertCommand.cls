VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxPictureInsertCommand"
Option Explicit
Implements INxFeatureCommand

Private mRequest As CNxPictureInsertRequest, mPreview As CNxPictureInsertPreview, mJournal As CNxPictureInsertJournal
Private mConfigured As Boolean, mPlanBuilt As Boolean

Public Sub Configure(ByVal request As CNxPictureInsertRequest, ByVal preview As CNxPictureInsertPreview, ByVal journal As CNxPictureInsertJournal)
    If mConfigured Or request Is Nothing Or preview Is Nothing Or journal Is Nothing Then NxRaiseContractError "그림 삽입 명령 구성이 잘못되었습니다."
    If Not preview.IsSealed Then NxRaiseContractError "그림 삽입 명령에는 봉인된 미리보기가 필요합니다."
    Set mRequest = request: Set mPreview = preview: Set mJournal = journal: mConfigured = True
End Sub

Private Property Get INxFeatureCommand_FeatureId() As String
    EnsureConfigured: INxFeatureCommand_FeatureId = "NX-DRAW-INSERT-PICTURE"
End Property

Private Function INxFeatureCommand_BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
    Dim plan As CNxPlan
    EnsureConfigured: If mPlanBuilt Or context Is Nothing Then NxRaiseContractError "그림 삽입 계획은 유효한 context에서 한 번만 만들 수 있습니다."
    If Not mPreview.Revalidate(mRequest) Then NxRaiseContractError "그림 삽입 미리보기의 대상 또는 파일이 변경되었습니다."
    If mPreview.InsertCount < 1 Or mPreview.InsertCount > 500 Then NxRaiseContractError "그림 삽입 batch는 1개 이상 500개 이하이어야 합니다."
    Set plan = NxCreatePlan("NX-DRAW-INSERT-PICTURE", context.Fingerprint, context.SourceStateDigest)
    plan.DeclareObservedCells 1
    ' Central Packet B supplies this exact bounded typed effect factory and dispatches NxShapeInsert.
    ' This owner command is execution grade NxExecutionPlanned only.
    plan.AddSideEffect NxCreateFeaturePictureInsertBatchEffect("NX-DRAW-INSERT-PICTURE", NxShapeInsert, mPreview.InsertCount, mPreview.TargetDigest, "delete only journalled shapes")
    plan.Seal: mPlanBuilt = True: Set INxFeatureCommand_BuildPlan = plan
End Function

Private Function INxFeatureCommand_Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
    Dim result As CNxResult, normalized As New CNxResult, completedItem As Variant
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    EnsureConfigured: If Not mPlanBuilt Or context Is Nothing Or approvedPlan Is Nothing Then NxRaiseContractError "그림 삽입 명령에는 승인된 계획이 필요합니다."
    If approvedPlan.FeatureId <> "NX-DRAW-INSERT-PICTURE" Then NxRaiseContractError "그림 삽입 승인 계획의 기능 ID가 다릅니다."
    If Not mPreview.Revalidate(mRequest) Then NxRaiseContractError "그림 삽입 실행 전에 preview drift가 발생했습니다."
    On Error GoTo Failed
    Set result = NxPictureInsertExecute(mRequest, mPreview, mJournal)
    ' The approved workbook identity is session-scoped for an unsaved workbook.
    ' Preserve the engine receipt while replacing only its unapproved FullName target.
    normalized.Configure result.FeatureId, result.Outcome, result.Stage, context.WorkbookIdentity, _
        result.SourceChanged, result.Recovery, result.MessageKey
    For Each completedItem In result.Completed
        normalized.AddCompleted CStr(completedItem)
    Next completedItem
    normalized.Seal
    Set INxFeatureCommand_Execute = normalized: Exit Function
Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    mJournal.Rollback
    Err.Raise failureNumber, failureSource, failureDescription
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "그림 삽입 명령이 구성되지 않았습니다."
End Sub
