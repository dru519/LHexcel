Attribute VB_Name = "NxFolderCreateController"
Option Explicit
Public Sub NxFolderCreateOpen(ByVal context As CNxFeatureDialogContext)
    Dim form As New FNxFolderCreate: form.BindContext context: form.Show vbModal
End Sub
Public Function NxFolderCreateRun(ByVal baseFolder As String, ByVal items As Collection, Optional ByVal confirmedPreview As CNxFolderCreatePreview) As CNxResult
    Dim preview As CNxFolderCreatePreview, commandObject As New CNxFolderCreateCommand, command As INxFeatureCommand, router As New CNxExecutionRouter, ticket As CNxExecutionTicket, definition As CNxFeatureDefinition
    If Len(baseFolder) = 0 Or items Is Nothing Then NxRaiseContractError "Folder create requires base folder and items"
    If confirmedPreview Is Nothing Then Set preview = New CNxFolderCreatePreview: preview.Snapshot baseFolder, items Else Set preview = confirmedPreview
    If Not preview.IsValid Then NxRaiseContractError "Folder preview snapshot invalidated"
    commandObject.Configure baseFolder, items, preview: Set command = commandObject: Set definition = NxFileFeatureDefinition("NX-FILE-FOLDER-CREATE")
    Set ticket = router.Prepare(definition, command): If ticket.Decision.ResolvedGrade <> NxExecutionPlanned Then NxRaiseContractError "Folder execution grade is invalid"
    Set NxFolderCreateRun = NxRunPlannedFile(ticket, command, definition)
End Function
Public Function NxFolderCreateSummary() As String: NxFolderCreateSummary = "승인된 기준 폴더 아래에 상대 경로 폴더를 미리보기 후 최대 500개까지 생성합니다.": End Function
