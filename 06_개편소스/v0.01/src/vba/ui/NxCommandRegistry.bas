Attribute VB_Name = "NxCommandRegistry"
Option Explicit

Private mCommandRegistry As CNxCommandRegistry

Public Function NxCreateCommandRegistry() As CNxCommandRegistry
    Dim registry As New CNxCommandRegistry
    NxRegisterGeneratedCommands registry
    Set NxCreateCommandRegistry = registry
End Function

Public Function NxCommandDefinition(ByVal commandId As String) As CNxCommandDefinition
    Dim definition As CNxCommandDefinition
    Set definition = NxCommandCatalog().CommandById(commandId)
    If definition Is Nothing Then NxRaiseContractError "알 수 없는 내엑셀 명령입니다."
    Set NxCommandDefinition = definition
End Function

Public Function NxCommandsForCategory(ByVal categoryId As String) As Collection
    Set NxCommandsForCategory = NxCommandCatalog().CommandsForCategory(categoryId)
End Function

Public Sub NxExecuteRegisteredCommand(ByVal definition As CNxCommandDefinition)
    NxDistributionEnsureExecutable
    Dim guard As CNxStateGuard
    Dim errorNumber As Long
    Dim errorDescription As String
    If definition Is Nothing Then NxRaiseContractError "내엑셀 명령 정의가 필요합니다."
    If Not definition.IsSealed Then NxRaiseContractError "봉인되지 않은 내엑셀 명령입니다."
    If definition.CommandId = "NX-CMD-SHEET-RB-SHEET-SAVE-TO-FILE" Then NxRaiseContractError "삭제된 기능입니다."
    If definition.MutatesDocument Then
        If StrComp(definition.RecoveryPolicy, "command_scoped_guard", vbBinaryCompare) <> 0 Then NxRaiseContractError "Mutation command recovery policy is not approved"
    ElseIf StrComp(definition.RecoveryPolicy, "none", vbBinaryCompare) <> 0 Then
        NxRaiseContractError "Read-only command recovery policy is invalid"
    End If
    On Error GoTo Failed
    If definition.MutatesDocument Then Set guard = New CNxStateGuard
    Select Case definition.Handler
        Case "NxCmdClipboard": NxCmdClipboard definition.ArgumentKey
        Case "NxCmdPrint": NxCmdPrint definition.ArgumentKey
        Case "NxCmdStyle": NxCmdStyle definition.ArgumentKey
        Case "NxCmdNumberFormat": NxCmdNumberFormat definition.ArgumentKey
        Case "NxCmdSelection": NxCmdSelection definition.ArgumentKey
        Case "NxCmdRowsColumns": NxCmdRowsColumns definition.ArgumentKey
        Case "NxCmdAlignment": NxCmdAlignment definition.ArgumentKey
        Case "NxCmdInsertDelete": NxCmdInsertDelete definition.ArgumentKey
        Case "NxCmdFilterSort": NxCmdFilterSort definition.ArgumentKey
        Case "NxCmdFormula": NxCmdFormula definition.ArgumentKey
        Case "NxCmdSheet": NxCmdSheet definition.ArgumentKey
        Case "NxCmdView": NxCmdView definition.ArgumentKey
        Case "NxCmdInfo": NxCmdInfo definition.ArgumentKey
        Case "NxCmdPivot": NxCmdPivot definition.ArgumentKey
        Case Else: NxRaiseContractError "허용되지 않은 내엑셀 명령 처리기입니다."
    End Select
    GoTo FinallyRestore
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
FinallyRestore:
    If Not guard Is Nothing Then
        guard.Restore
        If Len(guard.Recovery) > 0 Then
            If errorNumber = 0 Then errorNumber = vbObjectError + 2163
            errorDescription = NxJoinFailure(errorDescription, "command restore recovery required: " & guard.Recovery)
        End If
    End If
    If errorNumber <> 0 Then Err.Raise errorNumber, "NxExecuteRegisteredCommand", errorDescription
    If definition.ArgumentKey = "RB_EDIT_MEMO_ADD_LHEXCELFORMULA" Then NxFormulaNotesUndoArm
    If definition.ArgumentKey = "RB_FORMULA_PRECEDENTS_LIST" Then NxFormulaReferenceReveal
    If definition.Handler = "NxCmdStyle" And Left$(definition.ArgumentKey, 9) = "NX_STYLE_" Then NxQuickFormatUndoArm
    If definition.Handler = "NxCmdFormula" And Left$(definition.ArgumentKey, 12) = "NX_FUNCTION_" Then NxFunctionUndoArm
End Sub

Public Sub NxResetCommandRegistry()
    Set mCommandRegistry = Nothing
End Sub

Private Function NxCommandCatalog() As CNxCommandRegistry
    If mCommandRegistry Is Nothing Then Set mCommandRegistry = NxCreateCommandRegistry()
    Set NxCommandCatalog = mCommandRegistry
End Function
