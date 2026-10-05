Attribute VB_Name = "NxExecutionGradePolicy"
Option Explicit

Public Function ResolveGrade(ByVal baseGrade As NxExecutionGrade, ByVal actualCellCount As Double, _
    ByVal plannedCellThreshold As Long, ByVal aiProcessing As Boolean, _
    ByVal externalDataTransfer As Boolean, ByVal fileMutation As Boolean, _
    ByVal externalDocumentMutation As Boolean, ByVal manualRecoveryMutation As Boolean, _
    Optional ByVal imageInsertion As Boolean = False, _
    Optional ByVal riskyFormula As Boolean = False, Optional ByVal resultSheetCount As Long = 0, _
    Optional ByVal existingContentOverwrite As Boolean = False, Optional ByVal fileCreation As Boolean = False) As NxExecutionGrade

    Dim resolved As NxExecutionGrade
    ValidateInputs baseGrade, actualCellCount, plannedCellThreshold
    If resultSheetCount < 0 Then NxRaiseContractError "Execution policy result sheet count is invalid"
    resolved = baseGrade
    If aiProcessing Or externalDataTransfer Then resolved = RaiseGrade(resolved, NxExecutionGuarded)
    If actualCellCount > plannedCellThreshold Then resolved = RaiseGrade(resolved, NxExecutionPlanned)
    If resultSheetCount = 1 Then resolved = RaiseGrade(resolved, NxExecutionGuarded)
    If resultSheetCount > 1 Or existingContentOverwrite Or fileCreation Then resolved = RaiseGrade(resolved, NxExecutionPlanned)
    If fileMutation Or externalDocumentMutation Or manualRecoveryMutation Or imageInsertion Or riskyFormula Then
        resolved = RaiseGrade(resolved, NxExecutionPlanned)
    End If
    ResolveGrade = resolved
End Function

Public Function NxResolveRuntimeGrade(ByVal baseGrade As NxExecutionGrade, ByVal actualCellCount As Double, _
    ByVal resultSheetCount As Long, ByVal existingContentOverwrite As Boolean, ByVal fileCreation As Boolean) As NxExecutionGrade
    NxResolveRuntimeGrade = ResolveGrade(baseGrade, actualCellCount, 10000, False, False, False, False, False, _
        False, False, resultSheetCount, existingContentOverwrite, fileCreation)
End Function

Public Function ResolveDecision(ByVal definition As CNxFeatureDefinition, ByVal actualCellCount As Double, _
    ByVal fileMutation As Boolean, ByVal externalDocumentMutation As Boolean, _
    ByVal manualRecoveryMutation As Boolean, Optional ByVal imageInsertion As Boolean = False, _
    Optional ByVal riskyFormula As Boolean = False, Optional ByVal resultSheetCount As Long = 0, _
    Optional ByVal existingContentOverwrite As Boolean = False, Optional ByVal fileCreation As Boolean = False) As CNxExecutionDecision

    Dim resolved As NxExecutionGrade
    Dim decision As New CNxExecutionDecision
    Dim reasonCode As String
    If definition Is Nothing Or Not definition.IsSealed Then NxRaiseContractError "Execution policy requires a sealed feature definition"
    resolved = ResolveGrade(definition.ExecutionGrade, actualCellCount, definition.PlannedCellThreshold, _
        definition.AiProcessing, definition.NetworkTransfer, fileMutation, externalDocumentMutation, _
        manualRecoveryMutation, imageInsertion, riskyFormula, resultSheetCount, existingContentOverwrite, fileCreation)
    reasonCode = DecisionReason(definition.ExecutionGrade, resolved, actualCellCount, definition.PlannedCellThreshold, _
        definition.AiProcessing, definition.NetworkTransfer, fileMutation, externalDocumentMutation, _
        manualRecoveryMutation, imageInsertion, riskyFormula)
    decision.Configure definition.ExecutionGrade, resolved, reasonCode, actualCellCount, _
        definition.RequiresPrivacyInspection
    decision.Seal
    Set ResolveDecision = decision
End Function

Public Function RaiseGrade(ByVal currentGrade As NxExecutionGrade, ByVal candidate As NxExecutionGrade) As NxExecutionGrade
    If candidate > currentGrade Then
        RaiseGrade = candidate
    Else
        RaiseGrade = currentGrade
    End If
End Function

Private Sub ValidateInputs(ByVal baseGrade As NxExecutionGrade, ByVal actualCellCount As Double, ByVal plannedCellThreshold As Long)
    If baseGrade < NxExecutionFast Or baseGrade > NxExecutionPlanned Then NxRaiseContractError "Execution policy base grade is invalid"
    If actualCellCount < 0 Or actualCellCount <> Fix(actualCellCount) Then NxRaiseContractError "Execution policy cell count is invalid"
    If plannedCellThreshold <= 0 Or plannedCellThreshold > 10000 Then NxRaiseContractError "Execution policy threshold is invalid"
End Sub

Private Function DecisionReason(ByVal baseGrade As NxExecutionGrade, ByVal resolvedGrade As NxExecutionGrade, _
    ByVal actualCellCount As Double, ByVal plannedCellThreshold As Long, ByVal aiProcessing As Boolean, _
    ByVal externalDataTransfer As Boolean, ByVal fileMutation As Boolean, _
    ByVal externalDocumentMutation As Boolean, ByVal manualRecoveryMutation As Boolean, _
    ByVal imageInsertion As Boolean, ByVal riskyFormula As Boolean) As String

    If imageInsertion Then
        DecisionReason = "image_insertion"
    ElseIf fileMutation Then
        DecisionReason = "file_mutation"
    ElseIf externalDocumentMutation Then
        DecisionReason = "external_document_mutation"
    ElseIf manualRecoveryMutation Then
        DecisionReason = "manual_recovery_mutation"
    ElseIf riskyFormula Then
        DecisionReason = "risky_formula"
    ElseIf actualCellCount > plannedCellThreshold Then
        DecisionReason = "bulk_mutation"
    ElseIf externalDataTransfer Then
        DecisionReason = "external_data_transfer"
    ElseIf aiProcessing Then
        DecisionReason = "ai_processing"
    ElseIf resolvedGrade = baseGrade Then
        DecisionReason = "registry_base_grade"
    Else
        DecisionReason = "runtime_escalation"
    End If
End Function
