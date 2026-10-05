Attribute VB_Name = "T_R69RunnerSafety"
Option Explicit
Public Function Names() As String
    Names = "TestRunnerNormalReturn|TestRunnerForcedExecuteError|TestRunnerUserCancel|TestBuildPlanMutationIsRejectedBeforeGuard|TestApprovalBindsRunSnapshotSourceDigest|TestFinalStateResultMatrix|TestPostChangeExecuteExceptionAndRollbackFailureIsPartial"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    book.Worksheets(1).Range("A1").Value2 = "before"
    book.Worksheets(1).Range("A1").Select
    T_Core.RunCase name
    RunCase = "PASS|" & name
    GoTo CleanUp
Failed:
    detail = Err.Number & "|" & Err.Description
    RunCase = "FAIL|" & name & "|" & detail
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close False
End Function
