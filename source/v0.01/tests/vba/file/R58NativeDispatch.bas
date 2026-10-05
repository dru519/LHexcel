Attribute VB_Name = "R58NativeDispatch"
Option Explicit

Public Function NxR58TestDispatch(ByVal kind As String, ByVal caseName As String) As String
    On Error GoTo Failed
    Select Case kind
        Case "PING"
        Case "COMPARE_PREPARE"
            NxR58TestDispatch = "REPORT|" & T_R58Compare.PrepareConstantsReport
            Exit Function
        Case "COMPARE_VERIFY_OPEN"
            T_R58Compare.VerifyOpenedConstantsReport caseName
        Case "COMPARE_CLEANUP"
            T_R58Compare.CleanupPreparedConstantsReport caseName
        Case "COMPARE_REMAINING"
            T_R58Compare.RunRemaining
        Case "RECOVERY"
            T_R58Recovery.RunCase caseName
        Case "RECOVERY_CLEANUP"
            NxBatchRenameTestClearMoveFailure
        Case "DATA"
            T_R58Data.RunAll
        Case "MODEL_MERGE"
            T_R58ModelMerge.RunAll
        Case Else
            NxRaiseContractError "Unknown r58 native dispatch kind"
    End Select
    NxR58TestDispatch = "PASS"
    Exit Function
Failed:
    NxR58TestDispatch = "FAIL|" & CStr(Err.Number) & "|" & Replace(Replace(Err.Description, vbCr, " "), vbLf, " ")
End Function
