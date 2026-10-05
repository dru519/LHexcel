Attribute VB_Name = "T_R84Ai"
Option Explicit
Public CopiedPrompt As String
Public Function Names() As String
    Names = "first_execute|changed_execute|masked_execute|incomplete_block"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim session As New CNxAiInputSession, privacy As New CNxAiPrivacySnapshot, view As FNxAi
    Dim state As NxPrivacyState, original As String, masked As String
    On Error GoTo Failed
    state = NxPrivacySafe: original = "synthetic-input": masked = original
    If name = "masked_execute" Then state = NxPrivacyBlocked: masked = "[masked]"
    If name = "incomplete_block" Then state = NxPrivacyIncomplete
    privacy.Configure state, original, masked, "test"
    privacy.Seal
    session.ConfigureSelection original, "Sheet1!A1", "1셀", privacy
    Set view = New FNxAi
    view.BindSession session
    CopiedPrompt = vbNullString
    If name = "incomplete_block" Then
        If view.Controls("cmdExecute").Enabled Then Err.Raise 5, , "Incomplete input enabled"
        view.RequestCancel
    Else
        If Not view.Controls("cmdExecute").Enabled Then Err.Raise 5, , "Preview still required"
        If name = "changed_execute" Then
            view.Controls("txtInstruction").Value = "updated-request"
            If Not view.Controls("cmdExecute").Enabled Then Err.Raise 5, , "Edit disabled execution"
        End If
        view.R84Execute
        If Len(CopiedPrompt) = 0 Then Err.Raise 5, , "No prompt created"
        If name = "changed_execute" And InStr(CopiedPrompt, "updated-request") = 0 Then Err.Raise 5, , "Stale instruction"
        If name = "masked_execute" Then
            If InStr(CopiedPrompt, "synthetic-input") > 0 Or InStr(CopiedPrompt, "[masked]") = 0 Then Err.Raise 5, , "Masking bypassed"
        End If
    End If
    RunCase = "PASS|" & name
    Exit Function
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
    On Error Resume Next
    If Not view Is Nothing Then view.RequestCancel
End Function
