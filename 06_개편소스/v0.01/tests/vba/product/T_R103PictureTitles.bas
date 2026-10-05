Attribute VB_Name = "T_R103PictureTitles"
Option Explicit

Public Function Names() As String
    Names = "defaults|toggle|payload"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim form As FNxHangulPictures, paths As New Collection, payload As String
    On Error GoTo Failed
    Select Case name
        Case "defaults"
            Set form = New FNxHangulPictures
            Require Not form.Controls("chkTitles").Value, "title rows default off"
            Require form.Controls("chkTitleFileNames").Value, "filename default on"
            Require Not form.Controls("chkTitleFileNames").Enabled, "filename disabled without rows"
        Case "toggle"
            Set form = New FNxHangulPictures
            form.Controls("chkTitles").Value = True
            Require form.Controls("chkTitleFileNames").Enabled, "filename enabled"
            form.Controls("chkTitleFileNames").Value = False
            Require InStr(form.Controls("nxPhoto_1").Caption, "빈 제목칸") > 0, "blank preview"
            form.Controls("chkTitles").Value = False
            Require Not form.Controls("chkTitleFileNames").Enabled, "filename disabled"
            form.Controls("chkTitles").Value = True
            Require Not form.Controls("chkTitleFileNames").Value, "choice preserved"
            form.Controls("chkTitleFileNames").Value = True
            Require InStr(form.Controls("nxPhoto_1").Caption, "파일명") > 0, "filename preview"
        Case "payload"
            paths.Add "C:\test.png"
            payload = NxHangulPicturePayload(paths, 7, 6, 2, 1, True, "title", False)
            Require InStr(payload, """titles"":true") > 0, "row flag"
            Require InStr(payload, """title_filenames"":false") > 0, "blank flag"
            payload = NxHangulPicturePayload(paths, 7, 6, 2, 1, True, "title")
            Require InStr(payload, """title_filenames"":true") > 0, "legacy default"
        Case Else
            Err.Raise vbObjectError + 103, , "unknown case"
    End Select
    If Not form Is Nothing Then Unload form
    RunCase = "PASS|" & name
    Exit Function
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & ":" & Err.Description
    On Error Resume Next
    If Not form Is Nothing Then Unload form
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 103, , detail
End Sub
