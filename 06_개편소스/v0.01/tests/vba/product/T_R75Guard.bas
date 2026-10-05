Attribute VB_Name = "T_R75Guard"
Option Explicit
Private Const HASH_ABC As String = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad"
Public Function Names() As String
    Names = "normal|normal_locked|changed|missing|wrong_version|missing_manifest_hash|local_url|remote_url"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim path As String, handle As Integer, bytes(0 To 2) As Byte, expected As String, caught As Long, detail As String
    On Error GoTo Failed
    path = Environ$("LHEXCEL_PROFILE_ROOT") & "\guard " & name & ".bin"
    expected = HASH_ABC
    bytes(0) = 97: bytes(1) = 98: bytes(2) = 99
    If name = "changed" Then bytes(2) = 100
    If name = "wrong_version" Then expected = String$(64, "0")
    If name = "missing_manifest_hash" Then expected = vbNullString
    If name <> "missing" Then
        handle = FreeFile: Open path For Binary As #handle: Put #handle, , bytes: Close #handle: handle = 0
    End If
    If name = "normal_locked" Then
        handle = FreeFile: Open path For Binary Access Read Lock Write As #handle
    End If
    On Error Resume Next
    Select Case name
        Case "local_url"
            detail = NxHostProbeUrl("file:///C:/folder%20name/NxHost64.dll")
        Case "remote_url": detail = NxHostProbeUrl("file://server/share/NxHost64.dll")
        Case Else: NxHostProbeFile path, expected
    End Select
    caught = Err.Number: Err.Clear
    On Error GoTo Failed
    If handle <> 0 Then Close #handle: handle = 0
    If name = "normal" Or name = "normal_locked" Or name = "local_url" Then
        If caught <> 0 Then Err.Raise 5, , "valid file/url rejected"
        If name = "local_url" And detail <> "C:\folder name\NxHost64.dll" Then Err.Raise 5, , "decoded path mismatch"
    Else
        If caught = 0 Then Err.Raise 5, , "invalid file/url accepted"
    End If
    RunCase = "PASS|" & name: Exit Function
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
    On Error Resume Next
    If handle <> 0 Then Close #handle
End Function
