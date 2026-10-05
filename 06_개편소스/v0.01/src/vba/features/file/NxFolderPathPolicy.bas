Attribute VB_Name = "NxFolderPathPolicy"
Option Explicit
Public Const NX_FOLDER_CREATE_LIMIT As Long = 500
Public Function NxFolderPathValidate(ByVal baseFolder As String, ByVal relativePath As String, ByRef finalPath As String) As Boolean
    Dim fso As Object, part As Variant, current As String, parts As Variant
    finalPath = vbNullString: baseFolder = Replace$(Trim$(baseFolder), "/", "\")
    Do While Len(baseFolder) > 3 And Right$(baseFolder, 1) = "\": baseFolder = Left$(baseFolder, Len(baseFolder) - 1): Loop
    If Len(baseFolder) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseFolder) Then Exit Function ' approved base folder
    If Len(relativePath) = 0 Or Len(relativePath) > 240 Then Exit Function
    relativePath = Replace$(relativePath, "/", "\")
    If Left$(relativePath, 1) = "\" Or Left$(relativePath, 2) = "\\" Or InStr(relativePath, ":") > 0 Then Exit Function ' absolute
    parts = Split(relativePath, "\"): current = baseFolder
    For Each part In parts
        If Len(CStr(part)) = 0 Or CStr(part) = "." Or CStr(part) = ".." Then Exit Function
        If NxFolderInvalidPart(CStr(part)) Or NxFolderReserved(CStr(part)) Then Exit Function ' reserved
        current = current & "\" & CStr(part)
        If Len(current) > 248 Then Exit Function
        If fso.FileExists(current) Then Exit Function ' file conflict
    Next part
    finalPath = current: NxFolderPathValidate = True
End Function
Private Function NxFolderInvalidPart(ByVal value As String) As Boolean
    Dim token As Variant, i As Long
    For Each token In Array("<", ">", ":", Chr$(34), "/", "\", "|", "?", "*"): If InStr(value, CStr(token)) > 0 Then NxFolderInvalidPart = True: Exit Function
    Next token
    For i = 1 To Len(value): If (AscW(Mid$(value, i, 1)) And &HFFFF&) < 32 Then NxFolderInvalidPart = True: Exit Function
    Next i
End Function
Public Function NxFolderNameIsValid(ByVal value As String) As Boolean
    If Len(value) = 0 Or value = "." Or value = ".." Then Exit Function
    NxFolderNameIsValid = Not NxFolderInvalidPart(value) And Not NxFolderReserved(value)
End Function
Private Function NxFolderReserved(ByVal value As String) As Boolean
    Dim stem As String, dot As Long, token As Variant
    stem = UCase$(value): dot = InStr(stem, "."): If dot > 0 Then stem = Left$(stem, dot - 1)
    For Each token In Array("CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"): If stem = token Then NxFolderReserved = True: Exit Function
    Next token
    NxFolderReserved = (Right$(value, 1) = "." Or Right$(value, 1) = " ")
End Function
