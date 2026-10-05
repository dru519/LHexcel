Attribute VB_Name = "NxHostIntegrity"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function NxPathFromUrl Lib "shlwapi.dll" Alias "PathCreateFromUrlW" (ByVal url As LongPtr, ByVal path As LongPtr, ByRef length As Long, ByVal flags As Long) As Long
#Else
Private Declare Function NxPathFromUrl Lib "shlwapi.dll" Alias "PathCreateFromUrlW" (ByVal url As Long, ByVal path As Long, ByRef length As Long, ByVal flags As Long) As Long
#End If

' Distribution generation pins both DLL hashes after the native build.
Private Const NX_HOST_VERIFY_FILES As Boolean = False
Private Const NX_HOST_SHA256_X86 As String = ""
Private Const NX_HOST_SHA256_X64 As String = ""
Private mLastError As String

Public Function NxHostIntegrityLastError() As String
    NxHostIntegrityLastError = mLastError
    If Len(mLastError) = 0 Then NxHostIntegrityLastError = "아래한글 DLL을 사용할 수 없습니다. 같은 버전의 고도화판을 다시 설치해 주세요."
End Function

Public Function NxHostVerifiedCreate(ByVal progId As String) As Object
    Dim registry As Object, clsid As String, expectedClsid As String, url As String, path As String
    Dim expected As String, handle As Integer, detail As String
    On Error GoTo Failed
    mLastError = vbNullString
    Select Case progId
        Case "LH.NxHost.BridgeService": expectedClsid = "{9A90294C-81C8-4AEE-A2D8-14B078A61CD2}"
        Case "LH.NxHost.HwpxExportService": expectedClsid = "{1434C649-18AA-4435-B0D2-2BD81B548A01}"
        Case "LH.NxHost.PicturePreviewService": expectedClsid = "{C812F023-E912-49AA-98B7-75DC86501902}"
        Case "LH.NxHost.WorkbookCompareService": expectedClsid = "{C812F023-E912-49AA-98B7-75DC86501904}"
        Case "LH.NxHost.WorkbookCompareResultsService": expectedClsid = "{C812F023-E912-49AA-98B7-75DC86501906}"
        Case Else: NxRaiseContractError "지원하지 않는 DLL 서비스입니다."
    End Select
    If NX_HOST_VERIFY_FILES Then
        Set registry = CreateObject("WScript.Shell")
        clsid = CStr(registry.RegRead("HKEY_CLASSES_ROOT\" & progId & "\CLSID\"))
        If StrComp(clsid, expectedClsid, vbTextCompare) <> 0 Then NxRaiseContractError "DLL 등록 정보가 일치하지 않습니다."
        url = CStr(registry.RegRead("HKEY_CLASSES_ROOT\CLSID\" & clsid & "\InprocServer32\CodeBase"))
        path = NxHostLocalCodeBase(url)
        #If Win64 Then
            expected = NX_HOST_SHA256_X64
        #Else
            expected = NX_HOST_SHA256_X86
        #End If
        handle = FreeFile
        Open path For Binary Access Read Lock Write As #handle
        NxHostRequireFileHash path, expected
    End If
    Set NxHostVerifiedCreate = CreateObject(progId)
    If handle <> 0 Then Close #handle
    Exit Function
Failed:
    detail = Err.Description
    On Error Resume Next
    If handle <> 0 Then Close #handle
    On Error GoTo 0
    mLastError = "내엑셀 DLL 연결을 확인하지 못했습니다. 같은 버전의 설치 파일로 다시 설치해 주세요."
    If Len(path) > 0 Then mLastError = mLastError & vbCrLf & path
    If Len(detail) > 0 Then mLastError = mLastError & vbCrLf & detail
    Set NxHostVerifiedCreate = Nothing
End Function

Private Function NxHostLocalCodeBase(ByVal url As String) As String
    Dim buffer As String, length As Long, path As String
    If LCase$(Left$(url, 8)) <> "file:///" Then NxRaiseContractError "DLL은 Windows 로컬 폴더에 설치해 주세요."
    buffer = String$(32768, vbNullChar): length = Len(buffer)
    If NxPathFromUrl(StrPtr(url), StrPtr(buffer), length, 0) <> 0 Then NxRaiseContractError "DLL 경로를 확인할 수 없습니다."
    path = Left$(buffer, InStr(buffer, vbNullChar) - 1)
    If Len(path) < 4 Or Mid$(path, 2, 2) <> ":\" Or InStr(3, path, ":") > 0 Then NxRaiseContractError "DLL 로컬 경로가 올바르지 않습니다."
    NxHostLocalCodeBase = path
End Function

Private Sub NxHostRequireFileHash(ByVal path As String, ByVal expected As String)
    If Len(expected) <> 64 Then NxRaiseContractError "배포본 DLL 검증 정보가 없습니다."
    If StrComp(NxTemplateFileSha256(path), expected, vbBinaryCompare) <> 0 Then _
        NxRaiseContractError "DLL 파일이 변경되었거나 다른 버전입니다."
End Sub
