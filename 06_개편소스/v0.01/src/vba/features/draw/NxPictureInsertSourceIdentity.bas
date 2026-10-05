Attribute VB_Name = "NxPictureInsertSourceIdentity"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt.dll" (ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, ByRef pbOutput As Any, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, ByRef pbHashObject As Any, ByVal cbHashObject As Long, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByRef pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As LongPtr, ByRef pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As LongPtr) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function GetFullPathNameW Lib "kernel32.dll" (ByVal lpFileName As LongPtr, ByVal nBufferLength As Long, ByVal lpBuffer As LongPtr, ByVal lpFilePart As LongPtr) As Long
Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As LongPtr) As Long
Private Declare PtrSafe Function NormalizeString Lib "Normaliz.dll" (ByVal normForm As Long, ByVal lpSrcString As LongPtr, ByVal cwSrcLength As Long, ByVal lpDstString As LongPtr, ByVal cwDstLength As Long) As Long
#Else
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt.dll" (ByRef phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGetProperty Lib "bcrypt.dll" (ByVal hObject As Long, ByVal pszProperty As Long, ByRef pbOutput As Any, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCreateHash Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByRef phHash As Long, ByRef pbHashObject As Any, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptHashData Lib "bcrypt.dll" (ByVal hHash As Long, ByRef pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptFinishHash Lib "bcrypt.dll" (ByVal hHash As Long, ByRef pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyHash Lib "bcrypt.dll" (ByVal hHash As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt.dll" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetFullPathNameW Lib "kernel32.dll" (ByVal lpFileName As Long, ByVal nBufferLength As Long, ByVal lpBuffer As Long, ByVal lpFilePart As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function NormalizeString Lib "Normaliz.dll" (ByVal normForm As Long, ByVal lpSrcString As Long, ByVal cwSrcLength As Long, ByVal lpDstString As Long, ByVal cwDstLength As Long) As Long
#End If

Private Const NX_SOURCE_STATUS_SUCCESS As Long = 0
Private Const NX_SOURCE_ATTRIBUTE_INVALID As Long = -1
Private Const NX_SOURCE_REPARSE_POINT As Long = &H400
Private Const FILE_ATTRIBUTE_REPARSE_POINT As Long = &H400
Private Const NX_SOURCE_HASH_CHUNK_SIZE As Long = 65536
Private Const NX_SOURCE_REPARSE_POLICY As String = "REPARSE_REJECTED"
Private Const NX_SOURCE_OBJECT_LENGTH As String = "ObjectLength"
Private Const NX_SOURCE_HASH_LENGTH As String = "HashDigestLength"
Private Const NX_SOURCE_NORM_FORM_C As Long = 1

Public Function NxPictureInsertBuildSourceIdentity(ByVal rawPath As String) As String
    Dim canonicalPath As String, byteLength As Long, modifiedAt As Date, contentSha256 As String
    canonicalPath = NxPictureInsertCanonicalPath(rawPath)
    If NxPictureInsertHasReparsePath(canonicalPath) Then NxRaiseContractError "그림 원본의 reparse 경로는 허용되지 않습니다."
    byteLength = FileLen(canonicalPath)
    If byteLength <= 0 Then NxRaiseContractError "빈 그림 파일은 삽입할 수 없습니다."
    modifiedAt = FileDateTime(canonicalPath)
    contentSha256 = NxPictureInsertFileSha256(canonicalPath)
    If Len(contentSha256) <> 64 Then NxRaiseContractError "그림 원본 SHA-256을 확인할 수 없습니다."
    NxPictureInsertBuildSourceIdentity = canonicalPath & "|" & contentSha256 & "|" & NxPictureInsertCanonicalFileLength(byteLength) & "|" & NxPictureInsertCanonicalLastWrite(modifiedAt) & "|" & NX_SOURCE_REPARSE_POLICY
End Function

Private Function NxPictureInsertCanonicalFileLength(ByVal value As Long) As String
    NxPictureInsertCanonicalFileLength = Trim$(Str$(value))
End Function

Private Function NxPictureInsertCanonicalLastWrite(ByVal value As Date) As String
    NxPictureInsertCanonicalLastWrite = NxPictureInsertPaddedInteger(Year(value), 4) & "-" & _
        NxPictureInsertPaddedInteger(Month(value), 2) & "-" & NxPictureInsertPaddedInteger(Day(value), 2) & " " & _
        NxPictureInsertPaddedInteger(Hour(value), 2) & ":" & NxPictureInsertPaddedInteger(Minute(value), 2) & ":" & _
        NxPictureInsertPaddedInteger(Second(value), 2)
End Function

Private Function NxPictureInsertPaddedInteger(ByVal value As Long, ByVal width As Long) As String
    NxPictureInsertPaddedInteger = Right$(String$(width, "0") & Trim$(Str$(value)), width)
End Function

Public Function NxPictureInsertCanonicalPath(ByVal rawPath As String) As String
    Dim buffer As String, length As Long
    If Len(Trim$(rawPath)) = 0 Then NxRaiseContractError "그림 원본 경로가 비어 있습니다."
    buffer = Space$(32768)
    length = GetFullPathNameW(StrPtr(rawPath), Len(buffer), StrPtr(buffer), 0)
    If length <= 0 Or length >= Len(buffer) Then NxRaiseContractError "그림 원본 경로를 canonicalize할 수 없습니다."
    NxPictureInsertCanonicalPath = Left$(buffer, length)
End Function

Public Function NxPictureInsertCanonicalPathKey(ByVal rawPath As String) As String
    NxPictureInsertCanonicalPathKey = LCase$(NxPictureInsertNormalizeC(NxPictureInsertCanonicalPath(rawPath)))
End Function

Private Function NxPictureInsertNormalizeC(ByVal value As String) As String
    Dim requiredLength As Long, writtenLength As Long, buffer As String
    If Len(value) = 0 Then NxRaiseContractError "그림 원본 경로가 비어 있습니다."
    requiredLength = NormalizeString(NX_SOURCE_NORM_FORM_C, StrPtr(value), Len(value), 0, 0)
    If requiredLength <= 0 Then NxRaiseContractError "그림 원본 경로 Unicode 정규화를 수행할 수 없습니다."
    buffer = Space$(requiredLength)
    writtenLength = NormalizeString(NX_SOURCE_NORM_FORM_C, StrPtr(value), Len(value), StrPtr(buffer), requiredLength)
    If writtenLength <= 0 Then NxRaiseContractError "그림 원본 경로 Unicode 정규화에 실패했습니다."
    NxPictureInsertNormalizeC = Left$(buffer, writtenLength)
End Function

Public Function NxPictureInsertSourceReparsePolicy() As String
    NxPictureInsertSourceReparsePolicy = NX_SOURCE_REPARSE_POLICY
End Function

Private Function NxPictureInsertHasReparsePath(ByVal canonicalPath As String) As Boolean
    Dim candidate As String, slash As Long, attributes As Long
    candidate = canonicalPath
    Do
        attributes = GetFileAttributesW(StrPtr(candidate))
        If attributes = NX_SOURCE_ATTRIBUTE_INVALID Then NxRaiseContractError "그림 원본 경로의 속성을 확인할 수 없습니다."
        If (attributes And FILE_ATTRIBUTE_REPARSE_POINT) <> 0 Then NxPictureInsertHasReparsePath = True: Exit Function
        slash = InStrRev(candidate, "\")
        If slash <= 3 Then Exit Do
        candidate = Left$(candidate, slash - 1)
    Loop
End Function

Private Function NxPictureInsertFileSha256(ByVal path As String) As String
#If VBA7 Then
    Dim algorithm As LongPtr, hashHandle As LongPtr
#Else
    Dim algorithm As Long, hashHandle As Long
#End If
    Dim handle As Integer, remaining As Long, chunkLength As Long
    Dim objectLength As Long, hashLength As Long, resultLength As Long, index As Long
    Dim hashObject() As Byte, buffer() As Byte, digest() As Byte
    On Error GoTo Failed
    handle = FreeFile
    Open path For Binary Access Read Lock Write As #handle
    If BCryptOpenAlgorithmProvider(algorithm, StrPtr("SHA256"), 0, 0) <> NX_SOURCE_STATUS_SUCCESS Then GoTo Failed
    If BCryptGetProperty(algorithm, StrPtr(NX_SOURCE_OBJECT_LENGTH), objectLength, 4, resultLength, 0) <> NX_SOURCE_STATUS_SUCCESS Then GoTo Failed
    If BCryptGetProperty(algorithm, StrPtr(NX_SOURCE_HASH_LENGTH), hashLength, 4, resultLength, 0) <> NX_SOURCE_STATUS_SUCCESS Then GoTo Failed
    ReDim hashObject(0 To objectLength - 1)
    ReDim digest(0 To hashLength - 1)
    If BCryptCreateHash(algorithm, hashHandle, hashObject(0), objectLength, 0, 0, 0) <> NX_SOURCE_STATUS_SUCCESS Then GoTo Failed
    remaining = LOF(handle)
    Do While remaining > 0
        chunkLength = remaining
        If chunkLength > NX_SOURCE_HASH_CHUNK_SIZE Then chunkLength = NX_SOURCE_HASH_CHUNK_SIZE
        ReDim buffer(0 To chunkLength - 1)
        Get #handle, , buffer
        If BCryptHashData(hashHandle, buffer(0), chunkLength, 0) <> NX_SOURCE_STATUS_SUCCESS Then GoTo Failed
        remaining = remaining - chunkLength
    Loop
    If BCryptFinishHash(hashHandle, digest(0), hashLength, 0) <> NX_SOURCE_STATUS_SUCCESS Then GoTo Failed
    For index = 0 To hashLength - 1
        NxPictureInsertFileSha256 = NxPictureInsertFileSha256 & LCase$(Right$("0" & Hex$(digest(index)), 2))
    Next index
CleanExit:
    On Error Resume Next
    If handle <> 0 Then Close #handle
    If hashHandle <> 0 Then BCryptDestroyHash hashHandle
    If algorithm <> 0 Then BCryptCloseAlgorithmProvider algorithm, 0
    On Error GoTo 0
    If Len(NxPictureInsertFileSha256) <> 64 Then NxRaiseContractError "Windows CNG 그림 원본 SHA-256이 실패했습니다."
    Exit Function
Failed:
    NxPictureInsertFileSha256 = vbNullString
    Resume CleanExit
End Function
