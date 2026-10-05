VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxBatchRenameJournal"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function CreateFileW Lib "kernel32" (ByVal fileNamePointer As LongPtr, ByVal desiredAccess As Long, _
    ByVal shareMode As Long, ByVal securityAttributes As LongPtr, ByVal creationDisposition As Long, _
    ByVal flagsAndAttributes As Long, ByVal templateFile As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFileInformationByHandle Lib "kernel32" (ByVal fileHandle As LongPtr, _
    ByRef information As NX_BY_HANDLE_FILE_INFORMATION) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal objectHandle As LongPtr) As Long
#Else
Private Declare Function CreateFileW Lib "kernel32" (ByVal fileNamePointer As Long, ByVal desiredAccess As Long, _
    ByVal shareMode As Long, ByVal securityAttributes As Long, ByVal creationDisposition As Long, _
    ByVal flagsAndAttributes As Long, ByVal templateFile As Long) As Long
Private Declare Function GetFileInformationByHandle Lib "kernel32" (ByVal fileHandle As Long, _
    ByRef information As NX_BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal objectHandle As Long) As Long
#End If

Private Type NX_FILETIME
    lowDateTime As Long
    highDateTime As Long
End Type

Private Type NX_BY_HANDLE_FILE_INFORMATION
    fileAttributes As Long
    creationTime As NX_FILETIME
    lastAccessTime As NX_FILETIME
    lastWriteTime As NX_FILETIME
    volumeSerialNumber As Long
    fileSizeHigh As Long
    fileSizeLow As Long
    numberOfLinks As Long
    fileIndexHigh As Long
    fileIndexLow As Long
End Type

Private Const NX_RENAME_AT_ORIGINAL As Long = 0
Private Const NX_RENAME_AT_TEMPORARY As Long = 1
Private Const NX_RENAME_AT_TARGET As Long = 2
Private Const NX_OPEN_EXISTING As Long = 3
Private Const NX_FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const NX_FILE_SHARE_ALL As Long = &H7
Private Const NX_INVALID_HANDLE_VALUE As Long = -1

Private mOriginalPaths As Collection
Private mTemporaryPaths As Collection
Private mTargetPaths As Collection
Private mFileIdentities As Collection
Private mStages As Collection
Private mSealed As Boolean

Private Sub Class_Initialize()
    Set mOriginalPaths = New Collection
    Set mTemporaryPaths = New Collection
    Set mTargetPaths = New Collection
    Set mFileIdentities = New Collection
    Set mStages = New Collection
End Sub

Public Sub Add(ByVal originalPath As String, ByVal temporaryPath As String, ByVal targetPath As String)
    If mSealed Then NxRaiseContractError "Rename journal is sealed"
    If Len(originalPath) = 0 Or Len(temporaryPath) = 0 Or Len(targetPath) = 0 Then _
        NxRaiseContractError "Rename journal requires exact paths"
    mOriginalPaths.Add originalPath
    mTemporaryPaths.Add temporaryPath
    mTargetPaths.Add targetPath
    ' Capture once while the original is known. Creation time/size are not a
    ' substitute for this kernel identity, which survives a same-volume rename.
    mFileIdentities.Add NxRenameJournalFileIdentity(originalPath)
    mStages.Add NX_RENAME_AT_ORIGINAL
End Sub

Public Sub Seal()
    If mSealed Or mOriginalPaths.Count = 0 Then NxRaiseContractError "Rename journal cannot be sealed"
    mSealed = True
End Sub

Public Property Get Count() As Long
    Count = mOriginalPaths.Count
End Property

Public Property Get OriginalPath(ByVal index As Long) As String
    OriginalPath = CStr(mOriginalPaths(index))
End Property

Public Property Get TemporaryPath(ByVal index As Long) As String
    TemporaryPath = CStr(mTemporaryPaths(index))
End Property

Public Property Get TargetPath(ByVal index As Long) As String
    TargetPath = CStr(mTargetPaths(index))
End Property

Public Property Get Stage(ByVal index As Long) As Long
    Stage = CLng(mStages(index))
End Property

Public Function FileIdentity(ByVal candidate As String) As String
    FileIdentity = NxRenameJournalFileIdentity(candidate)
End Function

Public Sub RequireRollbackSafe()
    Dim index As Long, ownedPaths As Object, currentPath As String
    Set ownedPaths = CreateObject("Scripting.Dictionary"): ownedPaths.CompareMode = vbTextCompare
    ' Validate the entire journal before undo moves even one file.
    For index = 1 To Count
        Select Case Stage(index)
            Case NX_RENAME_AT_ORIGINAL: currentPath = OriginalPath(index)
            Case NX_RENAME_AT_TEMPORARY: currentPath = TemporaryPath(index)
            Case NX_RENAME_AT_TARGET: currentPath = TargetPath(index)
            Case Else: NxRaiseContractError "되돌리기 단계가 올바르지 않습니다."
        End Select
        NxRenameJournalRequireOwnedIdentity currentPath, CStr(mFileIdentities(index)), "undo-preflight"
        ownedPaths.Add currentPath, True
    Next index
    For index = 1 To Count
        If NxRenameJournalPathExists(OriginalPath(index)) And Not ownedPaths.Exists(OriginalPath(index)) Then _
            NxRaiseContractError "되돌릴 원래 이름에 다른 파일이 있습니다: " & OriginalPath(index)
        If NxRenameJournalPathExists(TemporaryPath(index)) And Not ownedPaths.Exists(TemporaryPath(index)) Then _
            NxRaiseContractError "되돌리기 임시 이름에 다른 파일이 있습니다: " & TemporaryPath(index)
    Next index
End Sub

Public Sub MoveOriginalToTemporary(ByVal index As Long)
    If Stage(index) <> NX_RENAME_AT_ORIGINAL Then NxRaiseContractError "Rename original stage is not pending"
    NxRenameJournalMoveOwned OriginalPath(index), TemporaryPath(index), CStr(mFileIdentities(index)), "original-to-temporary"
    NxRenameJournalSetStage index, NX_RENAME_AT_TEMPORARY
End Sub

Public Sub MoveTemporaryToTarget(ByVal index As Long)
    If Stage(index) <> NX_RENAME_AT_TEMPORARY Then NxRaiseContractError "Rename temporary stage is not pending"
    NxRenameJournalMoveOwned TemporaryPath(index), TargetPath(index), CStr(mFileIdentities(index)), "temporary-to-target"
    NxRenameJournalSetStage index, NX_RENAME_AT_TARGET
End Sub

' This journal is intentionally in-process only. A later process cannot safely
' infer ownership from a leftover path after another process may have replaced it.
Public Sub Rollback()
    Dim index As Long, failures As String
    If Not mSealed Then NxRaiseContractError "Rename rollback requires a sealed journal"

    ' Vacate every completed target before restoring any original. This keeps
    ' chains such as a->aa, aa->aaa from moving another item's file by mistake.
    For index = Count To 1 Step -1
        If Stage(index) = NX_RENAME_AT_TARGET Then
            On Error GoTo TargetRollbackFailed
            NxRenameJournalMoveOwned TargetPath(index), TemporaryPath(index), CStr(mFileIdentities(index)), "rollback-target-to-temporary"
            NxRenameJournalSetStage index, NX_RENAME_AT_TEMPORARY
        End If
ContinueTargetRollback:
        On Error GoTo 0
    Next index

    For index = Count To 1 Step -1
        If Stage(index) = NX_RENAME_AT_TEMPORARY Then
            On Error GoTo OriginalRollbackFailed
            NxRenameJournalMoveOwned TemporaryPath(index), OriginalPath(index), CStr(mFileIdentities(index)), "rollback-temporary-to-original"
            NxRenameJournalSetStage index, NX_RENAME_AT_ORIGINAL
        End If
ContinueOriginalRollback:
        On Error GoTo 0
    Next index

    If Len(failures) > 0 Then NxRaiseContractError "이름 변경 복구 실패: " & failures
    Exit Sub

TargetRollbackFailed:
    failures = NxRenameJournalJoinFailure(failures, "target-to-temporary|residual=" & TargetPath(index) & "|" & CStr(Err.Number) & "|" & Err.Description)
    Err.Clear
    Resume ContinueTargetRollback
OriginalRollbackFailed:
    failures = NxRenameJournalJoinFailure(failures, "temporary-to-original|residual=" & TemporaryPath(index) & "|" & CStr(Err.Number) & "|" & Err.Description)
    Err.Clear
    Resume ContinueOriginalRollback
End Sub

Private Sub NxRenameJournalMoveOwned(ByVal sourcePath As String, ByVal destinationPath As String, _
    ByVal expectedIdentity As String, ByVal stageName As String)
    NxRenameJournalRequireOwnedIdentity sourcePath, expectedIdentity, stageName
    If NxRenameJournalPathExists(destinationPath) Then _
        NxRaiseContractError "이름 변경 " & stageName & " 대상이 비어 있지 않습니다: " & destinationPath
    Name sourcePath As destinationPath
    ' Do not mark a stage merely because the source disappeared: confirm that
    ' the renamed object is still the object captured at journal creation.
    NxRenameJournalRequireOwnedIdentity destinationPath, expectedIdentity, stageName & " 확인"
End Sub

Private Sub NxRenameJournalRequireOwnedIdentity(ByVal candidate As String, ByVal expectedIdentity As String, ByVal stageName As String)
    Dim actualIdentity As String
    actualIdentity = NxRenameJournalFileIdentity(candidate)
    If StrComp(actualIdentity, expectedIdentity, vbBinaryCompare) <> 0 Then _
        NxRaiseContractError "이름 변경 " & stageName & " 파일 정체성이 일치하지 않습니다: " & candidate
End Sub

Private Function NxRenameJournalFileIdentity(ByVal candidate As String) As String
#If VBA7 Then
    Dim fileHandle As LongPtr
#Else
    Dim fileHandle As Long
#End If
    Dim information As NX_BY_HANDLE_FILE_INFORMATION, lastError As Long
    fileHandle = CreateFileW(StrPtr(candidate), 0, NX_FILE_SHARE_ALL, 0, NX_OPEN_EXISTING, NX_FILE_ATTRIBUTE_NORMAL, 0)
    If fileHandle = NX_INVALID_HANDLE_VALUE Then
        lastError = Err.LastDllError
        NxRaiseContractError "파일 정체성을 확인할 수 없습니다: " & candidate & " (Win32 " & CStr(lastError) & ")"
    End If
    If GetFileInformationByHandle(fileHandle, information) = 0 Then
        lastError = Err.LastDllError
        CloseHandle fileHandle
        NxRaiseContractError "파일 정체성을 확인할 수 없습니다: " & candidate & " (Win32 " & CStr(lastError) & ")"
    End If
    CloseHandle fileHandle
    NxRenameJournalFileIdentity = Hex$(information.volumeSerialNumber) & ":" & Hex$(information.fileIndexHigh) & ":" & Hex$(information.fileIndexLow)
End Function

Private Sub NxRenameJournalSetStage(ByVal index As Long, ByVal value As Long)
    ' Insert while Before still identifies an existing item, including the last.
    mStages.Add value, , index
    mStages.Remove index + 1
End Sub

Private Function NxRenameJournalPathExists(ByVal candidate As String) As Boolean
    On Error GoTo Missing
    NxRenameJournalPathExists = (Len(Dir$(candidate, vbNormal Or vbHidden Or vbSystem Or vbReadOnly Or vbDirectory)) > 0)
    Exit Function
Missing:
    NxRaiseContractError "이름 변경 경로 존재 여부를 확인할 수 없습니다: " & candidate
End Function

Private Function NxRenameJournalJoinFailure(ByVal aggregate As String, ByVal detail As String) As String
    If Len(aggregate) > 0 Then aggregate = aggregate & "; "
    NxRenameJournalJoinFailure = aggregate & detail
End Function
