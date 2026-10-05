VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxPictureInsertRequest"
Option Explicit

Private mTarget As Range, mFiles As Collection, mSourceIdentities As Collection, mSizeMode As String
Private mTargetDigest As String, mSourceIdentityDigest As String, mConfigured As Boolean

Public Sub Configure(ByVal target As Range, ByVal filePaths As Collection, ByVal sizeMode As String)
    If mConfigured Then NxRaiseContractError "그림 삽입 요청은 한 번만 구성할 수 있습니다."
    ValidateTarget target
    mSizeMode = NormalizeSizeMode(sizeMode)
    Set mTarget = target: Set mFiles = CanonicalizeAndSortFiles(filePaths)
    Set mSourceIdentities = BuildSourceIdentities(mFiles)
    mSourceIdentityDigest = BuildSourceIdentityDigest(mSourceIdentities)
    mTargetDigest = BuildDigest(mSourceIdentities)
    mConfigured = True
End Sub

Public Sub Revalidate()
    EnsureConfigured
    ValidateTarget mTarget
    ValidateFiles mFiles
    Dim currentFiles As Collection, currentIdentities As Collection
    Set currentFiles = CanonicalizeAndSortFiles(mFiles)
    If Not PictureFileCollectionsMatch(currentFiles, mFiles) Then NxRaiseContractError "그림 삽입 미리보기 이후 파일 순서 또는 경로가 변경되었습니다."
    Set currentIdentities = BuildSourceIdentities(currentFiles)
    If Not SourceIdentitiesMatch(currentIdentities) Then NxRaiseContractError "그림 삽입 미리보기 이후 원본 파일이 변경되었습니다."
    If BuildDigest(currentIdentities) <> mTargetDigest Then NxRaiseContractError "그림 삽입 미리보기 이후 대상 또는 파일이 변경되었습니다."
End Sub

Public Property Get Target() As Range: EnsureConfigured: Set Target = mTarget: End Property
Public Property Get Files() As Collection: EnsureConfigured: Set Files = CopyFiles(mFiles): End Property
Public Property Get SizeMode() As String: EnsureConfigured: SizeMode = mSizeMode: End Property
Public Property Get TargetDigest() As String: EnsureConfigured: TargetDigest = mTargetDigest: End Property
Public Property Get SourceIdentityDigest() As String: EnsureConfigured: SourceIdentityDigest = mSourceIdentityDigest: End Property
Public Property Get SourceIdentities() As Collection: EnsureConfigured: Set SourceIdentities = CopyFiles(mSourceIdentities): End Property
Public Property Get IsConfigured() As Boolean: IsConfigured = mConfigured: End Property

Private Sub ValidateTarget(ByVal target As Range)
    If target Is Nothing Then NxRaiseContractError "그림 삽입 대상 범위를 선택해야 합니다."
    If target.Parent Is Nothing Then NxRaiseContractError "그림 삽입 대상 시트를 확인할 수 없습니다."
    If Not target.Parent Is ActiveSheet Then NxRaiseContractError "현재 시트의 범위만 그림 삽입 대상이 될 수 있습니다."
    If target.Areas.Count <> 1 Then NxRaiseContractError "그림 삽입 대상은 연속된 한 범위여야 합니다."
    If target.Worksheet.ProtectContents Or target.Worksheet.ProtectDrawingObjects Then NxRaiseContractError "보호된 시트에는 그림을 삽입할 수 없습니다."
    If target.CountLarge > 500 Then NxRaiseContractError "그림 삽입 대상은 500개를 초과할 수 없습니다."
End Sub

Private Sub ValidateFiles(ByVal filePaths As Collection)
    Dim ignored As Collection
    Set ignored = CanonicalizeAndSortFiles(filePaths)
End Sub

Private Function AllowedExtension(ByVal path As String) As Boolean
    Dim extension As String
    extension = LCase$(Mid$(path, InStrRev(path, ".") + 1))
    AllowedExtension = (extension = "png" Or extension = "jpg" Or extension = "jpeg" Or extension = "gif" Or extension = "bmp" Or extension = "tif" Or extension = "tiff")
End Function

Private Function NormalizeSizeMode(ByVal value As String) As String
    Select Case UCase$(Trim$(value))
        Case "FIT": NormalizeSizeMode = "FIT"
        Case "ORIGINAL": NormalizeSizeMode = "ORIGINAL"
        Case Else: NxRaiseContractError "그림 크기 모드는 FIT 또는 ORIGINAL이어야 합니다."
    End Select
End Function

Private Function CopyFiles(ByVal source As Collection) As Collection
    Dim answer As New Collection, item As Variant
    If source Is Nothing Then NxRaiseContractError "그림 파일 선택이 없습니다."
    For Each item In source: answer.Add CStr(item): Next item
    Set CopyFiles = answer
End Function

Private Function CanonicalizeAndSortFiles(ByVal source As Collection) As Collection
    Dim values() As String, index As Long, scan As Long, current As String
    Dim path As String, key As String, keys As Object, answer As New Collection
    If source Is Nothing Or source.Count < 1 Or source.Count > 500 Then NxRaiseContractError "그림 파일은 1개 이상 500개 이하로 선택해야 합니다."
    ReDim values(1 To source.Count)
    Set keys = CreateObject("Scripting.Dictionary"): keys.CompareMode = 0
    For index = 1 To source.Count
        path = Trim$(CStr(source(index)))
        If Not IsAbsolutePicturePath(path) Then NxRaiseContractError "그림 파일은 절대 경로여야 합니다."
        If HasDotPathSegment(path) Then NxRaiseContractError "그림 파일 경로에 상대 경로 구간이 포함되어 있습니다."
        path = NxPictureInsertCanonicalPath(path)
        If Len(Dir$(path, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) = 0 Then NxRaiseContractError "그림 파일을 찾을 수 없습니다."
        If (GetAttr(path) And vbDirectory) <> 0 Then NxRaiseContractError "그림 경로는 일반 파일이어야 합니다."
        If Not AllowedExtension(path) Then NxRaiseContractError "허용되지 않은 그림 확장자입니다."
        If FileLen(path) <= 0 Then NxRaiseContractError "빈 그림 파일은 삽입할 수 없습니다."
        key = NxPictureInsertCanonicalPathKey(path)
        If keys.Exists(key) Then NxRaiseContractError "중복되거나 case/Unicode alias인 그림 파일이 있습니다."
        keys.Add key, True
        values(index) = path
    Next index
    For index = 2 To UBound(values)
        current = values(index): scan = index - 1
        Do While scan >= 1
            If ComparePicturePaths(values(scan), current) <= 0 Then Exit Do
            values(scan + 1) = values(scan): scan = scan - 1
        Loop
        values(scan + 1) = current
    Next index
    For index = 1 To UBound(values): answer.Add values(index): Next index
    Set CanonicalizeAndSortFiles = answer
End Function

Private Function PictureFileCollectionsMatch(ByVal first As Collection, ByVal second As Collection) As Boolean
    Dim index As Long
    If first Is Nothing Or second Is Nothing Then Exit Function
    If first.Count <> second.Count Then Exit Function
    For index = 1 To first.Count
        If StrComp(CStr(first(index)), CStr(second(index)), vbBinaryCompare) <> 0 Then Exit Function
    Next index
    PictureFileCollectionsMatch = True
End Function

Private Function IsAbsolutePicturePath(ByVal value As String) As Boolean
    IsAbsolutePicturePath = (Left$(value, 2) = "\\" Or (Len(value) >= 3 And Mid$(value, 2, 2) = ":\"))
End Function

Private Function HasDotPathSegment(ByVal value As String) As Boolean
    Dim part As Variant
    For Each part In Split(Replace(value, "/", "\"), "\")
        If CStr(part) = "." Or CStr(part) = ".." Then HasDotPathSegment = True: Exit Function
    Next part
End Function

Private Function ComparePicturePaths(ByVal firstPath As String, ByVal secondPath As String) As Long
    Dim firstKey As String, secondKey As String, firstName As String, secondName As String
    firstKey = NxPictureInsertCanonicalPathKey(firstPath): secondKey = NxPictureInsertCanonicalPathKey(secondPath)
    firstName = Mid$(firstKey, InStrRev(firstKey, "\") + 1): secondName = Mid$(secondKey, InStrRev(secondKey, "\") + 1)
    ComparePicturePaths = CompareNaturalText(firstName, secondName)
    If ComparePicturePaths = 0 Then ComparePicturePaths = StrComp(firstKey, secondKey, vbBinaryCompare)
    If ComparePicturePaths = 0 Then ComparePicturePaths = StrComp(firstPath, secondPath, vbBinaryCompare)
End Function

Private Function CompareNaturalText(ByVal firstValue As String, ByVal secondValue As String) As Long
    Dim firstPosition As Long, secondPosition As Long, firstChunk As String, secondChunk As String
    Dim firstNumeric As Boolean, secondNumeric As Boolean, compared As Long
    firstPosition = 1: secondPosition = 1
    Do While firstPosition <= Len(firstValue) And secondPosition <= Len(secondValue)
        firstChunk = NextNaturalChunk(firstValue, firstPosition, firstNumeric)
        secondChunk = NextNaturalChunk(secondValue, secondPosition, secondNumeric)
        If firstNumeric And secondNumeric Then
            compared = CompareNumberChunks(firstChunk, secondChunk)
        Else
            compared = StrComp(firstChunk, secondChunk, vbBinaryCompare)
        End If
        If compared <> 0 Then CompareNaturalText = compared: Exit Function
    Loop
    If firstPosition <= Len(firstValue) Then
        CompareNaturalText = 1
    ElseIf secondPosition <= Len(secondValue) Then
        CompareNaturalText = -1
    End If
End Function

Private Function NextNaturalChunk(ByVal value As String, ByRef position As Long, ByRef numericChunk As Boolean) As String
    Dim startPosition As Long, digit As Boolean
    startPosition = position: digit = IsAsciiDigit(Mid$(value, position, 1)): numericChunk = digit
    Do While position <= Len(value)
        If IsAsciiDigit(Mid$(value, position, 1)) <> digit Then Exit Do
        position = position + 1
    Loop
    NextNaturalChunk = Mid$(value, startPosition, position - startPosition)
End Function

Private Function IsAsciiDigit(ByVal value As String) As Boolean
    IsAsciiDigit = (Len(value) = 1 And InStr(1, "0123456789", value, vbBinaryCompare) > 0)
End Function

Private Function CompareNumberChunks(ByVal firstValue As String, ByVal secondValue As String) As Long
    Dim firstTrimmed As String, secondTrimmed As String
    firstTrimmed = TrimLeadingZeros(firstValue): secondTrimmed = TrimLeadingZeros(secondValue)
    If Len(firstTrimmed) <> Len(secondTrimmed) Then CompareNumberChunks = Sgn(Len(firstTrimmed) - Len(secondTrimmed)): Exit Function
    If StrComp(firstTrimmed, secondTrimmed, vbBinaryCompare) <> 0 Then CompareNumberChunks = Sgn(StrComp(firstTrimmed, secondTrimmed, vbBinaryCompare)): Exit Function
    CompareNumberChunks = Sgn(Len(firstValue) - Len(secondValue))
End Function

Private Function TrimLeadingZeros(ByVal value As String) As String
    Dim position As Long: position = 1
    Do While position < Len(value) And Mid$(value, position, 1) = "0": position = position + 1: Loop
    TrimLeadingZeros = Mid$(value, position)
End Function

Private Function BuildDigest(ByVal identities As Collection) As String
    Dim item As Variant, digest As String
    digest = CStr(mTarget.Parent.Parent.FullName) & "|" & CStr(mTarget.Parent.Name) & "|" & CStr(mTarget.Parent.CodeName) & "|" & mTarget.Address(External:=True) & "|" & mSizeMode
    digest = digest & "|" & BuildTargetTopologyDigest(mTarget)
    For Each item In identities
        digest = digest & "|" & CStr(item)
    Next item
    BuildDigest = digest
End Function

Private Function BuildTargetTopologyDigest(ByVal target As Range) As String
    Dim cell As Range, mergeArea As Range, digest As String
    Dim targetSheet As Worksheet
    Set targetSheet = target.Worksheet
    digest = "topology|areas=" & CanonicalTopologyNumber(CDbl(target.Areas.Count)) & "|cells=" & CanonicalTopologyNumber(CDbl(target.CountLarge)) & _
        "|protected=" & CanonicalTopologyFlag(targetSheet.ProtectContents) & ":" & CanonicalTopologyFlag(targetSheet.ProtectDrawingObjects) & ":" & CanonicalTopologyFlag(targetSheet.ProtectScenarios) & _
        "|filter=" & CanonicalTopologyFlag(targetSheet.FilterMode) & ":" & CanonicalTopologyFlag(targetSheet.AutoFilterMode) & _
        "|rect=" & CanonicalTopologyNumber(target.Left) & ":" & CanonicalTopologyNumber(target.Top) & ":" & CanonicalTopologyNumber(target.Width) & ":" & CanonicalTopologyNumber(target.Height)
    For Each cell In target.Cells
        If cell.MergeCells Then
            Set mergeArea = cell.MergeArea
        Else
            Set mergeArea = cell
        End If
        digest = digest & "|cell=" & cell.Address(False, False) & _
            "|merge=" & mergeArea.Address(False, False) & ":" & CanonicalTopologyFlag(cell.MergeCells) & _
            "|rowHidden=" & CanonicalTopologyFlag(cell.EntireRow.Hidden) & "|columnHidden=" & CanonicalTopologyFlag(cell.EntireColumn.Hidden) & _
            "|rowHeight=" & CanonicalTopologyNumber(cell.EntireRow.RowHeight) & "|columnWidth=" & CanonicalTopologyNumber(cell.EntireColumn.ColumnWidth)
    Next cell
    BuildTargetTopologyDigest = digest
End Function

Private Function CanonicalTopologyNumber(ByVal value As Double) As String
    CanonicalTopologyNumber = Trim$(Str$(value))
End Function

Private Function CanonicalTopologyFlag(ByVal value As Boolean) As String
    If value Then CanonicalTopologyFlag = "1" Else CanonicalTopologyFlag = "0"
End Function

Private Function BuildSourceIdentities(ByVal files As Collection) As Collection
    Dim answer As New Collection, item As Variant
    For Each item In files
        answer.Add NxPictureInsertBuildSourceIdentity(CStr(item))
    Next item
    Set BuildSourceIdentities = answer
End Function

Private Function BuildSourceIdentityDigest(ByVal identities As Collection) As String
    Dim item As Variant, canonical As String
    For Each item In identities: canonical = canonical & CStr(item) & vbLf: Next item
    BuildSourceIdentityDigest = canonical
End Function

Private Function SourceIdentitiesMatch(ByVal currentIdentities As Collection) As Boolean
    Dim index As Long
    If currentIdentities Is Nothing Or mSourceIdentities Is Nothing Then Exit Function
    If currentIdentities.Count <> mSourceIdentities.Count Then Exit Function
    For index = 1 To mSourceIdentities.Count
        If CStr(currentIdentities(index)) <> CStr(mSourceIdentities(index)) Then Exit Function
    Next index
    SourceIdentitiesMatch = True
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "그림 삽입 요청이 구성되지 않았습니다."
End Sub
