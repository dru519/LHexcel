Attribute VB_Name = "NxPictureInsertEngine"
Option Explicit

Public Function NxPictureInsertBuildPreview(ByVal request As CNxPictureInsertRequest) As CNxPictureInsertPreview
    Dim target As Range, cell As Range, area As Range, addresses As New Collection, names As New Collection, skipped As New Collection
    Dim seen As Object, reservedNames As Object, files As Collection
    Dim insertCount As Long, visibleCount As Long, ordinal As Long, key As String, shapeName As String
    If request Is Nothing Or Not request.IsConfigured Then NxRaiseContractError "그림 삽입 preview에는 유효한 요청이 필요합니다. contiguous 대상만 허용합니다."
    request.Revalidate: Set target = request.Target: Set files = request.Files
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1
    Set reservedNames = PictureShapeNameReservations(target.Worksheet)
    For Each cell In target.Cells
        Set area = cell.MergeArea
        If Not (target Is area) Then
            If Intersect(target, area).Address <> area.Address Then NxRaiseContractError "partial merged area는 그림 삽입 대상이 될 수 없습니다."
        End If
        key = area.Address(External:=True)
        If Not seen.Exists(key) Then
            seen.Add key, True
            If PictureTargetIsUsable(area) Then
                addresses.Add key: visibleCount = visibleCount + 1
            Else
                skipped.Add key & "|hidden_or_zero"
            End If
        End If
    Next cell
    If target.Parent Is Nothing Then NxRaiseContractError "그림 삽입 대상 시트를 확인할 수 없습니다."
    If target.Parent.CodeName <> ActiveSheet.CodeName Then NxRaiseContractError "그림 삽입 대상은 same worksheet여야 합니다."
    If visibleCount = 0 Then NxRaiseContractError "보이는 그림 삽입 대상이 없습니다."
    insertCount = files.Count
    If insertCount > visibleCount Then insertCount = visibleCount
    For ordinal = 1 To insertCount
        shapeName = NextAvailablePictureShapeName(reservedNames, ordinal)
        names.Add shapeName
    Next ordinal
    Dim preview As New CNxPictureInsertPreview
    preview.Configure request.TargetDigest, PictureFileFingerprint(files), request.SizeMode, target.Address(External:=True), addresses, names, skipped, visibleCount, insertCount
    Set NxPictureInsertBuildPreview = preview
End Function

Public Function NxPictureInsertExecute(ByVal request As CNxPictureInsertRequest, ByVal preview As CNxPictureInsertPreview, ByVal journal As CNxPictureInsertJournal) As CNxResult
    Dim files As Collection, addresses As Collection, names As Collection, index As Long
    Dim picture As Shape, target As Range, targetArea As Range
    Dim result As New CNxResult
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    If request Is Nothing Or preview Is Nothing Or journal Is Nothing Then NxRaiseContractError "그림 삽입 실행 구성이 잘못되었습니다."
    If Not preview.Revalidate(request) Then NxRaiseContractError "그림 삽입 preview drift가 발생했습니다."
    Set files = request.Files: Set addresses = preview.TargetAddresses: Set names = preview.ShapeNames: Set target = request.Target
    EnsurePictureShapeNamesAvailable target.Worksheet, names
    journal.Bind target
    On Error GoTo Failed
    For index = 1 To preview.InsertCount
        Set targetArea = target.Worksheet.Range(CStr(addresses(index)))
        Set picture = Nothing
        NxPictureInsertRequireSupportedHeader CStr(files(index))
        Set picture = target.Worksheet.Shapes.AddPicture(FileName:=CStr(files(index)), LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
            Left:=targetArea.Left, Top:=targetArea.Top, Width:=-1, Height:=-1)
        picture.Name = CStr(names(index))
        journal.Capture picture
        If request.SizeMode = "FIT" Then
            NxDrawFitPicture picture, targetArea, 0#, True
        Else
            picture.LockAspectRatio = msoTrue
            picture.Left = targetArea.Left: picture.Top = targetArea.Top
            picture.Placement = xlMove
        End If
    Next index
    result.Configure "NX-DRAW-INSERT-PICTURE", NxSuccess, "complete", target.Parent.Parent.FullName, True, vbNullString, "picture_inserted"
    result.AddCompleted "그림 " & CStr(preview.InsertCount) & "개 삽입 완료"
    result.Seal
    Set NxPictureInsertExecute = result
    Exit Function
Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    journal.Rollback
    On Error Resume Next
    If Not picture Is Nothing Then picture.Delete
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Function

Public Sub NxPictureInsertRequireSupportedHeader(ByVal path As String)
    Dim handle As Integer, opened As Boolean, bytes() As Byte, index As Long
    Dim fileLength As Long, readCount As Long, signature As String, recognized As Boolean
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    On Error GoTo Failed
    handle = FreeFile
    Open path For Binary Access Read Lock Write As #handle
    opened = True
    fileLength = LOF(handle)
    If fileLength < 4 Then NxRaiseContractError "지원하지 않거나 손상된 그림 파일 헤더입니다."
    readCount = fileLength
    If readCount > 16 Then readCount = 16
    ReDim bytes(0 To readCount - 1)
    Get #handle, , bytes
    Close #handle
    opened = False
    For index = 0 To readCount - 1
        signature = signature & Right$("0" & Hex$(bytes(index)), 2)
    Next index
    ' Signature/file-header sanity only, not full image decoding; Excel remains the decoder.
    ' Format references are recorded in docs/r62-picture-header-references.md.
    If fileLength >= 8 And Left$(signature, 16) = "89504E470D0A1A0A" Then recognized = True
    If fileLength >= 4 And Left$(signature, 6) = "FFD8FF" Then recognized = True
    If fileLength >= 6 And Left$(signature, 12) = "474946383761" Then recognized = True
    If fileLength >= 6 And Left$(signature, 12) = "474946383961" Then recognized = True
    If fileLength >= 14 And Left$(signature, 4) = "424D" Then recognized = True
    If fileLength >= 8 And Left$(signature, 8) = "49492A00" Then recognized = True
    If fileLength >= 8 And Left$(signature, 8) = "4D4D002A" Then recognized = True
    If fileLength >= 16 And Left$(signature, 16) = "49492B0008000000" Then recognized = True
    If fileLength >= 16 And Left$(signature, 16) = "4D4D002B00080000" Then recognized = True
    If Not recognized Then NxRaiseContractError "지원하지 않거나 손상된 그림 파일 헤더입니다."
    Exit Sub
Failed:
    failureNumber = Err.Number: failureSource = Err.Source: failureDescription = Err.Description
    On Error Resume Next
    If opened Then Close #handle
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Sub

Private Function PictureTargetIsUsable(ByVal targetArea As Range) As Boolean
    Dim visibleCell As Range
    If targetArea Is Nothing Then Exit Function
    If targetArea.Width <= 0 Or targetArea.Height <= 0 Then Exit Function
    For Each visibleCell In targetArea.Cells
        If visibleCell.EntireRow.Hidden Or visibleCell.EntireColumn.Hidden Then Exit Function
    Next visibleCell
    PictureTargetIsUsable = True
End Function

Private Function PictureShapeNameReservations(ByVal targetSheet As Worksheet) As Object
    Dim reservedNames As Object, existingShape As Shape
    Set reservedNames = CreateObject("Scripting.Dictionary")
    reservedNames.CompareMode = 1
    For Each existingShape In targetSheet.Shapes
        If Not reservedNames.Exists(CStr(existingShape.Name)) Then reservedNames.Add CStr(existingShape.Name), True
    Next existingShape
    Set PictureShapeNameReservations = reservedNames
End Function

Private Function NextAvailablePictureShapeName(ByVal reservedNames As Object, ByVal ordinal As Long) As String
    Dim baseName As String, candidate As String, suffix As Long
    baseName = "NxPictureInsert_" & Format$(ordinal, "000")
    candidate = baseName
    Do While reservedNames.Exists(candidate)
        suffix = suffix + 1
        candidate = baseName & "_" & Format$(suffix, "000")
    Loop
    reservedNames.Add candidate, True
    NextAvailablePictureShapeName = candidate
End Function

Private Sub EnsurePictureShapeNamesAvailable(ByVal targetSheet As Worksheet, ByVal names As Collection)
    Dim reservedNames As Object, item As Variant
    Set reservedNames = PictureShapeNameReservations(targetSheet)
    For Each item In names
        If reservedNames.Exists(CStr(item)) Then NxRaiseContractError "그림 삽입 preview 이후 도형 이름이 충돌했습니다."
        reservedNames.Add CStr(item), True
    Next item
End Sub

Private Function PictureFileFingerprint(ByVal files As Collection) As String
    Dim item As Variant
    For Each item In files: PictureFileFingerprint = PictureFileFingerprint & CStr(item) & "|" & CStr(FileLen(CStr(item))) & "|" & CStr(FileDateTime(CStr(item))) & ";": Next item
End Function
