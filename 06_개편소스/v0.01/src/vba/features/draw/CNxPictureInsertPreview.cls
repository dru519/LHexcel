VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxPictureInsertPreview"
Option Explicit

Private mTargetDigest As String, mTargetAddress As String, mFileFingerprint As String, mSizeMode As String
Private mTargetAddresses As Collection, mShapeNames As Collection, mSkipped As Collection
Private mVisibleTargetCount As Long, mInsertCount As Long, mSealed As Boolean

Public Sub Configure(ByVal targetDigest As String, ByVal fileFingerprint As String, ByVal sizeMode As String, _
    ByVal targetAddress As String, ByVal targetAddresses As Collection, ByVal shapeNames As Collection, ByVal skipped As Collection, _
    ByVal visibleTargetCount As Long, ByVal insertCount As Long)
    If mSealed Then NxRaiseContractError "그림 삽입 미리보기는 변경할 수 없습니다."
    If Len(targetDigest) = 0 Or Len(targetAddress) = 0 Or Len(fileFingerprint) = 0 Then NxRaiseContractError "그림 삽입 미리보기 식별자가 필요합니다."
    If targetAddresses Is Nothing Or shapeNames Is Nothing Or skipped Is Nothing Then NxRaiseContractError "그림 삽입 미리보기 목록이 필요합니다."
    If insertCount < 1 Or insertCount > 500 Or insertCount > targetAddresses.Count Then NxRaiseContractError "그림 삽입 미리보기 개수가 잘못되었습니다."
    mTargetDigest = targetDigest: mTargetAddress = targetAddress: mFileFingerprint = fileFingerprint: mSizeMode = sizeMode
    Set mTargetAddresses = CopyCollection(targetAddresses): Set mShapeNames = CopyCollection(shapeNames): Set mSkipped = CopyCollection(skipped)
    mVisibleTargetCount = visibleTargetCount: mInsertCount = insertCount: mSealed = True
End Sub

Public Function Revalidate(ByVal request As CNxPictureInsertRequest) As Boolean
    If Not mSealed Or request Is Nothing Then Exit Function
    request.Revalidate
    If Len(request.SourceIdentityDigest) = 0 Then NxRaiseContractError "그림 삽입 원본 identity가 봉인되지 않았습니다."
    Revalidate = (request.TargetDigest = mTargetDigest)
End Function

Public Property Get TargetDigest() As String: TargetDigest = mTargetDigest: End Property
Public Property Get TargetAddress() As String: TargetAddress = mTargetAddress: End Property
Public Property Get FileFingerprint() As String: FileFingerprint = mFileFingerprint: End Property
Public Property Get SizeMode() As String: SizeMode = mSizeMode: End Property
Public Property Get TargetAddresses() As Collection: Set TargetAddresses = CopyCollection(mTargetAddresses): End Property
Public Property Get ShapeNames() As Collection: Set ShapeNames = CopyCollection(mShapeNames): End Property
Public Property Get Skipped() As Collection: Set Skipped = CopyCollection(mSkipped): End Property
Public Property Get VisibleTargetCount() As Long: VisibleTargetCount = mVisibleTargetCount: End Property
Public Property Get InsertCount() As Long: InsertCount = mInsertCount: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property

Public Function Summary() As String
    Summary = "visible=" & CStr(mVisibleTargetCount) & "; insert=" & CStr(mInsertCount) & "; skip=" & CStr(mSkipped.Count) & "; mode=" & mSizeMode
End Function

Private Function CopyCollection(ByVal source As Collection) As Collection
    Dim answer As New Collection, item As Variant
    For Each item In source: answer.Add CStr(item): Next item
    Set CopyCollection = answer
End Function
