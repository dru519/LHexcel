VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDrawingRequest"
Option Explicit

Private mFeatureId As String, mTarget As Range, mPicture As Shape
Private mTotalRow As Boolean, mPreserveAlignment As Boolean, mIncludeHeader As Boolean, mAutoFitColumns As Boolean
Private mDisplayMode As String, mTableStyle As String
Private mHeaderRowCount As Long, mHeaderFontSize As Double, mHeaderAlignment As Long
Private mBodyFontSize As Double, mBodyAlignment As Long
Private mMargin As Double, mMoveAndSize As Boolean
Private mFingerprint As String, mTargetRangeDigest As String
Private mPictureWorkbookIdentity As String, mPictureSheetIdentity As String, mPictureName As String, mPictureObjectIdentity As String
Private mPictureLeft As Double, mPictureTop As Double, mPictureWidth As Double, mPictureHeight As Double
Private mPicturePlacement As XlPlacement, mPictureLockAspectRatio As MsoTriState
Private mConfigured As Boolean

Public Sub ConfigureRange(ByVal featureId As String, ByVal target As Range, ByVal totalRow As Boolean, _
    ByVal preserveAlignment As Boolean, Optional ByVal includeHeader As Boolean = True, _
    Optional ByVal autoFitColumns As Boolean = False, Optional ByVal displayMode As String = "MONO", _
    Optional ByVal tableStyle As String = "열린 표", Optional ByVal headerRowCount As Long = 1, _
    Optional ByVal headerFontSize As Double = 9#, Optional ByVal headerAlignment As Long = xlCenter, _
    Optional ByVal bodyFontSize As Double = 9#, Optional ByVal bodyAlignment As Long = xlGeneral)

    ValidateFeature featureId
    If target Is Nothing Or target.Areas.Count <> 1 Then NxRaiseContractError "표/선 기능은 연속된 한 범위가 필요합니다."
    If featureId = NX_FEATURE_DRAW_TITLE_TABLE Or featureId = NX_FEATURE_DRAW_BUSINESS_TABLE Then _
        ValidateTableOptions featureId, target, displayMode, tableStyle, headerRowCount, headerFontSize, headerAlignment, bodyFontSize, bodyAlignment
    Set mTarget = target: mFeatureId = featureId: mTotalRow = totalRow: mPreserveAlignment = preserveAlignment
    mIncludeHeader = includeHeader: mAutoFitColumns = autoFitColumns
    mDisplayMode = UCase$(displayMode): mTableStyle = tableStyle
    mHeaderRowCount = headerRowCount: mHeaderFontSize = headerFontSize: mHeaderAlignment = headerAlignment
    mBodyFontSize = bodyFontSize: mBodyAlignment = bodyAlignment
    mFingerprint = NxContextFactory.CaptureCurrentFingerprint(): mTargetRangeDigest = CaptureTargetRangeDigest(target): mConfigured = True
End Sub

Public Sub ConfigurePicture(ByVal pictureShape As Shape, ByVal target As Range, ByVal pictureMargin As Double, ByVal moveAndSize As Boolean)
    If pictureShape Is Nothing Or target Is Nothing Or target.Areas.Count <> 1 Or pictureMargin < 0 Then NxRaiseContractError "그림과 연속 대상 범위가 필요합니다."
    On Error GoTo InvalidPictureTarget
    If Not (pictureShape.Parent Is target.Worksheet) Then GoTo InvalidPictureTarget
    On Error GoTo 0
    Set mPicture = pictureShape: Set mTarget = target: mFeatureId = NX_FEATURE_DRAW_FIT_PICTURE
    mMargin = pictureMargin: mMoveAndSize = moveAndSize: mFingerprint = NxContextFactory.CaptureCurrentFingerprint()
    CapturePictureIdentity pictureShape
    mTargetRangeDigest = CaptureTargetRangeDigest(target)
    mConfigured = True
    Exit Sub
InvalidPictureTarget:
    NxRaiseContractError "그림과 대상 범위는 같은 시트에 있어야 합니다."
End Sub

Public Sub RequireContext(ByVal context As CNxExecutionContext)
    EnsureConfigured
    If context Is Nothing Or context.Fingerprint <> mFingerprint Then NxRaiseContractError "표/그리기 선택이 변경되었습니다."
    If context.ActualCellCount <> CDbl(mTarget.CountLarge) Then NxRaiseContractError "표/그리기 대상 크기가 변경되었습니다."
    If CaptureTargetRangeDigest(mTarget) <> mTargetRangeDigest Then NxRaiseContractError "표/그리기 대상 범위가 변경되었습니다."
    If mFeatureId = NX_FEATURE_DRAW_FIT_PICTURE Then ValidatePictureIdentity
End Sub

Public Property Get FeatureId() As String: EnsureConfigured: FeatureId = mFeatureId: End Property
Public Property Get Target() As Range: EnsureConfigured: Set Target = mTarget: End Property
Public Property Get Picture() As Shape: EnsureConfigured: Set Picture = mPicture: End Property
Public Property Get TotalRow() As Boolean: EnsureConfigured: TotalRow = mTotalRow: End Property
Public Property Get PreserveAlignment() As Boolean: EnsureConfigured: PreserveAlignment = mPreserveAlignment: End Property
Public Property Get IncludeHeader() As Boolean: EnsureConfigured: IncludeHeader = mIncludeHeader: End Property
Public Property Get AutoFitColumns() As Boolean: EnsureConfigured: AutoFitColumns = mAutoFitColumns: End Property
Public Property Get DisplayMode() As String: EnsureConfigured: DisplayMode = mDisplayMode: End Property
Public Property Get TableStyle() As String: EnsureConfigured: TableStyle = mTableStyle: End Property
Public Property Get HeaderRowCount() As Long: EnsureConfigured: HeaderRowCount = mHeaderRowCount: End Property
Public Property Get HeaderFontSize() As Double: EnsureConfigured: HeaderFontSize = mHeaderFontSize: End Property
Public Property Get HeaderAlignment() As Long: EnsureConfigured: HeaderAlignment = mHeaderAlignment: End Property
Public Property Get BodyFontSize() As Double: EnsureConfigured: BodyFontSize = mBodyFontSize: End Property
Public Property Get BodyAlignment() As Long: EnsureConfigured: BodyAlignment = mBodyAlignment: End Property
Public Property Get Margin() As Double: EnsureConfigured: Margin = mMargin: End Property
Public Property Get MoveAndSize() As Boolean: EnsureConfigured: MoveAndSize = mMoveAndSize: End Property
Public Property Get ActualCellCount() As Long: EnsureConfigured: ActualCellCount = CLng(CDbl(mTarget.CountLarge)): End Property
Public Property Get TargetRangeDigest() As String: EnsureConfigured: TargetRangeDigest = mTargetRangeDigest: End Property
Public Property Get PictureWorkbookIdentity() As String: EnsureConfigured: PictureWorkbookIdentity = mPictureWorkbookIdentity: End Property
Public Property Get PictureSheetIdentity() As String: EnsureConfigured: PictureSheetIdentity = mPictureSheetIdentity: End Property
Public Property Get PictureName() As String: EnsureConfigured: PictureName = mPictureName: End Property
Public Property Get PictureObjectIdentity() As String: EnsureConfigured: PictureObjectIdentity = mPictureObjectIdentity: End Property
Public Property Get PictureLeft() As Double: EnsureConfigured: PictureLeft = mPictureLeft: End Property
Public Property Get PictureTop() As Double: EnsureConfigured: PictureTop = mPictureTop: End Property
Public Property Get PictureWidth() As Double: EnsureConfigured: PictureWidth = mPictureWidth: End Property
Public Property Get PictureHeight() As Double: EnsureConfigured: PictureHeight = mPictureHeight: End Property
Public Property Get PicturePlacement() As XlPlacement: EnsureConfigured: PicturePlacement = mPicturePlacement: End Property
Public Property Get PictureLockAspectRatio() As MsoTriState: EnsureConfigured: PictureLockAspectRatio = mPictureLockAspectRatio: End Property
Public Function PictureMutationTarget(ByVal selectionIdentity As String) As String
    EnsureConfigured
    If mFeatureId <> NX_FEATURE_DRAW_FIT_PICTURE Then NxRaiseContractError "Picture mutation target is FIT-PICTURE only"
    PictureMutationTarget = selectionIdentity & "|PW=" & mPictureWorkbookIdentity & "|PS=" & mPictureSheetIdentity & "|PN=" & mPictureName & "|PO=" & mPictureObjectIdentity
End Function

Private Sub ValidateFeature(ByVal featureId As String)
    Select Case featureId
        Case NX_FEATURE_DRAW_TITLE_TABLE, NX_FEATURE_DRAW_BUSINESS_TABLE, NX_FEATURE_DRAW_CLEAR_INNER
        Case Else: NxRaiseContractError "표/그리기 기능이 닫힌 계약에 없습니다."
    End Select
End Sub

Private Sub ValidateTableOptions(ByVal featureId As String, ByVal target As Range, ByVal displayMode As String, _
    ByVal tableStyle As String, ByVal headerRowCount As Long, ByVal headerFontSize As Double, _
    ByVal headerAlignment As Long, ByVal bodyFontSize As Double, ByVal bodyAlignment As Long)

    Select Case UCase$(displayMode)
        Case NX_ROLE_STYLE_MODE_MONO, NX_ROLE_STYLE_MODE_COLOR
        Case Else: NxRaiseContractError "표시 방식은 단색 또는 색상이어야 합니다."
    End Select
    If headerRowCount < 0 Or headerRowCount > 5 Or headerRowCount > target.Rows.Count Then _
        NxRaiseContractError "머리글 행 수는 선택 범위 안에서 0~5여야 합니다."
    If headerFontSize < 8 Or headerFontSize > 72 Or bodyFontSize < 8 Or bodyFontSize > 72 Then _
        NxRaiseContractError "표 글꼴 크기는 8~72pt여야 합니다."
    Select Case headerAlignment
        Case xlLeft, xlCenter, xlRight
        Case Else: NxRaiseContractError "머리글 정렬 값이 올바르지 않습니다."
    End Select
    Select Case bodyAlignment
        Case xlGeneral, xlLeft, xlCenter, xlRight
        Case Else: NxRaiseContractError "본문 정렬 값이 올바르지 않습니다."
    End Select
    If featureId = NX_FEATURE_DRAW_TITLE_TABLE Then
        If tableStyle <> "제목" And tableStyle <> "부제목" Then NxRaiseContractError "제목 표 스타일이 올바르지 않습니다."
    ElseIf tableStyle <> "닫힌 표" And tableStyle <> "열린 표" Then
        NxRaiseContractError "실무 표 스타일이 올바르지 않습니다."
    End If
End Sub

Private Sub CapturePictureIdentity(ByVal pictureShape As Shape)
    Dim parentSheet As Object, parentBook As Object
    On Error GoTo Failed
    Set parentSheet = pictureShape.Parent: Set parentBook = parentSheet.Parent
    mPictureWorkbookIdentity = CStr(parentBook.FullName)
    mPictureSheetIdentity = CStr(parentSheet.CodeName)
    mPictureName = CStr(pictureShape.Name)
    mPictureObjectIdentity = CStr(ObjPtr(pictureShape))
    mPictureLeft = pictureShape.Left: mPictureTop = pictureShape.Top: mPictureWidth = pictureShape.Width: mPictureHeight = pictureShape.Height
    mPicturePlacement = pictureShape.Placement: mPictureLockAspectRatio = pictureShape.LockAspectRatio
    Exit Sub
Failed:
    NxRaiseContractError "그림의 통합문서·시트·이름·객체 식별자를 봉인할 수 없습니다."
End Sub

Private Sub ValidatePictureIdentity()
    Dim currentSheet As Object, currentBook As Object
    On Error GoTo Failed
    Set currentSheet = mPicture.Parent: Set currentBook = currentSheet.Parent
    If CStr(currentBook.FullName) <> mPictureWorkbookIdentity Or CStr(currentSheet.CodeName) <> mPictureSheetIdentity Then GoTo Failed
    If CStr(mPicture.Name) <> mPictureName Or CStr(ObjPtr(mPicture)) <> mPictureObjectIdentity Then GoTo Failed
    If pictureShapeChanged(mPicture.Left, mPictureLeft) Or pictureShapeChanged(mPicture.Top, mPictureTop) Or pictureShapeChanged(mPicture.Width, mPictureWidth) Or pictureShapeChanged(mPicture.Height, mPictureHeight) Then GoTo Failed
    If mPicture.Placement <> mPicturePlacement Or mPicture.LockAspectRatio <> mPictureLockAspectRatio Then GoTo Failed
    Exit Sub
Failed:
    NxRaiseContractError "그림 선택 또는 그림 속성이 실행 전에 변경되었습니다."
End Sub

Private Function pictureShapeChanged(ByVal currentValue As Double, ByVal sealedValue As Double) As Boolean
    pictureShapeChanged = (Abs(currentValue - sealedValue) > 0.0001)
End Function

Private Function CaptureTargetRangeDigest(ByVal target As Range) As String
    CaptureTargetRangeDigest = target.Worksheet.Parent.FullName & "|" & target.Worksheet.CodeName & "|" & target.Address(True, True, xlA1, True) & "|" & CStr(target.CountLarge) & "|" & CStr(target.Width) & "|" & CStr(target.Height)
End Function

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "표/그리기 요청이 구성되지 않았습니다."
End Sub
