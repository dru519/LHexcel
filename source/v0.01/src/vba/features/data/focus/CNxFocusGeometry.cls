VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFocusGeometry"
Option Explicit

Private Const NX_SEGMENT_COUNT As Long = 8

Private mLeft(1 To NX_SEGMENT_COUNT) As Long
Private mTop(1 To NX_SEGMENT_COUNT) As Long
Private mWidth(1 To NX_SEGMENT_COUNT) As Long
Private mHeight(1 To NX_SEGMENT_COUNT) As Long
Private mVisible(1 To NX_SEGMENT_COUNT) As Boolean
Private mGridLeft As Long
Private mGridTop As Long
Private mGridRight As Long
Private mGridBottom As Long
Private mSelectionLeft As Long
Private mSelectionTop As Long
Private mSelectionRight As Long
Private mSelectionBottom As Long
Private mDpi As Long
Private mZoom As Long
Private mPaneIndex As Long
Private mSelectionSignature As String
Private mSealed As Boolean

Public Sub ConfigureGrid(ByVal leftPixel As Long, ByVal topPixel As Long, _
    ByVal rightPixel As Long, ByVal bottomPixel As Long, ByVal dpiValue As Long, _
    ByVal zoomValue As Long, ByVal paneIndex As Long)
    If mSealed Then NxRaiseContractError "포커스셀 좌표는 이미 확정되었습니다."
    If rightPixel <= leftPixel Or bottomPixel <= topPixel Then NxRaiseContractError "포커스셀 그리드 좌표가 올바르지 않습니다."
    mGridLeft = leftPixel
    mGridTop = topPixel
    mGridRight = rightPixel
    mGridBottom = bottomPixel
    mDpi = dpiValue
    mZoom = zoomValue
    mPaneIndex = paneIndex
End Sub

Public Sub ConfigureSelection(ByVal leftPixel As Long, ByVal topPixel As Long, _
    ByVal rightPixel As Long, ByVal bottomPixel As Long, ByVal signature As String)
    If mSealed Then NxRaiseContractError "포커스셀 좌표는 이미 확정되었습니다."
    If rightPixel <= leftPixel Or bottomPixel <= topPixel Then NxRaiseContractError "포커스셀 선택 좌표가 올바르지 않습니다."
    mSelectionLeft = leftPixel
    mSelectionTop = topPixel
    mSelectionRight = rightPixel
    mSelectionBottom = bottomPixel
    mSelectionSignature = signature
End Sub

Public Sub SetSegment(ByVal index As Long, ByVal leftPixel As Long, ByVal topPixel As Long, _
    ByVal rightPixel As Long, ByVal bottomPixel As Long)
    If mSealed Then NxRaiseContractError "포커스셀 좌표는 이미 확정되었습니다."
    If index < 1 Or index > NX_SEGMENT_COUNT Then NxRaiseContractError "포커스셀 구간 번호가 올바르지 않습니다."
    If rightPixel <= leftPixel Or bottomPixel <= topPixel Then Exit Sub
    mLeft(index) = leftPixel
    mTop(index) = topPixel
    mWidth(index) = rightPixel - leftPixel
    mHeight(index) = bottomPixel - topPixel
    mVisible(index) = True
End Sub

Public Sub Seal()
    If mGridRight <= mGridLeft Or mGridBottom <= mGridTop Then NxRaiseContractError "포커스셀 그리드 좌표가 없습니다."
    If mSelectionRight <= mSelectionLeft Or mSelectionBottom <= mSelectionTop Then NxRaiseContractError "포커스셀 선택 좌표가 없습니다."
    mSealed = True
End Sub

Private Sub RequireSegment(ByVal index As Long)
    If index < 1 Or index > NX_SEGMENT_COUNT Then NxRaiseContractError "포커스셀 구간 번호가 올바르지 않습니다."
End Sub

Public Property Get SegmentCount() As Long: SegmentCount = NX_SEGMENT_COUNT: End Property
Public Property Get SegmentVisible(ByVal index As Long) As Boolean: RequireSegment index: SegmentVisible = mVisible(index): End Property
Public Property Get SegmentLeft(ByVal index As Long) As Long: RequireSegment index: SegmentLeft = mLeft(index): End Property
Public Property Get SegmentTop(ByVal index As Long) As Long: RequireSegment index: SegmentTop = mTop(index): End Property
Public Property Get SegmentWidth(ByVal index As Long) As Long: RequireSegment index: SegmentWidth = mWidth(index): End Property
Public Property Get SegmentHeight(ByVal index As Long) As Long: RequireSegment index: SegmentHeight = mHeight(index): End Property
Public Property Get GridLeft() As Long: GridLeft = mGridLeft: End Property
Public Property Get GridTop() As Long: GridTop = mGridTop: End Property
Public Property Get GridRight() As Long: GridRight = mGridRight: End Property
Public Property Get GridBottom() As Long: GridBottom = mGridBottom: End Property
Public Property Get SelectionLeft() As Long: SelectionLeft = mSelectionLeft: End Property
Public Property Get SelectionTop() As Long: SelectionTop = mSelectionTop: End Property
Public Property Get SelectionRight() As Long: SelectionRight = mSelectionRight: End Property
Public Property Get SelectionBottom() As Long: SelectionBottom = mSelectionBottom: End Property
Public Property Get Dpi() As Long: Dpi = mDpi: End Property
Public Property Get Zoom() As Long: Zoom = mZoom: End Property
Public Property Get PaneIndex() As Long: PaneIndex = mPaneIndex: End Property
Public Property Get SelectionSignature() As String: SelectionSignature = mSelectionSignature: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property
