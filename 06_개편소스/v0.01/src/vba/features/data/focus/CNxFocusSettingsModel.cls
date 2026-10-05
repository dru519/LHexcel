VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFocusSettingsModel"
Option Explicit

Private mShape As String
Private mStyle As String
Private mColorHex As String
Private mIntensityPercent As Long
Private mHighlightSelectedArea As Boolean
Private mEnableOnApply As Boolean
Private mSealed As Boolean

Public Sub Configure(ByVal shape As String, ByVal style As String, ByVal colorHex As String, _
                     ByVal intensityPercent As Long, ByVal highlightSelectedArea As Boolean, _
                     ByVal enableOnApply As Boolean)
    If mSealed Then NxRaiseContractError "포커스셀 설정은 이미 확정되었습니다."
    shape = NxFocusNormalizeShape(shape)
    style = NxFocusNormalizeStyle(style)
    colorHex = NxFocusNormalizeHexColor(colorHex)
    If Len(shape) = 0 Or Len(style) = 0 Or Len(colorHex) = 0 Then _
        NxRaiseContractError "포커스셀 설정값을 확인하세요."
    If intensityPercent < 0 Or intensityPercent > 100 Then _
        NxRaiseContractError "강조 농도는 0부터 100 사이여야 합니다."
    mShape = shape
    mStyle = style
    mColorHex = colorHex
    mIntensityPercent = intensityPercent
    mHighlightSelectedArea = highlightSelectedArea
    mEnableOnApply = enableOnApply
    mSealed = True
End Sub

Public Property Get Shape() As String: Shape = mShape: End Property
Public Property Get Style() As String: Style = mStyle: End Property
Public Property Get ColorHex() As String: ColorHex = mColorHex: End Property
Public Property Get IntensityPercent() As Long: IntensityPercent = mIntensityPercent: End Property
Public Property Get HighlightSelectedArea() As Boolean: HighlightSelectedArea = mHighlightSelectedArea: End Property
Public Property Get EnableOnApply() As Boolean: EnableOnApply = mEnableOnApply: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property
