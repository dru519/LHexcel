VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxHangulSettings"
Option Explicit

Private mFontName As String
Private mFontType As String
Private mFontSizePt As Long
Private mTitleRow As Boolean
Private mIncludeHidden As Boolean
Private mTableStyle As String
Private mSealed As Boolean

Public Sub Configure(ByVal fontName As String, ByVal fontType As String, ByVal fontSizePt As Long, _
    ByVal titleRow As Boolean, ByVal includeHidden As Boolean, ByVal tableStyle As String)
    If mSealed Then NxRaiseContractError "한글 표 설정은 한 번만 구성할 수 있습니다."
    fontName = Trim$(fontName)
    Select Case fontName
        Case "중고딕", "함초롬돋움", "맑은 고딕", "굴림"
        Case Else: NxRaiseContractError "지원하는 한글 글꼴을 선택하세요."
    End Select
    If UCase$(Trim$(fontType)) <> "HFT" Then NxRaiseContractError "한글 표 글꼴 유형은 HFT로 고정됩니다."
    If fontSizePt < 9 Or fontSizePt > 24 Then NxRaiseContractError "글자 크기는 9~24pt 범위입니다."
    If tableStyle <> "공문형" And tableStyle <> "간결형" Then NxRaiseContractError "표 모양을 선택하세요."
    mFontName = fontName
    mFontType = "HFT"
    mFontSizePt = fontSizePt
    mTitleRow = titleRow
    mIncludeHidden = includeHidden
    mTableStyle = tableStyle
End Sub

Public Sub Seal()
    If Len(mFontName) = 0 Or Len(mFontType) = 0 Or mFontSizePt = 0 Or Len(mTableStyle) = 0 Then _
        NxRaiseContractError "한글 표 설정이 완전하지 않습니다."
    mSealed = True
End Sub

Public Property Get FontName() As String: FontName = mFontName: End Property
Public Property Get FontType() As String: FontType = mFontType: End Property
Public Property Get FontSizePt() As Long: FontSizePt = mFontSizePt: End Property
Public Property Get TitleRow() As Boolean: TitleRow = mTitleRow: End Property
Public Property Get IncludeHidden() As Boolean: IncludeHidden = mIncludeHidden: End Property
Public Property Get TableStyle() As String: TableStyle = mTableStyle: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property
