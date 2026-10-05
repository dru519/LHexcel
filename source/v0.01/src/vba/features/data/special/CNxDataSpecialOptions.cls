VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDataSpecialOptions"
Option Explicit

Private mConfigured As Boolean
Private mReferenceDate As Date
Private mOutputMode As String
Private mHeader As Boolean
Private mAgeFormat As String
Private mPrefix As Boolean, mSuffix As Boolean, mMixed As Boolean
Private mAgePrefix As Boolean, mAgeSuffix As Boolean
Private mZeroMode As String, mMaskType As String, mMaskChar As String

Public Sub Configure(ByVal outputMode As String, ByVal hasHeader As Boolean, ByVal referenceText As String, _
    ByVal ageFormat As String, ByVal usePrefix As Boolean, ByVal useSuffix As Boolean, ByVal mixed As Boolean, _
    ByVal zeroMode As String, ByVal maskType As String, ByVal maskChar As String, _
    Optional ByVal agePrefix As Boolean = True, Optional ByVal ageSuffix As Boolean = True)
    If mConfigured Then NxRaiseContractError "데이터 옵션은 한 번만 구성할 수 있습니다."
    Select Case outputMode
        Case "오른쪽 새 열 삽입", "빈 결과 열/행", "원본 변경", "새 시트 만들기", "새 통합문서"
        Case Else: NxRaiseContractError "결과 위치를 선택하세요."
    End Select
    mReferenceDate = Date
    If Len(Trim$(referenceText)) > 0 Then
        If Not NxAgeTryParseDate(referenceText, Date, mReferenceDate) Then NxRaiseContractError "기준일을 인식하지 못했습니다."
    End If
    If ageFormat <> "만 00세" And ageFormat <> "숫자" And ageFormat <> "사용자 지정" Then NxRaiseContractError "만나이 출력양식을 선택하세요."
    If zeroMode <> "빈칸" And zeroMode <> "영원" And zeroMode <> "error" Then NxRaiseContractError "0원 처리 방식을 선택하세요."
    Select Case maskType
        Case "자동", "주민등록번호", "휴대폰번호", "이메일", "이름"
        Case Else: NxRaiseContractError "마스킹 종류를 선택하세요."
    End Select
    If maskChar <> "*" And maskChar <> "O" Then NxRaiseContractError "마스킹 문자는 * 또는 O만 지원합니다."
    mOutputMode = outputMode: mHeader = hasHeader: mAgeFormat = ageFormat
    mAgePrefix = agePrefix: mAgeSuffix = ageSuffix
    mPrefix = usePrefix: mSuffix = useSuffix: mMixed = mixed
    mZeroMode = zeroMode: mMaskType = maskType: mMaskChar = maskChar
    mConfigured = True
End Sub

Public Property Get IsConfigured() As Boolean: IsConfigured = mConfigured: End Property
Public Property Get OutputMode() As String: OutputMode = mOutputMode: End Property
Public Property Get HasHeader() As Boolean: HasHeader = mHeader: End Property
Public Property Get ReferenceDate() As Date: ReferenceDate = mReferenceDate: End Property
Public Property Get AgeFormat() As String: AgeFormat = mAgeFormat: End Property

Public Function TransformValue(ByVal featureId As String, ByVal value As Variant) As Variant
    Dim result As Variant
    If Not mConfigured Then NxRaiseContractError "데이터 옵션이 구성되지 않았습니다."
    If NxDataValueIsBlank(value) Then TransformValue = vbNullString: Exit Function
    If IsError(value) Then TransformValue = vbNullString: Exit Function
    Select Case featureId
        Case NX_FEATURE_DATA_AGE
            result = NxAgeValue(value, mReferenceDate)
            If IsError(result) Then
                result = vbNullString
            ElseIf mAgeFormat = "만 00세" Then
                result = "만 " & CStr(result) & "세"
            ElseIf mAgeFormat = "사용자 지정" Then
                result = IIf(mAgePrefix, "만 ", vbNullString) & CStr(result) & IIf(mAgeSuffix, "세", vbNullString)
            End If
        Case NX_FEATURE_DATA_KOREAN_MONEY
            result = NxKoreanMoneyWithOptions(value, mPrefix, mSuffix, mMixed, mZeroMode)
        Case NX_FEATURE_DATA_PRIVACY_MASK
            result = NxPrivacyMaskWithOptions(value, mMaskType, mMaskChar)
        Case Else: NxRaiseContractError "지원하지 않는 데이터 변환 기능입니다."
    End Select
    If IsError(result) Then result = vbNullString
    If VarType(result) = vbString Then
        If result = "error" Then result = vbNullString
    End If
    TransformValue = result
End Function
