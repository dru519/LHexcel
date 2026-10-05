Attribute VB_Name = "NxPrivacyMask"
Option Explicit

' The masking workflow follows donor options. The audit's NxPrivacyMaskedText
' and NxPrivacyKind APIs below retain their existing detection semantics.
Public Function NxPrivacyMaskWithOptions(ByVal value As Variant, ByVal maskType As String, ByVal maskChar As String) As String
    Dim text As String, digits As String, separator As Long, kind As String, maskLength As Long
    NxPrivacyMaskWithOptions = "error"
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If maskChar <> "*" And maskChar <> "O" Then NxRaiseContractError "마스킹 문자를 확인하세요."
    text = Trim$(CStr(value)): digits = NxPrivacyDigits(text): kind = maskType
    If kind = "자동" Then
        If Len(digits) = 13 Then
            kind = "주민등록번호"
        ElseIf (Len(digits) = 10 Or Len(digits) = 11) And Left$(digits, 1) = "0" Then
            kind = "휴대폰번호"
        ElseIf NxPrivacyLooksLikeEmail(text) Then
            kind = "이메일"
        ElseIf NxPrivacyLooksLikeKoreanName(Replace(text, " ", "")) Then
            kind = "이름"
        End If
    End If
    Select Case kind
        Case "주민등록번호"
            If Len(digits) = 13 Then NxPrivacyMaskWithOptions = Left$(digits, 6) & "-" & String$(7, maskChar)
        Case "휴대폰번호"
            If Left$(digits, 1) <> "0" Then Exit Function
            If Len(digits) = 11 Then
                NxPrivacyMaskWithOptions = Left$(digits, 3) & "-" & String$(4, maskChar) & "-" & Right$(digits, 4)
            ElseIf Len(digits) = 10 Then
                If Left$(digits, 2) = "02" Then
                    NxPrivacyMaskWithOptions = "02-" & String$(4, maskChar) & "-" & Right$(digits, 4)
                Else
                    NxPrivacyMaskWithOptions = Left$(digits, 3) & "-" & String$(3, maskChar) & "-" & Right$(digits, 4)
                End If
            End If
        Case "이메일"
            If Not NxPrivacyLooksLikeEmail(text) Then Exit Function
            separator = InStr(text, "@"): maskLength = separator - 2
            If maskLength < 3 Then maskLength = 3
            NxPrivacyMaskWithOptions = Left$(text, 1) & String$(maskLength, maskChar) & Mid$(text, separator)
        Case "이름"
            text = Replace(text, " ", "")
            If Not NxPrivacyLooksLikeKoreanName(text) Then Exit Function
            If Len(text) = 2 Then
                NxPrivacyMaskWithOptions = Left$(text, 1) & maskChar
            Else
                NxPrivacyMaskWithOptions = Left$(text, 1) & String$(Len(text) - 2, maskChar) & Right$(text, 1)
            End If
    End Select
End Function

Public Function NxPrivacyMaskedText(ByVal sourceValue As Variant, Optional ByRef detectedKind As String) As Variant
    Dim textValue As String
    Dim digits As String
    detectedKind = vbNullString
    If IsError(sourceValue) Or IsEmpty(sourceValue) Or IsNull(sourceValue) Then NxPrivacyMaskedText = sourceValue: Exit Function
    textValue = Trim$(CStr(sourceValue))
    If Len(textValue) = 0 Then NxPrivacyMaskedText = sourceValue: Exit Function
    digits = NxPrivacyDigits(textValue)
    If Len(digits) = 13 And Mid$(digits, 7, 1) >= "1" And Mid$(digits, 7, 1) <= "8" Then
        detectedKind = "주민등록번호 후보"
        NxPrivacyMaskedText = Left$(digits, 6) & "-" & Mid$(digits, 7, 1) & "******"
    ElseIf (Len(digits) = 10 Or Len(digits) = 11) And Left$(digits, 1) = "0" Then
        detectedKind = "연락처 후보"
        NxPrivacyMaskedText = Left$(digits, 3) & "-****-" & Right$(digits, 4)
    ElseIf NxPrivacyLooksLikeEmail(textValue) Then
        detectedKind = "이메일 후보"
        NxPrivacyMaskedText = NxPrivacyMaskEmail(textValue)
    ElseIf NxPrivacyLooksLikeKoreanName(textValue) Then
        detectedKind = "이름 후보"
        NxPrivacyMaskedText = Left$(textValue, 1) & String$(Len(textValue) - 1, "*")
    Else
        NxPrivacyMaskedText = sourceValue
    End If
End Function

Public Function NxPrivacyKind(ByVal sourceValue As Variant) As String
    Dim kind As String
    Dim ignored As Variant
    ignored = NxPrivacyMaskedText(sourceValue, kind)
    NxPrivacyKind = kind
End Function

Private Function NxPrivacyDigits(ByVal textValue As String) As String
    Dim index As Long
    Dim character As String
    For index = 1 To Len(textValue)
        character = Mid$(textValue, index, 1)
        If character >= "0" And character <= "9" Then NxPrivacyDigits = NxPrivacyDigits & character
    Next index
End Function

Private Function NxPrivacyLooksLikeEmail(ByVal textValue As String) As Boolean
    Dim separator As Long
    separator = InStr(2, textValue, "@", vbBinaryCompare)
    NxPrivacyLooksLikeEmail = (separator > 1 And separator < Len(textValue) And InStr(separator + 2, textValue, ".", vbBinaryCompare) > separator + 1)
End Function

Private Function NxPrivacyMaskEmail(ByVal textValue As String) As String
    Dim separator As Long
    Dim localPart As String
    separator = InStr(1, textValue, "@", vbBinaryCompare)
    localPart = Left$(textValue, separator - 1)
    If Len(localPart) = 1 Then
        localPart = "*"
    Else
        localPart = Left$(localPart, 1) & String$(Len(localPart) - 1, "*")
    End If
    NxPrivacyMaskEmail = localPart & Mid$(textValue, separator)
End Function

Private Function NxPrivacyLooksLikeKoreanName(ByVal textValue As String) As Boolean
    Dim index As Long
    Dim codePoint As Long
    If Len(textValue) < 2 Or Len(textValue) > 5 Or InStr(textValue, " ") > 0 Then Exit Function
    For index = 1 To Len(textValue)
        codePoint = AscW(Mid$(textValue, index, 1))
        If codePoint < 0 Then codePoint = codePoint + 65536
        If codePoint < 44032 Or codePoint > 55203 Then Exit Function
    Next index
    NxPrivacyLooksLikeKoreanName = True
End Function
