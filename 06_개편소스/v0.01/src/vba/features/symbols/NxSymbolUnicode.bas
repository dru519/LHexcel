Attribute VB_Name = "NxSymbolUnicode"
Option Explicit

Public Function NxSymbolIsAllowedCodePoint(ByVal codePoint As Long) As Boolean
    If codePoint < 0 Or codePoint > &H10FFFF Then Exit Function
    If codePoint <= &H1F Or (codePoint >= &H7F And codePoint <= &H9F) Then Exit Function
    If codePoint >= &HD800& And codePoint <= &HDFFF& Then Exit Function
    If codePoint >= &HFDD0& And codePoint <= &HFDEF& Then Exit Function
    If (codePoint And &HFFFF&) = &HFFFE& Or (codePoint And &HFFFF&) = &HFFFF& Then Exit Function
    NxSymbolIsAllowedCodePoint = True
End Function

Public Function NxSymbolIsDisplayableCodePoint(ByVal codePoint As Long) As Boolean
    If Not NxSymbolIsAllowedCodePoint(codePoint) Then Exit Function
    Select Case codePoint
        Case &HAD&, &H34F&, &H61C&, &H115F&, &H1160&, &H17B4&, &H17B5&
            Exit Function
        Case &H180B& To &H180F&, &H2000& To &H200F&, &H2028& To &H202E&, _
             &H2060& To &H206F&, &HFE00& To &HFE0F&, &HFEFF&
            Exit Function
    End Select
    NxSymbolIsDisplayableCodePoint = True
End Function

Public Function NxSymbolFromCodePoint(ByVal codePoint As Long) As String
    Dim adjusted As Long, highUnit As Long, lowUnit As Long
    If Not NxSymbolIsAllowedCodePoint(codePoint) Then NxRaiseContractError "사용할 수 없는 유니코드 코드입니다."
    If codePoint <= &HFFFF& Then
        NxSymbolFromCodePoint = NxSymbolChrW(codePoint)
    Else
        adjusted = codePoint - &H10000
        highUnit = &HD800& + (adjusted \ &H400&)
        lowUnit = &HDC00& + (adjusted And &H3FF&)
        NxSymbolFromCodePoint = NxSymbolChrW(highUnit) & NxSymbolChrW(lowUnit)
    End If
End Function

Public Function NxSymbolFirstCodePoint(ByVal value As String) As Long
    Dim firstUnit As Long, secondUnit As Long
    If Len(value) = 0 Then NxRaiseContractError "문자가 비어 있습니다."
    firstUnit = NxSymbolUnsignedCodeUnit(AscW(Left$(value, 1)))
    If firstUnit >= &HD800& And firstUnit <= &HDBFF& Then
        If Len(value) < 2 Then NxRaiseContractError "완성되지 않은 유니코드 문자입니다."
        secondUnit = NxSymbolUnsignedCodeUnit(AscW(Mid$(value, 2, 1)))
        If secondUnit < &HDC00& Or secondUnit > &HDFFF& Then NxRaiseContractError "완성되지 않은 유니코드 문자입니다."
        NxSymbolFirstCodePoint = &H10000 + ((firstUnit - &HD800&) * &H400&) + (secondUnit - &HDC00&)
    ElseIf firstUnit >= &HDC00& And firstUnit <= &HDFFF& Then
        NxRaiseContractError "완성되지 않은 유니코드 문자입니다."
    Else
        If Not NxSymbolIsAllowedCodePoint(firstUnit) Then NxRaiseContractError "사용할 수 없는 유니코드 문자입니다."
        NxSymbolFirstCodePoint = firstUnit
    End If
End Function

Public Function NxSymbolDropLastScalar(ByVal value As String) As String
    Dim length As Long, lastUnit As Long, previousUnit As Long
    length = Len(value)
    If length = 0 Then Exit Function
    lastUnit = NxSymbolUnsignedCodeUnit(AscW(Right$(value, 1)))
    If lastUnit >= &HDC00& And lastUnit <= &HDFFF& And length >= 2 Then
        previousUnit = NxSymbolUnsignedCodeUnit(AscW(Mid$(value, length - 1, 1)))
        If previousUnit >= &HD800& And previousUnit <= &HDBFF& Then
            NxSymbolDropLastScalar = Left$(value, length - 2)
            Exit Function
        End If
    End If
    NxSymbolDropLastScalar = Left$(value, length - 1)
End Function

Public Function NxSymbolFormatCode(ByVal codePoint As Long) As String
    Dim hexText As String
    If Not NxSymbolIsAllowedCodePoint(codePoint) Then NxRaiseContractError "사용할 수 없는 유니코드 코드입니다."
    hexText = Hex$(codePoint)
    If Len(hexText) < 4 Then hexText = Right$("0000" & hexText, 4)
    NxSymbolFormatCode = "U+" & hexText
End Function

Public Function NxSymbolDecodeUtf16Hex(ByVal encoded As String) As String
    Dim parts As Variant, item As Variant, codeUnit As Long, result As String
    If Len(encoded) = 0 Then Exit Function
    parts = Split(encoded, ".")
    For Each item In parts
        If Len(CStr(item)) <> 4 Then NxRaiseContractError "기호 카탈로그 문자열 형식이 잘못되었습니다."
        On Error GoTo InvalidEncoding
        codeUnit = CLng("&H" & CStr(item))
        On Error GoTo 0
        result = result & NxSymbolChrW(codeUnit)
    Next item
    NxSymbolDecodeUtf16Hex = result
    Exit Function
InvalidEncoding:
    Err.Clear
    NxRaiseContractError "기호 카탈로그 문자열 형식이 잘못되었습니다."
End Function

Private Function NxSymbolChrW(ByVal codeUnit As Long) As String
    If codeUnit < 0 Or codeUnit > &HFFFF& Then NxRaiseContractError "UTF-16 코드 단위가 잘못되었습니다."
    If codeUnit > &H7FFF Then codeUnit = codeUnit - &H10000
    NxSymbolChrW = ChrW$(codeUnit)
End Function

Private Function NxSymbolUnsignedCodeUnit(ByVal signedUnit As Long) As Long
    If signedUnit < 0 Then signedUnit = signedUnit + &H10000
    NxSymbolUnsignedCodeUnit = signedUnit
End Function
