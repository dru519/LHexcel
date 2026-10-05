Attribute VB_Name = "NxKoreanMoney"
Option Explicit

' Keep decimal strings intact: Excel text amounts can exceed Double precision.
Public Function NxKoreanMoneyWithOptions(ByVal value As Variant, ByVal usePrefix As Boolean, ByVal useSuffix As Boolean, _
    ByVal mixed As Boolean, ByVal zeroMode As String) As String
    Dim digits As String, korean As String, numericText As String, part As String, groupText As String
    Dim units As Variant, names As Variant, small As Variant, groupIndex As Long, i As Long, n As Long
    On Error GoTo InvalidValue
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then GoTo InvalidValue
    digits = Trim$(CStr(value))
    digits = Replace(Replace(Replace(Replace(Replace(digits, ",", ""), "원", ""), "₩", ""), "￦", ""), " ", "")
    If Len(digits) = 0 Or Len(digits) > 20 Then GoTo InvalidValue
    For i = 1 To Len(digits)
        If Mid$(digits, i, 1) < "0" Or Mid$(digits, i, 1) > "9" Then GoTo InvalidValue
    Next i
    Do While Len(digits) > 1 And Left$(digits, 1) = "0": digits = Mid$(digits, 2): Loop
    If digits = "0" Then
        If zeroMode = "빈칸" Then Exit Function
        If zeroMode = "error" Then GoTo InvalidValue
        korean = "영"
    End If
    numericText = digits
    For i = Len(numericText) - 3 To 1 Step -3
        numericText = Left$(numericText, i) & "," & Mid$(numericText, i + 1)
    Next i
    units = Array("", "만", "억", "조", "경")
    names = Array("", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구")
    small = Array("천", "백", "십", "")
    Do While Len(digits) > 0
        part = Right$("0000" & Right$(digits, 4), 4)
        digits = Left$(digits, Len(digits) - Len(Right$(digits, 4)))
        groupText = ""
        For i = 1 To 4
            n = CLng(Mid$(part, i, 1))
            If n > 0 Then groupText = groupText & CStr(names(n)) & CStr(small(i - 1))
        Next i
        If Len(groupText) > 0 Then korean = groupText & CStr(units(groupIndex)) & korean
        groupIndex = groupIndex + 1
    Loop
    korean = korean & "원": numericText = numericText & "원"
    If usePrefix Then korean = "금" & korean: numericText = "금" & numericText
    If useSuffix Then korean = korean & "정": numericText = numericText & "정"
    If mixed Then korean = numericText & "(" & korean & ")"
    NxKoreanMoneyWithOptions = korean
    Exit Function
InvalidValue:
    Err.Clear
    NxKoreanMoneyWithOptions = "error"
End Function

Public Function NxKoreanMoneyText(ByVal sourceValue As Variant) As Variant
    Dim numericValue As Double
    Dim digits As String
    Dim padded As String
    Dim groupCount As Long
    Dim groupIndex As Long
    Dim groupText As String
    Dim result As String
    Dim prefix As String
    Dim bigUnits As Variant
    On Error GoTo InvalidValue
    If IsError(sourceValue) Or IsEmpty(sourceValue) Or IsNull(sourceValue) Or Not IsNumeric(sourceValue) Then GoTo InvalidValue
    numericValue = CDbl(sourceValue)
    If Abs(numericValue) > 9999999999999999# Or numericValue <> Fix(numericValue) Then GoTo InvalidValue
    If numericValue = 0 Then NxKoreanMoneyText = "금 영원": Exit Function
    If numericValue < 0 Then prefix = "마이너스 ": numericValue = Abs(numericValue)
    digits = Format$(numericValue, "0")
    Do While Len(digits) Mod 4 <> 0
        digits = "0" & digits
    Loop
    padded = digits
    groupCount = Len(padded) \ 4
    bigUnits = Array("", "만", "억", "조", "경")
    For groupIndex = 1 To groupCount
        groupText = NxKoreanFourDigits(Mid$(padded, (groupIndex - 1) * 4 + 1, 4))
        If Len(groupText) > 0 Then result = result & groupText & CStr(bigUnits(groupCount - groupIndex))
    Next groupIndex
    NxKoreanMoneyText = "금 " & prefix & result & "원"
    Exit Function
InvalidValue:
    Err.Clear
    NxKoreanMoneyText = CVErr(xlErrValue)
End Function

Private Function NxKoreanFourDigits(ByVal digits As String) As String
    Dim digitNames As Variant
    Dim unitNames As Variant
    Dim index As Long
    Dim digitValue As Long
    digitNames = Array("", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구")
    unitNames = Array("천", "백", "십", "")
    For index = 1 To 4
        digitValue = CLng(Mid$(digits, index, 1))
        If digitValue <> 0 Then
            If digitValue <> 1 Or index = 4 Then NxKoreanFourDigits = NxKoreanFourDigits & CStr(digitNames(digitValue))
            NxKoreanFourDigits = NxKoreanFourDigits & CStr(unitNames(index - 1))
        End If
    Next index
End Function
