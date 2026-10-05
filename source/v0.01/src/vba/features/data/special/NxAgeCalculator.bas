Attribute VB_Name = "NxAgeCalculator"
Option Explicit

Public Function NxAgeValue(ByVal sourceValue As Variant, ByVal referenceDate As Date) As Variant
    Dim birthDate As Date
    Dim age As Long
    On Error GoTo InvalidValue
    If IsError(sourceValue) Or IsEmpty(sourceValue) Or IsNull(sourceValue) Then GoTo InvalidValue
    If Not NxAgeTryParseDate(sourceValue, referenceDate, birthDate) Then GoTo InvalidValue
    birthDate = DateValue(birthDate)
    referenceDate = DateValue(referenceDate)
    If birthDate > referenceDate Or Year(birthDate) < 1900 Then GoTo InvalidValue
    age = DateDiff("yyyy", birthDate, referenceDate)
    If DateAdd("yyyy", age, birthDate) > referenceDate Then age = age - 1
    If age < 0 Or age > 150 Then GoTo InvalidValue
    NxAgeValue = age
    Exit Function
InvalidValue:
    Err.Clear
    NxAgeValue = CVErr(xlErrValue)
End Function

' Six-digit dates are interpreted before Excel serial dates, using the reference
' century; explicit resident-number gender codes override that century.
Public Function NxAgeTryParseDate(ByVal value As Variant, ByVal referenceDate As Date, ByRef parsed As Date) As Boolean
    NxAgeTryParseDate = NxDateTryParseValue(value, referenceDate, parsed)
End Function

' Shared interpretation for date conversion, normalization and age calculation.
Public Function NxDateTryParseValue(ByVal value As Variant, ByVal referenceDate As Date, ByRef parsed As Date) As Boolean
    Dim text As String, digits As String, parts As Variant
    Dim y As Long, m As Long, d As Long, gender As String
    Dim clockText As String, clockParts As Variant, clockAt As Long, timeValue As Date
    Dim h As Long, minute As Long, second As Long, number As Double
    On Error GoTo InvalidDate
    If IsError(value) Or IsEmpty(value) Or IsNull(value) Then Exit Function
    If VarType(value) = vbDate Then parsed = CDate(value): NxDateTryParseValue = True: Exit Function
    text = Trim$(CStr(value))
    If Len(text) = 0 Then Exit Function
    ' Numeric Excel serials retain the time fraction. Compact calendar numbers
    ' take priority; strings are never interpreted through the machine locale.
    If IsNumeric(value) And VarType(value) <> vbString Then
        number = CDbl(value)
        If number <> Fix(number) Or (Len(text) <> 6 And Len(text) <> 8 And Len(text) <> 13) Then
            If number < 1 Or number >= 2958466# Then Exit Function
            parsed = CDate(value): NxDateTryParseValue = True: Exit Function
        End If
    End If
    text = Replace(text, "T", " ")
    If InStr(text, ":") > 0 Then
        clockAt = InStrRev(text, " ")
        If clockAt = 0 Then Exit Function
        clockText = Mid$(text, clockAt + 1): text = Trim$(Left$(text, clockAt - 1))
        clockParts = Split(clockText, ":")
        If UBound(clockParts) < 1 Or UBound(clockParts) > 2 Then Exit Function
        If Not NxDateOnlyDigits(CStr(clockParts(0))) Or Not NxDateOnlyDigits(CStr(clockParts(1))) Then Exit Function
        h = CLng(clockParts(0)): minute = CLng(clockParts(1))
        If UBound(clockParts) = 2 Then
            If Not NxDateOnlyDigits(CStr(clockParts(2))) Then Exit Function
            second = CLng(clockParts(2))
        End If
        If h > 23 Or minute > 59 Or second > 59 Then Exit Function
        timeValue = TimeSerial(h, minute, second)
    End If
    text = Replace(Replace(Replace(text, "년", "-"), "월", "-"), "일", "")
    text = Replace(Replace(Replace(text, ".", "-"), "/", "-"), " ", "")
    If Right$(text, 1) = "-" Then text = Left$(text, Len(text) - 1)
    parts = Split(text, "-")
    If UBound(parts) = 2 Then
        If Len(parts(0)) <> 2 And Len(parts(0)) <> 4 Then Exit Function
        If Not NxDateOnlyDigits(CStr(parts(0))) Or Not NxDateOnlyDigits(CStr(parts(1))) Or Not NxDateOnlyDigits(CStr(parts(2))) Then Exit Function
        y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
        If Len(parts(0)) = 2 Then y = NxAgeFullYear(y, Year(referenceDate))
    Else
        If InStr(text, "-") > 0 Then
            If UBound(parts) <> 1 Or Len(parts(0)) <> 6 Or Len(parts(1)) <> 7 Then Exit Function
        End If
        If Len(Replace(Replace(Replace(text, "-", ""), "*", ""), " ", "")) = 0 Then Exit Function
        If Not NxDateOnlyDigits(Replace(Replace(text, "-", ""), "*", "")) Then Exit Function
        digits = NxAgeDigits(text)
        If InStr(text, "*") > 0 Then
            If Len(Replace(text, "-", "")) <> 13 Then Exit Function
            If Len(digits) = 6 Then
                If Right$(text, 7) <> "*******" Then Exit Function
            ElseIf Len(digits) = 7 Then
                If Right$(text, 6) <> "******" Then Exit Function
            Else
                Exit Function
            End If
        End If
        If Len(digits) = 8 Then
            y = CLng(Left$(digits, 4)): m = CLng(Mid$(digits, 5, 2)): d = CLng(Right$(digits, 2))
        ElseIf Len(digits) = 6 Or Len(digits) = 13 Or (Len(digits) = 7 And InStr(text, "*") > 0) Then
            y = NxAgeFullYear(CLng(Left$(digits, 2)), Year(referenceDate))
            If Len(digits) >= 7 Then
                gender = Mid$(digits, 7, 1)
                Select Case gender
                    Case "1", "2", "5", "6": y = 1900 + CLng(Left$(digits, 2))
                    Case "3", "4", "7", "8": y = 2000 + CLng(Left$(digits, 2))
                    Case "9", "0": y = 1800 + CLng(Left$(digits, 2))
                End Select
            End If
            m = CLng(Mid$(digits, 3, 2)): d = CLng(Mid$(digits, 5, 2))
        Else
            Exit Function
        End If
    End If
    If y < 100 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    parsed = DateSerial(y, m, d) + timeValue
    NxDateTryParseValue = (Year(parsed) = y And Month(parsed) = m And Day(parsed) = d)
    Exit Function
InvalidDate:
    Err.Clear
End Function

Private Function NxDateOnlyDigits(ByVal text As String) As Boolean
    If Len(text) = 0 Then Exit Function
    NxDateOnlyDigits = Not text Like "*[!0-9]*"
End Function

Private Function NxAgeFullYear(ByVal yearPart As Long, ByVal referenceYear As Long) As Long
    NxAgeFullYear = (referenceYear \ 100) * 100 + yearPart
    If NxAgeFullYear > referenceYear Then NxAgeFullYear = NxAgeFullYear - 100
End Function


Private Function NxAgeDigits(ByVal textValue As String) As String
    Dim index As Long
    Dim character As String
    For index = 1 To Len(textValue)
        character = Mid$(textValue, index, 1)
        If character >= "0" And character <= "9" Then NxAgeDigits = NxAgeDigits & character
    Next index
End Function
