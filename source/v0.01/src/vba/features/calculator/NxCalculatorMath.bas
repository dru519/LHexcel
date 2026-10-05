Attribute VB_Name = "NxCalculatorMath"
Option Explicit

Private Const NX_CALCULATOR_ERROR_BASE As Long = vbObjectError + 940

Public Function NxCalculatorApplyBinary(ByVal leftValue As Double, ByVal operationId As String, ByVal rightValue As Double) As Double
    On Error GoTo ArithmeticFailed
    Select Case operationId
        Case "ADD": NxCalculatorApplyBinary = leftValue + rightValue
        Case "SUBTRACT": NxCalculatorApplyBinary = leftValue - rightValue
        Case "MULTIPLY": NxCalculatorApplyBinary = leftValue * rightValue
        Case "DIVIDE"
            If rightValue = 0# Then NxCalculatorMathError "0으로 나눌 수 없습니다."
            NxCalculatorApplyBinary = leftValue / rightValue
        Case "POWER"
            If leftValue = 0# And rightValue < 0# Then NxCalculatorMathError "0의 음수 제곱은 계산할 수 없습니다."
            If leftValue < 0# And rightValue <> Fix(rightValue) Then NxCalculatorMathError "음수 밑에는 정수 지수만 사용할 수 있습니다."
            NxCalculatorApplyBinary = leftValue ^ rightValue
        Case "YROOT"
            If rightValue = 0# Or rightValue <> Fix(rightValue) Then NxCalculatorMathError "근의 차수는 0이 아닌 정수여야 합니다."
            If leftValue < 0# Then
                If Abs(CLng(rightValue)) Mod 2 = 0 Then NxCalculatorMathError "음수에는 홀수 차수의 근만 사용할 수 있습니다."
                NxCalculatorApplyBinary = -((-leftValue) ^ (1# / rightValue))
            Else
                NxCalculatorApplyBinary = leftValue ^ (1# / rightValue)
            End If
        Case Else
            NxCalculatorMathError "지원하지 않는 이항 연산입니다."
    End Select
    Exit Function
ArithmeticFailed:
    NxCalculatorReraise Err.Number, Err.Description
End Function

Public Function NxCalculatorApplyUnary(ByVal operationId As String, ByVal value As Double, Optional ByVal argument As Double = 0#) As Double
    On Error GoTo ArithmeticFailed
    Select Case operationId
        Case "RECIPROCAL"
            If value = 0# Then NxCalculatorMathError "0의 역수는 계산할 수 없습니다."
            NxCalculatorApplyUnary = 1# / value
        Case "SQUARE": NxCalculatorApplyUnary = value * value
        Case "CUBE": NxCalculatorApplyUnary = value * value * value
        Case "SQRT"
            If value < 0# Then NxCalculatorMathError "음수의 제곱근은 실수 범위에서 계산할 수 없습니다."
            NxCalculatorApplyUnary = Sqr(value)
        Case "CBRT"
            If value = 0# Then
                NxCalculatorApplyUnary = 0#
            Else
                NxCalculatorApplyUnary = Sgn(value) * (Abs(value) ^ (1# / 3#))
            End If
        Case "TEN_POWER": NxCalculatorApplyUnary = 10# ^ value
        Case "FACTORIAL": NxCalculatorApplyUnary = NxCalculatorFactorial(value)
        Case Else
            NxCalculatorMathError "지원하지 않는 단항 연산입니다."
    End Select
    Exit Function
ArithmeticFailed:
    NxCalculatorReraise Err.Number, Err.Description
End Function

Public Function NxCalculatorFactorial(ByVal value As Double) As Double
    Dim index As Long
    If value <> Fix(value) Or value < 0# Or value > 170# Then NxCalculatorMathError "팩토리얼은 0부터 170까지의 정수만 지원합니다."
    NxCalculatorFactorial = 1#
    For index = 2 To CLng(value)
        NxCalculatorFactorial = NxCalculatorFactorial * CDbl(index)
    Next index
End Function

Public Function NxCalculatorIsBinaryOperation(ByVal operationId As String) As Boolean
    Select Case operationId
        Case "ADD", "SUBTRACT", "MULTIPLY", "DIVIDE", "POWER", "YROOT"
            NxCalculatorIsBinaryOperation = True
        Case Else
            NxCalculatorIsBinaryOperation = False
    End Select
End Function

Public Function NxCalculatorIsUnaryOperation(ByVal operationId As String) As Boolean
    Select Case operationId
        Case "RECIPROCAL", "SQUARE", "CUBE", "SQRT", "CBRT", "TEN_POWER", "FACTORIAL"
            NxCalculatorIsUnaryOperation = True
        Case Else
            NxCalculatorIsUnaryOperation = False
    End Select
End Function

Public Function NxCalculatorOperationSymbol(ByVal operationId As String) As String
    Select Case operationId
        Case "ADD": NxCalculatorOperationSymbol = "+"
        Case "SUBTRACT": NxCalculatorOperationSymbol = "−"
        Case "MULTIPLY": NxCalculatorOperationSymbol = "×"
        Case "DIVIDE": NxCalculatorOperationSymbol = "÷"
        Case "POWER": NxCalculatorOperationSymbol = "^"
        Case "YROOT": NxCalculatorOperationSymbol = "√ 차수"
        Case "RECIPROCAL": NxCalculatorOperationSymbol = "1/x"
        Case "SQUARE": NxCalculatorOperationSymbol = "x²"
        Case "CUBE": NxCalculatorOperationSymbol = "x³"
        Case "SQRT": NxCalculatorOperationSymbol = "√x"
        Case "CBRT": NxCalculatorOperationSymbol = "∛x"
        Case "TEN_POWER": NxCalculatorOperationSymbol = "10ˣ"
        Case "FACTORIAL": NxCalculatorOperationSymbol = "x!"
        Case Else: NxCalculatorMathError "지원하지 않는 연산입니다."
    End Select
End Function

Private Sub NxCalculatorMathError(ByVal message As String)
    Err.Raise NX_CALCULATOR_ERROR_BASE, "NxCalculatorMath", message
End Sub

Private Sub NxCalculatorReraise(ByVal number As Long, ByVal description As String)
    If number >= NX_CALCULATOR_ERROR_BASE And number < NX_CALCULATOR_ERROR_BASE + 20 Then
        Err.Raise number, "NxCalculatorMath", description
    End If
    Err.Raise NX_CALCULATOR_ERROR_BASE + 1, "NxCalculatorMath", "계산 결과가 숫자 범위를 벗어났습니다."
End Sub
