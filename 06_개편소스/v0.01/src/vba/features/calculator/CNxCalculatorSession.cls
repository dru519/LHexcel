VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxCalculatorSession"
Option Explicit

Private mHistory As CNxCalculatorHistory
Private mEntry As String
Private mAccumulator As Double
Private mHasAccumulator As Boolean
Private mPendingOperation As String
Private mStartNewEntry As Boolean
Private mHasValidResult As Boolean
Private mResultValue As Double
Private mFormulaText As String
Private mStatusText As String

Private Sub Class_Initialize()
    ResetCurrent
End Sub

Public Sub BindHistory(ByVal history As CNxCalculatorHistory)
    If history Is Nothing Then NxRaiseContractError "계산 이력이 필요합니다."
    Set mHistory = history
End Sub

Public Sub InputDigit(ByVal digit As String)
    If Len(digit) <> 1 Or digit < "0" Or digit > "9" Then NxRaiseContractError "숫자 한 자리만 입력할 수 있습니다."
    If mStartNewEntry Or mEntry = "0" Then
        mEntry = digit
    ElseIf mEntry = "-0" Then
        mEntry = "-" & digit
    Else
        mEntry = mEntry & digit
    End If
    mStartNewEntry = False
    MarkCurrentValue
End Sub

Public Sub InputDecimal()
    If mStartNewEntry Then mEntry = "0": mStartNewEntry = False
    If InStr(1, mEntry, ".", vbBinaryCompare) = 0 Then mEntry = mEntry & "."
    MarkCurrentValue
End Sub

Public Sub ToggleSign()
    If mStartNewEntry Then mEntry = "0": mStartNewEntry = False
    If Left$(mEntry, 1) = "-" Then
        mEntry = Mid$(mEntry, 2)
    Else
        mEntry = "-" & mEntry
    End If
    MarkCurrentValue
End Sub

Public Sub Backspace()
    If mStartNewEntry Then Exit Sub
    If Len(mEntry) <= 1 Or (Len(mEntry) = 2 And Left$(mEntry, 1) = "-") Then
        mEntry = "0"
    Else
        mEntry = Left$(mEntry, Len(mEntry) - 1)
    End If
    MarkCurrentValue
End Sub

Public Sub ClearEntry()
    mEntry = "0"
    mStartNewEntry = False
    mHasValidResult = False
    mStatusText = "현재 숫자를 지웠습니다."
End Sub

Public Sub ClearAll()
    ResetCurrent
End Sub

Public Sub SetBinaryOperation(ByVal operationId As String)
    Dim currentValue As Double, calculated As Double
    On Error GoTo Failed
    If Not NxCalculatorIsBinaryOperation(operationId) Then NxRaiseContractError "지원하지 않는 이항 연산입니다."
    If Len(mPendingOperation) > 0 And mStartNewEntry Then
        mPendingOperation = operationId
        mFormulaText = NxCalculatorDisplayValue(mAccumulator) & " " & NxCalculatorOperationSymbol(operationId)
        Exit Sub
    End If
    currentValue = EntryValue()
    If mHasAccumulator And Len(mPendingOperation) > 0 Then
        calculated = NxCalculatorApplyBinary(mAccumulator, mPendingOperation, currentValue)
    Else
        calculated = currentValue
    End If
    mAccumulator = calculated
    mHasAccumulator = True
    mPendingOperation = operationId
    mEntry = NxCalculatorStorageValue(calculated)
    mResultValue = calculated
    mHasValidResult = True
    mStartNewEntry = True
    mFormulaText = NxCalculatorDisplayValue(calculated) & " " & NxCalculatorOperationSymbol(operationId)
    mStatusText = "다음 숫자를 입력하세요."
    Exit Sub
Failed:
    mStatusText = Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ApplyUnary(ByVal operationId As String, Optional ByVal argument As Double = 0#)
    Dim currentValue As Double, calculated As Double, expressionText As String
    On Error GoTo Failed
    If Not NxCalculatorIsUnaryOperation(operationId) Then NxRaiseContractError "지원하지 않는 단항 연산입니다."
    currentValue = EntryValue()
    calculated = NxCalculatorApplyUnary(operationId, currentValue, argument)
    expressionText = NxCalculatorOperationSymbol(operationId) & "(" & NxCalculatorDisplayValue(currentValue) & ")"
    mEntry = NxCalculatorStorageValue(calculated)
    mResultValue = calculated
    mHasValidResult = True
    mStartNewEntry = True
    mFormulaText = expressionText & " = " & NxCalculatorDisplayValue(calculated)
    mStatusText = "계산했습니다. = 키를 눌러 이력에 추가할 수 있습니다."
    Exit Sub
Failed:
    mStatusText = Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub CompleteEquals()
    Dim leftValue As Double, rightValue As Double, calculated As Double, expressionText As String
    On Error GoTo Failed
    If mHistory Is Nothing Then NxRaiseContractError "계산 이력이 연결되지 않았습니다."
    rightValue = EntryValue()
    If Len(mPendingOperation) > 0 Then
        leftValue = mAccumulator
        calculated = NxCalculatorApplyBinary(leftValue, mPendingOperation, rightValue)
        expressionText = NxCalculatorDisplayValue(leftValue) & " " & NxCalculatorOperationSymbol(mPendingOperation) & " " & NxCalculatorDisplayValue(rightValue)
    Else
        calculated = rightValue
        expressionText = NxCalculatorDisplayValue(rightValue)
    End If
    mHistory.Add expressionText & " = " & NxCalculatorDisplayValue(calculated)
    mEntry = NxCalculatorStorageValue(calculated)
    mAccumulator = calculated
    mHasAccumulator = True
    mPendingOperation = vbNullString
    mStartNewEntry = True
    mResultValue = calculated
    mHasValidResult = True
    mFormulaText = expressionText & " = " & NxCalculatorDisplayValue(calculated)
    mStatusText = "계산을 완료했습니다."
    Exit Sub
Failed:
    mStatusText = Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub LoadValue(ByVal value As Double)
    mEntry = NxCalculatorStorageValue(value)
    mAccumulator = value
    mHasAccumulator = True
    mPendingOperation = vbNullString
    mStartNewEntry = True
    mResultValue = value
    mHasValidResult = True
    mFormulaText = "셀 값 " & NxCalculatorDisplayValue(value)
    mStatusText = "활성 셀의 숫자값을 가져왔습니다."
End Sub

Public Property Get DisplayText() As String
    DisplayText = Replace$(mEntry, ".", Application.International(xlDecimalSeparator))
End Property

Public Property Get FormulaText() As String
    FormulaText = mFormulaText
End Property

Public Property Get HasValidResult() As Boolean
    HasValidResult = mHasValidResult
End Property

Public Property Get ResultValue() As Double
    If Not mHasValidResult Then NxRaiseContractError "정상 계산 결과가 없습니다."
    ResultValue = mResultValue
End Property

Public Property Get StatusText() As String
    StatusText = mStatusText
End Property

Private Sub ResetCurrent()
    mEntry = "0"
    mAccumulator = 0#
    mHasAccumulator = False
    mPendingOperation = vbNullString
    mStartNewEntry = False
    mHasValidResult = False
    mResultValue = 0#
    mFormulaText = vbNullString
    mStatusText = "계산할 숫자를 입력하세요."
End Sub

Private Sub MarkCurrentValue()
    mResultValue = EntryValue()
    mHasValidResult = True
    mStatusText = "계산할 연산을 선택하세요."
End Sub

Private Function EntryValue() As Double
    EntryValue = Val(mEntry)
End Function

Private Function NxCalculatorStorageValue(ByVal value As Double) As String
    NxCalculatorStorageValue = Replace$(CStr(value), Application.International(xlDecimalSeparator), ".")
End Function

Private Function NxCalculatorDisplayValue(ByVal value As Double) As String
    NxCalculatorDisplayValue = CStr(value)
End Function
