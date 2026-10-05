Attribute VB_Name = "T_Calculator"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestDigitAndDecimalInput"
    names.Add "TestLocalDecimalSeparatorAccepted"
    names.Add "TestToggleSignAndBackspace"
    names.Add "TestClearEntryKeepsPendingOperation"
    names.Add "TestClearAllKeepsHistory"
    names.Add "TestBasicOperationsAreLeftToRight"
    names.Add "TestDivideByZeroPreservesPreviousResult"
    names.Add "TestReciprocal"
    names.Add "TestSquare"
    names.Add "TestCube"
    names.Add "TestSquareRootDomain"
    names.Add "TestCubeRootNegative"
    names.Add "TestPower"
    names.Add "TestIntegerRootContract"
    names.Add "TestTenPowerOverflow"
    names.Add "TestFactorialBounds"
    names.Add "TestEqualsAddsOneHistoryRow"
    names.Add "TestHistoryKeepsLatestThree"
    names.Add "TestHistoryAllowsDuplicates"
    names.Add "TestNonEqualsDoesNotAddHistory"
    names.Add "TestLoadValueDoesNotAddHistory"
    names.Add "TestReadNumericCellAndFormulaResult"
    names.Add "TestRejectBlankBooleanDateAndNumericText"
    names.Add "TestEmptyCellWriteIsFast"
    names.Add "TestOccupiedCellRequiresConfirmation"
    names.Add "TestWriteRevalidatesTargetIdentity"
    names.Add "TestProtectedReadOnlyMergedWriteRejected"
    names.Add "TestCalculatorRegistryAndFormContract"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestDigitAndDecimalInput": TestDigitAndDecimalInput
        Case "TestLocalDecimalSeparatorAccepted": TestLocalDecimalSeparatorAccepted
        Case "TestToggleSignAndBackspace": TestToggleSignAndBackspace
        Case "TestClearEntryKeepsPendingOperation": TestClearEntryKeepsPendingOperation
        Case "TestClearAllKeepsHistory": TestClearAllKeepsHistory
        Case "TestBasicOperationsAreLeftToRight": TestBasicOperationsAreLeftToRight
        Case "TestDivideByZeroPreservesPreviousResult": TestDivideByZeroPreservesPreviousResult
        Case "TestReciprocal": TestReciprocal
        Case "TestSquare": TestSquare
        Case "TestCube": TestCube
        Case "TestSquareRootDomain": TestSquareRootDomain
        Case "TestCubeRootNegative": TestCubeRootNegative
        Case "TestPower": TestPower
        Case "TestIntegerRootContract": TestIntegerRootContract
        Case "TestTenPowerOverflow": TestTenPowerOverflow
        Case "TestFactorialBounds": TestFactorialBounds
        Case "TestEqualsAddsOneHistoryRow": TestEqualsAddsOneHistoryRow
        Case "TestHistoryKeepsLatestThree": TestHistoryKeepsLatestThree
        Case "TestHistoryAllowsDuplicates": TestHistoryAllowsDuplicates
        Case "TestNonEqualsDoesNotAddHistory": TestNonEqualsDoesNotAddHistory
        Case "TestLoadValueDoesNotAddHistory": TestLoadValueDoesNotAddHistory
        Case "TestReadNumericCellAndFormulaResult": TestReadNumericCellAndFormulaResult
        Case "TestRejectBlankBooleanDateAndNumericText": TestRejectBlankBooleanDateAndNumericText
        Case "TestEmptyCellWriteIsFast": TestEmptyCellWriteIsFast
        Case "TestOccupiedCellRequiresConfirmation": TestOccupiedCellRequiresConfirmation
        Case "TestWriteRevalidatesTargetIdentity": TestWriteRevalidatesTargetIdentity
        Case "TestProtectedReadOnlyMergedWriteRejected": TestProtectedReadOnlyMergedWriteRejected
        Case "TestCalculatorRegistryAndFormContract": TestCalculatorRegistryAndFormContract
        Case Else: NxRaiseContractError "Unknown Calculator test"
    End Select
End Sub

Private Function NewSession(ByRef history As CNxCalculatorHistory) As CNxCalculatorSession
    Dim session As New CNxCalculatorSession
    Set history = New CNxCalculatorHistory
    session.BindHistory history
    Set NewSession = session
End Function

Private Sub AssertNear(ByVal expected As Double, ByVal actual As Double, ByVal message As String)
    NxTestHarness.AssertTrue Abs(expected - actual) < 0.0000001, message
End Sub

Private Sub TestDigitAndDecimalInput()
    Dim history As CNxCalculatorHistory, session As CNxCalculatorSession
    Set session = NewSession(history)
    session.InputDigit "1": session.InputDigit "2": session.InputDecimal: session.InputDigit "5"
    AssertNear 12.5, session.ResultValue, "Digit and decimal input changed"
End Sub

Private Sub TestLocalDecimalSeparatorAccepted()
    Dim history As CNxCalculatorHistory, session As CNxCalculatorSession
    Set session = NewSession(history)
    session.InputDigit "1": session.InputDecimal: session.InputDigit "5"
    NxTestHarness.AssertTrue InStr(session.DisplayText, Application.International(xlDecimalSeparator)) > 0, "Local decimal separator was not displayed"
End Sub

Private Sub TestToggleSignAndBackspace()
    Dim history As CNxCalculatorHistory, session As CNxCalculatorSession
    Set session = NewSession(history)
    session.InputDigit "1": session.InputDigit "2": session.ToggleSign: session.Backspace
    AssertNear -1#, session.ResultValue, "Sign or backspace changed"
End Sub

Private Sub TestClearEntryKeepsPendingOperation()
    Dim history As CNxCalculatorHistory, session As CNxCalculatorSession
    Set session = NewSession(history)
    session.InputDigit "1": session.InputDigit "2": session.SetBinaryOperation "ADD"
    session.InputDigit "3": session.ClearEntry: session.InputDigit "4": session.CompleteEquals
    AssertNear 16#, session.ResultValue, "ClearEntry removed pending operation"
End Sub

Private Sub TestClearAllKeepsHistory()
    Dim history As CNxCalculatorHistory, session As CNxCalculatorSession
    Set session = NewSession(history)
    session.InputDigit "2": session.CompleteEquals: session.ClearAll
    NxTestHarness.AssertTrue history.Count = 1, "ClearAll removed session history"
End Sub

Private Sub TestBasicOperationsAreLeftToRight()
    Dim history As CNxCalculatorHistory, session As CNxCalculatorSession
    Set session = NewSession(history)
    session.InputDigit "2": session.SetBinaryOperation "ADD": session.InputDigit "3"
    session.SetBinaryOperation "MULTIPLY": session.InputDigit "4": session.CompleteEquals
    AssertNear 20#, session.ResultValue, "Basic operations were not evaluated left to right"
End Sub

Private Sub TestDivideByZeroPreservesPreviousResult()
    Dim history As CNxCalculatorHistory, session As CNxCalculatorSession, failed As Boolean
    Set session = NewSession(history)
    session.InputDigit "8": session.SetBinaryOperation "DIVIDE": session.InputDigit "0"
    On Error Resume Next: session.CompleteEquals: failed = Err.Number <> 0: Err.Clear: On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Division by zero did not fail"
    NxTestHarness.AssertTrue history.Count = 0, "Division by zero added history"
End Sub

Private Sub TestReciprocal()
    AssertNear 0.25, NxCalculatorApplyUnary("RECIPROCAL", 4#), "Reciprocal changed"
End Sub

Private Sub TestSquare()
    AssertNear 25#, NxCalculatorApplyUnary("SQUARE", 5#), "Square changed"
End Sub

Private Sub TestCube()
    AssertNear -8#, NxCalculatorApplyUnary("CUBE", -2#), "Cube changed"
End Sub

Private Sub TestSquareRootDomain()
    Dim failed As Boolean
    AssertNear 3#, NxCalculatorApplyUnary("SQRT", 9#), "Square root changed"
    On Error Resume Next: Call NxCalculatorApplyUnary("SQRT", -1#): failed = Err.Number <> 0: Err.Clear: On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Negative square root did not fail"
End Sub

Private Sub TestCubeRootNegative()
    AssertNear -3#, NxCalculatorApplyUnary("CBRT", -27#), "Negative cube root changed"
End Sub

Private Sub TestPower()
    Dim failed As Boolean
    AssertNear 256#, NxCalculatorApplyBinary(2#, "POWER", 8#), "Power changed"
    On Error Resume Next: Call NxCalculatorApplyBinary(-2#, "POWER", 0.5): failed = Err.Number <> 0: Err.Clear: On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Fractional power of negative base did not fail"
End Sub

Private Sub TestIntegerRootContract()
    Dim failed As Boolean
    AssertNear -3#, NxCalculatorApplyBinary(-27#, "YROOT", 3#), "Odd integer root changed"
    On Error Resume Next: Call NxCalculatorApplyBinary(-16#, "YROOT", 2#): failed = Err.Number <> 0: Err.Clear: On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Even root of negative value did not fail"
End Sub

Private Sub TestTenPowerOverflow()
    Dim failed As Boolean
    AssertNear 1000#, NxCalculatorApplyUnary("TEN_POWER", 3#), "Ten power changed"
    On Error Resume Next: Call NxCalculatorApplyUnary("TEN_POWER", 400#): failed = Err.Number <> 0: Err.Clear: On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Ten power overflow did not fail"
End Sub

Private Sub TestFactorialBounds()
    Dim failed As Boolean
    AssertNear 120#, NxCalculatorFactorial(5#), "Factorial changed"
    On Error Resume Next: Call NxCalculatorFactorial(171#): failed = Err.Number <> 0: Err.Clear: On Error GoTo 0
    NxTestHarness.AssertTrue failed, "Factorial upper bound did not fail"
End Sub

Private Sub TestEqualsAddsOneHistoryRow()
    Dim history As CNxCalculatorHistory, session As CNxCalculatorSession
    Set session = NewSession(history)
    session.InputDigit "3": session.SetBinaryOperation "ADD": session.InputDigit "4": session.CompleteEquals
    NxTestHarness.AssertTrue history.Count = 1 And InStr(history.Item(1), "7") > 0, "Equals did not add one history row"
End Sub

Private Sub TestHistoryKeepsLatestThree()
    Dim history As New CNxCalculatorHistory
    history.Add "1": history.Add "2": history.Add "3": history.Add "4"
    NxTestHarness.AssertTrue history.Count = 3 And history.Item(1) = "4" And history.Item(3) = "2", "History did not keep latest three"
End Sub

Private Sub TestHistoryAllowsDuplicates()
    Dim history As New CNxCalculatorHistory
    history.Add "same": history.Add "same"
    NxTestHarness.AssertTrue history.Count = 2, "History removed a duplicate"
End Sub

Private Sub TestNonEqualsDoesNotAddHistory()
    Dim history As CNxCalculatorHistory, session As CNxCalculatorSession
    Set session = NewSession(history)
    session.InputDigit "9": session.ApplyUnary "SQRT"
    NxTestHarness.AssertTrue history.Count = 0, "Non-equals action added history"
End Sub

Private Sub TestLoadValueDoesNotAddHistory()
    Dim history As CNxCalculatorHistory, session As CNxCalculatorSession
    Set session = NewSession(history)
    session.LoadValue 42#
    NxTestHarness.AssertTrue history.Count = 0, "Cell import added history"
    AssertNear 42#, session.ResultValue, "Loaded value changed"
End Sub

Private Sub TestReadNumericCellAndFormulaResult()
    Dim book As Workbook, target As Range, number As Double, status As String, detail As String
    On Error GoTo Failed
    Set book = Application.ActiveWorkbook
    If book Is Nothing Or book Is ThisWorkbook Then NxRaiseContractError "Calculator test data workbook is unavailable"
    Set target = book.Worksheets(1).Range("B2")
    target.Clear: target.Value2 = 12#: target.Select
    NxTestHarness.AssertTrue NxCalculatorTryReadActiveNumber(number, status), "Numeric cell was rejected"
    AssertNear 12#, number, "Numeric cell result changed"
    target.Formula = "=6*7": Application.Calculate
    NxTestHarness.AssertTrue NxCalculatorTryReadActiveNumber(number, status), "Numeric formula result was rejected"
    AssertNear 42#, number, "Formula result changed"
    target.Clear: book.Worksheets(1).Range("A1").Select
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not target Is Nothing Then target.Clear: If Not book Is Nothing Then book.Worksheets(1).Range("A1").Select: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestRejectBlankBooleanDateAndNumericText()
    Dim book As Workbook, target As Range, number As Double, status As String, detail As String
    On Error GoTo Failed
    Set book = Application.ActiveWorkbook
    If book Is Nothing Or book Is ThisWorkbook Then NxRaiseContractError "Calculator test data workbook is unavailable"
    Set target = book.Worksheets(1).Range("B2"): target.Clear: target.Select
    target.Clear: NxTestHarness.AssertTrue Not NxCalculatorTryReadActiveNumber(number, status), "Blank cell was accepted"
    target.Value2 = True: NxTestHarness.AssertTrue Not NxCalculatorTryReadActiveNumber(number, status), "Boolean cell was accepted"
    target.Value = DateSerial(2026, 8, 9): target.NumberFormat = "yyyy-mm-dd": NxTestHarness.AssertTrue Not NxCalculatorTryReadActiveNumber(number, status), "Date cell was accepted"
    target.NumberFormat = "@": target.Value2 = "123": NxTestHarness.AssertTrue Not NxCalculatorTryReadActiveNumber(number, status), "Numeric text was accepted"
    target.Clear: target.NumberFormat = "General": book.Worksheets(1).Range("A1").Select
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not target Is Nothing Then target.Clear: target.NumberFormat = "General": If Not book Is Nothing Then book.Worksheets(1).Range("A1").Select: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestEmptyCellWriteIsFast()
    Dim book As Workbook, target As Range, prepared As CNxCalculatorCellWrite, detail As String
    On Error GoTo Failed
    Set book = Application.ActiveWorkbook
    If book Is Nothing Or book Is ThisWorkbook Then NxRaiseContractError "Calculator test data workbook is unavailable"
    Set target = book.Worksheets(1).Range("B2"): target.Clear: target.Select
    Set prepared = NxCalculatorPrepareActiveWrite(15#)
    NxTestHarness.AssertTrue Not prepared.RequiresConfirmation, "Empty cell required confirmation"
    prepared.Commit: AssertNear 15#, CDbl(target.Value2), "Empty cell write changed"
    target.Clear: book.Worksheets(1).Range("A1").Select
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not target Is Nothing Then target.Clear: If Not book Is Nothing Then book.Worksheets(1).Range("A1").Select: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestOccupiedCellRequiresConfirmation()
    Dim book As Workbook, target As Range, prepared As CNxCalculatorCellWrite, detail As String
    On Error GoTo Failed
    Set book = Application.ActiveWorkbook
    If book Is Nothing Or book Is ThisWorkbook Then NxRaiseContractError "Calculator test data workbook is unavailable"
    Set target = book.Worksheets(1).Range("B2"): target.Value2 = 1#: target.Select
    Set prepared = NxCalculatorPrepareActiveWrite(15#)
    NxTestHarness.AssertTrue prepared.RequiresConfirmation And prepared.ExistingKind = "기존 값", "Occupied cell confirmation contract changed"
    target.Clear: book.Worksheets(1).Range("A1").Select
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not target Is Nothing Then target.Clear: If Not book Is Nothing Then book.Worksheets(1).Range("A1").Select: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestWriteRevalidatesTargetIdentity()
    Dim book As Workbook, prepared As CNxCalculatorCellWrite, failed As Boolean, detail As String
    On Error GoTo Failed
    Set book = Application.ActiveWorkbook
    If book Is Nothing Or book Is ThisWorkbook Then NxRaiseContractError "Calculator test data workbook is unavailable"
    book.Worksheets(1).Range("B2:C2").Clear: book.Worksheets(1).Range("B2").Select
    Set prepared = NxCalculatorPrepareActiveWrite(15#): book.Worksheets(1).Range("C2").Select
    On Error Resume Next: prepared.Commit: failed = Err.Number <> 0: Err.Clear: On Error GoTo Failed
    NxTestHarness.AssertTrue failed And IsEmpty(book.Worksheets(1).Range("B2").Value2), "Target drift was not rejected"
    book.Worksheets(1).Range("B2:C2").Clear: book.Worksheets(1).Range("A1").Select
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not book Is Nothing Then book.Worksheets(1).Range("B2:C2").Clear: book.Worksheets(1).Range("A1").Select: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestProtectedReadOnlyMergedWriteRejected()
    Dim book As Workbook, readOnlyBook As Workbook, failed As Boolean, path As String, detail As String
    On Error GoTo Failed
    Set book = Application.ActiveWorkbook
    If book Is Nothing Or book Is ThisWorkbook Then NxRaiseContractError "Calculator test data workbook is unavailable"
    book.Worksheets(1).Range("B2:C2").Clear: book.Worksheets(1).Range("B2:C2").Merge: book.Worksheets(1).Range("B2").Select
    On Error Resume Next: Call NxCalculatorPrepareActiveWrite(1#): failed = Err.Number <> 0: Err.Clear: On Error GoTo Failed
    NxTestHarness.AssertTrue failed, "Merged target was accepted"
    book.Worksheets(1).Range("B2:C2").UnMerge: book.Worksheets(1).Protect: book.Worksheets(1).Range("B2").Select: failed = False
    On Error Resume Next: Call NxCalculatorPrepareActiveWrite(1#): failed = Err.Number <> 0: Err.Clear: On Error GoTo Failed
    NxTestHarness.AssertTrue failed, "Protected locked target was accepted"
    book.Worksheets(1).Unprotect
    path = Environ$("TEMP") & "\NaeExcel-calculator-readonly-" & NxCreateRunUuid() & ".xlsx"
    book.SaveCopyAs path
    Set readOnlyBook = ThisWorkbook.Application.Workbooks.Open(path, ReadOnly:=True): readOnlyBook.Activate: readOnlyBook.Worksheets(1).Range("A1").Select: failed = False
    On Error Resume Next: Call NxCalculatorPrepareActiveWrite(1#): failed = Err.Number <> 0: Err.Clear: On Error GoTo Failed
    NxTestHarness.AssertTrue failed, "Read-only target was accepted"
    readOnlyBook.Close SaveChanges:=False: Set readOnlyBook = Nothing: book.Activate: book.Worksheets(1).Range("B2:C2").Clear: book.Worksheets(1).Range("A1").Select: Kill path
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not readOnlyBook Is Nothing Then readOnlyBook.Close SaveChanges:=False: If Not book Is Nothing Then book.Worksheets(1).Unprotect: book.Worksheets(1).Range("B2:C2").UnMerge: book.Worksheets(1).Range("B2:C2").Clear: book.Activate: book.Worksheets(1).Range("A1").Select: If Len(path) > 0 And Len(Dir$(path)) > 0 Then Kill path: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestCalculatorRegistryAndFormContract()
    Dim registry As CNxFeatureRegistry, definition As CNxFeatureDefinition, form As FNxCalculator
    Set registry = CreateBaseRegistry()
    NxCalculatorRegisterFeatures registry
    Set definition = registry.FeatureById(NX_FEATURE_UTIL_CALCULATOR)
    NxTestHarness.AssertTrue definition.CategoryId = "NX-CAT-CALCULATOR", "Calculator registry category changed"
    Set form = New FNxCalculator
    NxTestHarness.AssertTrue form.Controls("txtResult").Locked And form.Controls("cmdEquals").Default, "Calculator form contract changed"
    Unload form
End Sub
