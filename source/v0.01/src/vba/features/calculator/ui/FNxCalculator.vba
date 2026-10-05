Option Explicit

Private mSession As CNxCalculatorSession
Private mHistory As CNxCalculatorHistory
Private mHandlers As Collection
Private mAllowClose As Boolean

Private Sub UserForm_Initialize()
    cmdEquals.Default = True
End Sub

Public Sub Bind(ByVal session As CNxCalculatorSession, ByVal history As CNxCalculatorHistory)
    Dim control As Object, handler As CNxCalculatorButtonHandler
    If session Is Nothing Or history Is Nothing Then NxRaiseContractError "계산기 세션과 이력이 필요합니다."
    Set mSession = session
    Set mHistory = history
    Set mHandlers = New Collection
    For Each control In Me.Controls
        If TypeName(control) = "CommandButton" Then
            Set handler = New CNxCalculatorButtonHandler
            handler.Bind control, Me
            mHandlers.Add handler
        End If
    Next control
    RefreshView
End Sub

Public Sub HandleCommand(ByVal commandId As String)
    If Not NxDistributionCanExecute() Then
        ShowCalculatorError NxDistributionMessage()
        Exit Sub
    End If
    Dim imported As Double, statusText As String, prepared As CNxCalculatorCellWrite
    On Error GoTo Failed
    If mSession Is Nothing Then NxRaiseContractError "계산기 세션이 준비되지 않았습니다."
    If Left$(commandId, 6) = "DIGIT:" Then
        mSession.InputDigit Mid$(commandId, 7)
    ElseIf NxCalculatorIsBinaryOperation(commandId) Then
        mSession.SetBinaryOperation commandId
    ElseIf NxCalculatorIsUnaryOperation(commandId) Then
        mSession.ApplyUnary commandId
    Else
        Select Case commandId
            Case "DECIMAL": mSession.InputDecimal
            Case "SIGN": mSession.ToggleSign
            Case "BACKSPACE": mSession.Backspace
            Case "CLEAR_ENTRY": mSession.ClearEntry
            Case "CLEAR_ALL": mSession.ClearAll
            Case "EQUALS": mSession.CompleteEquals
            Case "IMPORT"
                If Not NxCalculatorTryReadActiveNumber(imported, statusText) Then ShowCalculatorError statusText: Exit Sub
                mSession.LoadValue imported
            Case "COPY"
                If Not mSession.HasValidResult Then ShowCalculatorError "정상 계산 결과가 없어 복사할 수 없습니다.": Exit Sub
                NxClipboardWriteUnicode mSession.DisplayText
                RefreshView
                Exit Sub
            Case "PASTE"
                If Not mSession.HasValidResult Then ShowCalculatorError "정상 계산 결과가 없어 셀에 넣을 수 없습니다.": Exit Sub
                Set prepared = NxCalculatorPrepareActiveWrite(mSession.ResultValue)
                If prepared.RequiresConfirmation Then
                    If MsgBox(prepared.TargetLabel & "의 " & prepared.ExistingKind & "을 계산 결과로 바꾸시겠습니까?", vbYesNo + vbExclamation + vbDefaultButton2, "내엑셀 - 계산기") <> vbYes Then
                        Exit Sub
                    End If
                End If
                prepared.Commit
                RefreshView
                Exit Sub
            Case Else: NxRaiseContractError "지원하지 않는 계산기 명령입니다."
        End Select
    End If
    RefreshView
    Exit Sub
Failed:
    statusText = Err.Description
    Err.Clear
    ShowCalculatorError NxUserErrorText(statusText)
    RefreshButtons
End Sub

Public Sub HandleKey(ByRef keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    Dim active As Object
    If (shift And 2) <> 0 And keyCode = vbKeyC Then
        On Error Resume Next
        Set active = Me.ActiveControl
        If Not active Is Nothing Then
            If TypeName(active) = "TextBox" And active.SelLength > 0 Then Exit Sub
        End If
        On Error GoTo 0
        HandleCommand "COPY": keyCode = 0: Exit Sub
    End If
    Select Case keyCode
        Case 56
            If (shift And 1) <> 0 Then HandleCommand "MULTIPLY" Else HandleCommand "DIGIT:8"
            keyCode = 0
        Case vbKey0 To vbKey7, vbKey9: HandleCommand "DIGIT:" & Chr$(keyCode): keyCode = 0
        Case vbKeyNumpad0 To vbKeyNumpad9: HandleCommand "DIGIT:" & CStr(keyCode - vbKeyNumpad0): keyCode = 0
        Case vbKeyAdd: HandleCommand "ADD": keyCode = 0
        Case vbKeySubtract: HandleCommand "SUBTRACT": keyCode = 0
        Case vbKeyMultiply: HandleCommand "MULTIPLY": keyCode = 0
        Case vbKeyDivide: HandleCommand "DIVIDE": keyCode = 0
        Case 187
            If (shift And 1) <> 0 Then HandleCommand "ADD" Else HandleCommand "EQUALS"
            keyCode = 0
        Case 189: HandleCommand "SUBTRACT": keyCode = 0
        Case 191: HandleCommand "DIVIDE": keyCode = 0
        Case vbKeyDecimal, 190: HandleCommand "DECIMAL": keyCode = 0
        Case vbKeyReturn: HandleCommand "EQUALS": keyCode = 0
        Case vbKeyBack: HandleCommand "BACKSPACE": keyCode = 0
        Case vbKeyDelete: HandleCommand "CLEAR_ENTRY": keyCode = 0
        Case vbKeyC: HandleCommand "CLEAR_ALL": keyCode = 0
        Case vbKeyEscape: RequestClose: keyCode = 0
    End Select
End Sub

Public Sub RefreshView()
    txtResult.Value = mSession.DisplayText
    If Len(mSession.FormulaText) = 0 Then txtFormula.Value = "현재 계산식" Else txtFormula.Value = mSession.FormulaText
    txtHistory1.Value = HistoryAt(1)
    txtHistory2.Value = HistoryAt(2)
    txtHistory3.Value = HistoryAt(3)
    RefreshButtons
End Sub

Private Sub ShowCalculatorError(ByVal detail As String)
    If Len(Trim$(detail)) = 0 Then detail = "계산을 완료하지 못했습니다."
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 계산기"
End Sub

Private Sub RefreshButtons()
    Dim enabled As Boolean
    enabled = Not mSession Is Nothing And mSession.HasValidResult
    cmdCopy.Enabled = enabled
    cmdPaste.Enabled = enabled
End Sub

Private Function HistoryAt(ByVal index As Long) As String
    If mHistory Is Nothing Then Exit Function
    If index <= mHistory.Count Then HistoryAt = mHistory.Item(index)
End Function

Private Sub UserForm_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer)
    HandleKey keyCode, shift
End Sub

Private Sub txtResult_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer): HandleKey keyCode, shift: End Sub
Private Sub txtFormula_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer): HandleKey keyCode, shift: End Sub
Private Sub txtHistory1_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer): HandleKey keyCode, shift: End Sub
Private Sub txtHistory2_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer): HandleKey keyCode, shift: End Sub
Private Sub txtHistory3_KeyDown(ByVal keyCode As MSForms.ReturnInteger, ByVal shift As Integer): HandleKey keyCode, shift: End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    mAllowClose = True
    NxCalculatorNotifyClosed Me
End Sub

Private Sub RequestClose()
    If mAllowClose Then Exit Sub
    mAllowClose = True
    NxCalculatorNotifyClosed Me
    Unload Me
End Sub
