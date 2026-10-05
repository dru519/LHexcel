Option Explicit

Private mKind As String
Private mFileMode As Boolean
Private mResult As CNxRenameRule
Private mAccepted As Boolean

Private Sub UserForm_Initialize()
    cmdApply.Default = True
    cmdCancel.Cancel = True
End Sub

Public Function EditRule(ByVal source As CNxRenameRule, ByVal fileMode As Boolean) As CNxRenameRule
    mKind = source.Kind
    mFileMode = fileMode
    mAccepted = False
    ConfigureFields
    txtText.Value = source.Text
    txtReplacement.Value = source.Replacement
    txtStart.Value = CStr(source.Start)
    txtStep.Value = CStr(source.StepValue)
    txtDigits.Value = CStr(source.Digits)
    chkCase.Value = source.CaseSensitive
    chkFirst.Value = source.FirstOnly
    chkEmptyExtension.Value = source.AllowEmptyExtension
    If mKind = "CASE" Then
        Select Case source.Position
            Case "UPPER": cboPosition.ListIndex = 1
            Case "PROPER": cboPosition.ListIndex = 2
            Case Else: cboPosition.ListIndex = 0
        End Select
    ElseIf cboPosition.ListCount > 0 Then
        cboPosition.ListIndex = IIf(source.Position = "FRONT", 0, 1)
    End If
    Me.Show vbModal
    If mAccepted Then Set EditRule = mResult
End Function

Private Sub ConfigureFields()
    Dim control As Object
    For Each control In Me.Controls
        control.Visible = False
    Next control
    lblTitle.Visible = True
    cmdApply.Visible = True
    cmdCancel.Visible = True
    Select Case mKind
        Case "REPLACE"
            lblTitle.Caption = "문자열 바꾸기"
            ShowText "찾을 문구"
            lblReplacement.Visible = True: txtReplacement.Visible = True
            chkCase.Visible = True: chkFirst.Visible = True
        Case "PREFIX"
            lblTitle.Caption = "앞에 문구 추가": ShowText "추가할 문구"
        Case "SUFFIX"
            lblTitle.Caption = "뒤에 문구 추가": ShowText "추가할 문구"
        Case "NUMBER"
            lblTitle.Caption = "번호 붙이기": ShowText "구분 문자"
            ShowPosition Array("앞", "뒤")
            lblStart.Visible = True: txtStart.Visible = True
            lblStep.Visible = True: txtStep.Visible = True
            lblDigits.Visible = True: txtDigits.Visible = True
        Case "DELETE"
            lblTitle.Caption = "일부 문자 지우기"
            ShowPosition Array("앞에서", "뒤에서")
            lblStart.Caption = "지울 글자 수"
            lblStart.Visible = True: txtStart.Visible = True
        Case "FULL"
            lblTitle.Caption = "이름 전체 바꾸기": ShowText "새 이름 (파일 확장자는 유지)"
            If Not mFileMode Then lblText.Caption = "새 시트 이름"
        Case "CASE"
            lblTitle.Caption = "영문 대소문자"
            lblPosition.Caption = "변환"
            ShowPosition Array("모두 소문자", "모두 대문자", "단어 첫 글자")
        Case "EXTENSION"
            If Not mFileMode Then NxRaiseContractError "확장자 규칙은 파일 전용입니다."
            lblTitle.Caption = "확장자 변경": ShowText "새 확장자"
            chkEmptyExtension.Visible = True
        Case Else
            NxRaiseContractError "지원하지 않는 이름 변경 규칙입니다."
    End Select
End Sub

Private Sub ShowText(ByVal text As String)
    lblText.Caption = text
    lblText.Visible = True
    txtText.Visible = True
End Sub

Private Sub ShowPosition(ByVal values As Variant)
    Dim value As Variant
    lblPosition.Visible = True
    cboPosition.Visible = True
    cboPosition.Clear
    For Each value In values
        cboPosition.AddItem CStr(value)
    Next value
    cboPosition.ListIndex = 0
End Sub

Private Sub cmdApply_Click()
    Dim rule As New CNxRenameRule
    On Error GoTo Failed
    rule.Kind = mKind
    rule.Text = CStr(txtText.Value)
    rule.Replacement = CStr(txtReplacement.Value)
    rule.CaseSensitive = CBool(chkCase.Value)
    rule.FirstOnly = CBool(chkFirst.Value)
    rule.AllowEmptyExtension = CBool(chkEmptyExtension.Value)
    Select Case mKind
        Case "NUMBER", "DELETE"
            rule.Position = IIf(cboPosition.ListIndex = 0, "FRONT", "BACK")
            rule.Start = ParseInteger(CStr(txtStart.Value), "시작/글자 수")
            If mKind = "NUMBER" Then
                rule.StepValue = ParseInteger(CStr(txtStep.Value), "증가값")
                rule.Digits = ParseInteger(CStr(txtDigits.Value), "자릿수")
            End If
        Case "CASE"
            rule.Position = CStr(Array("LOWER", "UPPER", "PROPER")(cboPosition.ListIndex))
    End Select
    rule.Validate mFileMode
    Set mResult = rule
    mAccepted = True
    Me.Hide
    Exit Sub
Failed:
    MsgBox NxUserErrorText(Err.Description), vbExclamation, "내엑셀 - 규칙"
End Sub

Private Function ParseInteger(ByVal value As String, ByVal field As String) As Long
    Dim number As Double
    If Not IsNumeric(value) Then NxRaiseContractError field & "에 정수를 입력하세요."
    number = CDbl(value)
    If number <> Fix(number) Or number < -2147483648# Or number > 2147483647# Then NxRaiseContractError field & "의 정수 범위를 확인하세요."
    ParseInteger = CLng(number)
End Function

Private Sub cmdCancel_Click()
    mAccepted = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        cmdCancel_Click
    End If
End Sub
