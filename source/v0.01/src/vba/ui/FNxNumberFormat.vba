Option Explicit
Private mTarget As Range, mPercentage As Boolean, mEmphasis As Boolean

Public Sub BindTarget(ByVal target As Range, ByVal percentage As Boolean, Optional ByVal emphasis As Boolean = False)
    Set mTarget = target: mPercentage = percentage: mEmphasis = emphasis
    Me.Caption = "내엑셀 - " & IIf(percentage, "백분율", "숫자") & " 표시"
    If emphasis Then Me.Caption = "내엑셀 - 숫자 강조(▲▼)"
    chkPercentage.Visible = emphasis
    chkPercentage.Value = percentage
    lblGuide.Caption = "소수점 아래에 표시할 자릿수를 선택하세요. (0~15)"
    txtPlaces.Value = "2"
    cmdExecute.Default = True: cmdCancel.Cancel = True
    UpdateSample
End Sub
Private Function ReadPlaces() As Long
    Dim value As Double
    If Not IsNumeric(txtPlaces.Value) Then NxRaiseContractError "소수 자릿수는 숫자로 입력하세요."
    value = CDbl(txtPlaces.Value)
    If value < 0 Or value > 15 Or value <> Fix(value) Then NxRaiseContractError "소수 자릿수는 0~15 사이의 정수로 입력하세요."
    ReadPlaces = CLng(value)
End Function
Private Sub txtPlaces_Change(): UpdateSample: End Sub
Private Sub chkPercentage_Click()
    If mEmphasis Then mPercentage = CBool(chkPercentage.Value)
    UpdateSample
End Sub
Private Sub UpdateSample()
    On Error GoTo Invalid
    If mEmphasis Then
        lblSample.Caption = "▲ " & Format$(IIf(mPercentage, 0.123456, 1234.56789), NxDecimalDisplayFormat(ReadPlaces(), mPercentage)) & " (빨강) / ▼ (파랑)"
    Else
        lblSample.Caption = "예: " & Format$(IIf(mPercentage, 0.123456, 1234.56789), NxDecimalDisplayFormat(ReadPlaces(), mPercentage))
    End If
    Exit Sub
Invalid:
    lblSample.Caption = NxUserErrorText(Err.Description): Err.Clear
End Sub
Private Sub cmdExecute_Click()
    On Error GoTo Failed
    If mTarget Is Nothing Then NxRaiseContractError "셀 범위를 다시 선택하세요."
    If mTarget.Worksheet.ProtectContents Or mTarget.Worksheet.Parent.ReadOnly Then NxRaiseContractError "편집 가능한 시트에서 실행하세요."
    If mEmphasis Then
        mTarget.NumberFormat = NxNumberEmphasisFormat(ReadPlaces(), mPercentage)
    Else
        mTarget.NumberFormat = NxDecimalDisplayFormat(ReadPlaces(), mPercentage)
    End If
    Unload Me
    Exit Sub
Failed:
    lblSample.Caption = NxUserErrorText(Err.Description): Err.Clear
End Sub
Private Sub cmdCancel_Click(): Unload Me: End Sub
