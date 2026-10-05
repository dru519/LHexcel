Option Explicit
Private mTarget As Range, mErrorMode As Boolean, mBound As Boolean

Public Sub BindTarget(ByVal target As Range, ByVal errorMode As Boolean)
    Dim item As Variant
    mBound = False
    Set mTarget = target: mErrorMode = errorMode
    Me.Caption = "내엑셀 - " & IIf(errorMode, "IFERROR", "ROUND") & " 감싸기"
    lblRange.Caption = target.Address(External:=True)
    cboMode.Clear
    cboPrecision.Clear
    For Each item In Array("소수 4자리 (4)", "소수 3자리 (3)", "소수 2자리 (2)", "소수 1자리 (1)", "정수 (0)", "십 단위 (-1)", "백 단위 (-2)", "천 단위 (-3)", "만 단위 (-4)", "직접 입력")
        cboPrecision.AddItem CStr(item)
    Next item
    cboPrecision.ListIndex = 2
    cboPrecision.Visible = Not errorMode
    lblSample.BackColor = vbWhite
    lblSample.BackStyle = fmBackStyleOpaque
    If errorMode Then
        cboMode.AddItem "오류 시 공란"
        cboMode.AddItem "오류 시 숫자"
        cboMode.AddItem "오류 시 문자"
        lblArgument.Caption = "대체 값"
        txtArgument.Value = "0"
    Else
        cboMode.AddItem "반올림 (ROUND)"
        cboMode.AddItem "올림 (ROUNDUP)"
        cboMode.AddItem "버림 (ROUNDDOWN)"
        lblArgument.Caption = "자릿수"
        txtArgument.Value = "2"
    End If
    lblGuide.Caption = IIf(errorMode, "숫자·수식에 적용하며 빈 셀과 문자 상수는 제외합니다.", "기본 제안값 외 조정이 필요하면 직접 입력을 선택하세요. 양수는 소수 자릿수, 음수는 십(-1)·백(-2) 등의 단위입니다.") & vbCrLf & "날짜·백분율도 실제 숫자값에 적용합니다. 배열·스필·표 셀은 제외합니다."
    cboMode.ListIndex = 0
    cmdExecute.Default = True: cmdCancel.Cancel = True
    mBound = True
    UpdateSample
End Sub

Private Function FunctionName() As String
    If mErrorMode Then FunctionName = "IFERROR" Else FunctionName = Choose(cboMode.ListIndex + 1, "ROUND", "ROUNDUP", "ROUNDDOWN")
End Function

Private Function Argument() As String
    Argument = NxFunctionArgument(mErrorMode, cboMode.ListIndex, CStr(txtArgument.Value))
End Function

Private Sub UpdateSample()
    If Not mBound Then Exit Sub
    On Error GoTo Invalid
    If mErrorMode Then
        txtArgument.Enabled = (cboMode.ListIndex <> 0)
    Else
        txtArgument.Enabled = (cboPrecision.ListIndex = 9)
    End If
    lblSample.Caption = NxFunctionPreview(mTarget, FunctionName(), Argument(), CBool(chkHidden.Value))
    cmdExecute.Enabled = True
    Exit Sub
Invalid:
    lblSample.Caption = NxUserErrorText(Err.Description): Err.Clear
    cmdExecute.Enabled = False
End Sub

Private Sub cboMode_Change(): UpdateSample: End Sub
Private Sub cboPrecision_Change()
    If Not mBound Then Exit Sub
    If cboPrecision.ListIndex >= 0 And cboPrecision.ListIndex < 9 Then txtArgument.Value = CStr(4 - cboPrecision.ListIndex)
    UpdateSample
End Sub
Private Sub txtArgument_Change(): UpdateSample: End Sub
Private Sub chkHidden_Click(): UpdateSample: End Sub

Private Sub cmdExecute_Click()
    On Error GoTo Failed
    NxFunctionApply mTarget, FunctionName(), Argument(), CBool(chkHidden.Value)
    Unload Me
    Exit Sub
Failed:
    lblSample.Caption = NxUserErrorText(Err.Description): Err.Clear
End Sub
Private Sub cmdCancel_Click(): Unload Me: End Sub
