Option Explicit

Private mAllowClose As Boolean
Private mTransferMode As Boolean
Private mAccepted As Boolean
Private mRangeSession As CNxRangeSelectionSession

Private Sub UserForm_Initialize()
    NxInitializeStaticComboChoices Me
    lblRange.Visible = False
    txtRange.Visible = False
    cmdUseSelection.Visible = False
End Sub

Public Property Get Accepted() As Boolean: Accepted = mAccepted: End Property

Public Sub SetTransferMode()
    mTransferMode = True
    Me.Caption = "내엑셀 - 아래한글 표 전송"
    cmdApply.Caption = "전송"
    lblRange.Visible = True
    txtRange.Visible = True
    cmdUseSelection.Visible = True
    Set mRangeSession = New CNxRangeSelectionSession
    If TypeName(Application.Selection) = "Range" Then
        mRangeSession.UpdateFromRange Application.Selection
        txtRange.Value = mRangeSession.DisplayAddress
    End If
End Sub

Private Sub cmdUseSelection_Click()
    On Error GoTo Failed
    If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "시트에서 표 범위를 선택하세요."
    mRangeSession.UpdateFromRange Application.Selection
    txtRange.Value = mRangeSession.DisplayAddress
    Exit Sub
Failed:
    MsgBox NxUserErrorText(Err.Description), vbExclamation, "내엑셀 - 범위 선택"
End Sub

Public Sub BindSettings(ByVal settings As CNxHangulSettings)
    If settings Is Nothing Then Set settings = NxHangulDefaultSettings()
    cboFont.Value = settings.FontName
    BindChoice cboSize, CStr(settings.FontSizePt)
    cboStyle.Value = "스타일: " & settings.TableStyle
    BindChoice cboHeaderRows, CStr(settings.HeaderRows)
    chkTotalRow.Value = False
    chkIncludeHidden.Value = False
End Sub

Private Sub BindChoice(ByVal control As Object, ByVal value As String)
    Dim index As Long
    For index = 0 To control.ListCount - 1
        If CStr(control.List(index)) = value Then
            control.ListIndex = index
            Exit Sub
        End If
    Next index
    ' Valid persisted settings can fall outside the suggested dropdown values.
    control.AddItem value
    control.ListIndex = control.ListCount - 1
End Sub

Private Sub cmdApply_Click()
    Dim source As Range, settings As CNxHangulSettings
    On Error GoTo Failed
    If Not IsNumeric(cboHeaderRows.Value) Then NxRaiseContractError "제목행 수를 숫자로 입력하세요."
    If CDbl(cboHeaderRows.Value) <> Fix(CDbl(cboHeaderRows.Value)) Then NxRaiseContractError "제목행 수는 정수로 입력하세요."
    Set settings = New CNxHangulSettings
    settings.Configure CStr(cboFont.Value), "HFT", CLng(cboSize.Value), _
        (CLng(cboHeaderRows.Value) > 0), CBool(chkIncludeHidden.Value), CStr(cboStyle.Value), _
        CLng(cboHeaderRows.Value), CBool(chkTotalRow.Value)
    settings.Seal
    If mTransferMode Then
        NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
        Set source = mRangeSession.ResolveRange()
        cmdApply.Enabled = False
        If Not NxHangulTransferRange(source, settings) Then NxRaiseContractError NxHangulLastError()
    End If
    NxHangulSaveSettings settings
    mAccepted = True
    RequestClose
    Exit Sub
Failed:
    cmdApply.Enabled = True
    MsgBox NxUserErrorText(Err.Description), vbExclamation + vbOKOnly, "내엑셀 - 아래한글 표 설정"
End Sub

Private Sub cmdCancel_Click(): RequestClose: End Sub

Private Sub RequestClose()
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 1 Then Exit Sub
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestClose
End Sub
