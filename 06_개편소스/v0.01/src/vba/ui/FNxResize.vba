Option Explicit
Private mTarget As Range

Public Sub BindTarget(ByVal target As Range)
    Set mTarget = target
    txtTarget.Value = target.Address(External:=True)
    txtWidth.Value = target.Cells(1, 1).ColumnWidth
    txtHeight.Value = target.Cells(1, 1).RowHeight
    txtFont.Value = target.Cells(1, 1).Font.Size
    txtShape.Value = "100"
    cmdExecute.Default = True: cmdCancel.Cancel = True
End Sub

Private Sub cmdExecute_Click()
    On Error GoTo Failed
    NxResizeApply mTarget, CBool(chkWidth.Value), ReadNumber(txtWidth, chkWidth), _
        CBool(chkHeight.Value), ReadNumber(txtHeight, chkHeight), CBool(chkFont.Value), ReadNumber(txtFont, chkFont), _
        CBool(chkShape.Value), ReadNumber(txtShape, chkShape)
    Unload Me
    Exit Sub
Failed:
    lblStatus.Caption = NxUserErrorText(Err.Description): Err.Clear
End Sub

Private Function ReadNumber(ByVal field As Object, ByVal toggle As Object) As Double
    If Not toggle.Value Then Exit Function
    If Not IsNumeric(field.Value) Then NxRaiseContractError toggle.Caption & " 항목에 숫자를 입력하세요."
    ReadNumber = CDbl(field.Value)
End Function
Private Sub cmdCancel_Click(): Unload Me: End Sub
