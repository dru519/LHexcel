Option Explicit

Public Sub BindSettings()
    Dim values As Variant, i As Long
    values = NxViewPresetValues()
    For i = 1 To 5
        Me.Controls("txtZoom" & CStr(i)).Value = CStr(values(i))
    Next i
    cmdExecute.Default = True: cmdCancel.Cancel = True
End Sub

Private Sub cmdExecute_Click()
    On Error GoTo Failed
    Dim values(1 To 5) As Long, i As Long
    For i = 1 To 5
        values(i) = NxViewPresetZoom(CStr(Me.Controls("txtZoom" & CStr(i)).Value))
    Next i
    NxViewPresetSave values
    Unload Me
    Exit Sub
Failed:
    lblGuide.Caption = NxUserErrorText(Err.Description)
    Err.Clear
End Sub

Private Sub cmdCancel_Click(): Unload Me: End Sub
