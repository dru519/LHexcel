Option Explicit

Private mAllowClose As Boolean

Public Sub BindSettings(ByVal settings As CNxHangulSettings)
    If settings Is Nothing Then Set settings = NxHangulDefaultSettings()
    cboFont.Value = settings.FontName
    cboSize.Value = CStr(settings.FontSizePt)
    cboStyle.Value = settings.TableStyle
    chkTitleRow.Value = settings.TitleRow
    chkIncludeHidden.Value = settings.IncludeHidden
End Sub

Private Sub cmdApply_Click()
    On Error GoTo Failed
    NxHangulApplySettingsValues CStr(cboFont.Value), CLng(cboSize.Value), _
        CBool(chkTitleRow.Value), CBool(chkIncludeHidden.Value), CStr(cboStyle.Value)
    RequestClose
    Exit Sub
Failed:
    MsgBox NxUserErrorText(Err.Description), vbExclamation + vbOKOnly, "내엑셀 - 아래한글 표 설정"
End Sub

Private Sub cmdCancel_Click(): RequestClose: End Sub

Private Sub RequestClose()
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestClose
End Sub
