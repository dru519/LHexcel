Option Explicit

Private mPath As String
Public Sub BindPath(ByVal path As String)
    mPath = path
    lblPath.Caption = "사본을 저장했습니다." & vbCrLf & CreateObject("Scripting.FileSystemObject").GetFileName(path)
    lblPath.ControlTipText = path
    cmdClose.Cancel = True
    cmdClose.Default = True
End Sub
Private Sub cmdFolder_Click()
    On Error GoTo Failed
    CreateObject("Shell.Application").Open CreateObject("Scripting.FileSystemObject").GetParentFolderName(mPath)
    Unload Me
    Exit Sub
Failed:
    MsgBox NxUserErrorText(Err.Description), vbExclamation, "내엑셀 - 폴더 열기"
End Sub
Private Sub cmdFile_Click()
    On Error GoTo Failed
    Application.Workbooks.Open Filename:=mPath, UpdateLinks:=0
    Unload Me
    Exit Sub
Failed:
    MsgBox NxUserErrorText(Err.Description), vbExclamation, "내엑셀 - 파일 열기"
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
