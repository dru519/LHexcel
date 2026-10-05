Attribute VB_Name = "NxHangulTableTransfer"
Option Explicit

Private Const NX_HANGUL_TITLE As String = "내엑셀 - 아래한글 표 전송하기"
Private Const NX_HANGUL_MAX_CELLS As Long = 16384
Private mHangulLastError As String
Private mHangulLastOutputPath As String

Public Sub NxHangulSendSelection()
    Dim source As Range
    On Error GoTo Failed
    mHangulLastError = vbNullString
    mHangulLastOutputPath = vbNullString
    Set source = NxHangulPickSourceRange()
    If source Is Nothing Then Exit Sub
    If Not NxHangulTransferRange(source) Then NxRaiseContractError mHangulLastError
    Exit Sub
Failed:
    If Len(Err.Description) = 0 Then
        MsgBox "아래한글 표 전송을 완료하지 못했습니다.", vbExclamation + vbOKOnly, NX_HANGUL_TITLE
    Else
        MsgBox NxUserErrorText(Err.Description), vbExclamation + vbOKOnly, NX_HANGUL_TITLE
    End If
End Sub

Private Function NxHangulPickSourceRange() As Range
    Dim selected As Range, defaultAddress As String
    If TypeName(Application.Selection) = "Range" Then _
        defaultAddress = Application.Selection.Address(External:=True)
    ' Type 8 returns the dragged/typed Range. Cancel returns False (Set error 424).
    On Error Resume Next
    Set selected = Application.InputBox(Prompt:="아래한글로 보낼 표 범위를 드래그하거나 주소를 입력하세요.", _
        Title:=NX_HANGUL_TITLE & " - 범위 선택", Default:=defaultAddress, Type:=8)
    On Error GoTo 0
    Set NxHangulPickSourceRange = selected
End Function

Public Function NxHangulTransferCurrentSelection() As Boolean
    NxHangulTransferCurrentSelection = NxHangulTransferRange(Nothing)
End Function

Private Function NxHangulTransferRange(ByVal selected As Range) As Boolean
    Dim source As Range
    Dim settings As CNxHangulSettings
    On Error GoTo Failed
    mHangulLastError = vbNullString
    mHangulLastOutputPath = vbNullString
    Set source = NxHangulResolveSelection(selected)
    Set settings = NxHangulLoadSettings()
    mHangulLastOutputPath = NxHwpxSendTable(source, settings)
    NxHangulTransferRange = (Len(mHangulLastOutputPath) > 0)
    Exit Function
Failed:
    mHangulLastError = Err.Source & ": " & Err.Description
    If Len(mHangulLastError) = 0 Then mHangulLastError = "아래한글 표 전송을 완료하지 못했습니다."
    NxHangulTransferRange = False
End Function

Public Function NxHangulLastError() As String
    NxHangulLastError = mHangulLastError
End Function

Public Function NxHangulLastOutputPath() As String
    NxHangulLastOutputPath = mHangulLastOutputPath
End Function

Public Sub NxHangulOpenSettings()
    NxHangulSettings.NxHangulShowSettings
End Sub

Private Function NxHangulResolveSelection(ByVal selected As Range) As Range
    If selected Is Nothing Then
        If TypeName(Application.Selection) <> "Range" Then NxRaiseContractError "아래한글로 보낼 범위를 선택하세요."
        Set selected = Application.Selection
    End If
    If selected.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 전송할 수 있습니다."
    If selected.Parent.Parent Is ThisWorkbook Then NxRaiseContractError "일반 Excel 통합문서의 범위를 선택하세요."
    If selected.Rows.Count = selected.Parent.Rows.Count Or selected.Columns.Count = selected.Parent.Columns.Count Then _
        NxRaiseContractError "전체 행 또는 전체 열이 아닌 실제 표 범위를 선택하세요."
    If CDbl(selected.CountLarge) < 1 Or CDbl(selected.CountLarge) > NX_HANGUL_MAX_CELLS Then _
        NxRaiseContractError "아래한글 표는 최대 16,384셀입니다."
    Set NxHangulResolveSelection = selected
End Function

Public Sub NxHangulReleaseSession()
    ' HWPX is opened as an independent document; there is no COM session to retain.
End Sub

Public Sub NxHangulCloseSession()
    ' Kept as a compatibility no-op for older callers.
End Sub
