Attribute VB_Name = "NxUserErrors"
Option Explicit

' Converts internal (non-Korean) error text into a Korean message with a code.
' The original text is kept only in the local diagnostics log.
Private Const NX_USER_ERROR_LOG As String = "errors.log"

Public Function NxUserErrorText(ByVal detail As String) As String
    Dim code As String
    detail = Trim$(detail)
    If Len(detail) = 0 Then
        NxUserErrorText = "작업을 완료하지 못했습니다. 다시 실행해 주세요."
        Exit Function
    End If
    If NxUserErrorHasHangul(detail) Then
        NxUserErrorText = detail
        Exit Function
    End If
    code = NxUserErrorCode(detail)
    NxUserErrorLog code, detail
    NxUserErrorText = "작업을 완료하지 못했습니다. 선택 범위와 파일 상태를 확인한 뒤 다시 실행해 주세요." & vbCrLf & _
        "같은 문제가 반복되면 오류 코드 " & code & "를 알려 주세요."
End Function

Public Function NxUserErrorHasHangul(ByVal text As String) As Boolean
    Dim index As Long, code As Long
    For index = 1 To Len(text)
        code = AscW(Mid$(text, index, 1))
        If code < 0 Then code = code + 65536
        If code >= &HAC00& And code <= &HD7A3& Then
            NxUserErrorHasHangul = True
            Exit Function
        End If
    Next index
End Function

Public Function NxUserErrorCode(ByVal detail As String) As String
    Dim index As Long, hash As Double
    For index = 1 To Len(detail)
        hash = hash * 31# + AscW(Mid$(detail, index, 1))
        hash = hash - Int(hash / 999983#) * 999983#
    Next index
    NxUserErrorCode = "NX-" & Format$(hash, "000000")
End Function

Private Sub NxUserErrorLog(ByVal code As String, ByVal detail As String)
    Dim folder As String, handle As Integer
    On Error GoTo Done
    folder = NxLHexcelProfileRoot()
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    folder = folder & "\Diagnostics"
    If Len(Dir$(folder, vbDirectory)) = 0 Then MkDir folder
    handle = FreeFile
    Open folder & "\" & NX_USER_ERROR_LOG For Append Access Write As #handle
    Print #handle, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & code & vbTab & Replace$(Replace$(detail, vbCr, " "), vbLf, " ")
Done:
    On Error Resume Next
    If handle <> 0 Then Close #handle
End Sub

