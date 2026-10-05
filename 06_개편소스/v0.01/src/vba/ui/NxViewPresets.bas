Attribute VB_Name = "NxViewPresets"
Option Explicit

Private mStates As Object

Public Function NxViewPresetValues() As Variant
    Dim raw As String, parts As Variant, result(1 To 5) As Long, i As Long
    raw = GetSetting("LHExcel", "ViewPresets", "Zooms", "75,100,125,150,200")
    parts = Split(raw, ",")
    On Error GoTo Defaults
    If UBound(parts) <> 4 Then GoTo Defaults
    For i = 1 To 5
        result(i) = NxViewPresetZoom(CStr(parts(i - 1)))
    Next i
    NxViewPresetValues = result
    Exit Function
Defaults:
    Err.Clear
    result(1) = 75: result(2) = 100: result(3) = 125: result(4) = 150: result(5) = 200
    NxViewPresetValues = result
End Function

Public Function NxViewPresetZoom(ByVal text As String) As Long
    Dim value As Double
    If Not IsNumeric(text) Then NxRaiseContractError "배율은 10~400 사이의 정수로 입력하세요."
    value = CDbl(text)
    If value < 10 Or value > 400 Or value <> Fix(value) Then NxRaiseContractError "배율은 10~400 사이의 정수로 입력하세요."
    NxViewPresetZoom = CLng(value)
End Function

Public Sub NxViewPresetSave(ByVal values As Variant)
    Dim i As Long, raw As String
    For i = 1 To 5
        If i > 1 Then raw = raw & ","
        raw = raw & CStr(NxViewPresetZoom(CStr(values(i))))
    Next i
    ' One value contains all five presets: invalid input never partially saves.
    SaveSetting "LHExcel", "ViewPresets", "Zooms", raw
End Sub

Public Sub NxViewPresetApply(ByVal window As Excel.Window, ByVal index As Long)
    Dim key As String, state As Variant, values As Variant
    If window Is Nothing Then NxRaiseContractError "배율을 바꿀 엑셀 창이 없습니다."
    If index < 1 Or index > 5 Then NxRaiseContractError "화면 프리셋 번호가 올바르지 않습니다."
    If mStates Is Nothing Then Set mStates = CreateObject("Scripting.Dictionary")
    key = CStr(window.hwnd) & "|" & CStr(ObjPtr(window.ActiveSheet))
    values = NxViewPresetValues()
    If mStates.Exists(key) Then
        state = mStates(key)
        If CLng(state(0)) = index Then
            window.Zoom = CLng(state(1))
            mStates.Remove key
            Exit Sub
        End If
        ' Switching presets keeps the original pre-preset zoom for restoration.
    Else
        state = Array(index, CLng(window.Zoom))
    End If
    window.Zoom = CLng(values(index))
    state(0) = index
    mStates(key) = state
End Sub

Public Sub NxViewPresetForgetWorkbook(ByVal workbook As Excel.Workbook)
    Dim view As Excel.Window, key As Variant
    If mStates Is Nothing Then Exit Sub
    For Each view In workbook.Windows
        For Each key In mStates.Keys
            If Left$(CStr(key), Len(CStr(view.hwnd)) + 1) = CStr(view.hwnd) & "|" Then mStates.Remove key
        Next key
    Next view
End Sub

Public Sub NxViewPresetSettingsOpen()
    Dim view As New FNxViewPresets
    view.BindSettings
    view.Show vbModal
End Sub
