Attribute VB_Name = "NxQuickFormatUndo"
Option Explicit

Private mObjects As Collection
Private mStates As Collection
Private mKey As String
Private mSheet As Worksheet
Private mReady As Boolean

Public Sub NxQuickFormatUndoCapture(ByVal target As Range, ByVal key As String)
    Dim area As Range
    mReady = False
    Set mObjects = New Collection
    Set mStates = New Collection
    Set mSheet = target.Worksheet
    mKey = key
    For Each area In target.Areas
        CaptureRange area
    Next area
End Sub

Private Sub CaptureRange(ByVal target As Range)
    Dim item As Object, state As Variant, half As Long, i As Long
    If mKey = "fill_color" Then Set item = target.Interior Else Set item = target.Font
    state = ReadState(item)
    If StateUniform(state) Then
        mObjects.Add item: mStates.Add state
    ElseIf target.CountLarge = 1 Then
        ' Preserve differently formatted characters within a text cell.
        If mKey = "fill_color" Or target.HasFormula Or VarType(target.Value2) <> vbString Then _
            NxRaiseContractError "이 셀의 서식을 안전하게 보관하지 못했습니다."
        For i = 1 To Len(target.Value2)
            Set item = target.Characters(i, 1).Font
            state = ReadState(item)
            If Not StateUniform(state) Then NxRaiseContractError "문자 서식을 보관하지 못했습니다."
            mObjects.Add item: mStates.Add state
        Next i
    ElseIf target.Rows.Count >= target.Columns.Count Then
        half = target.Rows.Count \ 2
        CaptureRange target.Resize(half, target.Columns.Count)
        CaptureRange target.Cells(half + 1, 1).Resize(target.Rows.Count - half, target.Columns.Count)
    Else
        half = target.Columns.Count \ 2
        CaptureRange target.Resize(target.Rows.Count, half)
        CaptureRange target.Cells(1, half + 1).Resize(target.Rows.Count, target.Columns.Count - half)
    End If
End Sub

Private Function ReadTheme(ByVal item As Object, ByVal propertyName As String) As Variant
    On Error GoTo NoTheme
    ReadTheme = CallByName(item, propertyName, VbGet)
    Exit Function
NoTheme:
    Err.Clear
    ReadTheme = 0
End Function

Private Function ReadState(ByVal item As Object) As Variant
    Select Case mKey
        Case "font_size": ReadState = Array(item.Size)
        Case "font_color": ReadState = Array(item.Color, item.ColorIndex, ReadTheme(item, "ThemeColor"), item.TintAndShade)
        Case "fill_color"
            ReadState = Array(item.Color, item.ColorIndex, ReadTheme(item, "ThemeColor"), item.TintAndShade, _
                item.Pattern, item.PatternColor, item.PatternColorIndex, ReadTheme(item, "PatternThemeColor"), item.PatternTintAndShade)
        Case Else: NxRaiseContractError "알 수 없는 빠른 서식입니다."
    End Select
End Function

Private Function StateUniform(ByVal state As Variant) As Boolean
    Dim value As Variant
    For Each value In state
        If IsNull(value) Then Exit Function
    Next value
    StateUniform = True
End Function

Private Sub RestoreColor(ByVal item As Object, ByVal state As Variant, ByVal offset As Long, ByVal prefix As String)
    If state(offset + 2) > 0 Then
        CallByName item, prefix & "ThemeColor", VbLet, CLng(state(offset + 2))
    ElseIf state(offset + 1) = xlColorIndexAutomatic Or state(offset + 1) = xlColorIndexNone Then
        CallByName item, prefix & "ColorIndex", VbLet, CLng(state(offset + 1))
    Else
        CallByName item, prefix & "Color", VbLet, CLng(state(offset))
    End If
    CallByName item, prefix & "TintAndShade", VbLet, CDbl(state(offset + 3))
End Sub

Public Sub NxQuickFormatUndoRestore()
    Dim i As Long, item As Object, state As Variant
    If mObjects Is Nothing Or mSheet Is Nothing Then Exit Sub
    If mSheet.ProtectContents Or mSheet.Parent.ReadOnly Then NxRaiseContractError "편집 가능한 시트에서 되돌리세요."
    For i = mObjects.Count To 1 Step -1
        Set item = mObjects(i): state = mStates(i)
        If mKey = "font_size" Then
            item.Size = state(0)
        Else
            RestoreColor item, state, 0, ""
            If mKey = "fill_color" Then
                RestoreColor item, state, 5, "Pattern"
                item.Pattern = state(4)
            End If
        End If
    Next i
End Sub

Public Sub NxQuickFormatUndoCommit()
    mReady = True
    NxQuickFormatUndoArm
End Sub

Public Sub NxQuickFormatUndoArm()
    If Not mReady Then Exit Sub
    Application.OnUndo "내엑셀 서식 변경 취소", "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!NxQuickFormatUndoLast"
End Sub

Public Sub NxQuickFormatUndoLast()
    Dim guard As CNxStateGuard, detail As String
    If Not mReady Then Exit Sub
    On Error GoTo Failed
    Set guard = New CNxStateGuard
    NxQuickFormatUndoRestore
    mReady = False
    Set mObjects = Nothing: Set mStates = Nothing: Set mSheet = Nothing
    guard.Restore
    Exit Sub
Failed:
    detail = Err.Description
    If Not guard Is Nothing Then guard.Restore
    MsgBox "서식을 되돌리지 못했습니다. " & detail, vbExclamation, "내엑셀"
End Sub
