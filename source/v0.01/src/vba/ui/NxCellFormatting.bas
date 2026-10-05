Attribute VB_Name = "NxCellFormatting"
Option Explicit

Public Function NxDecimalDisplayFormat(ByVal places As Long, ByVal percentage As Boolean) As String
    If places < 0 Or places > 15 Then NxRaiseContractError "소수 자릿수는 0~15 사이의 정수로 입력하세요."
    NxDecimalDisplayFormat = IIf(percentage, "0", "#,##0")
    If places > 0 Then NxDecimalDisplayFormat = NxDecimalDisplayFormat & "." & String$(places, "0")
    If percentage Then NxDecimalDisplayFormat = NxDecimalDisplayFormat & "%"
End Function

Public Sub NxNumberFormatOpen(ByVal target As Range, ByVal percentage As Boolean, Optional ByVal emphasis As Boolean = False)
    Dim view As New FNxNumberFormat
    view.BindTarget target, percentage, emphasis
    view.Show vbModal
End Sub

Public Function NxNumberEmphasisFormat(ByVal places As Long, ByVal percentage As Boolean) As String
    Dim number As String
    number = NxDecimalDisplayFormat(places, percentage)
    NxNumberEmphasisFormat = "[Red]""" & ChrW(&H25B2) & " """ & number & ";[Blue]""" & ChrW(&H25BC) & " """ & number & ";" & number & ";@"
End Function

Public Sub NxResizeOpen(ByVal target As Range)
    Dim view As New FNxResize
    view.BindTarget target
    view.Show vbModal
End Sub

Public Sub NxResizeApply(ByVal target As Range, ByVal useWidth As Boolean, ByVal width As Double, _
    ByVal useHeight As Boolean, ByVal height As Double, ByVal useFont As Boolean, ByVal fontSize As Double, _
    ByVal useShapes As Boolean, ByVal shapePercent As Double)
    Dim rows As New Collection, columns As New Collection, fonts As New Collection, shapes As New Collection
    Dim item As Variant, cell As Range, shape As Shape, overlap As Range, geometry As Variant
    Dim failure As Long, detail As String, wasSaved As Boolean, mutated As Boolean, restoreFailure As String
    If target Is Nothing Then NxRaiseContractError "크기를 조정할 셀 범위를 선택하세요."
    If target.Areas.Count <> 1 Or target.CountLarge > 10000 Then NxRaiseContractError "연속된 범위를 10,000셀 이내로 선택하세요."
    If target.Worksheet.ProtectContents Or target.Worksheet.Parent.ReadOnly Then NxRaiseContractError "편집 가능한 시트에서 실행하세요."
    If Not (useWidth Or useHeight Or useFont Or useShapes) Then NxRaiseContractError "조정할 항목을 하나 이상 선택하세요."
    If useWidth And (width < 0.1 Or width > 255) Then NxRaiseContractError "열 너비는 0.1~255로 입력하세요."
    If useHeight And (height < 0.1 Or height > 409.5) Then NxRaiseContractError "행 높이는 0.1~409.5pt로 입력하세요."
    If useFont And (fontSize < 1 Or fontSize > 409) Then NxRaiseContractError "글자 크기는 1~409pt로 입력하세요."
    If useShapes And (shapePercent < 1 Or shapePercent > 1000) Then NxRaiseContractError "도형 배율은 1~1,000%로 입력하세요."
    If target.Worksheet.Shapes.Count > 500 Then NxRaiseContractError "도형이 500개를 넘는 시트에서는 크기 일괄 조정을 사용할 수 없습니다."
    wasSaved = target.Worksheet.Parent.Saved
    On Error GoTo Failed
    If useWidth Then
        For Each cell In target.Columns: columns.Add Array(cell.Column, cell.ColumnWidth): Next cell
    End If
    If useHeight Then
        For Each cell In target.Rows: rows.Add Array(cell.Row, cell.RowHeight): Next cell
    End If
    If useFont Then
        For Each cell In target.Cells: fonts.Add Array(cell.Address, cell.Font.Size): Next cell
    End If
    For Each shape In target.Worksheet.Shapes
        Set overlap = Application.Intersect(target, target.Worksheet.Range(shape.TopLeftCell, shape.BottomRightCell))
        shapes.Add Array(shape.Name, shape.Left, shape.Top, shape.Width, shape.Height, shape.LockAspectRatio, Not overlap Is Nothing)
    Next shape
    mutated = True
    If useWidth Then target.ColumnWidth = width
    If useHeight Then target.RowHeight = height
    If useFont Then target.Font.Size = fontSize
    If useShapes Then
        For Each geometry In shapes
            If geometry(6) Then
                Set shape = target.Worksheet.Shapes(CStr(geometry(0)))
                shape.LockAspectRatio = msoFalse
                shape.Width = CDbl(geometry(3)) * shapePercent / 100#
                shape.Height = CDbl(geometry(4)) * shapePercent / 100#
                shape.LockAspectRatio = geometry(5)
            End If
        Next geometry
    End If
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description
    Err.Clear
    If mutated Then
        On Error Resume Next
        For Each item In columns: target.Worksheet.Columns(CLng(item(0))).ColumnWidth = item(1): Next item
        For Each item In rows: target.Worksheet.Rows(CLng(item(0))).RowHeight = item(1): Next item
        For Each item In fonts: target.Worksheet.Range(CStr(item(0))).Font.Size = item(1): Next item
        For Each geometry In shapes
            Set shape = target.Worksheet.Shapes(CStr(geometry(0)))
            shape.LockAspectRatio = msoFalse
            shape.Left = geometry(1): shape.Top = geometry(2): shape.Width = geometry(3): shape.Height = geometry(4)
            shape.LockAspectRatio = geometry(5)
        Next geometry
        If Err.Number <> 0 Then restoreFailure = " 일부 크기를 복원하지 못했습니다: " & Err.Description
        If Len(restoreFailure) = 0 And wasSaved Then target.Worksheet.Parent.Saved = True
        On Error GoTo 0
    End If
    Err.Raise failure, "NxResizeApply", detail & restoreFailure
End Sub
