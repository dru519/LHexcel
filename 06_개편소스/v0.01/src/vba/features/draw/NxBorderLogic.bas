Attribute VB_Name = "NxBorderLogic"
Option Explicit

Public Sub NxDrawClearInner(ByVal target As Range)
    RequireContiguous target
    target.Borders(xlInsideHorizontal).LineStyle = xlNone
    target.Borders(xlInsideVertical).LineStyle = xlNone
End Sub

Public Sub NxDrawApplyOutline(ByVal target As Range)
    RequireContiguous target
    ApplyEdge target, xlEdgeLeft: ApplyEdge target, xlEdgeTop
    ApplyEdge target, xlEdgeRight: ApplyEdge target, xlEdgeBottom
End Sub

Private Sub ApplyEdge(ByVal target As Range, ByVal edge As XlBordersIndex)
    With target.Borders(edge)
        .LineStyle = xlContinuous: .Color = NxEmphasisLineColor(): .Weight = NxEmphasisLineWeight()
    End With
End Sub

Private Sub RequireContiguous(ByVal target As Range)
    If target Is Nothing Then NxRaiseContractError "표/그리기 대상 범위가 필요합니다."
    If target.Areas.Count <> 1 Then NxRaiseContractError "불연속 범위는 표/그리기에서 지원하지 않습니다."
    If target.Worksheet.ProtectContents Or target.Worksheet.ProtectDrawingObjects Then NxRaiseContractError "보호된 시트에는 적용할 수 없습니다."
End Sub
