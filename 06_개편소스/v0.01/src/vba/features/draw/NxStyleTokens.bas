Attribute VB_Name = "NxStyleTokens"
Option Explicit

Public Const NX_ROLE_STYLE_MODE_MONO As String = "MONO"
Public Const NX_ROLE_STYLE_MODE_COLOR As String = "COLOR"
Public Const NX_ROLE_STYLE_TITLE As String = "TITLE"
Public Const NX_ROLE_STYLE_SUBTITLE As String = "SUBTITLE"
Public Const NX_ROLE_STYLE_TABLE_HEADER As String = "TABLE_HEADER"
Public Const NX_ROLE_STYLE_TABLE_BODY As String = "TABLE_BODY"
Public Const NX_ROLE_STYLE_EMPHASIS_CELL As String = "EMPHASIS_CELL"
Public Const NX_ROLE_STYLE_TOTAL_ROW As String = "TOTAL_ROW"

Public Function NxBodyFontName() As String: NxBodyFontName = "맑은 고딕": End Function
Public Function NxHeaderFontName() As String: NxHeaderFontName = "맑은 고딕": End Function
Public Function NxBodyFontSize() As Double: NxBodyFontSize = 10#: End Function
Public Function NxHeaderFontSize() As Double: NxHeaderFontSize = 10#: End Function
Public Function NxBodyFontColor() As Long: NxBodyFontColor = RGB(31, 41, 51): End Function
Public Function NxHeaderFontColor() As Long: NxHeaderFontColor = RGB(23, 50, 77): End Function
Public Function NxHeaderFillColor() As Long: NxHeaderFillColor = RGB(232, 238, 247): End Function
Public Function NxTotalFillColor() As Long: NxTotalFillColor = RGB(243, 245, 247): End Function
Public Function NxBaseLineColor() As Long: NxBaseLineColor = RGB(91, 101, 115): End Function
Public Function NxEmphasisLineColor() As Long: NxEmphasisLineColor = RGB(51, 78, 104): End Function
Public Function NxBaseLineWeight() As XlBorderWeight: NxBaseLineWeight = xlThin: End Function
Public Function NxEmphasisLineWeight() As XlBorderWeight: NxEmphasisLineWeight = xlMedium: End Function

Public Function NxRoleStyleTokenColor(ByVal tokenName As String, ByVal displayMode As String) As Long
    If UCase$(displayMode) = NX_ROLE_STYLE_MODE_MONO Then
        Select Case UCase$(tokenName)
            Case "FONT": NxRoleStyleTokenColor = RGB(32, 32, 32)
            Case "FILL": NxRoleStyleTokenColor = RGB(242, 242, 242)
            Case Else: NxRoleStyleTokenColor = RGB(96, 96, 96)
        End Select
    Else
        Select Case UCase$(tokenName)
            Case "FONT": NxRoleStyleTokenColor = NxHeaderFontColor()
            Case "FILL": NxRoleStyleTokenColor = NxHeaderFillColor()
            Case "TOTAL": NxRoleStyleTokenColor = NxTotalFillColor()
            Case "EMPHASIS": NxRoleStyleTokenColor = NxEmphasisLineColor()
            Case Else: NxRoleStyleTokenColor = NxBodyFontColor()
        End Select
    End If
End Function

Public Sub NxApplyHeaderToken(ByVal target As Range)
    If target Is Nothing Then NxRaiseContractError "Header range is required"
    With target
        .Font.Name = NxHeaderFontName(): .Font.Size = NxHeaderFontSize()
        .Font.Bold = True: .Font.Color = NxHeaderFontColor()
        .Interior.Pattern = xlSolid: .Interior.Color = NxHeaderFillColor()
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

Public Sub NxApplyBodyToken(ByVal target As Range, Optional ByVal preserveAlignment As Boolean = True)
    If target Is Nothing Then NxRaiseContractError "Body range is required"
    With target
        .Font.Name = NxBodyFontName(): .Font.Size = NxBodyFontSize()
        .Font.Bold = False: .Font.Color = NxBodyFontColor()
        .Interior.Pattern = xlNone
        If Not preserveAlignment Then .VerticalAlignment = xlCenter
    End With
End Sub

Public Sub NxApplyTotalToken(ByVal target As Range)
    If target Is Nothing Then NxRaiseContractError "Total range is required"
    NxApplyBodyToken target, True
    target.Interior.Pattern = xlSolid: target.Interior.Color = NxTotalFillColor()
    With target.Borders(xlEdgeTop)
        .LineStyle = xlContinuous: .Color = NxEmphasisLineColor(): .Weight = NxEmphasisLineWeight()
    End With
End Sub
