VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxRoleStyleCellState"
Option Explicit

Private Const NX_NO_THEME_COLOR As Long = 0

Private mCell As Range
Private mFontName As Variant, mFontSize As Variant, mFontBold As Variant
Private mFontColor As Variant, mFontColorIndex As Variant, mFontTintAndShade As Variant, mFontThemeColor As Variant, mFontThemeAvailable As Boolean
Private mPattern As Variant, mInteriorColor As Variant, mInteriorColorIndex As Variant, mInteriorTintAndShade As Variant, mInteriorThemeColor As Variant, mInteriorThemeAvailable As Boolean
Private mHorizontal As Variant, mVertical As Variant, mWrap As Variant
Private mEdges As Object

Public Sub Capture(ByVal cell As Range)
    Dim edge As Variant, borderLine As Border, themeAvailable As Boolean, themeColor As Variant
    Set mCell = cell: Set mEdges = CreateObject("Scripting.Dictionary")
    mFontName = cell.Font.Name: mFontSize = cell.Font.Size: mFontBold = cell.Font.Bold
    mFontColor = cell.Font.Color: mFontColorIndex = cell.Font.ColorIndex: mFontTintAndShade = cell.Font.TintAndShade
    mFontThemeAvailable = TryReadThemeColor(cell.Font, mFontThemeColor)
    mPattern = cell.Interior.Pattern: mInteriorColor = cell.Interior.Color
    mInteriorColorIndex = cell.Interior.ColorIndex: mInteriorTintAndShade = cell.Interior.TintAndShade
    mInteriorThemeAvailable = TryReadThemeColor(cell.Interior, mInteriorThemeColor)
    mHorizontal = cell.HorizontalAlignment: mVertical = cell.VerticalAlignment: mWrap = cell.WrapText
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom, xlInsideHorizontal, xlInsideVertical)
        Set borderLine = cell.Borders(CLng(edge))
        themeColor = Empty: themeAvailable = TryReadThemeColor(borderLine, themeColor)
        mEdges.Add CStr(edge), Array(borderLine.Weight, borderLine.Color, borderLine.ColorIndex, borderLine.TintAndShade, themeAvailable, themeColor, borderLine.LineStyle)
    Next edge
End Sub

Public Sub Restore()
    Dim edge As Variant, values As Variant, borderLine As Border
    If mCell Is Nothing Then Exit Sub
    mCell.Font.Name = mFontName: mCell.Font.Size = mFontSize: mCell.Font.Bold = mFontBold
    RestoreColorState mCell.Font, mFontColor, mFontColorIndex, mFontTintAndShade, mFontThemeAvailable, mFontThemeColor
    RestoreColorState mCell.Interior, mInteriorColor, mInteriorColorIndex, mInteriorTintAndShade, mInteriorThemeAvailable, mInteriorThemeColor
    mCell.Interior.Pattern = mPattern
    mCell.HorizontalAlignment = mHorizontal: mCell.VerticalAlignment = mVertical: mCell.WrapText = mWrap
    For Each edge In mEdges.Keys
        values = mEdges(edge): Set borderLine = mCell.Borders(CLng(edge))
        With borderLine
            .Weight = values(0)
        End With
        RestoreColorState borderLine, values(1), values(2), values(3), CBool(values(4)), values(5)
        borderLine.LineStyle = values(6)
    Next edge
End Sub

Private Function TryReadThemeColor(ByVal colorObject As Object, ByRef themeColor As Variant) As Boolean
    On Error GoTo NotThemeBound
    themeColor = colorObject.ThemeColor
    If CLng(themeColor) = NX_NO_THEME_COLOR Then GoTo NotThemeBound
    TryReadThemeColor = True
    Exit Function
NotThemeBound:
    Err.Clear
    themeColor = Empty
End Function

Private Sub RestoreColorState(ByVal colorObject As Object, ByVal colorValue As Variant, ByVal colorIndex As Variant, ByVal tintAndShade As Variant, ByVal themeAvailable As Boolean, ByVal themeColor As Variant)
    If themeAvailable Then
        colorObject.ThemeColor = themeColor
        colorObject.TintAndShade = tintAndShade
        Exit Sub
    ElseIf Not IsNull(colorIndex) Then
        If CLng(colorIndex) = xlColorIndexAutomatic Or CLng(colorIndex) = xlColorIndexNone Then
            colorObject.ColorIndex = colorIndex
            Exit Sub
        End If
    Else
        colorObject.Color = colorValue
        colorObject.TintAndShade = tintAndShade
        Exit Sub
    End If
    colorObject.Color = colorValue
    colorObject.TintAndShade = tintAndShade
End Sub
