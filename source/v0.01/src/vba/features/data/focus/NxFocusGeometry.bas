Attribute VB_Name = "NxFocusGeometry"
Option Explicit

Public Const NX_FOCUS_SEGMENT_ROW_LEFT As Long = 1
Public Const NX_FOCUS_SEGMENT_ROW_RIGHT As Long = 2
Public Const NX_FOCUS_SEGMENT_COLUMN_TOP As Long = 3
Public Const NX_FOCUS_SEGMENT_COLUMN_BOTTOM As Long = 4
Public Const NX_FOCUS_SEGMENT_SELECT_TOP As Long = 5
Public Const NX_FOCUS_SEGMENT_SELECT_RIGHT As Long = 6
Public Const NX_FOCUS_SEGMENT_SELECT_BOTTOM As Long = 7
Public Const NX_FOCUS_SEGMENT_SELECT_LEFT As Long = 8
Public Const NX_FOCUS_SEGMENT_COUNT As Long = 8
Private Const NX_FOCUS_ROW_HEADER_LOGICAL_PX As Double = 31#
Private Const NX_FOCUS_COLUMN_HEADER_LOGICAL_PX As Double = 22.5

Private mGeometryStage As String
Private mGeometryDiagnostic As String

#If VBA7 Then
    Private Declare PtrSafe Function GetDpiForWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function EnumChildWindows Lib "user32" (ByVal hWndParent As LongPtr, ByVal lpEnumFunc As LongPtr, ByVal lParam As LongPtr) As Long
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As NxFocusWinRect) As Long
    Private mExcel7WindowHandle As LongPtr
#Else
    Private Declare Function GetDpiForWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As NxFocusWinRect) As Long
    Private mExcel7WindowHandle As Long
#End If

Private Type NxFocusWinRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Function NxFocusBuildGeometry(ByVal host As Application, ByVal selectedArea As Range) As CNxFocusGeometry
    Dim activeWindow As Window
    Dim panes As Panes
    Dim activePane As Pane
    Dim activeVisible As Range
    Dim logicalSelection As Range
    Dim geometry As New CNxFocusGeometry
    Dim excelRect As NxFocusWinRect
    Dim dpi As Long
    Dim zoom As Long
    Dim pointScale As Double
    Dim activeX0 As Long
    Dim activeY0 As Long
    Dim activeGridLeft As Long
    Dim activeGridTop As Long
    Dim gridRight As Long
    Dim gridBottom As Long
    Dim selectionLeft As Long
    Dim selectionTop As Long
    Dim selectionRight As Long
    Dim selectionBottom As Long
    Dim paneIndex As Long
    Dim divider As Long
    Dim rowHeaderWidth As Long
    Dim columnHeaderHeight As Long
    Dim isRightPane As Boolean
    Dim isBottomPane As Boolean
    Dim outline As Long
    Dim signature As String
    Dim visibleLeftPoints As Double
    Dim visibleTopPoints As Double
    Dim visibleRightPoints As Double
    Dim visibleBottomPoints As Double
    Dim selectionLeftPoints As Double
    Dim selectionTopPoints As Double
    Dim selectionRightPoints As Double
    Dim selectionBottomPoints As Double
    On Error GoTo Failed
    mGeometryStage = "input"
    mGeometryDiagnostic = vbNullString

    If host Is Nothing Or selectedArea Is Nothing Then Exit Function
    mGeometryStage = "active-window"
    Set activeWindow = host.ActiveWindow
    If activeWindow Is Nothing Then Exit Function
    Set panes = activeWindow.Panes
    Set activePane = activeWindow.ActivePane
    If activePane Is Nothing Or panes Is Nothing Then Exit Function
    paneIndex = activePane.Index
    Set activeVisible = activePane.VisibleRange
    If activeVisible Is Nothing Then Exit Function
    If Not activeVisible.Parent Is selectedArea.Parent Then Exit Function

    Set logicalSelection = selectedArea
    visibleLeftPoints = CDbl(activeVisible.Left)
    visibleTopPoints = CDbl(activeVisible.Top)
    visibleRightPoints = visibleLeftPoints + CDbl(activeVisible.Width)
    visibleBottomPoints = visibleTopPoints + CDbl(activeVisible.Height)
    selectionLeftPoints = CDbl(logicalSelection.Left)
    selectionTopPoints = CDbl(logicalSelection.Top)
    selectionRightPoints = selectionLeftPoints + CDbl(logicalSelection.Width)
    selectionBottomPoints = selectionTopPoints + CDbl(logicalSelection.Height)
    If selectionRightPoints <= visibleLeftPoints Or selectionBottomPoints <= visibleTopPoints Then Exit Function
    If selectionLeftPoints >= visibleRightPoints Or selectionTopPoints >= visibleBottomPoints Then Exit Function
    selectionLeftPoints = NxFocusMaxDouble(selectionLeftPoints, visibleLeftPoints)
    selectionTopPoints = NxFocusMaxDouble(selectionTopPoints, visibleTopPoints)
    selectionRightPoints = NxFocusMinDouble(selectionRightPoints, visibleRightPoints)
    selectionBottomPoints = NxFocusMinDouble(selectionBottomPoints, visibleBottomPoints)

    mGeometryStage = "scale"
    dpi = GetDpiForWindow(activeWindow.hWnd)
    If dpi <= 0 Then dpi = 96
    zoom = CLng(activeWindow.Zoom)
    If zoom <= 0 Then Exit Function
    pointScale = dpi / 72# * zoom / 100#

    mGeometryStage = "pane-origin"
    activeX0 = activeWindow.PointsToScreenPixelsX(0)
    activeY0 = activeWindow.PointsToScreenPixelsY(0)
    NxFocusPanePosition activeWindow, paneIndex, isRightPane, isBottomPane
    divider = 0
    If Not activeWindow.FreezePanes Then divider = NxFocusRoundPixel(2# * CDbl(dpi) / 96#)
    If activeWindow.DisplayHeadings Then
        rowHeaderWidth = NxFocusRoundPixel(NX_FOCUS_ROW_HEADER_LOGICAL_PX * CDbl(dpi) / 96#)
        columnHeaderHeight = NxFocusRoundPixel(NX_FOCUS_COLUMN_HEADER_LOGICAL_PX * CDbl(dpi) / 96#)
    End If
    activeGridLeft = activeX0 + NxFocusRoundPixel(CDbl(activeVisible.Left) * pointScale)
    activeGridTop = activeY0 + NxFocusRoundPixel(CDbl(activeVisible.Top) * pointScale)
    If isRightPane Then activeGridLeft = activeGridLeft + rowHeaderWidth + NxFocusRoundPixel(CDbl(activeWindow.SplitHorizontal) * pointScale) + divider
    If isBottomPane Then activeGridTop = activeGridTop + columnHeaderHeight + NxFocusRoundPixel(CDbl(activeWindow.SplitVertical) * pointScale) + divider

    mGeometryStage = "excel7"
    If Not NxFocusGetExcel7Rect(activeWindow.hWnd, excelRect) Then
        If Not NxFocusGetExcel7Rect(host.hWnd, excelRect) Then Exit Function
    End If
    gridRight = activeGridLeft + NxFocusRoundPixel(CDbl(activeVisible.Width) * pointScale)
    gridBottom = activeGridTop + NxFocusRoundPixel(CDbl(activeVisible.Height) * pointScale)
    NxFocusClampPixelRect activeGridLeft, activeGridTop, gridRight, gridBottom, excelRect
    If gridRight <= activeGridLeft Or gridBottom <= activeGridTop Then Exit Function

    mGeometryStage = "selection"
    selectionLeft = activeGridLeft + NxFocusRoundPixel((selectionLeftPoints - visibleLeftPoints) * pointScale)
    selectionTop = activeGridTop + NxFocusRoundPixel((selectionTopPoints - visibleTopPoints) * pointScale)
    selectionRight = selectionLeft + NxFocusRoundPixel((selectionRightPoints - selectionLeftPoints) * pointScale)
    selectionBottom = selectionTop + NxFocusRoundPixel((selectionBottomPoints - selectionTopPoints) * pointScale)
    If selectionLeft < activeGridLeft Then selectionLeft = activeGridLeft
    If selectionTop < activeGridTop Then selectionTop = activeGridTop
    If selectionRight > gridRight Then selectionRight = gridRight
    If selectionBottom > gridBottom Then selectionBottom = gridBottom
    If selectionRight <= selectionLeft Or selectionBottom <= selectionTop Then Exit Function

    mGeometryStage = "segments"
    signature = CStr(logicalSelection.Row) & ":" & CStr(logicalSelection.Row + logicalSelection.Rows.Count - 1) & "|" & _
        CStr(logicalSelection.Column) & ":" & CStr(logicalSelection.Column + logicalSelection.Columns.Count - 1)
    geometry.ConfigureGrid activeGridLeft, activeGridTop, gridRight, gridBottom, dpi, zoom, paneIndex
    geometry.ConfigureSelection selectionLeft, selectionTop, selectionRight, selectionBottom, signature

    geometry.SetSegment NX_FOCUS_SEGMENT_ROW_LEFT, activeGridLeft, selectionTop, selectionLeft, selectionBottom
    geometry.SetSegment NX_FOCUS_SEGMENT_ROW_RIGHT, selectionRight, selectionTop, gridRight, selectionBottom
    geometry.SetSegment NX_FOCUS_SEGMENT_COLUMN_TOP, selectionLeft, activeGridTop, selectionRight, selectionTop
    geometry.SetSegment NX_FOCUS_SEGMENT_COLUMN_BOTTOM, selectionLeft, selectionBottom, selectionRight, gridBottom

    outline = NxFocusRoundPixel(2# * CDbl(dpi) / 96#)
    If outline < 1 Then outline = 1
    geometry.SetSegment NX_FOCUS_SEGMENT_SELECT_TOP, selectionLeft, _
        NxFocusMaxLong(activeGridTop, selectionTop - outline), selectionRight, selectionTop
    geometry.SetSegment NX_FOCUS_SEGMENT_SELECT_RIGHT, selectionRight, selectionTop, _
        NxFocusMinLong(gridRight, selectionRight + outline), selectionBottom
    geometry.SetSegment NX_FOCUS_SEGMENT_SELECT_BOTTOM, selectionLeft, selectionBottom, selectionRight, _
        NxFocusMinLong(gridBottom, selectionBottom + outline)
    geometry.SetSegment NX_FOCUS_SEGMENT_SELECT_LEFT, NxFocusMaxLong(activeGridLeft, selectionLeft - outline), _
        selectionTop, selectionLeft, selectionBottom
    geometry.Seal
    mGeometryDiagnostic = "PASS"
    Set NxFocusBuildGeometry = geometry
    Exit Function
Failed:
    mGeometryDiagnostic = "FAIL|" & mGeometryStage & "|" & CStr(Err.Number) & "|" & Err.Description
    Set NxFocusBuildGeometry = Nothing
End Function

Public Function NxFocusGeometryLastDiagnostic() As String
    NxFocusGeometryLastDiagnostic = mGeometryDiagnostic
End Function

Private Sub NxFocusPanePosition(ByVal activeWindow As Window, ByVal paneIndex As Long, _
    ByRef isRightPane As Boolean, ByRef isBottomPane As Boolean)
    Dim splitColumns As Boolean
    Dim splitRows As Boolean
    splitColumns = (activeWindow.SplitColumn > 0 Or activeWindow.SplitHorizontal > 0)
    splitRows = (activeWindow.SplitRow > 0 Or activeWindow.SplitVertical > 0)
    If splitColumns And splitRows Then
        isRightPane = (paneIndex = 2 Or paneIndex = 4)
        isBottomPane = (paneIndex = 3 Or paneIndex = 4)
    ElseIf splitColumns Then
        isRightPane = (paneIndex = 2)
    ElseIf splitRows Then
        isBottomPane = (paneIndex = 2)
    End If
End Sub

Public Function NxFocusGeometryStateSignature(ByVal host As Application) As String
    Dim activeWindow As Window
    Dim activePane As Pane
    Dim visibleRange As Range
    Dim selectionRange As Range
    Dim excelRect As NxFocusWinRect
    Dim activeX0 As Long
    Dim activeY0 As Long
    On Error GoTo Failed
    If host Is Nothing Then Exit Function
    Set activeWindow = host.ActiveWindow
    If activeWindow Is Nothing Then Exit Function
    Set activePane = activeWindow.ActivePane
    Set visibleRange = activePane.VisibleRange
    If TypeName(host.Selection) = "Range" Then Set selectionRange = host.Selection
    activeX0 = activeWindow.PointsToScreenPixelsX(0)
    activeY0 = activeWindow.PointsToScreenPixelsY(0)
    If Not NxFocusGetExcel7Rect(activeWindow.hWnd, excelRect) Then _
        Call NxFocusGetExcel7Rect(host.hWnd, excelRect)
    NxFocusGeometryStateSignature = CStr(activeWindow.hWnd) & "|" & CStr(activePane.Index) & "|" & _
        CStr(activeWindow.Zoom) & "|" & CStr(activeWindow.SplitColumn) & "|" & CStr(activeWindow.SplitRow) & "|" & _
        CStr(activeWindow.SplitHorizontal) & "|" & CStr(activeWindow.SplitVertical) & "|" & _
        CStr(activeWindow.FreezePanes) & "|" & CStr(activeWindow.DisplayHeadings) & "|" & _
        visibleRange.Address(External:=True) & "|" & CStr(activeX0) & "|" & CStr(activeY0) & "|" & _
        CStr(excelRect.Left) & "|" & CStr(excelRect.Top) & "|" & CStr(excelRect.Right) & "|" & CStr(excelRect.Bottom)
    If Not selectionRange Is Nothing Then NxFocusGeometryStateSignature = NxFocusGeometryStateSignature & "|" & selectionRange.Address(External:=True)
    Exit Function
Failed:
    NxFocusGeometryStateSignature = vbNullString
End Function

Private Sub NxFocusClampPixelRect(ByRef leftPixel As Long, ByRef topPixel As Long, _
    ByRef rightPixel As Long, ByRef bottomPixel As Long, ByRef bounds As NxFocusWinRect)
    If leftPixel < bounds.Left Then leftPixel = bounds.Left
    If topPixel < bounds.Top Then topPixel = bounds.Top
    If rightPixel > bounds.Right Then rightPixel = bounds.Right
    If bottomPixel > bounds.Bottom Then bottomPixel = bounds.Bottom
End Sub

#If VBA7 Then
Private Function NxFocusGetExcel7Rect(ByVal parentHandle As LongPtr, ByRef bounds As NxFocusWinRect) As Boolean
#Else
Private Function NxFocusGetExcel7Rect(ByVal parentHandle As Long, ByRef bounds As NxFocusWinRect) As Boolean
#End If
    mExcel7WindowHandle = 0
    If parentHandle = 0 Then Exit Function
    Call EnumChildWindows(parentHandle, AddressOf NxFocusEnumExcel7Child, 0)
    If mExcel7WindowHandle = 0 Then Exit Function
    NxFocusGetExcel7Rect = (GetWindowRect(mExcel7WindowHandle, bounds) <> 0)
End Function

#If VBA7 Then
Public Function NxFocusEnumExcel7Child(ByVal childHandle As LongPtr, ByVal lParam As LongPtr) As Long
#Else
Public Function NxFocusEnumExcel7Child(ByVal childHandle As Long, ByVal lParam As Long) As Long
#End If
    Dim className As String
    Dim length As Long
    className = String$(64, vbNullChar)
    length = GetClassName(childHandle, className, Len(className))
    If length > 0 Then
        className = Left$(className, length)
        If StrComp(className, "EXCEL7", vbBinaryCompare) = 0 Then
            mExcel7WindowHandle = childHandle
            NxFocusEnumExcel7Child = 0
            Exit Function
        End If
    End If
    NxFocusEnumExcel7Child = 1
End Function

Private Function NxFocusRoundPixel(ByVal value As Double) As Long
    If value >= 0 Then
        NxFocusRoundPixel = CLng(Int(value + 0.5))
    Else
        NxFocusRoundPixel = -CLng(Int(-value + 0.5))
    End If
End Function

Private Function NxFocusMinLong(ByVal first As Long, ByVal second As Long) As Long
    If first < second Then NxFocusMinLong = first Else NxFocusMinLong = second
End Function

Private Function NxFocusMaxLong(ByVal first As Long, ByVal second As Long) As Long
    If first > second Then NxFocusMaxLong = first Else NxFocusMaxLong = second
End Function

Private Function NxFocusMinDouble(ByVal first As Double, ByVal second As Double) As Double
    If first < second Then NxFocusMinDouble = first Else NxFocusMinDouble = second
End Function

Private Function NxFocusMaxDouble(ByVal first As Double, ByVal second As Double) As Double
    If first > second Then NxFocusMaxDouble = first Else NxFocusMaxDouble = second
End Function
