Attribute VB_Name = "FxFocusOverlayProbe"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
#Else
    Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
#End If

Private Const NX_OVERLAY_TIMER_MS As Long = 100
Private Const NX_ROW_FIRST As Long = 1
Private Const NX_ROW_LAST As Long = 2
Private Const NX_COLUMN_FIRST As Long = 3
Private Const NX_COLUMN_LAST As Long = 4

Private mHost As Application
Private mEvents As CFxFocusOverlayEvents
Private mForms As Collection
#If VBA7 Then
Private mTimerId As LongPtr
#Else
Private mTimerId As Long
#End If
Private mExpectedLeft(1 To 4) As Long
Private mExpectedTop(1 To 4) As Long
Private mExpectedWidth(1 To 4) As Long
Private mExpectedHeight(1 To 4) As Long
Private mExpectedVisible(1 To 4) As Boolean
Private mRefreshing As Boolean
Private mStopping As Boolean
Private mRefreshCount As Long
Private mLastSelectionSignature As String
Private mLastError As String

Public Function FxOverlayStart(ByVal host As Object, ByVal selectedColor As Long) As Boolean
    On Error GoTo Failed
    FxOverlayStop
    Set mHost = host
    Set mForms = New Collection
    FxCreateOverlay "ROW_FIRST", selectedColor
    FxCreateOverlay "ROW_LAST", selectedColor
    FxCreateOverlay "COLUMN_FIRST", selectedColor
    FxCreateOverlay "COLUMN_LAST", selectedColor
    Set mEvents = New CFxFocusOverlayEvents
    mEvents.StartListening mHost
    FxOverlayRefresh
    mTimerId = SetTimer(0, 0, NX_OVERLAY_TIMER_MS, AddressOf FxOverlayTimerProc)
    If mTimerId = 0 Then Err.Raise vbObjectError + 9811, "FxFocusOverlayProbe", "Overlay timer unavailable"
    FxOverlayStart = (FxOverlayWindowCount() = 4)
    Exit Function
Failed:
    mLastError = Err.Description
    FxOverlayStop
End Function

Private Sub FxCreateOverlay(ByVal token As String, ByVal selectedColor As Long)
    Dim overlay As FNxFocusOverlayProbe
    Set overlay = New FNxFocusOverlayProbe
    overlay.Configure token, selectedColor, mHost.hWnd
    overlay.OpenOverlay
    mForms.Add overlay, token
End Sub

Public Sub FxOverlayRefresh()
    If mRefreshing Or mStopping Then Exit Sub
    If mHost Is Nothing Then Exit Sub
    If mForms Is Nothing Then Exit Sub

    On Error GoTo Failed
    mRefreshing = True
    Dim activeWindow As Window
    Dim selectionRange As Range
    Dim selectedArea As Range
    Dim visibleRange As Range
    Dim target As Range
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstColumn As Long
    Dim lastColumn As Long

    Set activeWindow = mHost.ActiveWindow
    If activeWindow Is Nothing Then GoTo Hidden
    If TypeName(mHost.Selection) <> "Range" Then GoTo Hidden
    Set selectionRange = mHost.Selection
    Set selectedArea = selectionRange.Areas(1)
    Set visibleRange = activeWindow.VisibleRange
    If visibleRange Is Nothing Then GoTo Hidden
    If Not selectedArea.Parent Is visibleRange.Parent Then GoTo Hidden

    firstRow = selectedArea.Row
    lastRow = selectedArea.Row + selectedArea.Rows.Count - 1
    firstColumn = selectedArea.Column
    lastColumn = selectedArea.Column + selectedArea.Columns.Count - 1
    mLastSelectionSignature = CStr(firstRow) & ":" & CStr(lastRow) & "|" & CStr(firstColumn) & ":" & CStr(lastColumn)

    Set target = mHost.Intersect(visibleRange, selectedArea.Parent.Rows(firstRow))
    FxMoveOverlay NX_ROW_FIRST, target, activeWindow
    Set target = Nothing
    If lastRow = firstRow Then
        FxHideOverlay NX_ROW_LAST
    Else
        Set target = mHost.Intersect(visibleRange, selectedArea.Parent.Rows(lastRow))
        FxMoveOverlay NX_ROW_LAST, target, activeWindow
    End If
    Set target = Nothing
    Set target = mHost.Intersect(visibleRange, selectedArea.Parent.Columns(firstColumn))
    FxMoveOverlay NX_COLUMN_FIRST, target, activeWindow
    Set target = Nothing
    If lastColumn = firstColumn Then
        FxHideOverlay NX_COLUMN_LAST
    Else
        Set target = mHost.Intersect(visibleRange, selectedArea.Parent.Columns(lastColumn))
        FxMoveOverlay NX_COLUMN_LAST, target, activeWindow
    End If

    mRefreshCount = mRefreshCount + 1
    mLastError = vbNullString
    mRefreshing = False
    Exit Sub
Hidden:
    FxOverlayHideAll
    mRefreshing = False
    Exit Sub
Failed:
    mLastError = Err.Description
    FxOverlayHideAll
    mRefreshing = False
End Sub

Private Sub FxMoveOverlay(ByVal index As Long, ByVal target As Range, ByVal activeWindow As Window)
    If target Is Nothing Then FxHideOverlay index: Exit Sub
    Dim leftPixel As Long
    Dim topPixel As Long
    Dim rightPixel As Long
    Dim bottomPixel As Long
    Dim overlay As FNxFocusOverlayProbe

    leftPixel = activeWindow.PointsToScreenPixelsX(target.Left)
    topPixel = activeWindow.PointsToScreenPixelsY(target.Top)
    rightPixel = activeWindow.PointsToScreenPixelsX(target.Left + target.Width)
    bottomPixel = activeWindow.PointsToScreenPixelsY(target.Top + target.Height)
    If rightPixel <= leftPixel Then rightPixel = leftPixel + 1
    If bottomPixel <= topPixel Then bottomPixel = topPixel + 1

    Set overlay = mForms.Item(index)
    overlay.MoveOverlay leftPixel, topPixel, rightPixel - leftPixel, bottomPixel - topPixel
    mExpectedLeft(index) = leftPixel
    mExpectedTop(index) = topPixel
    mExpectedWidth(index) = rightPixel - leftPixel
    mExpectedHeight(index) = bottomPixel - topPixel
    mExpectedVisible(index) = True
End Sub

Private Sub FxHideOverlay(ByVal index As Long)
    Dim overlay As FNxFocusOverlayProbe
    If mForms Is Nothing Then Exit Sub
    Set overlay = mForms.Item(index)
    overlay.HideOverlay
    mExpectedVisible(index) = False
End Sub

Public Sub FxOverlayHideAll()
    On Error Resume Next
    Dim index As Long
    If mForms Is Nothing Then Exit Sub
    For index = 1 To mForms.Count
        FxHideOverlay index
    Next index
    On Error GoTo 0
End Sub

Public Sub FxOverlayStop()
    On Error Resume Next
    mStopping = True
    If mTimerId <> 0 Then Call KillTimer(0, mTimerId)
    mTimerId = 0
    If Not mEvents Is Nothing Then mEvents.StopListening
    Set mEvents = Nothing

    Dim index As Long
    Dim overlay As FNxFocusOverlayProbe
    If Not mForms Is Nothing Then
        For index = mForms.Count To 1 Step -1
            Set overlay = mForms.Item(index)
            overlay.HideOverlay
            overlay.ReleaseHandle
            Unload overlay
            Set overlay = Nothing
        Next index
    End If
    Set mForms = Nothing
    Set mHost = Nothing
    mRefreshing = False
    mStopping = False
    On Error GoTo 0
End Sub

#If VBA7 Then
Public Sub FxOverlayTimerProc(ByVal hWnd As LongPtr, ByVal message As Long, ByVal timerIdentifier As LongPtr, ByVal elapsed As Long)
#Else
Public Sub FxOverlayTimerProc(ByVal hWnd As Long, ByVal message As Long, ByVal timerIdentifier As Long, ByVal elapsed As Long)
#End If
    On Error Resume Next
    FxOverlayRefresh
    On Error GoTo 0
End Sub

Public Function FxOverlayWindowCount() As Long
    Dim index As Long
    Dim overlay As FNxFocusOverlayProbe
    On Error GoTo Done
    If mForms Is Nothing Then Exit Function
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        If overlay.IsAlive Then FxOverlayWindowCount = FxOverlayWindowCount + 1
    Next index
Done:
End Function

Public Function FxOverlayVisibleWindowCount() As Long
    Dim index As Long
    Dim overlay As FNxFocusOverlayProbe
    On Error GoTo Done
    If mForms Is Nothing Then Exit Function
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        If overlay.IsVisible Then FxOverlayVisibleWindowCount = FxOverlayVisibleWindowCount + 1
    Next index
Done:
End Function

Public Function FxOverlayStyleContractSatisfied() As Boolean
    Dim index As Long
    Dim overlay As FNxFocusOverlayProbe
    On Error GoTo Failed
    If FxOverlayWindowCount() <> 4 Then Exit Function
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        If Not overlay.StyleContractSatisfied Then Exit Function
    Next index
    FxOverlayStyleContractSatisfied = True
    Exit Function
Failed:
End Function

Public Function FxOverlayExpectedBoundsMatch(Optional ByVal tolerance As Long = 2) As Boolean
    Dim index As Long
    Dim overlay As FNxFocusOverlayProbe
    On Error GoTo Failed
    If mForms Is Nothing Then Exit Function
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        If mExpectedVisible(index) Then
            If Not overlay.IsVisible Then Exit Function
            If Not overlay.BoundsMatch(mExpectedLeft(index), mExpectedTop(index), mExpectedWidth(index), mExpectedHeight(index), tolerance) Then Exit Function
        ElseIf overlay.IsVisible Then
            Exit Function
        End If
    Next index
    FxOverlayExpectedBoundsMatch = True
    Exit Function
Failed:
End Function

Public Function FxOverlayBoundsSignature() As String
    Dim index As Long
    Dim overlay As FNxFocusOverlayProbe
    Dim result As String
    On Error GoTo Done
    If mForms Is Nothing Then Exit Function
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        If Len(result) > 0 Then result = result & "|"
        result = result & overlay.BoundsSignature
    Next index
Done:
    FxOverlayBoundsSignature = result
End Function

Public Function FxOverlaySelectionSignature() As String
    FxOverlaySelectionSignature = mLastSelectionSignature
End Function

Public Function FxOverlayRefreshCount() As Long
    FxOverlayRefreshCount = mRefreshCount
End Function

Public Function FxOverlayLastError() As String
    FxOverlayLastError = mLastError
End Function
