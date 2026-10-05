Attribute VB_Name = "NxFocusOverlay"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
#Else
    Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
#End If

Private Const NX_OVERLAY_TIMER_MS As Long = 100
Private Const NX_FOCUS_MAX_SELECTION_CELLS As Double = 100000#

Private mHost As Application
Private mForms As Collection
Private mSettings As CNxFocusSettingsModel
#If VBA7 Then
Private mTimerId As LongPtr
#Else
Private mTimerId As Long
#End If
Private mExpectedLeft(1 To NX_FOCUS_SEGMENT_COUNT) As Long
Private mExpectedTop(1 To NX_FOCUS_SEGMENT_COUNT) As Long
Private mExpectedWidth(1 To NX_FOCUS_SEGMENT_COUNT) As Long
Private mExpectedHeight(1 To NX_FOCUS_SEGMENT_COUNT) As Long
Private mExpectedVisible(1 To NX_FOCUS_SEGMENT_COUNT) As Boolean
Private mRefreshing As Boolean
Private mStopping As Boolean
Private mSuspended As Boolean
Private mRefreshPending As Boolean
Private mRefreshCount As Long
Private mLastStateSignature As String
Private mLastSelectionSignature As String
Private mLastError As String

Public Function NxFocusOverlayStart(ByVal host As Application, ByVal settings As CNxFocusSettingsModel) As Boolean
    Dim fillColor As Long
    Dim alphaValue As Byte
    Dim index As Long
    On Error GoTo Failed
    NxFocusOverlayStop
    If host Is Nothing Then NxRaiseContractError "Excel 실행 상태를 확인하세요."
    If settings Is Nothing Then NxRaiseContractError "포커스셀 설정을 확인하세요."
    If Not settings.IsSealed Then NxRaiseContractError "포커스셀 설정을 확인하세요."
    Set mHost = host
    Set mSettings = settings
    fillColor = NxFocusEffectiveColor(mSettings)
    alphaValue = NxFocusAlphaByte(mSettings)
    Set mForms = New Collection
    For index = 1 To NX_FOCUS_SEGMENT_COUNT
        NxCreateOverlay NxFocusSegmentToken(index), fillColor, alphaValue
    Next index
    mRefreshCount = 0
    mLastStateSignature = vbNullString
    mRefreshPending = True
    NxFocusOverlayRefresh
    NxStartOverlayTimer
    NxFocusOverlayStart = NxFocusOverlayCanRender()
    Exit Function
Failed:
    mLastError = Err.Description
    NxFocusOverlayStop
End Function

Private Sub NxCreateOverlay(ByVal token As String, ByVal fillColor As Long, ByVal alphaValue As Byte)
    Dim overlay As FNxFocusOverlay
    Dim activeWindow As Window
    Set activeWindow = mHost.ActiveWindow
    If activeWindow Is Nothing Then Err.Raise vbObjectError + 2142, "NxFocusOverlay", "Active Excel window unavailable"
    Set overlay = New FNxFocusOverlay
    overlay.Configure token, fillColor, alphaValue, activeWindow.hWnd
    overlay.OpenOverlay
    overlay.HideOverlay
    mForms.Add overlay, token
End Sub

Public Sub NxFocusOverlayApplySettings(ByVal settings As CNxFocusSettingsModel)
    Dim index As Long
    Dim overlay As FNxFocusOverlay
    Dim fillColor As Long
    Dim alphaValue As Byte
    If settings Is Nothing Then NxRaiseContractError "포커스셀 설정을 확인하세요."
    If Not settings.IsSealed Then NxRaiseContractError "포커스셀 설정을 확인하세요."
    If mForms Is Nothing Then NxRaiseContractError "포커스셀 화면이 시작되지 않았습니다."
    fillColor = NxFocusEffectiveColor(settings)
    alphaValue = NxFocusAlphaByte(settings)
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        overlay.ApplyAppearance fillColor, alphaValue
    Next index
    Set mSettings = settings
    NxFocusOverlayRequestRefresh
    NxFocusOverlayRefresh
End Sub

#If VBA7 Then
Private Sub NxRebindOverlayOwner(ByVal ownerHandle As LongPtr)
#Else
Private Sub NxRebindOverlayOwner(ByVal ownerHandle As Long)
#End If
    Dim index As Long
    Dim overlay As FNxFocusOverlay
    If ownerHandle = 0 Then Err.Raise vbObjectError + 2143, "NxFocusOverlay", "Active Excel window handle unavailable"
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        overlay.SetOwnerHandle ownerHandle
    Next index
End Sub

Private Sub NxStartOverlayTimer()
    If mTimerId <> 0 Then Exit Sub
    mTimerId = SetTimer(0, 0, NX_OVERLAY_TIMER_MS, AddressOf NxFocusOverlayTimerProc)
    If mTimerId = 0 Then Err.Raise vbObjectError + 2141, "NxFocusOverlay", "Overlay timer unavailable"
End Sub

Private Sub NxStopOverlayTimer()
    If mTimerId <> 0 Then Call KillTimer(0, mTimerId)
    mTimerId = 0
End Sub

Public Sub NxFocusOverlayRequestRefresh()
    mRefreshPending = True
End Sub

Public Sub NxFocusOverlayRefresh()
    Dim activeWindow As Window
    Dim selectionRange As Range
    Dim selectedArea As Range
    Dim selectedSheet As Worksheet
    Dim geometry As CNxFocusGeometry
    Dim index As Long
    If mRefreshing Or mStopping Or mSuspended Then Exit Sub
    If mHost Is Nothing Or mForms Is Nothing Or mSettings Is Nothing Then Exit Sub
    On Error GoTo Failed
    mRefreshing = True
    Set activeWindow = mHost.ActiveWindow
    If activeWindow Is Nothing Then GoTo Hidden
    NxRebindOverlayOwner activeWindow.hWnd
    If TypeName(mHost.ActiveSheet) <> "Worksheet" Then GoTo Hidden
    Set selectedSheet = mHost.ActiveSheet
    If selectedSheet.Parent Is ThisWorkbook Then GoTo Hidden
    If selectedSheet.Visible <> xlSheetVisible Then GoTo Hidden
    If TypeName(mHost.Selection) <> "Range" Then GoTo Hidden
    Set selectionRange = mHost.Selection
    If selectionRange.Areas.Count <> 1 Then GoTo Hidden
    If CDbl(selectionRange.CountLarge) > NX_FOCUS_MAX_SELECTION_CELLS Then GoTo Hidden
    If CDbl(selectionRange.Rows.CountLarge) = CDbl(selectedSheet.Rows.Count) Then GoTo Hidden
    If CDbl(selectionRange.Columns.CountLarge) = CDbl(selectedSheet.Columns.Count) Then GoTo Hidden
    Set selectedArea = selectionRange.Areas(1)
    If Not selectedArea.Parent Is selectedSheet Then GoTo Hidden
    Set geometry = NxFocusBuildGeometry(mHost, selectedArea)
    If geometry Is Nothing Then mLastError = NxFocusGeometryLastDiagnostic(): GoTo Hidden

    For index = 1 To NX_FOCUS_SEGMENT_COUNT
        If NxFocusSegmentEnabled(index, mSettings) And geometry.SegmentVisible(index) Then
            NxMoveOverlay index, geometry
        Else
            NxHideOverlay index
        End If
    Next index
    mLastSelectionSignature = geometry.SelectionSignature
    mLastStateSignature = NxFocusGeometryStateSignature(mHost)
    mRefreshPending = False
    mRefreshCount = mRefreshCount + 1
    mLastError = vbNullString
    mRefreshing = False
    Exit Sub
Hidden:
    NxFocusOverlayHideAll
    mLastStateSignature = NxFocusGeometryStateSignature(mHost)
    mRefreshPending = False
    mRefreshing = False
    Exit Sub
Failed:
    mLastError = Err.Description
    NxFocusOverlayHideAll
    mRefreshPending = False
    mRefreshing = False
End Sub

Private Function NxFocusSegmentEnabled(ByVal index As Long, ByVal settings As CNxFocusSettingsModel) As Boolean
    Select Case index
        Case NX_FOCUS_SEGMENT_ROW_LEFT, NX_FOCUS_SEGMENT_ROW_RIGHT
            NxFocusSegmentEnabled = (settings.Shape = NX_FOCUS_SHAPE_CROSS Or settings.Shape = NX_FOCUS_SHAPE_HORIZONTAL)
        Case NX_FOCUS_SEGMENT_COLUMN_TOP, NX_FOCUS_SEGMENT_COLUMN_BOTTOM
            NxFocusSegmentEnabled = (settings.Shape = NX_FOCUS_SHAPE_CROSS Or settings.Shape = NX_FOCUS_SHAPE_VERTICAL)
        Case NX_FOCUS_SEGMENT_SELECT_TOP To NX_FOCUS_SEGMENT_SELECT_LEFT
            NxFocusSegmentEnabled = settings.HighlightSelectedArea
    End Select
End Function

Private Sub NxMoveOverlay(ByVal index As Long, ByVal geometry As CNxFocusGeometry)
    Dim overlay As FNxFocusOverlay
    Set overlay = mForms.Item(index)
    overlay.MoveOverlay geometry.SegmentLeft(index), geometry.SegmentTop(index), _
        geometry.SegmentWidth(index), geometry.SegmentHeight(index)
    mExpectedLeft(index) = geometry.SegmentLeft(index)
    mExpectedTop(index) = geometry.SegmentTop(index)
    mExpectedWidth(index) = geometry.SegmentWidth(index)
    mExpectedHeight(index) = geometry.SegmentHeight(index)
    mExpectedVisible(index) = True
End Sub

Private Sub NxHideOverlay(ByVal index As Long)
    Dim overlay As FNxFocusOverlay
    If mForms Is Nothing Then Exit Sub
    If index < 1 Or index > mForms.Count Then Exit Sub
    Set overlay = mForms.Item(index)
    overlay.HideOverlay
    mExpectedVisible(index) = False
End Sub

Public Sub NxFocusOverlayHideAll()
    Dim index As Long
    On Error Resume Next
    If mForms Is Nothing Then Exit Sub
    For index = 1 To mForms.Count
        NxHideOverlay index
    Next index
    On Error GoTo 0
End Sub

Public Sub NxFocusOverlaySuspend()
    If mSuspended Then Exit Sub
    mSuspended = True
    NxStopOverlayTimer
    NxFocusOverlayHideAll
End Sub

Public Sub NxFocusOverlayResume()
    If Not mSuspended Then Exit Sub
    mSuspended = False
    NxFocusOverlayRequestRefresh
    NxFocusOverlayRefresh
    NxStartOverlayTimer
End Sub

Public Sub NxFocusOverlayStop()
    Dim index As Long
    Dim overlay As FNxFocusOverlay
    On Error Resume Next
    mStopping = True
    NxStopOverlayTimer
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
    Set mSettings = Nothing
    mRefreshing = False
    mSuspended = False
    mRefreshPending = False
    mLastStateSignature = vbNullString
    mStopping = False
    On Error GoTo 0
End Sub

#If VBA7 Then
Public Sub NxFocusOverlayTimerProc(ByVal hWnd As LongPtr, ByVal message As Long, ByVal timerIdentifier As LongPtr, ByVal elapsed As Long)
#Else
Public Sub NxFocusOverlayTimerProc(ByVal hWnd As Long, ByVal message As Long, ByVal timerIdentifier As Long, ByVal elapsed As Long)
#End If
    Dim currentSignature As String
    On Error Resume Next
    If mStopping Or mSuspended Or mHost Is Nothing Then Exit Sub
    currentSignature = NxFocusGeometryStateSignature(mHost)
    If mRefreshPending Or StrComp(currentSignature, mLastStateSignature, vbBinaryCompare) <> 0 Then NxFocusOverlayRefresh
    On Error GoTo 0
End Sub

Public Function NxFocusOverlayWindowCount() As Long
    Dim index As Long
    Dim overlay As FNxFocusOverlay
    On Error GoTo Done
    If mForms Is Nothing Then Exit Function
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        If overlay.IsAlive Then NxFocusOverlayWindowCount = NxFocusOverlayWindowCount + 1
    Next index
Done:
End Function

Public Function NxFocusOverlayVisibleWindowCount() As Long
    Dim index As Long
    Dim overlay As FNxFocusOverlay
    On Error GoTo Done
    If mForms Is Nothing Then Exit Function
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        If overlay.IsVisible Then NxFocusOverlayVisibleWindowCount = NxFocusOverlayVisibleWindowCount + 1
    Next index
Done:
End Function

Public Function NxFocusOverlayTimerCount() As Long
    If mTimerId <> 0 Then NxFocusOverlayTimerCount = 1
End Function

Public Function NxFocusOverlayStyleContractSatisfied() As Boolean
    Dim index As Long
    Dim overlay As FNxFocusOverlay
    On Error GoTo Failed
    If NxFocusOverlayWindowCount() <> NX_FOCUS_SEGMENT_COUNT Then Exit Function
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        If Not overlay.StyleContractSatisfied Then Exit Function
    Next index
    NxFocusOverlayStyleContractSatisfied = True
    Exit Function
Failed:
End Function

Public Function NxFocusOverlayCanRender() As Boolean
    NxFocusOverlayCanRender = (NxFocusOverlayWindowCount() = NX_FOCUS_SEGMENT_COUNT And _
        NxFocusOverlayStyleContractSatisfied() And Len(mLastError) = 0)
End Function

Public Function NxFocusOverlayExpectedWindowCount() As Long
    NxFocusOverlayExpectedWindowCount = NX_FOCUS_SEGMENT_COUNT
End Function

Public Function NxFocusOverlayExpectedBoundsMatch(Optional ByVal tolerance As Long = 2) As Boolean
    Dim index As Long
    Dim overlay As FNxFocusOverlay
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
    NxFocusOverlayExpectedBoundsMatch = True
    Exit Function
Failed:
End Function

Public Function NxFocusOverlayBoundsSignature() As String
    Dim index As Long
    Dim overlay As FNxFocusOverlay
    Dim result As String
    On Error GoTo Done
    If mForms Is Nothing Then Exit Function
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        If Len(result) > 0 Then result = result & "|"
        result = result & overlay.BoundsSignature
    Next index
Done:
    NxFocusOverlayBoundsSignature = result
End Function

Public Function NxFocusOverlayAppearanceSignature() As String
    Dim index As Long
    Dim overlay As FNxFocusOverlay
    Dim result As String
    On Error GoTo Done
    If mForms Is Nothing Then Exit Function
    For index = 1 To mForms.Count
        Set overlay = mForms.Item(index)
        If Len(result) > 0 Then result = result & "|"
        result = result & overlay.AppearanceSignature
    Next index
Done:
    NxFocusOverlayAppearanceSignature = result
End Function

Public Function NxFocusOverlaySelectionSignature() As String
    NxFocusOverlaySelectionSignature = mLastSelectionSignature
End Function

Public Function NxFocusOverlayRefreshCount() As Long
    NxFocusOverlayRefreshCount = mRefreshCount
End Function

Public Function NxFocusOverlayLastError() As String
    NxFocusOverlayLastError = mLastError
End Function

Private Function NxFocusSegmentToken(ByVal index As Long) As String
    Select Case index
        Case NX_FOCUS_SEGMENT_ROW_LEFT: NxFocusSegmentToken = "ROW_LEFT"
        Case NX_FOCUS_SEGMENT_ROW_RIGHT: NxFocusSegmentToken = "ROW_RIGHT"
        Case NX_FOCUS_SEGMENT_COLUMN_TOP: NxFocusSegmentToken = "COLUMN_TOP"
        Case NX_FOCUS_SEGMENT_COLUMN_BOTTOM: NxFocusSegmentToken = "COLUMN_BOTTOM"
        Case NX_FOCUS_SEGMENT_SELECT_TOP: NxFocusSegmentToken = "SELECT_TOP"
        Case NX_FOCUS_SEGMENT_SELECT_RIGHT: NxFocusSegmentToken = "SELECT_RIGHT"
        Case NX_FOCUS_SEGMENT_SELECT_BOTTOM: NxFocusSegmentToken = "SELECT_BOTTOM"
        Case NX_FOCUS_SEGMENT_SELECT_LEFT: NxFocusSegmentToken = "SELECT_LEFT"
        Case Else: NxRaiseContractError "포커스셀 구간 번호가 올바르지 않습니다."
    End Select
End Function
