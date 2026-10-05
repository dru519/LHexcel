Attribute VB_Name = "NxSymbolsUi"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal className As LongPtr, ByVal windowName As LongPtr) As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
Private Declare Function FindWindowW Lib "user32" (ByVal className As Long, ByVal windowName As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

Private Const NX_SYMBOLS_CAPTION As String = "내엑셀 - 기호표"
Private mForm As FNxSymbols
Private mRecent As CNxSymbolRecent
Private mEvents As CNxSymbolsApplicationEvents

Public Sub NxSymbolsOpen()
    NxDistributionEnsureExecutable
    If mRecent Is Nothing Then Set mRecent = New CNxSymbolRecent
    If mForm Is Nothing Then
        Set mForm = New FNxSymbols
        mForm.BindSession mRecent
        Set mEvents = New CNxSymbolsApplicationEvents
        mEvents.Bind Application, mForm
        mForm.Show vbModeless
    Else
        mForm.Show vbModeless
        ActivateSymbolsWindow
    End If
End Sub

Public Sub NxSymbolsFormClosed(ByVal closedForm As FNxSymbols)
    NxSymbolsWheelDetach closedForm
    If Not mEvents Is Nothing Then mEvents.Unbind
    Set mEvents = Nothing
    Set mForm = Nothing
End Sub

Public Sub NxSymbolsClose()
    If mForm Is Nothing Then Exit Sub
    Unload mForm
End Sub

Public Function NxSymbolsSelectTab(ByVal tabId As String) As Boolean
    If mForm Is Nothing Then Exit Function
    NxSymbolsSelectTab = mForm.SelectTab(tabId)
End Function

Public Function NxSymbolsRuntimeGlyphCount() As Long
    If mForm Is Nothing Then Exit Function
    NxSymbolsRuntimeGlyphCount = mForm.RuntimeGlyphCount
End Function

Public Function NxSymbolsRuntimeBufferValue() As String
    If mForm Is Nothing Then Exit Function
    NxSymbolsRuntimeBufferValue = mForm.RuntimeBufferValue
End Function

Public Function NxSymbolsRuntimeScrollMaximum() As Long
    If mForm Is Nothing Then Exit Function
    NxSymbolsRuntimeScrollMaximum = mForm.RuntimeScrollMaximum
End Function

Public Function NxSymbolsRuntimeScrollValue() As Long
    If mForm Is Nothing Then Exit Function
    NxSymbolsRuntimeScrollValue = mForm.RuntimeScrollValue
End Function

Public Function NxSymbolsRuntimeActivateGlyph(ByVal oneBasedIndex As Long) As Boolean
    If mForm Is Nothing Then Exit Function
    NxSymbolsRuntimeActivateGlyph = mForm.RuntimeActivateGlyph(oneBasedIndex)
End Function

Public Function NxSymbolsRuntimeSetScrollValue(ByVal requestedValue As Long) As Boolean
    If mForm Is Nothing Then Exit Function
    NxSymbolsRuntimeSetScrollValue = mForm.RuntimeSetScrollValue(requestedValue)
End Function

Private Sub ActivateSymbolsWindow()
#If VBA7 Then
    Dim windowHandle As LongPtr
#Else
    Dim windowHandle As Long
#End If
    windowHandle = FindWindowW(0, StrPtr(NX_SYMBOLS_CAPTION))
    If windowHandle <> 0 Then SetForegroundWindow windowHandle
End Sub
