Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal nCmdShow As Long) As Long
    Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function IsWindowVisible Lib "user32" (ByVal hWnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, ByRef lpRect As FxOverlayRect) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As FxOverlayRect) As Long
#End If

Private Type FxOverlayRect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_HWNDPARENT As Long = -8
Private Const WS_EX_TRANSPARENT As Long = &H20
Private Const WS_EX_TOOLWINDOW As Long = &H80
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_NOACTIVATE As Long = &H8000000
Private Const LWA_ALPHA As Long = &H2
Private Const SW_HIDE As Long = 0
Private Const SW_SHOWNOACTIVATE As Long = 4
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40
Private Const NX_OVERLAY_ALPHA As Byte = 52

#If VBA7 Then
Private mWindowHandle As LongPtr
Private mOwnerHandle As LongPtr
#Else
Private mWindowHandle As Long
Private mOwnerHandle As Long
#End If
Private mToken As String

#If VBA7 Then
Public Sub Configure(ByVal token As String, ByVal fillColor As Long, ByVal ownerHandle As LongPtr)
#Else
Public Sub Configure(ByVal token As String, ByVal fillColor As Long, ByVal ownerHandle As Long)
#End If
    mToken = token
    mOwnerHandle = ownerHandle
    Me.Caption = "NX_FOCUS_OVERLAY_" & token & "_" & Replace(CStr(Timer), ".", "_")
    Me.BackColor = fillColor
End Sub

Public Sub OpenOverlay()
    Me.Show vbModeless
    DoEvents
    mWindowHandle = FindWindow(vbNullString, Me.Caption)
    If mWindowHandle = 0 Then Err.Raise vbObjectError + 9801, "FNxFocusOverlayProbe", "Overlay window handle unavailable"

    Dim extendedStyle As LongPtr
    extendedStyle = GetWindowLongPtr(mWindowHandle, GWL_EXSTYLE)
    extendedStyle = extendedStyle Or WS_EX_LAYERED Or WS_EX_TRANSPARENT Or WS_EX_TOOLWINDOW Or WS_EX_NOACTIVATE
    Call SetWindowLongPtr(mWindowHandle, GWL_EXSTYLE, extendedStyle)
    Call SetWindowLongPtr(mWindowHandle, GWL_HWNDPARENT, mOwnerHandle)
    If SetLayeredWindowAttributes(mWindowHandle, 0, NX_OVERLAY_ALPHA, LWA_ALPHA) = 0 Then
        Err.Raise vbObjectError + 9802, "FNxFocusOverlayProbe", "Overlay alpha configuration failed"
    End If
    Call ShowWindow(mWindowHandle, SW_SHOWNOACTIVATE)
End Sub

Public Sub MoveOverlay(ByVal leftPixel As Long, ByVal topPixel As Long, ByVal widthPixel As Long, ByVal heightPixel As Long)
    If mWindowHandle = 0 Then Exit Sub
    If widthPixel < 1 Then widthPixel = 1
    If heightPixel < 1 Then heightPixel = 1
    If SetWindowPos(mWindowHandle, 0, leftPixel, topPixel, widthPixel, heightPixel, SWP_NOACTIVATE Or SWP_SHOWWINDOW) = 0 Then
        Err.Raise vbObjectError + 9803, "FNxFocusOverlayProbe", "Overlay positioning failed"
    End If
End Sub

Public Sub HideOverlay()
    If mWindowHandle <> 0 Then Call ShowWindow(mWindowHandle, SW_HIDE)
End Sub

Public Function IsAlive() As Boolean
    If mWindowHandle = 0 Then Exit Function
    IsAlive = (IsWindow(mWindowHandle) <> 0)
End Function

Public Function IsVisible() As Boolean
    If Not IsAlive Then Exit Function
    IsVisible = (IsWindowVisible(mWindowHandle) <> 0)
End Function

Public Function StyleContractSatisfied() As Boolean
    If Not IsAlive Then Exit Function
    Dim extendedStyle As LongPtr
    Dim owner As LongPtr
    extendedStyle = GetWindowLongPtr(mWindowHandle, GWL_EXSTYLE)
    owner = GetWindowLongPtr(mWindowHandle, GWL_HWNDPARENT)
    StyleContractSatisfied = ((extendedStyle And WS_EX_LAYERED) <> 0) _
        And ((extendedStyle And WS_EX_TRANSPARENT) <> 0) _
        And ((extendedStyle And WS_EX_TOOLWINDOW) <> 0) _
        And ((extendedStyle And WS_EX_NOACTIVATE) <> 0) _
        And (owner = mOwnerHandle)
End Function

Public Function BoundsMatch(ByVal expectedLeft As Long, ByVal expectedTop As Long, ByVal expectedWidth As Long, ByVal expectedHeight As Long, ByVal tolerance As Long) As Boolean
    Dim bounds As FxOverlayRect
    If Not IsAlive Then Exit Function
    If GetWindowRect(mWindowHandle, bounds) = 0 Then Exit Function
    BoundsMatch = Abs(bounds.Left - expectedLeft) <= tolerance _
        And Abs(bounds.Top - expectedTop) <= tolerance _
        And Abs((bounds.Right - bounds.Left) - expectedWidth) <= tolerance _
        And Abs((bounds.Bottom - bounds.Top) - expectedHeight) <= tolerance
End Function

Public Function BoundsSignature() As String
    Dim bounds As FxOverlayRect
    If Not IsAlive Then Exit Function
    If GetWindowRect(mWindowHandle, bounds) = 0 Then Exit Function
    BoundsSignature = mToken & ":" & CStr(bounds.Left) & "," & CStr(bounds.Top) & "," _
        & CStr(bounds.Right - bounds.Left) & "," & CStr(bounds.Bottom - bounds.Top)
End Function

Public Sub ReleaseHandle()
    mWindowHandle = 0
    mOwnerHandle = 0
End Sub
