Attribute VB_Name = "NxFormWheel"
Option Explicit

Private Const WH_MOUSE As Long = 7
Private Const HC_ACTION As Long = 0
Private Const GA_ROOT As Long = 2
Private Const GWLP_HWNDPARENT As Long = -8
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const NX_FORM_WHEEL_ROWS As Long = 3

#If VBA7 Then
#If Win64 Then
Private Declare PtrSafe Function GetWindowLongPtrW Lib "user32" (ByVal hwnd As LongPtr, ByVal index As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtrW Lib "user32" (ByVal hwnd As LongPtr, ByVal index As Long, ByVal value As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function GetWindowLongPtrW Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As LongPtr, ByVal index As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLongPtrW Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As LongPtr, ByVal index As Long, ByVal value As LongPtr) As LongPtr
#End If
Private Declare PtrSafe Function FindWindowExW Lib "user32" (ByVal parent As LongPtr, ByVal after As LongPtr, ByVal className As LongPtr, ByVal windowName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef processId As Long) As Long
Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare PtrSafe Function GetAncestor Lib "user32" (ByVal hwnd As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SetWindowsHookExW Lib "user32" (ByVal kind As Long, ByVal callback As LongPtr, ByVal instance As LongPtr, ByVal threadId As Long) As LongPtr
Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hook As LongPtr) As Long
Private Declare PtrSafe Function CallNextHookEx Lib "user32" (ByVal hook As LongPtr, ByVal code As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByVal source As LongPtr, ByVal length As LongPtr)
Private mWindow As LongPtr
Private mHook As LongPtr
#Else
Private Declare Function GetWindowLongPtrW Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As Long, ByVal index As Long) As Long
Private Declare Function SetWindowLongPtrW Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal index As Long, ByVal value As Long) As Long
Private Declare Function FindWindowExW Lib "user32" (ByVal parent As Long, ByVal after As Long, ByVal className As Long, ByVal windowName As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef processId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetAncestor Lib "user32" (ByVal hwnd As Long, ByVal flags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowsHookExW Lib "user32" (ByVal kind As Long, ByVal callback As Long, ByVal instance As Long, ByVal threadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hook As Long, ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByVal source As Long, ByVal length As Long)
Private mWindow As Long
Private mHook As Long
#End If

Private mHost As Object
Private mAttached As Boolean
Private mHandling As Boolean

Public Function NxFormWheelAttach(ByVal host As Object) As Boolean
#If VBA7 Then
    Dim formWindow As LongPtr
#Else
    Dim formWindow As Long
#End If
    Dim processId As Long, caption As String
    If host Is Nothing Then Exit Function
    If mAttached Then
        If host Is mHost Then NxFormWheelAttach = True: Exit Function
        NxFormWheelDetach mHost
        If mAttached Then Exit Function
    End If
    On Error GoTo Failed
    caption = CStr(host.Caption)
    ' Captions can repeat in another Excel process. Match only our UI thread.
    Do
        formWindow = FindWindowExW(0, formWindow, 0, StrPtr(caption))
        If formWindow = 0 Then Exit Function
        If GetWindowThreadProcessId(formWindow, processId) = GetCurrentThreadId() Then Exit Do
    Loop
    If formWindow = 0 Then Exit Function
    Set mHost = host
    mWindow = formWindow
    ' MSForms child controls consume wheel messages before the form WndProc.
    ' A same-thread hook needs no external DLL and leaves other processes alone.
    mHook = SetWindowsHookExW(WH_MOUSE, AddressOf NxFormWheelProc, 0, GetCurrentThreadId())
    If mHook = 0 Then GoTo Failed
    mAttached = True
    NxFormWheelAttach = True
    Exit Function
Failed:
    NxFormWheelDetach host
    Err.Clear
End Function

Public Sub NxFormWheelDetach(ByVal host As Object)
    If host Is Nothing Then Exit Sub
    If Not host Is mHost Then Exit Sub
    On Error Resume Next
    If mHook <> 0 Then
        ' Keep ownership on an OS unhook failure; never abandon a live callback.
        If UnhookWindowsHookEx(mHook) = 0 Then Exit Sub
    End If
    mAttached = False
    mWindow = 0
    mHook = 0
    Set mHost = Nothing
    On Error GoTo 0
End Sub

Public Function NxFormWheelIsAttached() As Boolean
    NxFormWheelIsAttached = mAttached
End Function

Public Function NxFormWheelActivateHost(ByVal host As Object) As Boolean
#If VBA7 Then
    Dim ownerWindow As LongPtr
#Else
    Dim ownerWindow As Long
#End If
    Dim processId As Long
    On Error GoTo Failed
    If host Is Nothing Then Exit Function
    If Not host.Visible Then Exit Function
    ' Explicit navigation may put an SDI workbook above its modeless owner.
    ' Never raise the form over another application or another Excel UI thread.
    If GetWindowThreadProcessId(GetForegroundWindow(), processId) <> GetCurrentThreadId() Then Exit Function
    If Not NxFormWheelAttach(host) Then Exit Function
    If Not host Is mHost Or mWindow = 0 Then Exit Function
    ownerWindow = Application.ActiveWindow.Hwnd
    If ownerWindow = 0 Or ownerWindow = mWindow Then Exit Function
    If GetWindowThreadProcessId(ownerWindow, processId) <> GetCurrentThreadId() Then Exit Function
    ' Rebind only this top-level form, not Excel or a child control. Otherwise
    ' raising the form also raises its old SDI workbook behind the new target.
    If GetWindowLongPtrW(mWindow, GWLP_HWNDPARENT) <> ownerWindow Then
        Call SetWindowLongPtrW(mWindow, GWLP_HWNDPARENT, ownerWindow)
        If GetWindowLongPtrW(mWindow, GWLP_HWNDPARENT) <> ownerWindow Then Exit Function
    End If
    NxFormWheelActivateHost = (SetForegroundWindow(mWindow) <> 0)
    Exit Function
Failed:
    Err.Clear
End Function

Public Sub NxFormWheelScrollActiveList(ByVal host As Object, ByVal defaultControlName As String, ByVal wheelDelta As Long)
    Dim target As Object
    On Error Resume Next
    Set target = host.ActiveControl
    If target Is Nothing Then
        Set target = host.Controls(defaultControlName)
    ElseIf TypeName(target) <> "ListBox" Then
        Set target = host.Controls(defaultControlName)
    End If
    On Error GoTo 0
    If target Is Nothing Then Exit Sub
    If TypeName(target) <> "ListBox" Then Exit Sub
    NxFormWheelScrollList target, wheelDelta, NX_FORM_WHEEL_ROWS
End Sub

Public Sub NxFormWheelScrollList(ByVal target As Object, ByVal wheelDelta As Long, Optional ByVal rows As Long = NX_FORM_WHEEL_ROWS)
    Dim requestedTop As Long
    Dim maximumTop As Long
    If target Is Nothing Then Exit Sub
    If target.ListCount <= 0 Or wheelDelta = 0 Then Exit Sub
    If rows < 1 Then rows = 1
    requestedTop = CLng(target.TopIndex)
    If wheelDelta > 0 Then requestedTop = requestedTop - rows Else requestedTop = requestedTop + rows
    maximumTop = target.ListCount - 1
    If requestedTop < 0 Then requestedTop = 0
    If requestedTop > maximumTop Then requestedTop = maximumTop
    target.TopIndex = requestedTop
End Sub

#If VBA7 Then
Public Function NxFormWheelProc(ByVal code As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Dim targetWindow As LongPtr
#Else
Public Function NxFormWheelProc(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim targetWindow As Long
#End If
    Dim wheelDelta As Integer
    Dim handlingHere As Boolean
    ' Pass HC_NOREMOVE too: PeekMessage must not apply the same wheel twice.
    If code <> HC_ACTION Then GoTo ForwardMessage
    If wParam <> WM_MOUSEWHEEL Then GoTo ForwardMessage
    If mHandling Then GoTo ForwardMessage
    If Not mAttached Or mHost Is Nothing Then GoTo ForwardMessage
    If lParam = 0 Then GoTo ForwardMessage
    On Error GoTo ForwardMessage
    ' MOUSEHOOKSTRUCTEX: POINT precedes HWND on both architectures.
#If Win64 Then
    CopyMemory targetWindow, lParam + 8, 8
    CopyMemory wheelDelta, lParam + 34, 2
#Else
    CopyMemory targetWindow, lParam + 8, 4
    CopyMemory wheelDelta, lParam + 22, 2
#End If
    If GetAncestor(targetWindow, GA_ROOT) <> mWindow Then GoTo ForwardMessage
    If GetAncestor(GetForegroundWindow(), GA_ROOT) <> mWindow Then GoTo ForwardMessage
    If wheelDelta = 0 Then GoTo ForwardMessage
    mHandling = True
    handlingHere = True
    CallByName mHost, "HandleWheelDelta", VbMethod, CLng(wheelDelta)
    mHandling = False
    NxFormWheelProc = 1
    Exit Function
ForwardMessage:
    If handlingHere Then mHandling = False
    Err.Clear
    NxFormWheelProc = CallNextHookEx(mHook, code, wParam, lParam)
End Function
