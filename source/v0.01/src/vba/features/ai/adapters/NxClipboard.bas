Attribute VB_Name = "NxClipboard"
Option Explicit
Option Private Module

#If VBA7 Then
Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal windowHandle As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal formatId As Long, ByVal memoryHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal formatId As Long) As LongPtr
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal formatId As Long) As Long
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal byteCount As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal memoryHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal memoryHandle As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal memoryHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal valuePointer As LongPtr) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal destination As LongPtr, ByVal source As LongPtr, ByVal byteCount As LongPtr)
#Else
Private Declare Function OpenClipboard Lib "user32" (ByVal windowHandle As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal formatId As Long, ByVal memoryHandle As Long) As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal formatId As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal formatId As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal byteCount As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal memoryHandle As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal memoryHandle As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal memoryHandle As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal valuePointer As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal destination As Long, ByVal source As Long, ByVal byteCount As Long)
#End If

Private Const CF_UNICODETEXT As Long = 13
Private Const GMEM_MOVEABLE_ZEROINIT As Long = &H42
Private Const NX_CLIPBOARD_READ_ATTEMPTS As Long = 3

Private mTestMode As Boolean
Private mTestFailure As Boolean
Private mTestCallCount As Long
Private mTestLastValue As String
Private mReadTestMode As Boolean
Private mReadTestFailure As Boolean
Private mReadTestValue As String
Private mReadTestCallCount As Long

Public Sub NxClipboardSetTestFailure(ByVal enabled As Boolean)
    mTestMode = True
    mTestFailure = enabled
    mTestCallCount = 0
    mTestLastValue = vbNullString
End Sub

Public Function NxClipboardTestCallCount() As Long
    NxClipboardTestCallCount = mTestCallCount
End Function

Public Function NxClipboardTestLastValue() As String
    NxClipboardTestLastValue = mTestLastValue
End Function

Public Sub NxClipboardSetTestRead(ByVal value As String, Optional ByVal fail As Boolean = False)
    mReadTestMode = True
    mReadTestFailure = fail
    mReadTestValue = value
    mReadTestCallCount = 0
End Sub

Public Function NxClipboardTestReadCallCount() As Long
    NxClipboardTestReadCallCount = mReadTestCallCount
End Function

Public Sub NxClipboardResetTestMode()
    mTestMode = False
    mTestFailure = False
    mTestCallCount = 0
    mTestLastValue = vbNullString
    mReadTestMode = False
    mReadTestFailure = False
    mReadTestValue = vbNullString
    mReadTestCallCount = 0
End Sub

Public Function NxClipboardReadUnicode() As String
    mReadTestCallCount = mReadTestCallCount + 1
    If mReadTestMode Then
        If mReadTestFailure Then RaiseClipboardError "클립보드 읽기 시험 실패입니다."
        NxClipboardReadUnicode = mReadTestValue
        Exit Function
    End If

#If VBA7 Then
    Dim memoryHandle As LongPtr, memoryPointer As LongPtr
#Else
    Dim memoryHandle As Long, memoryPointer As Long
#End If
    Dim attempt As Long
    Dim characterCount As Long
    Dim value As String
    Dim clipboardOpened As Boolean
    Dim clipboardOpenedOnce As Boolean
    Dim errorNumber As Long, errorDescription As String
    On Error GoTo Failed

    For attempt = 1 To NX_CLIPBOARD_READ_ATTEMPTS
        If OpenClipboard(0) <> 0 Then
            clipboardOpened = True
            clipboardOpenedOnce = True
            If IsClipboardFormatAvailable(CF_UNICODETEXT) = 0 Then
                If CloseClipboard() = 0 Then RaiseClipboardError "Windows 클립보드를 닫지 못했습니다."
                clipboardOpened = False
                Exit Function
            End If
            memoryHandle = GetClipboardData(CF_UNICODETEXT)
            If memoryHandle <> 0 Then
                memoryPointer = GlobalLock(memoryHandle)
                If memoryPointer <> 0 Then
                    characterCount = lstrlenW(memoryPointer)
                    If characterCount > 0 Then
                        value = String$(characterCount, vbNullChar)
                        RtlMoveMemory StrPtr(value), memoryPointer, characterCount * 2
                    End If
                    Call GlobalUnlock(memoryHandle)
                    memoryPointer = 0
                    If CloseClipboard() = 0 Then RaiseClipboardError "Windows 클립보드를 닫지 못했습니다."
                    clipboardOpened = False
                    NxClipboardReadUnicode = value
                    Exit Function
                End If
            End If
            If memoryPointer <> 0 Then Call GlobalUnlock(memoryHandle)
            memoryPointer = 0
            If CloseClipboard() = 0 Then RaiseClipboardError "Windows 클립보드를 닫지 못했습니다."
            clipboardOpened = False
        End If
        DoEvents
    Next attempt

    If Not clipboardOpenedOnce Then RaiseClipboardError "Windows 클립보드를 열지 못했습니다."
    NxClipboardReadUnicode = vbNullString
    Exit Function

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    On Error Resume Next
    If memoryPointer <> 0 Then Call GlobalUnlock(memoryHandle)
    If clipboardOpened Then Call CloseClipboard
    On Error GoTo 0
    If errorNumber = 0 Then errorNumber = vbObjectError + 886
    Err.Raise errorNumber, "NxClipboard", errorDescription
End Function

Public Sub NxClipboardWriteUnicode(ByVal value As String)
    If Len(value) = 0 Then RaiseClipboardError "클립보드에 복사할 프롬프트가 없습니다."
    PutUnicodeText value
End Sub

Public Sub NxClipboardRestoreUnicode(ByVal value As String)
    PutUnicodeText value
End Sub

Private Sub PutUnicodeText(ByVal value As String)
    mTestCallCount = mTestCallCount + 1
    If mTestMode Then
        If mTestFailure Then RaiseClipboardError "클립보드 쓰기 시험 실패입니다."
        mTestLastValue = value
        Exit Sub
    End If

#If VBA7 Then
    Dim memoryHandle As LongPtr, memoryPointer As LongPtr
#Else
    Dim memoryHandle As Long, memoryPointer As Long
#End If
    Dim clipboardOpened As Boolean, ownershipTransferred As Boolean
    Dim errorNumber As Long, errorDescription As String
    On Error GoTo Failed

    memoryHandle = GlobalAlloc(GMEM_MOVEABLE_ZEROINIT, (Len(value) + 1) * 2)
    If memoryHandle = 0 Then RaiseClipboardError "Unicode 클립보드 메모리를 만들지 못했습니다."
    memoryPointer = GlobalLock(memoryHandle)
    If memoryPointer = 0 Then RaiseClipboardError "Unicode 클립보드 메모리를 잠그지 못했습니다."
    RtlMoveMemory memoryPointer, StrPtr(value), Len(value) * 2
    Call GlobalUnlock(memoryHandle)
    memoryPointer = 0

    If OpenClipboard(0) = 0 Then RaiseClipboardError "Windows 클립보드를 열지 못했습니다."
    clipboardOpened = True
    If EmptyClipboard() = 0 Then RaiseClipboardError "Windows 클립보드를 비우지 못했습니다."
    If SetClipboardData(CF_UNICODETEXT, memoryHandle) = 0 Then RaiseClipboardError "Unicode 텍스트를 클립보드에 넣지 못했습니다."
    ownershipTransferred = True
    memoryHandle = 0
    If CloseClipboard() = 0 Then RaiseClipboardError "Windows 클립보드를 닫지 못했습니다."
    clipboardOpened = False
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    On Error Resume Next
    If memoryPointer <> 0 Then Call GlobalUnlock(memoryHandle)
    If clipboardOpened Then Call CloseClipboard
    If memoryHandle <> 0 And Not ownershipTransferred Then Call GlobalFree(memoryHandle)
    On Error GoTo 0
    If errorNumber = 0 Then errorNumber = vbObjectError + 886
    Err.Raise errorNumber, "NxClipboard", errorDescription
End Sub

Private Sub RaiseClipboardError(ByVal message As String)
    Err.Raise vbObjectError + 886, "NxClipboard", message
End Sub
