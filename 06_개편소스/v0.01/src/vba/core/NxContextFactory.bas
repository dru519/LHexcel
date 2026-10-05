Attribute VB_Name = "NxContextFactory"
Option Explicit
Option Private Module

Private mUnsavedBooks As Collection
Private mUnsavedTokens As Collection

Public Function CaptureCurrent() As CNxExecutionContext
    Dim context As New CNxExecutionContext
    context.Capture NX_CONTEXT_FACTORY_TOKEN
    Set CaptureCurrent = context
End Function

Public Function CaptureCurrentFingerprint() As String
    Dim context As New CNxExecutionContext
    context.Capture NX_CONTEXT_FACTORY_TOKEN, False
    CaptureCurrentFingerprint = context.Fingerprint
End Function

Public Function CaptureForFeature(ByVal featureId As String) As CNxExecutionContext
    Dim context As New CNxExecutionContext
    Select Case featureId
        Case "NX-FILE-WORKBOOK-COMPARE", "NX-FILE-SHEET-COMPARE", "NX-FILE-FILE-COMPARE"
            ' Comparison creates a new workbook; it does not journal source edits.
            ' Retain the complete source snapshot and approval checks.
            context.Capture NX_CONTEXT_FACTORY_TOKEN, True, 1000000
        Case Else
            context.Capture NX_CONTEXT_FACTORY_TOKEN
    End Select
    Set CaptureForFeature = context
End Function

Public Function UnsavedWorkbookSessionToken(ByVal workbook As Workbook) As String
    Dim index As Long
    Dim candidate As Workbook
    If workbook Is Nothing Then NxRaiseContractError "Unsaved workbook identity requires a workbook"
    If mUnsavedBooks Is Nothing Then
        Set mUnsavedBooks = New Collection
        Set mUnsavedTokens = New Collection
    End If
    For index = 1 To mUnsavedBooks.Count
        Set candidate = mUnsavedBooks.Item(index)
        If candidate Is workbook Then
            UnsavedWorkbookSessionToken = CStr(mUnsavedTokens.Item(index))
            Exit Function
        End If
    Next index
    Randomize Timer
    UnsavedWorkbookSessionToken = "NXUNSAVED1|" & Hex$(CLng(Timer * 1000#)) & "|" & Hex$(CLng(Rnd() * 2147483646#)) & "|" & CStr(mUnsavedBooks.Count + 1)
    mUnsavedBooks.Add workbook
    mUnsavedTokens.Add UnsavedWorkbookSessionToken
End Function
