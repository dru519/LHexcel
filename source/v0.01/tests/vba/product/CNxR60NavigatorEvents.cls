VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "CNxR60NavigatorEvents"
Option Explicit

Private WithEvents mApp As Excel.Application
Attribute mApp.VB_VarHelpID = -1
Private mTargetBook As Workbook
Private mOtherWindow As Window
Private mItem As CNxDocumentNavigatorItem
Private mKind As String
Private mInside As Boolean
Public Hits As Long
Public ReentrantSucceeded As Boolean
Public ReentrantReason As String
Public Failure As String

Public Sub Start(ByVal kind As String, ByVal targetBook As Workbook, ByVal otherWindow As Window, ByVal item As CNxDocumentNavigatorItem)
    mKind = kind
    Set mTargetBook = targetBook
    Set mOtherWindow = otherWindow
    Set mItem = item
    Set mApp = Application
End Sub

Public Sub StopListening()
    Set mApp = Nothing
    Set mTargetBook = Nothing
    Set mOtherWindow = Nothing
    Set mItem = Nothing
End Sub

Private Sub mApp_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    If mInside Then Exit Sub
    If Not Wb Is mTargetBook Then Exit Sub
    On Error GoTo Failed
    mInside = True
    Hits = Hits + 1
    Select Case mKind
        Case "event_redirect"
            mOtherWindow.Activate
        Case "event_reenter"
            ReentrantSucceeded = NxDocumentNavigatorMove(mItem, ReentrantReason)
        Case "event_close"
            NxDocumentNavigatorShutdown
    End Select
    mInside = False
    Exit Sub
Failed:
    Failure = CStr(Err.Number) & "|" & Err.Description
    mInside = False
End Sub
