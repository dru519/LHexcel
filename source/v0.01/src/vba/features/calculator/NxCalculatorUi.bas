Attribute VB_Name = "NxCalculatorUi"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal command As Long) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal command As Long) As Long
    Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

Private Const NX_CALCULATOR_CAPTION As String = "내엑셀 - 계산기"
Private Const NX_WINDOW_RESTORE As Long = 9
Private mHistory As CNxCalculatorHistory
Private mSession As CNxCalculatorSession
Private mForm As FNxCalculator

Public Sub NxCalculatorOpen()
    NxDistributionEnsureExecutable
    If Not mForm Is Nothing Then
        NxCalculatorActivateWindow
        Exit Sub
    End If
    If mHistory Is Nothing Then Set mHistory = New CNxCalculatorHistory
    Set mSession = New CNxCalculatorSession
    mSession.BindHistory mHistory
    Set mForm = New FNxCalculator
    mForm.Bind mSession, mHistory
    mForm.Show vbModeless
End Sub

Public Sub NxCalculatorNotifyClosed(ByVal closedForm As Object)
    If closedForm Is Nothing Then Exit Sub
    If mForm Is Nothing Then Exit Sub
    If Not (closedForm Is mForm) Then Exit Sub
    Set mForm = Nothing
    Set mSession = Nothing
End Sub

Private Sub NxCalculatorActivateWindow()
#If VBA7 Then
    Dim hwnd As LongPtr
#Else
    Dim hwnd As Long
#End If
    hwnd = FindWindowW(0, StrPtr(NX_CALCULATOR_CAPTION))
    If hwnd <> 0 Then
        ShowWindow hwnd, NX_WINDOW_RESTORE
        SetForegroundWindow hwnd
    Else
        mForm.Show vbModeless
    End If
End Sub
