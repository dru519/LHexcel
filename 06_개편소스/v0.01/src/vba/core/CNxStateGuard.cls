VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxStateGuard"
Option Explicit

Private mScreenUpdating As Boolean
Private mEnableEvents As Boolean
Private mDisplayAlerts As Boolean
Private mCalculation As XlCalculation
Private mStatusBar As Variant
Private mWorkbook As Workbook
Private mSheet As Worksheet
Private mSelection As Range
Private mRestored As Boolean
Private mRecovery As String
Private mInjectedProperties As Collection

Private Sub Class_Initialize()
    Set mInjectedProperties = New Collection
    mScreenUpdating = Application.ScreenUpdating
    mEnableEvents = Application.EnableEvents
    mDisplayAlerts = Application.DisplayAlerts
    mCalculation = Application.Calculation
    mStatusBar = Application.StatusBar
    Set mWorkbook = Application.ActiveWorkbook
    Set mSheet = Application.ActiveSheet
    If TypeName(Application.Selection) = "Range" Then Set mSelection = Application.Selection
End Sub

Public Sub Restore()
    If mRestored Then Exit Sub
    mRestored = True
    RestoreScreenUpdating
    RestoreEnableEvents
    RestoreDisplayAlerts
    RestoreCalculation
    RestoreStatusBar
    RestoreActiveWorkbook
    RestoreActiveSheet
    RestoreSelection
End Sub

Public Sub EmergencyCleanup()
    On Error Resume Next
    Restore
    On Error GoTo 0
End Sub

Public Property Get Recovery() As String: Recovery = mRecovery: End Property
Public Property Get IsRestored() As Boolean: IsRestored = mRestored: End Property

Friend Sub InjectRestoreFailure(ByVal propertyName As String)
    If Len(Trim$(propertyName)) = 0 Then NxRaiseContractError "Restore property is required"
    mInjectedProperties.Add propertyName
End Sub

Private Sub RestoreScreenUpdating()
    On Error GoTo Failed
    RaiseIfInjected "ScreenUpdating"
    Application.ScreenUpdating = mScreenUpdating
    Exit Sub
Failed:
    AppendFailure "ScreenUpdating", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub RestoreEnableEvents()
    On Error GoTo Failed
    RaiseIfInjected "EnableEvents"
    Application.EnableEvents = mEnableEvents
    Exit Sub
Failed:
    AppendFailure "EnableEvents", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub RestoreDisplayAlerts()
    On Error GoTo Failed
    RaiseIfInjected "DisplayAlerts"
    Application.DisplayAlerts = mDisplayAlerts
    Exit Sub
Failed:
    AppendFailure "DisplayAlerts", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub RestoreCalculation()
    On Error GoTo Failed
    RaiseIfInjected "Calculation"
    ' Even assigning the same mode can mark saved workbooks dirty in Excel.
    If Application.Calculation <> mCalculation Then Application.Calculation = mCalculation
    Exit Sub
Failed:
    AppendFailure "Calculation", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub RestoreStatusBar()
    On Error GoTo Failed
    RaiseIfInjected "StatusBar"
    If IsEmpty(mStatusBar) Or VarType(mStatusBar) = vbBoolean Then
        Application.StatusBar = vbNullString
    Else
        Application.StatusBar = mStatusBar
    End If
    Exit Sub
Failed:
    AppendFailure "StatusBar", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub RestoreActiveWorkbook()
    On Error GoTo Failed
    RaiseIfInjected "ActiveWorkbook"
    If Not mWorkbook Is Nothing Then mWorkbook.Activate
    Exit Sub
Failed:
    AppendFailure "ActiveWorkbook", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub RestoreActiveSheet()
    On Error GoTo Failed
    RaiseIfInjected "ActiveSheet"
    If Not mSheet Is Nothing Then mSheet.Activate
    Exit Sub
Failed:
    AppendFailure "ActiveSheet", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub RestoreSelection()
    On Error GoTo Failed
    RaiseIfInjected "Selection"
    If Not mSelection Is Nothing Then mSelection.Select
    Exit Sub
Failed:
    AppendFailure "Selection", Err.Number, Err.Description
    Err.Clear
End Sub

Private Sub RaiseIfInjected(ByVal propertyName As String)
    Dim injected As Variant
    For Each injected In mInjectedProperties
        If CStr(injected) = propertyName Then Err.Raise vbObjectError + 2802, "CNxStateGuard", "Injected restore failure"
    Next injected
End Sub

Private Sub AppendFailure(ByVal propertyName As String, ByVal errorNumber As Long, ByVal description As String)
    mRecovery = NxJoinFailure(mRecovery, propertyName & "|" & CStr(errorNumber) & "|" & description)
End Sub

Private Sub Class_Terminate()
    EmergencyCleanup
End Sub
