Attribute VB_Name = "NxFocusController"
Option Explicit

Private mFocusState As NxFocusState
Private mFocusEvents As CNxFocusApplicationEvents
Private mSettings As CNxFocusSettingsModel
Private mRuleEngine As CNxFocusRuleEngine
Private mSuspendCount As Long
Private mPendingSaveWorkbook As Workbook
Private mFocusBackend As String
Private mBackendDiagnostic As String

Public Function NxFocusControllerIsEnabled() As Boolean
    NxFocusControllerEnsure
    NxFocusControllerIsEnabled = (mFocusState = NxFocusOn)
End Function

Public Function NxFocusControllerBackend() As String
    If NxFocusControllerIsEnabled() Then NxFocusControllerBackend = mFocusBackend Else NxFocusControllerBackend = "off"
End Function

Public Function NxFocusControllerBackendDiagnostic() As String
    NxFocusControllerBackendDiagnostic = mBackendDiagnostic
End Function

Public Sub NxFocusControllerSetEnabled(ByVal enabled As Boolean)
    If enabled Then
        NxFocusControllerEnable
    Else
        NxFocusControllerDisable
    End If
End Sub

Public Sub NxFocusControllerToggle()
    NxFocusControllerSetEnabled Not NxFocusControllerIsEnabled()
End Sub

Public Sub NxFocusControllerEnable()
    NxDistributionEnsureExecutable
    Dim failure As String
    Dim previousEvents As Boolean
    NxFocusControllerEnsure
    previousEvents = Application.EnableEvents
    On Error GoTo Failed
    If mFocusState = NxFocusOn Then
        If Not previousEvents Then
            Application.EnableEvents = True
            NxFocusControllerRefreshActiveSelection True
        End If
        Exit Sub
    End If
    mFocusState = NxFocusEnabling
    Set mSettings = NxFocusLoadSettings()
    If Not NxFocusCleanupAllOpenWorkbooks() Then _
        NxRaiseContractError "이전 포커스셀 규칙을 정리하지 못했습니다."
    If Not NxFocusDeleteAddinCacheNames() Then _
        NxRaiseContractError "이전 포커스셀 캐시를 정리하지 못했습니다."
    If Not NxFocusControllerStartBackend() Then _
        NxRaiseContractError "포커스셀 화면 표시를 시작하지 못했습니다."
    Set mFocusEvents = New CNxFocusApplicationEvents
    mFocusEvents.Start Application
    mSuspendCount = 0
    mFocusState = NxFocusOn
    ' Explicit activation needs selection events even after another macro left
    ' Excel events disabled. Otherwise the initial highlight never moves.
    Application.EnableEvents = True
    NxFocusControllerRefreshBackend
    NxRibbonInvalidateFocus
    Exit Sub
Failed:
    failure = Err.Description
    On Error Resume Next
    If Not mFocusEvents Is Nothing Then mFocusEvents.StopListening
    Set mFocusEvents = Nothing
    NxFocusControllerStopBackend
    mSuspendCount = 0
    Set mPendingSaveWorkbook = Nothing
    mFocusState = NxFocusOff
    Application.EnableEvents = previousEvents
    NxRibbonInvalidateFocus
    On Error GoTo 0
    If Len(failure) = 0 Then failure = "포커스셀을 켜지 못했습니다."
    NxRaiseContractError failure
End Sub

Public Sub NxFocusControllerDisable()
    Dim clean As Boolean
    NxFocusControllerEnsure
    If Not mFocusEvents Is Nothing Then mFocusEvents.StopListening
    Set mFocusEvents = Nothing
    NxFocusControllerStopBackend
    clean = NxFocusCleanupAllOpenWorkbooks()
    If Not NxFocusDeleteAddinCacheNames() Then clean = False
    mSuspendCount = 0
    Set mPendingSaveWorkbook = Nothing
    mFocusState = NxFocusOff
    mFocusBackend = vbNullString
    NxRibbonInvalidateFocus
    If Not clean Then NxRaiseContractError "포커스셀 관리 규칙을 모두 정리하지 못했습니다."
End Sub

Public Sub NxFocusControllerApplySettings(ByVal settings As CNxFocusSettingsModel)
    Dim wasEnabled As Boolean
    Dim failure As String
    Dim previousSettings As CNxFocusSettingsModel
    If settings Is Nothing Then NxRaiseContractError "포커스셀 설정을 확인하세요."
    If Not settings.IsSealed Then NxRaiseContractError "포커스셀 설정을 확인하세요."
    NxFocusControllerEnsure
    Set previousSettings = mSettings
    wasEnabled = NxFocusControllerIsEnabled()
    If wasEnabled Then
        On Error GoTo Failed
        If mFocusBackend = "dll" Then
            If Not NxHostFocusApply(settings) Then NxRaiseContractError "변경한 포커스셀 설정을 적용하지 못했습니다."
        ElseIf NxFocusControllerUsesConditionalBackend() Then
            Set mSettings = settings
            If mRuleEngine Is Nothing Then Set mRuleEngine = New CNxFocusRuleEngine
            mRuleEngine.Configure mSettings
            If Not NxFocusControllerRefreshConditionalSelection(Application.Selection, True) Then _
                NxRaiseContractError "변경한 포커스셀 설정을 적용하지 못했습니다."
        End If
        Set mSettings = settings
        NxRibbonInvalidateFocus
    ElseIf settings.EnableOnApply Then
        Set mSettings = settings
        NxFocusControllerEnable
    Else
        Set mSettings = settings
    End If
    Exit Sub
Failed:
    failure = Err.Description
    On Error Resume Next
    If Not mFocusEvents Is Nothing Then mFocusEvents.StopListening
    Set mFocusEvents = Nothing
    NxFocusControllerStopBackend
    On Error GoTo 0
    Set mSettings = previousSettings
    mFocusState = NxFocusOff
    NxRibbonInvalidateFocus
    If Len(failure) = 0 Then failure = "변경한 포커스셀 설정을 적용하지 못했습니다."
    NxRaiseContractError failure
End Sub

Public Sub NxFocusControllerCommitSettings(ByVal settings As CNxFocusSettingsModel)
    Dim previous As CNxFocusSettingsModel
    Dim wasEnabled As Boolean
    Dim failure As String
    If settings Is Nothing Then NxRaiseContractError "포커스셀 설정을 확인하세요."
    If Not settings.IsSealed Then NxRaiseContractError "포커스셀 설정을 확인하세요."
    Set previous = NxFocusLoadSettings()
    wasEnabled = NxFocusControllerIsEnabled()
    On Error GoTo Failed
    NxFocusControllerProbeSettings settings
    NxFocusStageSettingsModel settings
    NxFocusCommitStagedSettings
    NxFocusControllerApplySettings settings
    NxFocusCompleteSettingsTransaction
    Exit Sub
Failed:
    failure = Err.Description
    On Error Resume Next
    NxFocusRollbackSettingsFile
    NxFocusControllerRestoreSnapshot previous, wasEnabled
    On Error GoTo 0
    If Len(failure) = 0 Then failure = "변경한 포커스셀 설정을 적용하지 못했습니다."
    NxRaiseContractError failure
End Sub

Public Sub NxFocusControllerSelectionChanged(ByVal target As Range, Optional ByVal allowRepair As Boolean = True)
    If Not NxFocusControllerIsEnabled() Or mSuspendCount > 0 Then Exit Sub
    If NxFocusControllerUsesConditionalBackend() Then _
        NxFocusControllerRefreshConditionalSelection target, allowRepair
End Sub

Public Sub NxFocusControllerSuspend()
    If mSuspendCount < &H7FFFFFFF Then mSuspendCount = mSuspendCount + 1
    If mSuspendCount = 1 And NxFocusControllerIsEnabled() Then NxFocusControllerSuspendBackend
End Sub

Public Sub NxFocusControllerResume()
    If mSuspendCount > 0 Then mSuspendCount = mSuspendCount - 1
    If mSuspendCount = 0 And NxFocusControllerIsEnabled() Then NxFocusControllerResumeBackend
End Sub

Public Function NxFocusControllerBeforeSave(ByVal workbook As Workbook) As Boolean
    NxFocusControllerBeforeSave = True
    If Not NxFocusControllerIsEnabled() Then Exit Function
    If workbook Is Nothing Then Exit Function
    If Not mRuleEngine Is Nothing Then mRuleEngine.InvalidateInspection
    If NxFocusControllerUsesConditionalBackend() Then
        If Not NxFocusCleanupWorkbook(workbook) Then
            mBackendDiagnostic = "CF_CLEANUP_FAILED"
            NxFocusControllerBeforeSave = False
            Exit Function
        End If
    Else
        NxFocusControllerHideBackend
    End If
    Set mPendingSaveWorkbook = workbook
End Function

Public Sub NxFocusControllerAfterSave(ByVal workbook As Workbook, ByVal success As Boolean)
    If mPendingSaveWorkbook Is Nothing Then Exit Sub
    If workbook Is mPendingSaveWorkbook Then Set mPendingSaveWorkbook = Nothing
    If NxFocusControllerIsEnabled() Then NxFocusControllerRefreshBackend
End Sub

Public Function NxFocusControllerBeforeClose(ByVal workbook As Workbook) As Boolean
    NxFocusControllerBeforeClose = True
    If Not NxFocusControllerIsEnabled() Then Exit Function
    If workbook Is Nothing Then Exit Function
    If Not mRuleEngine Is Nothing Then mRuleEngine.InvalidateInspection
    If NxFocusControllerUsesConditionalBackend() Then
        If Not NxFocusCleanupWorkbook(workbook) Then
            mBackendDiagnostic = "CF_CLEANUP_FAILED"
            NxFocusControllerBeforeClose = False
            Exit Function
        End If
    Else
        NxFocusControllerHideBackend
    End If
    If Not mPendingSaveWorkbook Is Nothing Then
        If workbook Is mPendingSaveWorkbook Then Set mPendingSaveWorkbook = Nothing
    End If
End Function

Public Sub NxFocusControllerRefreshActiveSelection(Optional ByVal allowRepair As Boolean = True)
    If Not NxFocusControllerIsEnabled() Or mSuspendCount > 0 Then Exit Sub
    If Not allowRepair And Not mRuleEngine Is Nothing Then mRuleEngine.InvalidateInspection
    If NxFocusControllerUsesConditionalBackend() Then
        NxFocusControllerRefreshConditionalSelection Application.Selection, allowRepair
    Else
        NxFocusControllerRefreshBackend
    End If
End Sub

Public Sub NxFocusControllerWindowDeactivated(Optional ByVal workbook As Workbook)
    If Not NxFocusControllerIsEnabled() Then Exit Sub
    If Not mRuleEngine Is Nothing Then mRuleEngine.InvalidateInspection
    If NxFocusControllerUsesConditionalBackend() Then
        If workbook Is Nothing Then
            NxFocusControllerHideBackend
        ElseIf Not NxFocusCleanupWorkbook(workbook) Then
            mBackendDiagnostic = "CF_CLEANUP_FAILED"
        End If
    Else
        NxFocusControllerHideBackend
    End If
End Sub

Private Sub NxFocusControllerEnsure()
    If mSettings Is Nothing Then Set mSettings = NxFocusLoadSettings()
End Sub

Private Sub NxFocusControllerProbeSettings(ByVal candidate As CNxFocusSettingsModel)
    Dim previous As CNxFocusSettingsModel
    Dim wasEnabled As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String
    If candidate Is Nothing Then NxRaiseContractError "포커스셀 설정을 확인하세요."
    If Not candidate.IsSealed Then NxRaiseContractError "포커스셀 설정을 확인하세요."
    NxFocusControllerEnsure
    Set previous = mSettings
    wasEnabled = (mFocusState = NxFocusOn)
    On Error GoTo Failed
    If wasEnabled Then
        NxFocusControllerApplySettings candidate
        If Not NxFocusControllerBackendHealthy() Then NxRaiseContractError "변경한 포커스셀 설정을 표시할 수 없습니다."
        NxFocusControllerApplySettings previous
    Else
        Set mSettings = candidate
        If Not NxFocusControllerStartBackend() Then NxRaiseContractError "변경한 포커스셀 설정을 표시할 수 없습니다."
        If Not NxFocusControllerBackendHealthy() Then NxRaiseContractError "변경한 포커스셀 설정을 표시할 수 없습니다."
        NxFocusControllerStopBackend
        Set mSettings = previous
    End If
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    NxFocusControllerRestoreSnapshot previous, wasEnabled
    On Error GoTo 0
    Err.Raise failureNumber, "NxFocusController", failureDescription
End Sub

Private Function NxFocusControllerBackendHealthy() As Boolean
    If mFocusBackend = "dll" Then
        NxFocusControllerBackendHealthy = True
    ElseIf NxFocusControllerUsesConditionalBackend() Then
        NxFocusControllerBackendHealthy = _
            NxFocusControllerRefreshConditionalSelection(Application.Selection, False)
    End If
End Function

Private Function NxFocusControllerUsesConditionalBackend() As Boolean
    NxFocusControllerUsesConditionalBackend = _
        (mFocusBackend = "cf" Or mFocusBackend = "cf-fallback")
End Function

Private Sub NxFocusControllerRestoreSnapshot(ByVal previous As CNxFocusSettingsModel, ByVal wasEnabled As Boolean)
    On Error Resume Next
    If Not mFocusEvents Is Nothing Then mFocusEvents.StopListening
    Set mFocusEvents = Nothing
    NxFocusControllerStopBackend
    Set mSettings = previous
    mFocusState = NxFocusOff
    If wasEnabled Then
        If NxFocusControllerStartBackend() Then
            Set mFocusEvents = New CNxFocusApplicationEvents
            mFocusEvents.Start Application
            mFocusState = NxFocusOn
            NxFocusControllerRefreshBackend
        End If
    End If
    NxRibbonInvalidateFocus
    On Error GoTo 0
End Sub

Private Function NxFocusControllerStartBackend() As Boolean
    Dim hostCompatible As Boolean
    mFocusBackend = vbNullString
    mBackendDiagnostic = vbNullString
    hostCompatible = NxHostBridgeCompatible()
    If hostCompatible Then
        If NxHostFocusApply(mSettings) Then
            mFocusBackend = "dll"
            NxFocusControllerStartBackend = True
            Exit Function
        End If
    End If
    If NxFocusControllerStartConditionalBackend() Then
        If hostCompatible Then
            mFocusBackend = "cf-fallback"
            mBackendDiagnostic = "DLL_START_FAILED|CF_FALLBACK_ACTIVE"
        Else
            mFocusBackend = "cf"
        End If
        NxFocusControllerStartBackend = True
    End If
End Function

Private Sub NxFocusControllerStopBackend()
    If Not mRuleEngine Is Nothing Then mRuleEngine.InvalidateInspection
    If mFocusBackend = "dll" Then
        NxHostFocusStop
    ElseIf NxFocusControllerUsesConditionalBackend() Then
        If Not NxFocusCleanupAllOpenWorkbooks() Then mBackendDiagnostic = "CF_CLEANUP_FAILED"
        If Not NxFocusDeleteAddinCacheNames() Then mBackendDiagnostic = "CF_CLEANUP_FAILED"
    End If
    Set mRuleEngine = Nothing
    mFocusBackend = vbNullString
End Sub

Private Sub NxFocusControllerRefreshBackend()
    If mFocusBackend = "dll" Then
        If NxHostFocusApply(mSettings) Then Exit Sub
        NxHostFocusStop
        If Not NxFocusControllerStartConditionalBackend() Then _
            NxRaiseContractError "포커스셀 화면 표시를 갱신하지 못했습니다."
        mFocusBackend = "cf-fallback"
        mBackendDiagnostic = "DLL_REFRESH_FAILED|CF_FALLBACK_ACTIVE"
    ElseIf NxFocusControllerUsesConditionalBackend() Then
        NxFocusControllerRefreshConditionalSelection Application.Selection, True
    End If
End Sub

Private Sub NxFocusControllerSuspendBackend()
    If Not mRuleEngine Is Nothing Then mRuleEngine.InvalidateInspection
    If mFocusBackend = "dll" Then
        NxHostFocusStop
    ElseIf NxFocusControllerUsesConditionalBackend() Then
        If Not NxFocusCleanupAllOpenWorkbooks() Then mBackendDiagnostic = "CF_CLEANUP_FAILED"
    End If
End Sub

Private Sub NxFocusControllerResumeBackend()
    If mFocusBackend = "dll" Then
        If Not NxHostFocusApply(mSettings) Then
            If Not NxFocusControllerStartConditionalBackend() Then _
                NxRaiseContractError "포커스셀 화면 표시를 다시 시작하지 못했습니다."
            mFocusBackend = "cf-fallback"
            mBackendDiagnostic = "DLL_RESUME_FAILED|CF_FALLBACK_ACTIVE"
        End If
    ElseIf NxFocusControllerUsesConditionalBackend() Then
        NxFocusControllerRefreshConditionalSelection Application.Selection, True
    End If
End Sub

Private Sub NxFocusControllerHideBackend()
    Dim workbook As Workbook
    If Not mRuleEngine Is Nothing Then mRuleEngine.InvalidateInspection
    If mFocusBackend = "dll" Then
        NxHostFocusStop
    ElseIf NxFocusControllerUsesConditionalBackend() Then
        On Error Resume Next
        Set workbook = Application.ActiveWorkbook
        On Error GoTo 0
        If Not workbook Is Nothing Then
            If Not NxFocusCleanupWorkbook(workbook) Then mBackendDiagnostic = "CF_CLEANUP_FAILED"
        End If
    End If
End Sub

Private Function NxFocusControllerStartConditionalBackend() As Boolean
    Dim target As Object
    On Error GoTo Failed
    If mRuleEngine Is Nothing Then Set mRuleEngine = New CNxFocusRuleEngine
    mRuleEngine.Configure mSettings
    On Error Resume Next
    Set target = Application.Selection
    On Error GoTo Failed
    If target Is Nothing Then
        mBackendDiagnostic = "CF_UNSUPPORTED_SELECTION"
    Else
        If Not NxFocusControllerRefreshConditionalSelection(target, True) Then
            If Left$(mBackendDiagnostic, Len("CF_UNSUPPORTED_SELECTION")) <> _
               "CF_UNSUPPORTED_SELECTION" Then GoTo Failed
        End If
    End If
    NxFocusControllerStartConditionalBackend = True
    Exit Function
Failed:
    mBackendDiagnostic = "CF_INSTALL_FAILED"
    NxFocusControllerStartConditionalBackend = False
End Function

Private Function NxFocusControllerRefreshConditionalSelection(ByVal target As Object, _
    ByVal allowRepair As Boolean) As Boolean
    Dim selection As CNxFocusSelection
    Dim sheet As Worksheet
    Dim workbook As Workbook
    Dim wasSaved As Boolean
    On Error GoTo Failed
    If target Is Nothing Then
        mBackendDiagnostic = "CF_UNSUPPORTED_SELECTION"
        Exit Function
    End If
    If TypeName(target) <> "Range" Then
        mBackendDiagnostic = "CF_UNSUPPORTED_SELECTION"
        Exit Function
    End If
    Set sheet = target.Worksheet
    Set workbook = sheet.Parent
    Set selection = New CNxFocusSelection
    If Not selection.Configure(target) Then
        mBackendDiagnostic = "CF_UNSUPPORTED_SELECTION|" & selection.Reason
        If allowRepair And Not sheet.ProtectContents And Not workbook.ReadOnly Then _
            NxFocusDeleteManagedRules sheet
        Exit Function
    End If
    If mRuleEngine Is Nothing Then Set mRuleEngine = New CNxFocusRuleEngine
    mRuleEngine.Configure mSettings
    wasSaved = workbook.Saved
    If Not mRuleEngine.ApplySelection(selection, allowRepair) Then GoTo Failed
    If wasSaved Then workbook.Saved = True
    If mFocusBackend = "cf" Then mBackendDiagnostic = vbNullString
    NxFocusControllerRefreshConditionalSelection = True
    Exit Function
Failed:
    On Error Resume Next
    If Not workbook Is Nothing Then
        If wasSaved Then workbook.Saved = True
    End If
    On Error GoTo 0
    mBackendDiagnostic = "CF_INSTALL_FAILED"
    NxFocusControllerRefreshConditionalSelection = False
End Function

Public Function NxFocusTryCommitSettingsValues(ByVal shape As String, ByVal style As String, _
    ByVal colorHex As String, ByVal intensityPercent As Long, _
    ByVal highlightSelectedArea As Boolean, ByVal enableOnApply As Boolean) As String
    Dim settings As CNxFocusSettingsModel
    On Error GoTo Failed
    Set settings = NxFocusCreateSettings(shape, style, colorHex, intensityPercent, highlightSelectedArea, enableOnApply)
    NxFocusControllerCommitSettings settings
    NxFocusTryCommitSettingsValues = "PASS"
    Exit Function
Failed:
    NxFocusTryCommitSettingsValues = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function
