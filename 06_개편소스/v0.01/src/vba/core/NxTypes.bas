Attribute VB_Name = "NxTypes"
Option Explicit

Private Type NxGuid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef value As NxGuid) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef value As NxGuid, ByVal buffer As LongPtr, ByVal bufferLength As Long) As Long
#Else
Private Declare Function CoCreateGuid Lib "ole32" (ByRef value As NxGuid) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (ByRef value As NxGuid, ByVal buffer As Long, ByVal bufferLength As Long) As Long
#End If

Public Enum NxOutcome
    NxSuccess = 0
    NxCancelled = 1
    NxInputError = 2
    NxEnvironmentError = 3
    NxPartialFailure = 4
End Enum

Public Enum NxSideEffectKind
    NxCellMutation = 1
    NxFileCreate = 2
    NxFileReplace = 3
    NxClipboardWrite = 4
    NxExternalProcess = 5
    NxBrowserOpen = 6
    NxShapeInsert = 7
    NxBinaryPackageWrite = 8
    NxWorksheetCreate = 9
    NxShapeMutation = 10
    NxDirectoryBatchCreate = 11
    NxColumnInsert = 12
    ' NxFileWrite is represented by NxFileCreate/NxFileReplace; NxShellOpen by NxExternalProcess.
End Enum

Public Enum NxCellMutationMask
    NxCellValueMask = 1
    NxCellNumberFormatMask = 2
    NxCellFontMask = 4
    NxCellInteriorMask = 8
    NxCellAlignmentMask = 16
    NxCellProtectionMask = 32
    NxCellBorderMask = 64
    NxCellMergeMask = 128
    NxCellAllSupportedMask = 255
End Enum

Public Enum NxRollbackStrategy
    NxRollbackCellJournal = 1
    NxRollbackTypedCompensator = 2
    NxRollbackManualRecovery = 3
End Enum

Public Enum NxRiskLevel
    NxRiskL0 = 0
    NxRiskL1 = 1
    NxRiskL2 = 2
    NxRiskL3 = 3
    NxRiskL4 = 4
End Enum

Public Enum NxExecutionGrade
    NxExecutionFast = 0
    NxExecutionGuarded = 1
    NxExecutionPlanned = 2
End Enum

Public Enum NxCapability
    NxCapabilityNone = 0
    NxCapabilityRead = 1
    NxCapabilityCellMutation = 2
    NxCapabilityFileCreate = 4
    NxCapabilityFileReplace = 8
    NxCapabilityClipboardWrite = 16
    NxCapabilityExternalProcess = 32
    NxCapabilityBrowserOpen = 64
    NxCapabilitySensitiveOutput = 128
    NxCapabilityTemplateQuarantine = 256
    NxCapabilityAiProcessing = 512
    NxCapabilityExternalDataTransfer = 1024
    NxCapabilityShapeInsert = 2048
    NxCapabilityBinaryPackageWrite = 4096
    NxCapabilityWorksheetCreate = 8192
    NxCapabilityRiskyFormula = 16384
    NxCapabilityConditionalFormatting = 32768
    NxCapabilityDirectoryCreate = 65536
End Enum

Public Enum NxPrivacyState
    NxPrivacySafe = 0
    NxPrivacyCaution = 1
    NxPrivacyBlocked = 2
    NxPrivacyIncomplete = 3
End Enum

Public Enum NxCompatibilityState
    NxCompatibilityNative = 0
    NxCompatibilityFallback = 1
    NxCompatibilityUnavailable = 2
    NxCompatibilityIncomplete = 3
End Enum

Public Enum NxFrameState
    NxFrameInput = 0
    NxFrameChecking = 1
    NxFrameReady = 2
    NxFrameBlocked = 3
    NxFrameRunning = 4
    NxFrameSuccess = 5
    NxFrameCancelled = 6
    NxFrameError = 7
    NxFramePartialFailure = 8
End Enum

Public Enum NxFrameAction
    NxFrameBack = 0
    NxFrameCancelAction = 1
    NxFrameExecutePlan = 2
    NxFrameOpenResult = 3
    NxFrameCopyPath = 4
End Enum

Public Enum NxOfficeCapability
    NxOfficeCapabilityNone = 0
    NxOfficeCapabilityVba7 = 1
    NxOfficeCapabilityWindowsApi = 2
    NxOfficeCapabilityWin32Api = 2
    NxOfficeCapabilityWin64Api = 2
    NxOfficeCapabilityFileDialog = 8
    NxOfficeCapabilitySheetExport = 16
    NxOfficeCapabilityDynamicArray = 32
End Enum

' Closed payload vocabulary.  Target describes where an effect is recorded;
' adapter input is always the separately sealed payload.
Public Enum NxSideEffectPayloadKind
    NxPayloadNone = 0
    NxPayloadText = 1
    NxPayloadCommandLine = 2
    NxPayloadUri = 3
End Enum

Public Enum NxAiOperationKind
    NxAiPreviewOnly = 1
    NxAiSetValues = 2
    NxAiSetFormulas = 3
    NxAiSetNumberFormats = 4
    NxAiInsertImage = 5
End Enum

Public Const NX_CONTRACT_ERROR As Long = vbObjectError + 2801
Private Const NX_COMBO_TAG_PREFIX As String = "nxcombo1|"
Private Const NX_LIST_TAG_PREFIX As String = "nxlist1|"

Public Sub NxRaiseContractError(ByVal description As String)
    Err.Raise NX_CONTRACT_ERROR, "LHexcel.Core", description
End Sub

Public Function NxLHexcelProfileRoot() As String
    Dim localAppData As String
    Dim profileOverride As String
    profileOverride = Environ$("LHEXCEL_PROFILE_ROOT")
    If Len(profileOverride) > 0 Then
        If Len(profileOverride) > 240 Or InStr(profileOverride, vbNullChar) > 0 Then _
            NxRaiseContractError "LHEXCEL_PROFILE_ROOT is invalid"
        If Right$(profileOverride, 1) = "\" Then profileOverride = Left$(profileOverride, Len(profileOverride) - 1)
        NxLHexcelProfileRoot = profileOverride
        Exit Function
    End If
    localAppData = Environ$("LOCALAPPDATA")
    If Len(localAppData) = 0 Or InStr(localAppData, vbNullChar) > 0 Then NxRaiseContractError "LOCALAPPDATA is unavailable"
    If Right$(localAppData, 1) = "\" Then localAppData = Left$(localAppData, Len(localAppData) - 1)
    NxLHexcelProfileRoot = localAppData & "\LHexcel"
End Function

Public Sub NxInitializeStaticComboChoices(ByVal view As Object)
    Dim control As Object
    Dim metadata As String
    If view Is Nothing Then NxRaiseContractError "UserForm runtime view is required"
    For Each control In view.Controls
        If TypeName(control) = "ComboBox" Then
            metadata = CStr(control.Tag)
            If Left$(metadata, Len(NX_COMBO_TAG_PREFIX)) = NX_COMBO_TAG_PREFIX Then
                NxLoadStaticComboChoices control, Mid$(metadata, Len(NX_COMBO_TAG_PREFIX) + 1)
            End If
        End If
    Next control
End Sub

Public Sub NxInitializeStaticListChoices(ByVal view As Object)
    Dim control As Object
    Dim metadata As String
    If view Is Nothing Then NxRaiseContractError "UserForm runtime view is required"
    For Each control In view.Controls
        If TypeName(control) = "ListBox" Then
            metadata = CStr(control.Tag)
            If Left$(metadata, Len(NX_LIST_TAG_PREFIX)) = NX_LIST_TAG_PREFIX Then
                NxLoadStaticListChoices control, Mid$(metadata, Len(NX_LIST_TAG_PREFIX) + 1)
            End If
        End If
    Next control
End Sub

Private Sub NxLoadStaticComboChoices(ByVal combo As Object, ByVal metadata As String)
    Dim separator As Long
    Dim selectedIndexText As String
    Dim selectedIndex As Long
    Dim choicesPayload As String
    Dim choices As Variant
    Dim choiceIndex As Long

    separator = InStr(1, metadata, "|", vbBinaryCompare)
    If separator < 2 Then NxRaiseContractError "UserForm ComboBox metadata is invalid"
    selectedIndexText = Left$(metadata, separator - 1)
    choicesPayload = Mid$(metadata, separator + 1)
    If Len(choicesPayload) = 0 Or Not IsNumeric(selectedIndexText) Then NxRaiseContractError "UserForm ComboBox metadata is invalid"
    selectedIndex = CLng(selectedIndexText)
    If CStr(selectedIndex) <> selectedIndexText Then NxRaiseContractError "UserForm ComboBox metadata is invalid"

    choices = Split(choicesPayload, ChrW$(31), -1, vbBinaryCompare)
    If selectedIndex < LBound(choices) Or selectedIndex > UBound(choices) Then NxRaiseContractError "UserForm ComboBox selected index is invalid"
    combo.Clear
    For choiceIndex = LBound(choices) To UBound(choices)
        If Len(CStr(choices(choiceIndex))) = 0 Then NxRaiseContractError "UserForm ComboBox item is invalid"
        combo.AddItem CStr(choices(choiceIndex))
    Next choiceIndex
    combo.ListIndex = selectedIndex
End Sub

Private Sub NxLoadStaticListChoices(ByVal list As Object, ByVal metadata As String)
    Dim separator As Long
    Dim selectedIndexText As String
    Dim selectedIndex As Long
    Dim choicesPayload As String
    Dim choices As Variant
    Dim choiceIndex As Long

    separator = InStr(1, metadata, "|", vbBinaryCompare)
    If separator < 2 Then NxRaiseContractError "UserForm ListBox metadata is invalid"
    selectedIndexText = Left$(metadata, separator - 1)
    choicesPayload = Mid$(metadata, separator + 1)
    If Len(choicesPayload) = 0 Or Not IsNumeric(selectedIndexText) Then NxRaiseContractError "UserForm ListBox metadata is invalid"
    selectedIndex = CLng(selectedIndexText)
    If CStr(selectedIndex) <> selectedIndexText Then NxRaiseContractError "UserForm ListBox metadata is invalid"

    choices = Split(choicesPayload, ChrW$(31), -1, vbBinaryCompare)
    If selectedIndex < LBound(choices) Or selectedIndex > UBound(choices) Then NxRaiseContractError "UserForm ListBox selected index is invalid"
    list.Clear
    For choiceIndex = LBound(choices) To UBound(choices)
        If Len(CStr(choices(choiceIndex))) = 0 Then NxRaiseContractError "UserForm ListBox item is invalid"
        list.AddItem CStr(choices(choiceIndex))
    Next choiceIndex
    list.ListIndex = selectedIndex
End Sub

Public Function NxExecutionGradeWireName(ByVal grade As NxExecutionGrade) As String
    Select Case grade
        Case NxExecutionFast: NxExecutionGradeWireName = "fast"
        Case NxExecutionGuarded: NxExecutionGradeWireName = "guarded"
        Case NxExecutionPlanned: NxExecutionGradeWireName = "planned"
        Case Else: NxRaiseContractError "Unknown execution grade"
    End Select
End Function

Public Function NxSideEffectWireName(ByVal kind As NxSideEffectKind) As String
    Select Case kind
        Case NxCellMutation: NxSideEffectWireName = "cell_mutation"
        Case NxFileCreate: NxSideEffectWireName = "file_create"
        Case NxFileReplace: NxSideEffectWireName = "file_replace"
        Case NxClipboardWrite: NxSideEffectWireName = "clipboard_write"
        Case NxExternalProcess: NxSideEffectWireName = "external_process"
        Case NxBrowserOpen: NxSideEffectWireName = "browser_open"
        Case NxShapeInsert: NxSideEffectWireName = "shape_insert"
        Case NxShapeMutation: NxSideEffectWireName = "shape_mutation"
        Case NxBinaryPackageWrite: NxSideEffectWireName = "binary_package_write"
        Case NxWorksheetCreate: NxSideEffectWireName = "worksheet_create"
        Case NxColumnInsert: NxSideEffectWireName = "column_insert"
        Case NxDirectoryBatchCreate: NxSideEffectWireName = "directory_batch_create"
        Case Else: NxRaiseContractError "Unknown side effect kind"
    End Select
End Function

Public Function NxInvariantUnsigned(ByVal value As Double) As String
    Dim quotient As Double
    Dim digit As Long
    If value < 0 Or value <> Fix(value) Then NxRaiseContractError "Invariant value must be an unsigned integer"
    If value = 0 Then NxInvariantUnsigned = "0": Exit Function
    Do While value > 0
        quotient = Fix(value / 10#)
        digit = CLng(value - (quotient * 10#))
        NxInvariantUnsigned = Chr$(48 + digit) & NxInvariantUnsigned
        value = quotient
    Loop
End Function

Public Function NxJoinFailure(ByVal existingValue As String, ByVal nextValue As String) As String
    If Len(nextValue) = 0 Then
        NxJoinFailure = existingValue
    ElseIf Len(existingValue) = 0 Then
        NxJoinFailure = nextValue
    Else
        NxJoinFailure = existingValue & " | " & nextValue
    End If
End Function

Public Function NxCreateRunUuid() As String
    Dim value As NxGuid
    Dim raw As String
    Dim candidate As String
    On Error GoTo Failed
    If CoCreateGuid(value) <> 0 Then GoTo Failed
    raw = String$(39, vbNullChar)
    If StringFromGUID2(value, StrPtr(raw), Len(raw)) <> 39 Then GoTo Failed
    candidate = LCase$(Mid$(raw, 2, 36))
    If Not NxIsCanonicalRunUuid(candidate) Then GoTo Failed
    NxCreateRunUuid = candidate
    Exit Function
Failed:
    Err.Clear
    NxRaiseContractError "A fresh run UUID is required for approval"
End Function

Private Function NxIsCanonicalRunUuid(ByVal value As String) As Boolean
    Dim index As Long
    Dim character As String
    If Len(value) <> 36 Then Exit Function
    For index = 1 To 36
        character = Mid$(value, index, 1)
        Select Case index
            Case 9, 14, 19, 24
                If character <> "-" Then Exit Function
            Case Else
                If InStr(1, "0123456789abcdef", character, vbBinaryCompare) = 0 Then Exit Function
        End Select
    Next index
    NxIsCanonicalRunUuid = True
End Function
