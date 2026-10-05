Attribute VB_Name = "NxOfficeCapabilityProbe"
Option Explicit

Public Const NX_FALLBACK_FORMULA_LEGACY As String = "formula_legacy"

Public Function NxProbeOfficeCompatibility(ByVal definition As CNxFeatureDefinition) As CNxOfficeCompatibility
    If definition Is Nothing Then NxRaiseContractError "Office capability probe requires a feature definition"
    If Not definition.IsSealed Then NxRaiseContractError "Office capability probe requires a sealed feature definition"
    Set NxProbeOfficeCompatibility = NxEvaluateOfficeCompatibility( _
        definition.FeatureId, _
        definition.RequiredOfficeCapabilities, _
        definition.FallbackStrategy, _
        NxDetectedOfficeCapabilityMask(), _
        True, _
        CStr(Application.Version), _
        NxOfficeBitness())
End Function

Public Function NxEvaluateOfficeCompatibility( _
    ByVal featureId As String, _
    ByVal requiredCapabilities As Long, _
    ByVal fallbackStrategy As String, _
    ByVal detectedCapabilities As Long, _
    ByVal probeComplete As Boolean, _
    ByVal observedOfficeVersion As String, _
    ByVal observedBitness As Long) As CNxOfficeCompatibility

    Dim result As New CNxOfficeCompatibility
    Dim missing As Long
    Dim state As NxCompatibilityState
    Dim reason As String

    If requiredCapabilities < 0 Or detectedCapabilities < 0 Then NxRaiseContractError "Office capability masks cannot be negative"
    missing = requiredCapabilities And Not detectedCapabilities
    If Not probeComplete Then
        state = NxCompatibilityIncomplete
        reason = "Office capability probe is incomplete"
    ElseIf missing = 0 Then
        state = NxCompatibilityNative
    ElseIf StrComp(Trim$(fallbackStrategy), "disable_feature", vbTextCompare) = 0 Then
        state = NxCompatibilityUnavailable
        reason = "Required Office capability is unavailable"
    ElseIf StrComp(Trim$(fallbackStrategy), NX_FALLBACK_FORMULA_LEGACY, vbBinaryCompare) = 0 Then
        state = NxCompatibilityFallback
    Else
        state = NxCompatibilityUnavailable
        reason = "Office fallback strategy is not implemented"
    End If

    result.Configure featureId, state, missing, fallbackStrategy, reason, observedOfficeVersion, observedBitness
    result.Seal
    Set NxEvaluateOfficeCompatibility = result
End Function

Public Function NxDetectedOfficeCapabilityMask() As Long
    Dim detected As Long
#If VBA7 Then
    detected = detected Or NxOfficeCapabilityVba7
#End If
#If Win64 Then
    detected = detected Or NxOfficeCapabilityWindowsApi
#Else
    detected = detected Or NxOfficeCapabilityWindowsApi
#End If
    If NxCanCreateFileDialog() Then detected = detected Or NxOfficeCapabilityFileDialog
    If NxCanExportSheet() Then detected = detected Or NxOfficeCapabilitySheetExport
    If NxCanEvaluateDynamicArray() Then detected = detected Or NxOfficeCapabilityDynamicArray
    NxDetectedOfficeCapabilityMask = detected
End Function

Private Function NxCanExportSheet() As Boolean
    ' ExportAsFixedFormat mutates the filesystem; no side-effect-free COM probe exists.
    NxCanExportSheet = False
End Function

Private Function NxCanEvaluateDynamicArray() As Boolean
    Dim worksheetFunctions As Object
    Dim value As Variant
    On Error GoTo Failed
    Set worksheetFunctions = Application.WorksheetFunction
    value = CallByName(worksheetFunctions, "Sequence", VbMethod, 1, 1)
    NxCanEvaluateDynamicArray = Not IsError(value)
    Exit Function
Failed:
    Err.Clear
    NxCanEvaluateDynamicArray = False
End Function

Public Function NxOfficeBitness() As Long
#If Win64 Then
    NxOfficeBitness = 64
#Else
    NxOfficeBitness = 32
#End If
End Function

Private Function NxCanCreateFileDialog() As Boolean
    Dim candidate As Object
    On Error GoTo Failed
    Set candidate = Application.FileDialog(3)
    NxCanCreateFileDialog = Not candidate Is Nothing
    Exit Function
Failed:
    Err.Clear
    NxCanCreateFileDialog = False
End Function
