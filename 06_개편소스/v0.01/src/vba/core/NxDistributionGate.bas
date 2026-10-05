Attribute VB_Name = "NxDistributionGate"
Option Explicit

' Filled only in distribution copies. Development workbooks remain unrestricted.
Private Const NX_DISTRIBUTION_ENFORCE As Boolean = False
Private Const NX_POLICY_PAYLOAD As String = ""
Private Const NX_POLICY_PUBLIC As String = ""
Private Const NX_POLICY_SIGNATURE As String = ""
Private mVerified As Boolean
Private mChecked As Boolean
Private mBlockFrom As Date
Private mProfile As String
Private mLastState As String
Private mNextRefresh As Date
Private mRefreshScheduled As Boolean

#If VBA7 Then
Private Declare PtrSafe Function CryptAcquireContextW Lib "advapi32.dll" (ByRef provider As LongPtr, ByVal container As LongPtr, ByVal name As LongPtr, ByVal providerType As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptImportKey Lib "advapi32.dll" (ByVal provider As LongPtr, ByRef data As Byte, ByVal length As Long, ByVal parent As LongPtr, ByVal flags As Long, ByRef key As LongPtr) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" (ByVal provider As LongPtr, ByVal algorithm As Long, ByVal key As LongPtr, ByVal flags As Long, ByRef hash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" (ByVal hash As LongPtr, ByRef data As Byte, ByVal length As Long, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptVerifySignatureW Lib "advapi32.dll" (ByVal hash As LongPtr, ByRef signature As Byte, ByVal length As Long, ByVal key As LongPtr, ByVal description As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" (ByVal hash As LongPtr) As Long
Private Declare PtrSafe Function CryptDestroyKey Lib "advapi32.dll" (ByVal key As LongPtr) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal provider As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function CryptAcquireContextW Lib "advapi32.dll" (ByRef provider As Long, ByVal container As Long, ByVal name As Long, ByVal providerType As Long, ByVal flags As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" (ByVal provider As Long, ByRef data As Byte, ByVal length As Long, ByVal parent As Long, ByVal flags As Long, ByRef key As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal provider As Long, ByVal algorithm As Long, ByVal key As Long, ByVal flags As Long, ByRef hash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hash As Long, ByRef data As Byte, ByVal length As Long, ByVal flags As Long) As Long
Private Declare Function CryptVerifySignatureW Lib "advapi32.dll" (ByVal hash As Long, ByRef signature As Byte, ByVal length As Long, ByVal key As Long, ByVal description As Long, ByVal flags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hash As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal key As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal provider As Long, ByVal flags As Long) As Long
#End If

Public Function NxDistributionState() As String
    Dim fields As Variant, iso As String
    On Error GoTo Invalid
    If Not NX_DISTRIBUTION_ENFORCE Then
        NxDistributionState = "development"
        Exit Function
    End If
    If Not mChecked Then
        mChecked = True
        mProfile = vbNullString
        fields = Split(NX_POLICY_PAYLOAD, "|")
        If UBound(fields) <> 3 Then GoTo Invalid
        If fields(0) <> "NXL1" Then GoTo Invalid
        If fields(2) <> "internal-xlam" And fields(2) <> "enhanced-dll" Then GoTo Invalid
        iso = CStr(fields(3))
        If Len(iso) <> 10 Then GoTo Invalid
        mBlockFrom = DateSerial(CLng(Left$(iso, 4)), CLng(Mid$(iso, 6, 2)), CLng(Right$(iso, 2)))
        If Format$(mBlockFrom, "yyyy-mm-dd") <> iso Then GoTo Invalid
        If CStr(ThisWorkbook.CustomDocumentProperties("DistributionBlockFrom").Value) <> iso Then GoTo Invalid
        If CStr(ThisWorkbook.CustomDocumentProperties("DistributionValidUntil").Value) <> Format$(mBlockFrom - 1, "yyyy-mm-dd") Then GoTo Invalid
        If CStr(ThisWorkbook.CustomDocumentProperties("DistributionProfile").Value) <> CStr(fields(2)) Then GoTo Invalid
        mVerified = NxDistributionVerify()
        If mVerified Then mProfile = CStr(fields(2))
    End If
    If Not mVerified Then GoTo Invalid
    If Date >= mBlockFrom Then
        NxDistributionState = "expired"
    Else
        NxDistributionState = "active"
    End If
    Exit Function
Invalid:
    NxDistributionState = "invalid"
End Function

Public Function NxDistributionIsStandaloneXlam() As Boolean
    Dim state As String
    state = NxDistributionState()
    If state = "development" Then
        NxDistributionIsStandaloneXlam = True
    ElseIf state = "active" Or state = "expired" Then
        NxDistributionIsStandaloneXlam = (mVerified And mProfile = "internal-xlam")
    End If
End Function

Public Sub NxDistributionEnsureStandaloneXlam()
    If Not NxDistributionIsStandaloneXlam() Then _
        NxRaiseContractError "설치·제거는 단독 XLAM에서만 사용할 수 있습니다. DLL/EXE 배포판은 배포 패키지의 설치·제거 도구를 사용해 주세요."
End Sub

Public Function NxDistributionCanExecute() As Boolean
    Dim state As String
    state = NxDistributionState()
    NxDistributionCanExecute = (state = "active" Or state = "development")
End Function

Public Function NxDistributionValidUntilText() As String
    Select Case NxDistributionState()
        Case "active", "expired"
            NxDistributionValidUntilText = Format$(mBlockFrom - 1, "yyyy-mm-dd")
            If Date >= mBlockFrom Then NxDistributionValidUntilText = NxDistributionValidUntilText & " (만료)"
        Case "development"
            NxDistributionValidUntilText = "제한 없음 (개발용)"
        Case Else
            NxDistributionValidUntilText = "확인 불가 (보호 정보 확인 필요)"
    End Select
End Function

Public Function NxDistributionMessage() As String
    If NxDistributionState() = "expired" Then
        NxDistributionMessage = "사용기한이 만료되어 내엑셀 기능을 사용할 수 없습니다. 새 배포본은 배포자에게 문의해 주세요."
    Else
        NxDistributionMessage = "사용기한 또는 보호 정보가 변경되었거나 확인되지 않아 기능을 사용할 수 없습니다. 원본 배포본을 다시 설치해 주세요."
    End If
    NxDistributionMessage = NxDistributionMessage & vbCrLf & _
        "사용기한과 보호장치의 임의 수정은 금지됩니다. AI나 자동화 도구를 이용한 변경에도 동일하게 적용됩니다."
End Function

Public Sub NxDistributionEnsureExecutable()
    If Not NxDistributionCanExecute() Then NxRaiseContractError NxDistributionMessage()
End Sub

Public Function NxDistributionEntryAllowed(ByVal kind As String, ByVal identifier As String) As Boolean
    If kind = "entry" And (identifier = "NX-MGMT-INSTALL" Or identifier = "NX-MGMT-REMOVE") Then
        NxDistributionEntryAllowed = NxDistributionIsStandaloneXlam()
        If identifier = "NX-MGMT-INSTALL" Then NxDistributionEntryAllowed = NxDistributionEntryAllowed And NxDistributionCanExecute()
        Exit Function
    End If
    If NxDistributionCanExecute() Then
        NxDistributionEntryAllowed = True
    ElseIf kind = "entry" Then
        Select Case identifier
            Case "NX-ENTRY-MANAGEMENT", "NX-MGMT-DISTRIBUTION-STATUS", "NX-MGMT-FILE-INFO"
                NxDistributionEntryAllowed = True
        End Select
    End If
End Function

Public Sub NxDistributionStart()
    If Not NX_DISTRIBUTION_ENFORCE Then Exit Sub
    NxDistributionStop
    mLastState = vbNullString
    NxDistributionRefresh
End Sub

Public Sub NxDistributionRefresh()
    Dim state As String, form As Object, item As Object
    On Error GoTo Finish
    NxDistributionStop
    mChecked = False: mVerified = False
    state = NxDistributionState()
    If state <> mLastState Then
        mLastState = state
        If Not NxDistributionCanExecute() Then
            NxShortcutsShutdown
            NxFocusControllerDisable
            Call NxHostHideNavigator
            NxHostHideDocumentNavigator
            On Error Resume Next
            For Each form In VBA.UserForms
                For Each item In form.Controls
                    If item.Name <> "cmdClose" And item.Name <> "cmdCancel" Then item.Enabled = False
                Next item
            Next form
            On Error GoTo Finish
        End If
        NxRibbonInvalidateAvailability
    End If
Finish:
    ' One owned timer updates an idle ribbon after midnight; each action also
    ' checks Date independently, so a delayed Excel OnTime cannot permit a run.
    If NX_DISTRIBUTION_ENFORCE Then
        mNextRefresh = DateAdd("s", 30, Now)
        On Error Resume Next
        Application.OnTime mNextRefresh, "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!NxDistributionRefresh"
        mRefreshScheduled = (Err.Number = 0)
        On Error GoTo 0
    End If
End Sub

Public Sub NxDistributionStop()
    If Not mRefreshScheduled Then Exit Sub
    On Error Resume Next
    Application.OnTime mNextRefresh, "'" & Replace$(ThisWorkbook.Name, "'", "''") & "'!NxDistributionRefresh", Schedule:=False
    mRefreshScheduled = False
    On Error GoTo 0
End Sub

Private Function NxDistributionVerify() As Boolean
#If VBA7 Then
    Dim provider As LongPtr, key As LongPtr, hash As LongPtr
#Else
    Dim provider As Long, key As Long, hash As Long
#End If
    Dim blob() As Byte, signature() As Byte, payload() As Byte
    On Error GoTo Clean
    blob = NxDistributionHex(NX_POLICY_PUBLIC)
    signature = NxDistributionHex(NX_POLICY_SIGNATURE)
    payload = StrConv(NX_POLICY_PAYLOAD, vbFromUnicode)
    If CryptAcquireContextW(provider, 0, 0, 24, &HF0000000) = 0 Then GoTo Clean
    If CryptImportKey(provider, blob(0), UBound(blob) + 1, 0, 0, key) = 0 Then GoTo Clean
    If CryptCreateHash(provider, &H800C&, 0, 0, hash) = 0 Then GoTo Clean
    If CryptHashData(hash, payload(0), UBound(payload) + 1, 0) = 0 Then GoTo Clean
    NxDistributionVerify = (CryptVerifySignatureW(hash, signature(0), UBound(signature) + 1, key, 0, 0) <> 0)
Clean:
    On Error Resume Next
    If hash <> 0 Then CryptDestroyHash hash
    If key <> 0 Then CryptDestroyKey key
    If provider <> 0 Then CryptReleaseContext provider, 0
End Function

Private Function NxDistributionHex(ByVal value As String) As Byte()
    Dim result() As Byte, i As Long
    If Len(value) = 0 Or Len(value) Mod 2 <> 0 Then Err.Raise 5
    ReDim result(0 To Len(value) \ 2 - 1)
    For i = 0 To UBound(result)
        result(i) = CByte("&H" & Mid$(value, i * 2 + 1, 2))
    Next i
    NxDistributionHex = result
End Function
