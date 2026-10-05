Attribute VB_Name = "NxAiPrivacy"
Option Explicit

Private Const MASK_RRN As String = "[주민번호 마스킹]"
Private Const MASK_CARD As String = "[카드번호 마스킹]"
Private Const MASK_ACCOUNT As String = "[계좌 마스킹]"
Private Const MASK_SECRET As String = "[비밀값 마스킹]"
Private Const MASK_EMAIL As String = "[이메일 마스킹]"
Private Const MASK_PHONE As String = "[전화번호 마스킹]"
Private Const MASK_ADDRESS As String = "[주소 마스킹]"

Public Function NxAiScanPrivacy(ByVal serializedInput As String) As CNxAiPrivacySnapshot
    If Len(Trim$(serializedInput)) = 0 Then NxRaiseContractError "Privacy scan requires serialized input"
    On Error GoTo ScanFailed

    Dim masked As String
    Dim blockedCount As Long
    Dim cautionCount As Long
    masked = serializedInput
    masked = MaskRule(masked, "\b[0-9]{6}-?[1-8][0-9]{6}\b", MASK_RRN, blockedCount)
    masked = MaskLuhnCandidates(masked, blockedCount)
    masked = MaskRule(masked, "(계좌|account|bank)[^0-9\r\n]{0,12}[0-9][0-9 -]{5,20}[0-9]", MASK_ACCOUNT, blockedCount)
    masked = MaskRule(masked, "(비밀번호|password|passwd|api[ _-]?key|secret)[ \t]*[:=][ \t]*[^ \t\r\n|]{4,80}", MASK_SECRET, blockedCount)
    masked = MaskRule(masked, "\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b", MASK_EMAIL, cautionCount)
    masked = MaskRule(masked, "\b(01[016789][ -]?[0-9]{3,4}[ -]?[0-9]{4}|0[2-6][0-9]?[ -]?[0-9]{3,4}[ -]?[0-9]{4})\b", MASK_PHONE, cautionCount)
    masked = MaskRule(masked, "(서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주)[^|,\r\n]{0,40}(로|길|동|읍|면)[ ]*[0-9-]*", MASK_ADDRESS, cautionCount)

    Dim state As NxPrivacyState
    Dim summary As String
    If blockedCount > 0 Then
        state = NxPrivacyBlocked
        summary = "개인정보 차단 후보 " & CStr(blockedCount) & "건"
        If cautionCount > 0 Then summary = summary & ", 확인 후보 " & CStr(cautionCount) & "건"
    ElseIf cautionCount > 0 Then
        state = NxPrivacyCaution
        summary = "개인정보 확인 후보 " & CStr(cautionCount) & "건"
    Else
        state = NxPrivacySafe
        summary = "개인정보 위험 없음"
    End If
    Set NxAiScanPrivacy = CreateSnapshot(state, serializedInput, masked, summary)
    Exit Function

ScanFailed:
    Err.Clear
    Set NxAiScanPrivacy = CreateSnapshot(NxPrivacyIncomplete, "[개인정보 검사 미완료]", _
        "[개인정보 검사 미완료]", "개인정보 검사 미완료")
End Function

Private Function MaskRule(ByVal source As String, ByVal pattern As String, ByVal token As String, _
    ByRef matchCount As Long) As String

    Dim expression As Object
    Set expression = CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.IgnoreCase = True
    expression.MultiLine = True
    expression.Pattern = pattern
    Dim matches As Object
    Set matches = expression.Execute(source)
    matchCount = matchCount + matches.Count
    If matches.Count = 0 Then MaskRule = source Else MaskRule = expression.Replace(source, token)
End Function

Private Function MaskLuhnCandidates(ByVal source As String, ByRef matchCount As Long) As String
    Dim expression As Object
    Set expression = CreateObject("VBScript.RegExp")
    expression.Global = True
    expression.IgnoreCase = False
    expression.MultiLine = True
    expression.Pattern = "\b([0-9][ -]?){12,18}[0-9]\b"
    Dim matches As Object
    Dim item As Object
    Dim output As String
    output = source
    Set matches = expression.Execute(source)
    For Each item In matches
        If IsLuhnCandidate(CStr(item.Value)) Then
            output = Replace(output, CStr(item.Value), MASK_CARD, 1, -1, vbBinaryCompare)
            matchCount = matchCount + 1
        End If
    Next item
    MaskLuhnCandidates = output
End Function

Private Function IsLuhnCandidate(ByVal candidate As String) As Boolean
    Dim digits As String
    Dim index As Long
    Dim character As String
    For index = 1 To Len(candidate)
        character = Mid$(candidate, index, 1)
        If character >= "0" And character <= "9" Then digits = digits & character
    Next index
    If Len(digits) < 13 Or Len(digits) > 19 Then Exit Function
    If digits = String$(Len(digits), Left$(digits, 1)) Then Exit Function

    Dim total As Long
    Dim value As Long
    Dim doubleNext As Boolean
    For index = Len(digits) To 1 Step -1
        value = Asc(Mid$(digits, index, 1)) - 48
        If doubleNext Then
            value = value * 2
            If value > 9 Then value = value - 9
        End If
        total = total + value
        doubleNext = Not doubleNext
    Next index
    IsLuhnCandidate = (total Mod 10 = 0)
End Function

Private Function CreateSnapshot(ByVal state As NxPrivacyState, ByVal originalCopy As String, _
    ByVal maskedCopy As String, ByVal summary As String) As CNxAiPrivacySnapshot

    Dim snapshot As New CNxAiPrivacySnapshot
    snapshot.Configure state, originalCopy, maskedCopy, summary
    snapshot.Seal
    Set CreateSnapshot = snapshot
End Function
