Attribute VB_Name = "NxAiPrompt"
Option Explicit

Private Const NX_AI_DECISION_SAFE As String = "SAFE"
Private Const NX_AI_DECISION_MASK As String = "MASK"
Private Const NX_AI_DECISION_ORIGINAL As String = "ORIGINAL"

Public Function NxAiTaskChoices(ByVal featureId As String) As Variant
    Select Case featureId
        Case "NX-AI-SUMMARY": NxAiTaskChoices = Array("핵심요약", "리스크점검", "비교분석", "보고용 정리")
        Case "NX-AI-CLEAN": NxAiTaskChoices = Array("품질진단", "정제규칙 도출", "중복/누락 점검", "코드값 점검")
        Case "NX-AI-FORMULA": NxAiTaskChoices = Array("수식 작성", "수식 오류 검토", "구버전 호환식 변환", "함수 설명")
        Case "NX-AI-WRITE": NxAiTaskChoices = Array("보고문장 작성", "공문체 정리", "표 설명문 작성", "번역/윤문")
        Case "NX-AI-IMAGE": NxAiTaskChoices = Array("표 이미지 판독", "화면 오류 분석", "이미지 속 텍스트 정리")
        Case Else: RaisePromptError "지원하지 않는 AI 작업 모드입니다."
    End Select
End Function

Public Function NxAiTaskIsValid(ByVal featureId As String, ByVal taskName As String) As Boolean
    Dim choice As Variant
    For Each choice In NxAiTaskChoices(featureId)
        If StrComp(CStr(choice), taskName, vbBinaryCompare) = 0 Then NxAiTaskIsValid = True: Exit Function
    Next choice
End Function

Public Function NxAiTaskPurpose(ByVal featureId As String, ByVal taskName As String) As String
    If Not NxAiTaskIsValid(featureId, taskName) Then RaisePromptError "현재 모드의 세부 작업을 선택하세요."
    Select Case taskName
        Case "핵심요약": NxAiTaskPurpose = "핵심 내용과 주요 수치를 요약하고 확인할 사항을 정리해주세요."
        Case "리스크점검": NxAiTaskPurpose = "이상값, 누락과 위험 신호의 근거 및 후속 확인 항목을 우선순위로 정리해주세요."
        Case "비교분석": NxAiTaskPurpose = "항목별 차이와 공통점을 비교하고 비교 근거를 밝혀주세요."
        Case "보고용 정리": NxAiTaskPurpose = "선택 자료의 결론, 근거 수치와 조치사항을 보고용 문장과 표로 정리해주세요."
        Case "품질진단": NxAiTaskPurpose = "누락, 중복, 형식과 날짜·숫자 오류를 점검하고 심각도와 근거 위치를 표시해주세요."
        Case "정제규칙 도출": NxAiTaskPurpose = "원본을 보존하면서 적용할 데이터 정리 규칙과 예외를 제안해주세요."
        Case "중복/누락 점검": NxAiTaskPurpose = "중복과 누락의 판단 기준을 밝히고 원본 삭제 없이 확인할 위치를 정리해주세요."
        Case "코드값 점검": NxAiTaskPurpose = "코드와 분류값의 표기 차이를 찾고 표준값 후보와 치환 전 확인 기준을 제안해주세요."
        Case "수식 작성": NxAiTaskPurpose = "선택 자료의 셀 주소와 헤더를 근거로 수식과 적용 위치, 검증 예시를 제안해주세요."
        Case "수식 오류 검토": NxAiTaskPurpose = "수식의 오류 가능성과 수정 후보, 검증 방법을 설명해주세요."
        Case "구버전 호환식 변환": NxAiTaskPurpose = "최신 함수와 동적 배열 수식을 Office 2024 이하에서 쓸 수 있는 대체식이나 보조열 방식으로 바꿔주세요."
        Case "함수 설명": NxAiTaskPurpose = "현재 수식의 계산 순서와 참조 관계, 오류가 나는 조건을 설명해주세요."
        Case "보고문장 작성": NxAiTaskPurpose = "자료의 사실관계와 수치를 유지해 업무 보고 초안을 작성해주세요."
        Case "공문체 정리": NxAiTaskPurpose = "의미를 바꾸지 않고 공문과 업무보고에 맞는 문장으로 다듬고 변경 이유를 정리해주세요."
        Case "표 설명문 작성": NxAiTaskPurpose = "선택한 표의 핵심 의미를 설명하는 문장과 주석을 작성해주세요."
        Case "번역/윤문": NxAiTaskPurpose = "원문의 의미, 숫자와 고유명사를 유지하며 요청한 언어와 문체로 번역하거나 윤문해주세요."
        Case "표 이미지 판독": NxAiTaskPurpose = "사용자가 직접 첨부한 이미지의 표 구조와 숫자를 판독하고 불명확한 항목을 구분해주세요."
        Case "화면 오류 분석": NxAiTaskPurpose = "사용자가 직접 첨부한 화면의 오류 메시지와 상태를 읽고 가능한 원인과 조치를 구분해주세요."
        Case "이미지 속 텍스트 정리": NxAiTaskPurpose = "사용자가 직접 첨부한 이미지의 텍스트를 정리하고 판독 불가 항목을 명확히 표시해주세요."
    End Select
End Function

Public Function NxAiBuildPrompt( _
    ByVal featureId As String, _
    ByVal serializedInput As String, _
    ByVal userPurpose As String, _
    ByVal outputFormat As String, _
    ByVal additionalInstruction As String, _
    ByVal privacyState As NxPrivacyState, _
    ByVal privacyDecision As String, _
    ByVal inputWasMasked As Boolean, _
    ByVal originalUseConfirmed As Boolean) As String

    If Len(Trim$(serializedInput)) = 0 Then RaisePromptError "AI 입력이 비어 있습니다."
    If Len(Trim$(userPurpose)) = 0 Then RaisePromptError "작업 목적을 입력해야 합니다."
    If Len(Trim$(outputFormat)) = 0 Then RaisePromptError "원하는 출력 형식을 입력해야 합니다."
    NxAiValidatePrivacyDecision privacyState, privacyDecision, inputWasMasked, originalUseConfirmed

    Dim modePurpose As String, modeInput As String
    Dim modeConstraints As String, modeOutput As String
    ResolveMode featureId, modePurpose, modeInput, modeConstraints, modeOutput

    Dim prompt As String
    prompt = "[목적]" & vbLf & modePurpose & vbLf & "사용자 목적: " & userPurpose & vbLf & vbLf
    prompt = prompt & "[입력]" & vbLf & modeInput & vbLf
    If featureId = "NX-AI-IMAGE" Then prompt = prompt & "이미지 내용 검사 미완료" & vbLf
    prompt = prompt & serializedInput & vbLf & vbLf
    prompt = prompt & "[제약]" & vbLf & modeConstraints
    If Len(Trim$(additionalInstruction)) > 0 Then prompt = prompt & vbLf & "추가 지시: " & additionalInstruction
    prompt = prompt & vbLf & vbLf & "[원하는 출력]" & vbLf & modeOutput & vbLf & "출력 형식: " & outputFormat
    NxAiBuildPrompt = prompt
End Function

Public Sub NxAiValidatePrivacyDecision( _
    ByVal privacyState As NxPrivacyState, _
    ByVal privacyDecision As String, _
    ByVal inputWasMasked As Boolean, _
    ByVal originalUseConfirmed As Boolean)

    If privacyState < NxPrivacySafe Or privacyState > NxPrivacyIncomplete Then
        RaisePromptError "알 수 없는 개인정보 검사 상태입니다."
    End If
    If privacyState = NxPrivacyIncomplete Then
        RaisePromptError "개인정보 검사 미완료 상태에서는 프롬프트를 만들 수 없습니다."
    End If

    Dim decision As String
    decision = UCase$(Trim$(privacyDecision))
    If privacyState = NxPrivacySafe Then
        If decision <> NX_AI_DECISION_SAFE Then RaisePromptError "안전 상태의 개인정보 결정이 올바르지 않습니다."
        Exit Sub
    End If
    If decision = NX_AI_DECISION_MASK Then
        If Not inputWasMasked Then RaisePromptError "마스킹 결정에는 마스킹된 입력 복사본이 필요합니다."
        Exit Sub
    End If
    If decision = NX_AI_DECISION_ORIGINAL Then
        If privacyState = NxPrivacyBlocked Then RaisePromptError "차단된 개인정보 후보는 원문을 사용할 수 없습니다."
        If inputWasMasked Then RaisePromptError "원문 사용 결정과 마스킹 입력이 일치하지 않습니다."
        If Not originalUseConfirmed Then RaisePromptError "원문 사용은 두 번째 확인이 필요합니다."
        Exit Sub
    End If
    RaisePromptError "개인정보 후보는 마스킹, 확인된 원문 사용 또는 취소로 결정해야 합니다."
End Sub

Private Sub ResolveMode( _
    ByVal featureId As String, _
    ByRef modePurpose As String, _
    ByRef modeInput As String, _
    ByRef modeConstraints As String, _
    ByRef modeOutput As String)

    Select Case featureId
        Case "NX-AI-SUMMARY"
            modePurpose = "선택한 자료에서 핵심 내용과 확인할 지점을 정리합니다."
            modeInput = "사용자가 승인한 Excel 선택 범위의 구조와 값을 제공합니다."
            modeConstraints = "자료에 없는 사실을 만들지 말고 불명확한 항목은 확인 질문으로 구분합니다."
            modeOutput = "핵심 요약, 이상점, 확인 질문 순서로 작성합니다."
        Case "NX-AI-CLEAN"
            modePurpose = "선택한 자료를 일관된 기준으로 정리하는 방법을 제안합니다."
            modeInput = "사용자가 승인한 Excel 선택 범위의 구조와 값을 제공합니다."
            modeConstraints = "원본을 변경한다고 가정하지 말고 적용할 정규화 규칙을 먼저 밝힙니다."
            modeOutput = "정리 규칙, 변환 예시, 확인이 필요한 예외 순서로 작성합니다."
        Case "NX-AI-FORMULA"
            modePurpose = "선택한 자료에 맞는 Excel 수식과 검증 방법을 제안합니다."
            modeInput = "표시값과 FormulaR1C1을 구분한 Excel 선택 정보를 제공합니다."
            modeConstraints = "셀 주소와 헤더를 근거로 삼고 지원 버전을 벗어난 함수는 대안을 함께 제시합니다."
            modeOutput = "제안 수식, 적용 위치, 검증 예시, 주의점 순서로 작성합니다."
        Case "NX-AI-WRITE"
            modePurpose = "선택한 표 내용을 업무 문안으로 바꾸는 초안을 만듭니다."
            modeInput = "사용자가 승인한 Excel 선택 범위의 구조와 값을 제공합니다."
            modeConstraints = "원문의 사실관계와 수치를 유지하고 확인되지 않은 내용을 추가하지 않습니다."
            modeOutput = "요청한 문서 형식에 맞춘 초안과 확인할 항목을 구분해 작성합니다."
        Case "NX-AI-IMAGE"
            modePurpose = "선택한 그림 또는 범위 이미지의 활용 방향을 정리합니다."
            modeInput = "Excel 선택 범위의 구조와 값을 제공합니다. 이미지는 사용자가 직접 첨부한 경우에만 분석합니다."
            modeConstraints = "이미지 내용 검사 미완료 상태임을 전제로 민감정보를 추정하거나 재현하지 않습니다."
            modeOutput = "관찰 내용, 활용 제안, 사용자가 확인할 항목 순서로 작성합니다."
        Case Else
            RaisePromptError "지원하지 않는 AI 작업 모드입니다."
    End Select
End Sub

Private Sub RaisePromptError(ByVal message As String)
    Err.Raise vbObjectError + 884, "NxAiPrompt", message
End Sub
