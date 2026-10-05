VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxRenameRule"
Option Explicit

Public Kind As String
Public Text As String
Public Replacement As String
Public Position As String
Public Start As Long
Public StepValue As Long
Public Digits As Long
Public CaseSensitive As Boolean
Public FirstOnly As Boolean
Public AllowEmptyExtension As Boolean

Private Sub Class_Initialize()
    Position = "BACK"
    Start = 1
    StepValue = 1
End Sub

Public Sub Validate(ByVal fileMode As Boolean)
    Select Case Kind
        Case "REPLACE", "PREFIX", "SUFFIX", "FULL"
            If Len(Text) = 0 Then NxRaiseContractError "규칙에 사용할 문구를 입력하세요."
        Case "NUMBER"
            If StepValue = 0 Then NxRaiseContractError "번호 증가값은 0일 수 없습니다."
            If Digits < 0 Or Digits > 6 Then NxRaiseContractError "번호 자릿수는 0부터 6까지 입력하세요."
            If Position <> "FRONT" And Position <> "BACK" Then NxRaiseContractError "번호 위치를 선택하세요."
        Case "DELETE"
            If Start < 1 Or Start > 255 Then NxRaiseContractError "지울 글자 수는 1부터 255까지 입력하세요."
            If Position <> "FRONT" And Position <> "BACK" Then NxRaiseContractError "문자 삭제 방향을 선택하세요."
        Case "CASE"
            If Position <> "LOWER" And Position <> "UPPER" And Position <> "PROPER" Then NxRaiseContractError "대소문자 변환 방식을 선택하세요."
        Case "EXTENSION"
            If Not fileMode Then NxRaiseContractError "확장자 규칙은 파일 이름에만 사용할 수 있습니다."
            If Len(Trim$(Text)) = 0 And Not AllowEmptyExtension Then NxRaiseContractError "확장자를 삭제하려면 삭제 확인을 선택하세요."
        Case Else
            NxRaiseContractError "지원하지 않는 이름 변경 규칙입니다."
    End Select
End Sub

Public Property Get Summary() As String
    Select Case Kind
        Case "REPLACE": Summary = "바꾸기: " & Text & " → " & Replacement
        Case "PREFIX": Summary = "앞말: " & Text
        Case "SUFFIX": Summary = "뒷말: " & Text
        Case "NUMBER": Summary = "번호: " & Position & " / " & CStr(Start) & " / 증가 " & CStr(StepValue) & " / " & CStr(Digits) & "자리"
        Case "DELETE": Summary = "지우기: " & Position & " / " & CStr(Start) & "자"
        Case "FULL": Summary = "전체 이름: " & Text
        Case "CASE": Summary = "대소문자: " & Position
        Case "EXTENSION": Summary = "확장자: " & Text
    End Select
End Property
