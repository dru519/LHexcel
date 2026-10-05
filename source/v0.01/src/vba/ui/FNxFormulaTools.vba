Option Explicit
Private mTarget As Range, mReferences As Boolean

Public Sub BindTarget(ByVal target As Range, ByVal references As Boolean)
    Dim formulas As Range
    Set mTarget = target: mReferences = references
    Set formulas = NxFormulaCells(target)
    If formulas Is Nothing Then NxRaiseContractError "선택 범위에 수식이 없습니다."
    Me.Caption = "내엑셀 - " & IIf(references, "수식 참조표 만들기", "셀수식 메모 기록")
    lblRange.Caption = target.Address(External:=True) & vbCrLf & "수식 " & Format$(formulas.CountLarge, "#,##0") & "개"
    cboOption.Clear
    If references Then
        lblOption.Caption = "결과 위치"
        cboOption.AddItem "새 통합문서": cboOption.AddItem "현재 파일의 새 시트"
        chkShow.Visible = False
        lblGuide.Caption = "넓은 참조는 범위 단위로 정리합니다. 별도 확인 항목은 확인 상태에 표시합니다." & vbCrLf & "링크가 열리지 않으면 원본 파일을 열고 파일·시트 이름을 확인하세요."
    Else
        lblOption.Caption = "기존 메모"
        cboOption.AddItem "유지하고 수식 추가": cboOption.AddItem "수식으로 덮어쓰기"
        chkShow.Caption = "기록 후 메모 표시": chkShow.Value = False
        lblGuide.Caption = "수식이 있는 셀에만 기록합니다." & vbCrLf & "마지막 메모 기록은 되돌리기로 취소할 수 있습니다."
    End If
    cboOption.ListIndex = 0
    cmdExecute.Default = True: cmdCancel.Cancel = True
End Sub

Private Sub cmdExecute_Click()
    Dim report As Worksheet
    On Error GoTo Failed
    If mReferences Then
        Set report = NxFormulaReferenceReport(mTarget, cboOption.ListIndex = 0)
    Else
        NxFormulaNotesApply mTarget, CBool(chkShow.Value), cboOption.ListIndex = 1
    End If
    Unload Me
    If mReferences Then NxFormulaReferenceReveal
    Exit Sub
Failed:
    lblGuide.Caption = NxUserErrorText(Err.Description): Err.Clear
End Sub

Private Sub cmdCancel_Click(): Unload Me: End Sub
