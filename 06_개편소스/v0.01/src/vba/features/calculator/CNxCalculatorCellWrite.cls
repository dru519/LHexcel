VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxCalculatorCellWrite"
Option Explicit

Private mIdentity As String
Private mFingerprint As String
Private mTargetLabel As String
Private mExistingKind As String
Private mRequiresConfirmation As Boolean
Private mResultValue As Double
Private mSealed As Boolean

Public Sub Capture(ByVal target As Range, ByVal resultValue As Double)
    If mSealed Then NxRaiseContractError "셀 쓰기 계획은 한 번만 만들 수 있습니다."
    NxCalculatorRequireWritableTarget target
    If target.Worksheet.Parent Is Nothing Then NxRaiseContractError "대상 통합문서를 확인할 수 없습니다."
    If TypeName(Application.ActiveCell) <> "Range" Then NxRaiseContractError "활성 셀이 변경되었습니다."
    mIdentity = NxCalculatorCellIdentity(target)
    If Len(mIdentity) = 0 Then NxRaiseContractError "활성 셀 식별자를 확인할 수 없습니다."
    If NxCalculatorActiveCellIdentity() <> mIdentity Then NxRaiseContractError "활성 셀이 변경되었습니다."
    mFingerprint = NxCalculatorCellFingerprint(target)
    mTargetLabel = target.Worksheet.Parent.Name & " / " & target.Worksheet.Name & " / " & target.Address(False, False, xlA1)
    If IsEmpty(target.Value2) And Not target.HasFormula Then
        mExistingKind = "빈 셀"
        mRequiresConfirmation = False
    ElseIf target.HasFormula Then
        mExistingKind = "수식"
        mRequiresConfirmation = True
    Else
        mExistingKind = "기존 값"
        mRequiresConfirmation = True
    End If
    mResultValue = resultValue
    mSealed = True
End Sub

Public Property Get TargetLabel() As String
    RequireSealed
    TargetLabel = mTargetLabel
End Property

Public Property Get ExistingKind() As String
    RequireSealed
    ExistingKind = mExistingKind
End Property

Public Property Get RequiresConfirmation() As Boolean
    RequireSealed
    RequiresConfirmation = mRequiresConfirmation
End Property

Public Sub Commit()
    Dim target As Range
    RequireSealed
    If NxCalculatorActiveCellIdentity() <> mIdentity Then NxRaiseContractError "대상 셀이 바뀌었습니다. 새 셀에서 다시 실행하세요."
    If TypeName(Application.ActiveCell) <> "Range" Then NxRaiseContractError "결과를 입력할 셀을 확인할 수 없습니다."
    Set target = Application.ActiveCell
    NxCalculatorRequireWritableTarget target
    If NxCalculatorCellFingerprint(target) <> mFingerprint Then NxRaiseContractError "대상 셀 내용이 바뀌었습니다. 다시 확인하세요."
    target.Value2 = mResultValue
End Sub

Private Sub RequireSealed()
    If Not mSealed Then NxRaiseContractError "셀 쓰기 계획이 준비되지 않았습니다."
End Sub
