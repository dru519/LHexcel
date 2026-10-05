VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxSymbolCellWrite"
Option Explicit

Private mCellCount As Long
Private mFingerprints() As String
Private mIdentity As String
Private mNewValues() As Variant
Private mOriginalValues() As Variant
Private mSealed As Boolean
Private mTargetLabel As String

Public Sub Capture(ByVal bufferText As String, ByVal expectedIdentity As String)
    Dim index As Long
    Dim raw As Variant
    Dim target As Range
    Dim targetCell As Range

    If mSealed Then NxRaiseContractError "기호 셀 쓰기 계획은 한 번만 만들 수 있습니다."
    If Len(bufferText) = 0 Then NxRaiseContractError "셀에 입력할 기호가 없습니다."
    If Len(expectedIdentity) = 0 Or NxSymbolActiveSelectionIdentity() <> expectedIdentity Then _
        NxRaiseContractError "대상 셀 범위가 바뀌었습니다. 새 범위에서 다시 실행하세요."
    Set target = Application.Selection
    NxSymbolRequireWritableTarget target

    mIdentity = NxSymbolSelectionIdentity(target)
    mTargetLabel = NxSymbolActiveTargetLabel()
    mCellCount = CLng(target.Cells.CountLarge)
    ReDim mFingerprints(1 To mCellCount)
    ReDim mNewValues(1 To mCellCount)
    ReDim mOriginalValues(1 To mCellCount)

    index = 0
    For Each targetCell In target.Cells
        index = index + 1
        raw = targetCell.Value2
        If targetCell.HasFormula Then NxRaiseContractError "수식 셀에는 기호를 입력할 수 없습니다."
        If IsError(raw) Then NxRaiseContractError "오류 셀에는 기호를 입력할 수 없습니다."
        If VarType(raw) = vbBoolean Then NxRaiseContractError "논리값 셀에는 기호를 입력할 수 없습니다."
        If VarType(targetCell.Value) = vbDate Or NxSymbolIsDateFormatted(CStr(targetCell.NumberFormat)) Then _
            NxRaiseContractError "날짜·시간 셀에는 기호를 입력할 수 없습니다."
        If Not IsEmpty(raw) And VarType(raw) <> vbString Then _
            NxRaiseContractError "숫자 셀에는 기호를 입력할 수 없습니다."

        mFingerprints(index) = NxSymbolCellFingerprint(targetCell)
        mOriginalValues(index) = raw
        If IsEmpty(raw) Then
            mNewValues(index) = bufferText
        Else
            mNewValues(index) = CStr(raw) & bufferText
        End If
    Next targetCell
    mSealed = True
End Sub

Public Property Get TargetLabel() As String
    RequireSealed
    TargetLabel = mTargetLabel
End Property

Public Property Get NewValue() As String
    RequireSealed
    If mCellCount > 0 Then NewValue = CStr(mNewValues(1))
End Property

Public Property Get TargetCount() As Long
    RequireSealed
    TargetCount = mCellCount
End Property

Public Sub Commit()
    Dim committed As Long
    Dim failureDescription As String
    Dim failureNumber As Long
    Dim failureSource As String
    Dim index As Long
    Dim target As Range
    Dim targetCell As Range

    RequireSealed
    If NxSymbolActiveSelectionIdentity() <> mIdentity Then _
        NxRaiseContractError "대상 셀 범위가 바뀌었습니다. 새 범위에서 다시 실행하세요."
    Set target = Application.Selection
    NxSymbolRequireWritableTarget target
    If CLng(target.Cells.CountLarge) <> mCellCount Then NxRaiseContractError "대상 셀 개수가 바뀌었습니다."

    index = 0
    For Each targetCell In target.Cells
        index = index + 1
        If NxSymbolCellFingerprint(targetCell) <> mFingerprints(index) Then _
            NxRaiseContractError "대상 셀 내용이 바뀌었습니다. 다시 확인하세요."
    Next targetCell

    On Error GoTo Rollback
    index = 0
    For Each targetCell In target.Cells
        index = index + 1
        targetCell.Value2 = mNewValues(index)
        committed = index
    Next targetCell
    Exit Sub

Rollback:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    index = 0
    For Each targetCell In target.Cells
        index = index + 1
        If index > committed Then Exit For
        If IsEmpty(mOriginalValues(index)) Then
            targetCell.ClearContents
        Else
            targetCell.Value2 = mOriginalValues(index)
        End If
    Next targetCell
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Sub

Private Sub RequireSealed()
    If Not mSealed Then NxRaiseContractError "기호 셀 쓰기 계획이 준비되지 않았습니다."
End Sub
