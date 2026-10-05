VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxHangulTableRequest"
Option Explicit

Private Const NX_HANGUL_MAX_CELLS As Long = 16384

Private mSelection As CNxRangeSelectionSession
Private mCellCount As Long
Private mRowCount As Long
Private mColumnCount As Long
Private mConfigured As Boolean

Public Sub ConfigureFromSelection()
    Dim source As Range
    If mConfigured Then NxRaiseContractError "아래한글 표 전송 요청은 한 번만 준비할 수 있습니다."
    Set mSelection = NxRangeSessionFromSelection()
    Set source = mSelection.ResolveRange()
    ValidateSource source
    mCellCount = CLng(CDbl(source.CountLarge))
    mRowCount = source.Rows.Count
    mColumnCount = source.Columns.Count
    mConfigured = True
End Sub

Public Function ResolveSourceRange() As Range
    Dim source As Range
    EnsureConfigured
    Set source = mSelection.ResolveRange()
    ValidateSource source
    If CLng(CDbl(source.CountLarge)) <> mCellCount Then _
        NxRaiseContractError "선택한 표 범위가 변경되었습니다. 다시 선택해 주세요."
    Set ResolveSourceRange = source
End Function

Public Property Get DisplayAddress() As String
    EnsureConfigured
    DisplayAddress = mSelection.DisplayAddress
End Property

Public Property Get CellCount() As Long
    EnsureConfigured
    CellCount = mCellCount
End Property

Public Property Get RowCount() As Long
    EnsureConfigured
    RowCount = mRowCount
End Property

Public Property Get ColumnCount() As Long
    EnsureConfigured
    ColumnCount = mColumnCount
End Property

Private Sub ValidateSource(ByVal source As Range)
    If source Is Nothing Then NxRaiseContractError "아래한글로 보낼 셀 범위를 먼저 선택하세요."
    If source.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 선택하세요."
    If source.Parent.Parent Is ThisWorkbook Then NxRaiseContractError "일반 Excel 통합문서의 범위를 선택하세요."
    If source.Rows.Count = source.Parent.Rows.Count Or source.Columns.Count = source.Parent.Columns.Count Then _
        NxRaiseContractError "전체 행 또는 전체 열이 아닌 실제 표 범위를 선택하세요."
    If CDbl(source.CountLarge) < 1# Or CDbl(source.CountLarge) > CDbl(NX_HANGUL_MAX_CELLS) Then _
        NxRaiseContractError "아래한글 표는 16,384셀 이하만 전송할 수 있습니다."
End Sub

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "아래한글 표 전송 요청이 준비되지 않았습니다."
End Sub
