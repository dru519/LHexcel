VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFocusSelection"
Option Explicit

Private mSheet As Worksheet
Private mTarget As Range
Private mTopRow As Long
Private mBottomRow As Long
Private mLeftColumn As Long
Private mRightColumn As Long
Private mReason As String
Private mValid As Boolean

Public Function Configure(ByVal target As Range, Optional ByVal allowWholeAxis As Boolean = False) As Boolean
    Dim workbook As Workbook
    On Error GoTo Rejected
    If target Is Nothing Then mReason = "선택한 셀 범위가 없습니다.": Exit Function
    If target.Areas.Count <> 1 Then mReason = "연속된 한 범위를 선택하세요.": Exit Function
    Set mSheet = target.Worksheet
    Dim wholeAxis As Boolean
    wholeAxis = target.Rows.Count = mSheet.Rows.Count Or target.Columns.Count = mSheet.Columns.Count
    If target.Rows.Count = mSheet.Rows.Count And target.Columns.Count = mSheet.Columns.Count Then mReason = "전체 시트 선택은 제외합니다.": Exit Function
    If CDbl(target.CountLarge) > NX_FOCUS_MAX_SELECTION_CELLS And Not (allowWholeAxis And wholeAxis) Then mReason = "선택 범위가 너무 큽니다.": Exit Function
    Set workbook = mSheet.Parent
    If workbook Is ThisWorkbook Then mReason = "추가기능 내부 시트는 제외합니다.": Exit Function
    If workbook.ReadOnly Then mReason = "읽기 전용 통합문서는 제외합니다.": Exit Function
    If mSheet.ProtectContents Then mReason = "보호된 시트는 제외합니다.": Exit Function
    If workbook.MultiUserEditing Then mReason = "공유 통합문서는 제외합니다.": Exit Function
    If wholeAxis And Not allowWholeAxis Then mReason = "전체 행 또는 열 선택은 제외합니다.": Exit Function
    Set mTarget = target
    If wholeAxis Then
        Set mTarget = Application.Intersect(target, Application.ActiveWindow.VisibleRange)
        If mTarget Is Nothing Then mReason = "화면에 선택 범위가 없습니다.": Exit Function
    End If
    mTopRow = target.Row
    mBottomRow = target.Row + target.Rows.Count - 1
    mLeftColumn = target.Column
    mRightColumn = target.Column + target.Columns.Count - 1
    mValid = True
    Configure = True
    Exit Function
Rejected:
    mReason = "현재 선택에서는 포커스셀을 표시할 수 없습니다."
End Function

Public Property Get IsValid() As Boolean: IsValid = mValid: End Property
Public Property Get Reason() As String: Reason = mReason: End Property
Public Property Get Sheet() As Worksheet: Set Sheet = mSheet: End Property
Public Property Get Target() As Range: Set Target = mTarget: End Property
Public Property Get TopRow() As Long: TopRow = mTopRow: End Property
Public Property Get BottomRow() As Long: BottomRow = mBottomRow: End Property
Public Property Get LeftColumn() As Long: LeftColumn = mLeftColumn: End Property
Public Property Get RightColumn() As Long: RightColumn = mRightColumn: End Property
